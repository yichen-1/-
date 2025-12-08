import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# 设置页面配置
st.set_page_config(
    page_title="电力结算数据提取工具",
    page_icon="⚡",
    layout="wide"
)

# 页面标题
st.title("⚡ 电力结算数据提取工具")
st.markdown("---")

def extract_station_name(file_name):
    """从文件名提取场站名称（去除日期和扩展名）"""
    name_without_ext = os.path.splitext(file_name)[0]
    date_pattern = r'\d{4}-\d{2}-\d{2}'  # 匹配YYYY-MM-DD格式日期
    station_name = re.sub(date_pattern, '', name_without_ext).strip()
    return station_name if station_name else name_without_ext


def safe_convert_to_numeric(value, default=None):
    """安全转换为数值，排除无效值（兼容逗号分隔的金额）"""
    try:
        if pd.notna(value) and value is not None:
            str_val = str(value).strip()
            # 排除无效占位符
            if str_val in ['/', 'NA', 'None', '', '无', '——', '无数据']:
                return default
            # 移除金额中的逗号（如86,479.04 → 86479.04）
            cleaned_value = str_val.replace(',', '').replace(' ', '').strip()
            return pd.to_numeric(cleaned_value)
        return default
    except (ValueError, TypeError):
        return default


def extract_trade_data_by_column(trade_name, pdf_lines, header_col_index):
    """
    按表头列索引提取数据（核心修正：精准匹配三列）
    header_col_index：字典，包含'电量列索引'、'电价列索引'、'电费列索引'
    """
    quantity = None  # 结算电量/容量
    price = None     # 结算电价/均价
    fee = None       # 结算电费

    # 遍历所有行，查找包含目标科目的行
    for idx, line in enumerate(pdf_lines):
        line_lower = line.lower()
        # 模糊匹配科目名称（如“优先发购电量交易”包含在“010101 优先发购电量交易”中）
        if trade_name.lower() in line_lower:
            # 按空格分割当前行数据（兼容多空格分隔）
            line_cols = [col.strip() for col in re.split(r'\s+', line) if col.strip()]
            
            # 按表头定位的列索引取数
            # 电量：从“结算电量/容量”列索引获取
            if len(line_cols) > header_col_index['电量列索引']:
                quantity = safe_convert_to_numeric(line_cols[header_col_index['电量列索引']])
            # 电价：从“结算电价/均价”列索引获取
            if len(line_cols) > header_col_index['电价列索引']:
                price = safe_convert_to_numeric(line_cols[header_col_index['电价列索引']])
            # 电费：从“结算电费”列索引获取
            if len(line_cols) > header_col_index['电费列索引']:
                fee = safe_convert_to_numeric(line_cols[header_col_index['电费列索引']])
            break  # 找到目标科目后退出循环
    return [quantity, price, fee]


def extract_data_from_pdf(file_path, file_name, station_name):
    """从PDF提取数据（完全按列定位，精准匹配三列数据）"""
    try:
        # 1. 读取PDF所有页面并合并文本
        with pdfplumber.open(file_path) as pdf:
            if not pdf.pages:
                raise ValueError("PDF文件无页面")
            
            all_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    all_text += page_text + "\n"  # 保留页面分隔，避免数据粘连
            pdf_lines = [line.strip() for line in all_text.split('\n') if line.strip()]
            if not pdf_lines:
                raise ValueError("PDF为扫描件，无可用文本")

        # 2. 提取清分日期（从文件名获取，如“2025-10-01”）
        name_without_ext = os.path.splitext(file_name)[0]
        date_match = re.search(r'\d{4}-\d{2}-\d{2}', name_without_ext)
        date = date_match.group() if date_match else None

        # 3. 提取整体结算数据（第一个表格：清分日期、结算电量、结算电费）
        total_quantity = None  # 总结算电量（如283.029兆瓦时）
        total_amount = None    # 总结算电费（如86,479.04元）
        for line in pdf_lines:
            # 匹配日期行（如“2025-10-01 283.03 86,479.04”）
            if re.match(r'\d{4}-\d{2}-\d{2}', line):
                line_cols = re.split(r'\s+', line)
                if len(line_cols) >= 3:
                    total_quantity = safe_convert_to_numeric(line_cols[1])
                    total_amount = safe_convert_to_numeric(line_cols[2])
                break

        # 4. 定位“结算科目”表头行，并获取三列的索引（核心步骤）
        header_col_index = {
            '电量列索引': None,  # “结算电量/容量”列的位置
            '电价列索引': None,  # “结算电价/均价”列的位置
            '电费列索引': None   # “结算电费”列的位置
        }

        for idx, line in enumerate(pdf_lines):
            line_cols = [col.strip() for col in re.split(r'\s+', line) if col.strip()]
            # 匹配包含三个目标列名的表头行
            if (('结算科目' in line_cols or '交易类型' in line_cols) and
                '结算电量/容量' in line_cols and
                '结算电价/均价' in line_cols and
                '结算电费' in line_cols):
                
                # 记录三列的索引位置
                header_col_index['电量列索引'] = line_cols.index('结算电量/容量')
                header_col_index['电价列索引'] = line_cols.index('结算电价/均价')
                header_col_index['电费列索引'] = line_cols.index('结算电费')
                break  # 找到表头后退出

        # 检查表头是否定位成功
        if None in header_col_index.values():
            raise ValueError("未找到完整的表头列（需包含'结算电量/容量'、'结算电价/均价'、'结算电费'）")

        # 5. 定义需提取的科目
        target_trades = [
            '电量清分',                # 01总科目
            '优先发购电量交易',        # 010101
            '新能源现货保障性收购电量',# 0101010312
            '电力直接交易',            # 010102
            '常规直接交易',            # 0101020301
            '连续运营集中竞价交易',    # 0101020303
            '现货交易',                # 0102
            '省间现货交易',            # 010201
            '省内现货交易',            # 010202
            '实时交易正偏差结算电量',  # 0102020301
            '实时交易负偏差结算电量'   # 0102020302
        ]

        # 6. 按列索引提取所有科目数据
        all_trade_data = []
        for trade in target_trades:
            trade_data = extract_trade_data_by_column(trade, pdf_lines, header_col_index)
            all_trade_data.extend(trade_data)  # 每个科目追加“电量、电价、电费”

        # 7. 组装返回数据
        return [station_name, date, total_quantity, total_amount] + all_trade_data

    except Exception as e:
        st.warning(f"处理PDF {file_name} 出错: {str(e)}")
        # 确保返回长度一致
        return [station_name, None, None, None] + [None] * 33


def extract_data_from_excel(file_path, file_name, station_name):
    """Excel数据提取（按列名匹配，与PDF逻辑一致）"""
    try:
        df = pd.read_excel(file_path, dtype=object)
        
        # 提取日期
        name_without_ext = os.path.splitext(file_name)[0]
        date_match = re.search(r'\d{4}-\d{2}-\d{2}', name_without_ext)
        date = date_match.group() if date_match else None

        # 提取整体结算数据（第一个表格，按列名匹配）
        total_quantity = None
        total_amount = None
        if '整体结算电量' in df.columns and '整体结算电费' in df.columns:
            total_quantity = safe_convert_to_numeric(df['整体结算电量'].iloc[0] if len(df) > 0 else None)
            total_amount = safe_convert_to_numeric(df['整体结算电费'].iloc[0] if len(df) > 0 else None)

        # 定位科目数据行（按列名匹配三列）
        target_trades = [
            '电量清分', '优先发购电量交易', '新能源现货保障性收购电量',
            '电力直接交易', '常规直接交易', '连续运营集中竞价交易',
            '现货交易', '省间现货交易', '省内现货交易',
            '实时交易正偏差结算电量', '实时交易负偏差结算电量'
        ]

        all_trade_data = []
        # 检查Excel是否包含必要列
        if all(col in df.columns for col in ['结算科目', '结算电量/容量', '结算电价/均价', '结算电费']):
            for trade in target_trades:
                # 按“结算科目”列匹配目标科目
                trade_row = df[df['结算科目'] == trade]
                if not trade_row.empty:
                    quantity = safe_convert_to_numeric(trade_row['结算电量/容量'].iloc[0])
                    price = safe_convert_to_numeric(trade_row['结算电价/均价'].iloc[0])
                    fee = safe_convert_to_numeric(trade_row['结算电费'].iloc[0])
                else:
                    quantity, price, fee = None, None, None
                all_trade_data.extend([quantity, price, fee])
        else:
            # 若列名不匹配，填充空值
            all_trade_data = [None] * 33

        return [station_name, date, total_quantity, total_amount] + all_trade_data

    except Exception as e:
        st.warning(f"处理Excel {file_name} 出错: {str(e)}")
        return [station_name, None, None, None] + [None] * 33


def calculate_summary_row(data_df):
    """计算汇总行（仅对有效数值求和/求平均）"""
    # 需求和的列（电量、电费）
    sum_cols = [col for col in data_df.columns if any(key in col for key in ['电量', '电费'])]
    # 需求平均的列（电价）
    avg_cols = [col for col in data_df.columns if '电价' in col]

    summary_row = {'场站名称': '总计', '清分日期': ''}
    # 求和逻辑
    for col in sum_cols:
        if col in data_df.columns:
            # 仅对非空数值求和
            valid_vals = data_df[col].dropna()
            summary_row[col] = valid_vals.sum() if not valid_vals.empty else 0
    # 求平均逻辑
    for col in avg_cols:
        if col in data_df.columns:
            valid_vals = data_df[col].dropna()
            summary_row[col] = round(valid_vals.mean(), 3) if not valid_vals.empty else None

    return pd.DataFrame([summary_row])


def process_files(uploaded_files):
    """处理上传的文件并返回结果"""
    all_data = []
    
    if not uploaded_files:
        st.error("未选择任何文件")
        return None, None
    
    # 显示处理进度
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, uploaded_file in enumerate(uploaded_files):
        # 更新进度
        progress = (i + 1) / len(uploaded_files)
        progress_bar.progress(progress)
        status_text.text(f"处理中: {uploaded_file.name} ({i+1}/{len(uploaded_files)})")
        
        station_name = extract_station_name(uploaded_file.name)
        
        # 根据文件类型处理
        if uploaded_file.name.lower().endswith('.xlsx'):
            data = extract_data_from_excel(uploaded_file, uploaded_file.name, station_name)
        else:  # PDF
            data = extract_data_from_pdf(uploaded_file, uploaded_file.name, station_name)

        # 验证数据有效性
        if data[1] is not None and any(isinstance(val, (float, int)) for val in data[2:] if val is not None):
            all_data.append(data)
        else:
            st.warning(f"跳过文件 {uploaded_file.name}（数据无效）")
    
    # 清除进度条和状态
    progress_bar.empty()
    status_text.empty()
    
    if not all_data:
        st.warning("未提取到有效数据，请检查文件是否符合要求")
        return None, None
    
    # 定义列名
    result_columns = [
        '场站名称', '清分日期', '整体结算电量(兆瓦时)', '整体结算电费(元)',
        # 电量清分（01）
        '电量清分_电量(兆瓦时)', '电量清分_电价(元/兆瓦时)', '电量清分_电费(元)',
        # 优先发购相关
        '优先发购电量交易_电量', '优先发购电量交易_电价', '优先发购电量交易_电费',
        '新能源现货保障性收购_电量', '新能源现货保障性收购_电价', '新能源现货保障性收购_电费',
        # 电力直接交易相关
        '电力直接交易_电量', '电力直接交易_电价', '电力直接交易_电费',
        '常规直接交易_电量', '常规直接交易_电价', '常规直接交易_电费',
        '连续运营集中竞价_电量', '连续运营集中竞价_电价', '连续运营集中竞价_电费',
        # 现货交易相关
        '现货交易_电量', '现货交易_电价', '现货交易_电费',
        '省间现货交易_电量', '省间现货交易_电价', '省间现货交易_电费',
        '省内现货交易_电量', '省内现货交易_电价', '省内现货交易_电费',
        # 实时偏差相关
        '实时交易正偏差_电量', '实时交易正偏差_电价', '实时交易正偏差_电费',
        '实时交易负偏差_电量', '实时交易负偏差_电价', '实时交易负偏差_电费'
    ]

    # 构建DataFrame并处理数值类型
    result_df = pd.DataFrame(all_data, columns=result_columns)
    num_cols = result_df.columns[2:]  # 从“整体结算电量”开始均为数值列
    result_df[num_cols] = result_df[num_cols].apply(pd.to_numeric, errors='coerce')

    # 按“场站名称+清分日期”排序
    result_df['清分日期'] = pd.to_datetime(result_df['清分日期'])
    result_df = result_df.sort_values(['场站名称', '清分日期']).reset_index(drop=True)
    result_df['清分日期'] = result_df['清分日期'].dt.strftime('%Y-%m-%d')

    # 添加汇总行
    summary_row = calculate_summary_row(result_df)
    result_df = pd.concat([result_df, summary_row], ignore_index=True)

    # 生成统计报告
    total_files = len(uploaded_files)
    processed_files = len(all_data)
    success_rate = f"{processed_files / total_files:.2%}" if total_files > 0 else "0%"
    stations = result_df['场站名称'].unique()
    station_count = len(stations) - 1 if '总计' in stations else len(stations)
    valid_rows = len(result_df) - 1  # 排除汇总行
    
    report_df = pd.DataFrame({
        '统计项': ['文件总数', '成功处理文件数', '处理失败文件数', '处理成功率', '涉及场站数', '有效数据行数'],
        '数值': [total_files, processed_files, total_files - processed_files,
                 success_rate, station_count, valid_rows]
    })
    
    return result_df, report_df


def to_excel(df1, df2):
    """将数据框转换为Excel文件的字节流"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='按列精准结算数据', index=False)
        df2.to_excel(writer, sheet_name='数据处理报告', index=False)
    return output.getvalue()


# 主界面
with st.sidebar:
    st.header("文件上传")
    uploaded_files = st.file_uploader(
        "选择Excel或PDF文件", 
        type=['xlsx', 'pdf'], 
        accept_multiple_files=True
    )
    
    st.markdown("---")
    st.info("""
    支持格式：.xlsx、.pdf
    
    注意事项：
    1. PDF需为可复制文本格式（非扫描件）
    2. 文件需包含标准表头：'结算电量/容量'、'结算电价/均价'、'结算电费'
    """)

# 主内容区
if st.button("开始处理", disabled=not uploaded_files):
    with st.spinner("正在处理文件，请稍候..."):
        result_df, report_df = process_files(uploaded_files)
        
        if result_df is not None and report_df is not None:
            st.success("数据处理完成！")
            
            # 显示报告
            st.subheader("处理报告")
            st.dataframe(report_df, use_container_width=True)
            
            # 显示结果数据
            st.subheader("结算数据")
            st.dataframe(result_df, use_container_width=True)
            
            # 提供下载
            excel_data = to_excel(result_df, report_df)
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="下载Excel结果",
                data=excel_data,
                file_name=f"结算数据_精准按列提取_{current_time}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    if uploaded_files:
        st.info("点击上方的'开始处理'按钮开始提取数据")
    else:
        st.info("请在左侧上传Excel或PDF文件进行处理")
