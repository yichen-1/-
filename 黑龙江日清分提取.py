import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心工具函数 ----------------------
def safe_convert_to_numeric(value, default=None):
    """安全转换为数值，兼容逗号分隔的金额和空值"""
    try:
        if pd.notna(value) and value is not None:
            str_val = str(value).strip()
            if str_val in ['/', 'NA', 'None', '', '无', '——']:
                return default
            cleaned_value = str_val.replace(',', '').replace(' ', '').strip()
            return pd.to_numeric(cleaned_value)
        return default
    except (ValueError, TypeError):
        return default

def extract_company_name(pdf_lines):
    """从PDF提取公司名称"""
    for line in pdf_lines:
        if "公司名称:" in line:
            return re.sub(r'公司名称:\s*', '', line).strip()
    return "未知公司"

def extract_clear_date(pdf_lines):
    """提取清分日期"""
    date_pattern = r'清分日期\s*(\d{4}-\d{2}-\d{2})'
    for line in pdf_lines:
        date_match = re.search(date_pattern, line)
        if date_match:
            return date_match.group(1)
    return None

def extract_total_data(pdf_lines):
    """提取文件级合计电量和合计电费"""
    total_quantity = None
    total_amount = None
    for line in pdf_lines:
        line_cols = [col.strip() for col in re.split(r'\s+', line) if col.strip()]
        if "合计电量" in line_cols and "合计电费" in line_cols:
            if "合计电量" in line_cols:
                qty_idx = line_cols.index("合计电量") + 1
                if qty_idx < len(line_cols):
                    total_quantity = safe_convert_to_numeric(line_cols[qty_idx])
            if "合计电费" in line_cols:
                fee_idx = line_cols.index("合计电费") + 1
                if fee_idx < len(line_cols):
                    total_amount = safe_convert_to_numeric(line_cols[fee_idx])
            break
    return total_quantity, total_amount

# ---------------------- 核心提取逻辑（适配多场站+动态科目） ----------------------
def extract_station_data(pdf_lines, company_name, clear_date, total_quantity, total_amount):
    """提取单个PDF中的所有场站数据（动态识别科目）"""
    all_station_data = []
    station_pattern = r'机组\s+([^:：\s]+风电场)'  # 匹配"机组 双发A风电场"格式
    current_station = None
    current_station_meter_qty = None
    trade_data_started = False  # 标记是否进入交易数据区域
    all_trade_names = set()  # 收集所有动态识别的科目名称

    # 第一步：扫描所有交易科目名称和场站信息
    for line in pdf_lines:
        line = line.strip()
        # 识别场站切换
        station_match = re.search(station_pattern, line)
        if station_match:
            current_station = station_match.group(1)
            trade_data_started = False
            continue
        
        # 识别当前场站的计量电量
        if current_station and "计量电量" in line:
            meter_qty_match = re.search(r'计量电量\s*(\S+)', line)
            if meter_qty_match:
                current_station_meter_qty = safe_convert_to_numeric(meter_qty_match.group(1))
            continue
        
        # 标记交易数据开始（电能量电费下方）
        if "电能量电费" in line:
            trade_data_started = True
            continue
        
        # 提取交易科目（结算类型列）
        if trade_data_started and current_station:
            line_cols = [col.strip() for col in re.split(r'\s+', line) if col.strip()]
            # 交易数据行特征：至少5列，第2列为结算类型名称，第3-5列为数值
            if len(line_cols) >= 5 and line_cols[1] not in ['结算类型', '科目编码'] and not line_cols[1].isdigit():
                trade_name = line_cols[1]
                if trade_name not in ['电量', '电价', '电费', '小计']:
                    all_trade_names.add(trade_name)
    
    # 第二步：提取每个场站的具体交易数据
    current_station = None
    current_station_meter_qty = None
    trade_data_started = False
    station_trade_data = {}

    for line in pdf_lines:
        line = line.strip()
        line_cols = [col.strip() for col in re.split(r'\s+', line) if col.strip()]
        
        # 场站切换处理（保存上一个场站数据）
        station_match = re.search(station_pattern, line)
        if station_match:
            if current_station and station_trade_data:
                # 构建当前场站完整数据
                station_row = {
                    '公司名称': company_name,
                    '场站名称': current_station,
                    '清分日期': clear_date,
                    '文件合计电量(兆瓦时)': total_quantity,
                    '文件合计电费(元)': total_amount,
                    '场站计量电量(兆瓦时)': current_station_meter_qty
                }
                # 补充所有交易科目的数据
                for trade in all_trade_names:
                    station_row[f'{trade}_电量'] = station_trade_data.get(trade, {}).get('电量')
                    station_row[f'{trade}_电价'] = station_trade_data.get(trade, {}).get('电价')
                    station_row[f'{trade}_电费'] = station_trade_data.get(trade, {}).get('电费')
                all_station_data.append(station_row)
            
            # 初始化新场站
            current_station = station_match.group(1)
            station_trade_data = {}
            trade_data_started = False
            continue
        
        # 识别当前场站的计量电量
        if current_station and "计量电量" in line:
            meter_qty_match = re.search(r'计量电量\s*(\S+)', line)
            if meter_qty_match:
                current_station_meter_qty = safe_convert_to_numeric(meter_qty_match.group(1))
            continue
        
        # 标记交易数据开始
        if "电能量电费" in line:
            trade_data_started = True
            continue
        
        # 提取当前交易科目的数据
        if trade_data_started and current_station and len(line_cols) >= 5:
            trade_name = line_cols[1]
            if trade_name in all_trade_names:
                quantity = safe_convert_to_numeric(line_cols[2])
                price = safe_convert_to_numeric(line_cols[3])
                fee = safe_convert_to_numeric(line_cols[4])
                station_trade_data[trade_name] = {
                    '电量': quantity,
                    '电价': price,
                    '电费': fee
                }
    
    # 保存最后一个场站的数据
    if current_station and station_trade_data:
        station_row = {
            '公司名称': company_name,
            '场站名称': current_station,
            '清分日期': clear_date,
            '文件合计电量(兆瓦时)': total_quantity,
            '文件合计电费(元)': total_amount,
            '场站计量电量(兆瓦时)': current_station_meter_qty
        }
        for trade in all_trade_names:
            station_row[f'{trade}_电量'] = station_trade_data.get(trade, {}).get('电量')
            station_row[f'{trade}_电价'] = station_trade_data.get(trade, {}).get('电价')
            station_row[f'{trade}_电费'] = station_trade_data.get(trade, {}).get('电费')
        all_station_data.append(station_row)
    
    return all_station_data, list(all_trade_names)

def extract_data_from_pdf(file_obj, file_name):
    """从PDF文件对象提取数据（支持多场站+动态科目）"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            if not pdf.pages:
                raise ValueError("PDF无有效页面")
            
            # 读取所有页面文本
            all_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    all_text += page_text + "\n"
            pdf_lines = [line.strip() for line in all_text.split('\n') if line.strip()]
            if not pdf_lines:
                raise ValueError("PDF为扫描件，无可用文本")

        # 提取基础信息
        company_name = extract_company_name(pdf_lines)
        clear_date = extract_clear_date(pdf_lines)
        total_quantity, total_amount = extract_total_data(pdf_lines)
        
        # 提取所有场站数据和动态科目
        station_data_list, all_trade_names = extract_station_data(
            pdf_lines, company_name, clear_date, total_quantity, total_amount
        )
        
        return station_data_list, all_trade_names

    except Exception as e:
        st.warning(f"处理PDF {file_name} 出错: {str(e)}")
        return [], []

def extract_data_from_excel(file_obj, file_name):
    """Excel文件处理（保持原逻辑，适配动态科目）"""
    try:
        df = pd.read_excel(file_obj, dtype=object)
        company_name = "未知公司"
        # 从文件名提取公司名
        name_without_ext = file_name.split('.')[0]
        if "晶盛" in name_without_ext:
            company_name = "大庆晶盛光伏电站"
        
        # 提取日期
        date_match = re.search(r'\d{4}-\d{2}-\d{2}', name_without_ext)
        clear_date = date_match.group() if date_match else None

        # 提取合计数据
        total_quantity = safe_convert_to_numeric(df.iloc[0, 3] if len(df) > 0 else None)
        total_amount = safe_convert_to_numeric(df.iloc[0, 5] if len(df) > 0 else None)

        # 固定基础列（Excel暂按原逻辑处理，可根据实际格式调整）
        station_data = [{
            '公司名称': company_name,
            '场站名称': company_name.replace('有限公司', '场站'),
            '清分日期': clear_date,
            '文件合计电量(兆瓦时)': total_quantity,
            '文件合计电费(元)': total_amount,
            '场站计量电量(兆瓦时)': total_quantity
        }]
        
        # 默认Excel科目（可根据实际需求调整）
        excel_trade_names = [
            '优先发电交易', '电网企业代理购电交易', '省内电力直接交易'
        ]
        
        return station_data, excel_trade_names

    except Exception as e:
        st.warning(f"处理Excel {file_name} 出错: {str(e)}")
        return [], []

# ---------------------- 数据汇总与导出 ----------------------
def calculate_summary_row(data_df, all_trade_names):
    """计算汇总行（适配动态科目）"""
    summary_row = {
        '公司名称': '总计',
        '场站名称': '总计',
        '清分日期': '',
        '文件合计电量(兆瓦时)': data_df['文件合计电量(兆瓦时)'].dropna().sum(),
        '文件合计电费(元)': data_df['文件合计电费(元)'].dropna().sum(),
        '场站计量电量(兆瓦时)': data_df['场站计量电量(兆瓦时)'].dropna().sum()
    }
    
    # 汇总各交易科目的数据
    for trade in all_trade_names:
        summary_row[f'{trade}_电量'] = data_df[f'{trade}_电量'].dropna().sum()
        summary_row[f'{trade}_电费'] = data_df[f'{trade}_电费'].dropna().sum()
        # 电价取平均值（排除0值）
        price_vals = data_df[f'{trade}_电价'].dropna()
        price_vals = price_vals[price_vals > 0]
        summary_row[f'{trade}_电价'] = round(price_vals.mean(), 3) if not price_vals.empty else None
    
    return pd.DataFrame([summary_row])

def to_excel_bytes(df, report_df):
    """将DataFrame转为Excel字节流"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='结算数据明细', index=False)
        report_df.to_excel(writer, sheet_name='处理报告', index=False)
    output.seek(0)
    return output

# ---------------------- Streamlit 页面布局 ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取（多场站版）", layout="wide")
    
    # 页面标题
    st.title("📊 黑龙江日清分结算单数据提取工具（支持多场站+动态科目）")
    st.divider()

    # 1. 文件上传区域
    st.subheader("📁 上传文件")
    uploaded_files = st.file_uploader(
        "支持PDF/Excel格式，可批量上传（PDF自动识别多场站和动态科目）",
        type=['pdf', 'xlsx'],
        accept_multiple_files=True
    )

    # 2. 数据处理逻辑
    if uploaded_files and st.button("🚀 开始处理", type="primary"):
        st.divider()
        st.subheader("⚙️ 处理进度")
        
        all_station_data = []
        all_trade_names = set()  # 收集所有文件的交易科目
        total_files = len(uploaded_files)
        processed_files = 0

        # 批量处理上传的文件
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, file in enumerate(uploaded_files):
            file_name = file.name
            status_text.text(f"正在处理：{file_name}")
            
            # 根据文件类型调用对应提取函数
            if file_name.lower().endswith('.pdf'):
                station_data, trade_names = extract_data_from_pdf(file, file_name)
            else:
                station_data, trade_names = extract_data_from_excel(file, file_name)
            
            # 累积数据和科目
            if station_data:
                all_station_data.extend(station_data)
                all_trade_names.update(trade_names)
                processed_files += 1
            
            # 更新进度
            progress_bar.progress((idx + 1) / total_files)

        progress_bar.empty()
        status_text.text("处理完成！")

        # 3. 结果展示与导出
        if all_station_data and all_trade_names:
            st.divider()
            st.subheader("📈 提取结果")
            
            # 构建结果列名（基础列 + 动态科目列）
            base_columns = [
                '公司名称', '场站名称', '清分日期',
                '文件合计电量(兆瓦时)', '文件合计电费(元)', '场站计量电量(兆瓦时)'
            ]
            trade_columns = []
            for trade in sorted(all_trade_names):
                trade_columns.extend([f'{trade}_电量', f'{trade}_电价', f'{trade}_电费'])
            result_columns = base_columns + trade_columns

            # 构建DataFrame并格式化
            result_df = pd.DataFrame(all_station_data)
            # 补充缺失的列（不同文件可能有不同科目）
            for col in result_columns:
                if col not in result_df.columns:
                    result_df[col] = None
            # 只保留目标列
            result_df = result_df[result_columns]
            # 数值列格式化
            numeric_cols = [col for col in result_columns if any(key in col for key in ['电量', '电价', '电费'])]
            result_df[numeric_cols] = result_df[numeric_cols].apply(pd.to_numeric, errors='coerce')

            # 排序
            result_df['清分日期'] = pd.to_datetime(result_df['清分日期'], errors='coerce')
            result_df = result_df.sort_values(['公司名称', '场站名称', '清分日期']).reset_index(drop=True)
            result_df['清分日期'] = result_df['清分日期'].dt.strftime('%Y-%m-%d').fillna('')

            # 添加汇总行
            summary_row = calculate_summary_row(result_df, all_trade_names)
            result_df = pd.concat([result_df, summary_row], ignore_index=True)

            # 生成处理报告
            failed_files = total_files - processed_files
            success_rate = f"{processed_files / total_files:.2%}" if total_files > 0 else "0%"
            stations = result_df['场站名称'].unique()
            station_count = len(stations) - 1 if '总计' in stations else len(stations)
            valid_rows = len(result_df) - 1

            report_df = pd.DataFrame({
                '统计项': ['文件总数', '成功处理数', '失败数', '处理成功率', '涉及场站数', '有效数据行数', '识别科目数'],
                '数值': [total_files, processed_files, failed_files,
                         success_rate, station_count, valid_rows, len(all_trade_names)]
            })

            # 展示结果表格
            tab1, tab2 = st.tabs(["结算数据明细", "处理报告"])
            with tab1:
                st.dataframe(result_df, use_container_width=True)
            with tab2:
                st.dataframe(report_df, use_container_width=True)

            # 生成下载文件
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            download_filename = f"黑龙江结算数据提取_{current_time}.xlsx"
            excel_bytes = to_excel_bytes(result_df, report_df)

            # 下载按钮
            st.divider()
            st.download_button(
                label="📥 导出Excel文件",
                data=excel_bytes,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            # 显示统计信息
            st.info(
                f"""处理完成！
                - 总计上传 {total_files} 个文件，成功处理 {processed_files} 个（成功率 {success_rate}）
                - 识别到 {len(all_trade_names)} 个交易科目，涉及 {station_count} 个场站，{valid_rows} 行有效数据
                - PDF文件已自动拆分多场站数据，科目随结算单动态更新
                """
            )
        else:
            st.warning("⚠️ 未提取到有效数据！请检查：")
            st.markdown("""
                1. PDF是否为可复制文本（非扫描件）；
                2. 文件是否为黑龙江日清分格式（包含"机组 某某风电场"标识）；
                3. Excel文件格式是否匹配。
            """)

    # 无文件上传时的提示
    elif not uploaded_files and st.button("🚀 开始处理", disabled=True):
        st.warning("请先上传PDF/Excel文件！")

if __name__ == "__main__":
    main()
