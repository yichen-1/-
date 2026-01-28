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
            if str_val in ['/', 'NA', 'None', '', '无', '——', '0.00']:
                return default
            cleaned_value = str_val.replace(',', '').replace(' ', '').strip()
            return pd.to_numeric(cleaned_value)
        return default
    except (ValueError, TypeError):
        return default

def extract_company_name(pdf_lines):
    """从PDF提取公司名称（排除无关字符）"""
    for line in pdf_lines:
        if "公司名称:" in line:
            company = re.sub(r'公司名称:\s*', '', line).strip()
            return re.sub(r'[\u4e00-\u9fa5a-zA-Z0-9()（）有限公司]+', lambda m: m.group(), company)  # 只保留合法公司名字符
    return "未知公司"

def extract_clear_date(pdf_lines):
    """提取清分日期（精准匹配格式）"""
    date_pattern = r'清分日期\s*[:：]?\s*(\d{4}-\d{2}-\d{2})'
    for line in pdf_lines:
        date_match = re.search(date_pattern, line)
        if date_match:
            return date_match.group(1)
    return None

def extract_total_data(pdf_lines):
    """提取文件级合计电量和合计电费（避免误匹配）"""
    total_quantity = None
    total_amount = None
    for line in pdf_lines:
        line = line.replace('\\t', ' ').strip()
        line_cols = [col.strip() for col in re.split(r'\s{2,}', line) if col.strip()]  # 用2个以上空格分割，减少列拆分错误
        if len(line_cols) >= 4 and "合计电量" in line_cols and "合计电费" in line_cols:
            # 定位合计电量和电费的位置
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

# ---------------------- 核心提取逻辑（精准+顺序保留） ----------------------
def extract_station_data(pdf_lines, company_name, clear_date, total_quantity, total_amount):
    """
    提取单个PDF中的所有场站数据
    1. 按PDF原始顺序保留科目
    2. 过滤无效科目
    3. 严格匹配交易数据行结构
    """
    all_station_data = []
    station_pattern = r'机组\s+([^:：\s]{2,10}风电场)'  # 匹配2-10字的风电场名称（如“双发A风电场”）
    current_station = None
    current_station_meter_qty = None
    trade_data_start_flag = False  # 标记交易数据区域开始（需同时匹配表头）
    all_trade_names = []  # 用列表存储，保留原始顺序
    header_matched = False  # 标记是否匹配到“科目编码+结算类型”表头

    # 第一步：精准定位交易数据区域+提取科目（按顺序）
    for line_idx, line in enumerate(pdf_lines):
        line = line.replace('\\t', ' ').strip()
        line_cols = [col.strip() for col in re.split(r'\s{2,}', line) if col.strip()]

        # 1. 识别场站（仅匹配“机组 某某风电场”格式）
        station_match = re.search(station_pattern, line)
        if station_match:
            current_station = station_match.group(1)
            trade_data_start_flag = False
            header_matched = False
            continue

        # 2. 识别当前场站的计量电量（精准匹配“计量电量：XXX”格式）
        if current_station and "计量电量" in line and "：" in line:
            meter_qty_match = re.search(r'计量电量\s*[:：]\s*(\S+)', line)
            if meter_qty_match:
                current_station_meter_qty = safe_convert_to_numeric(meter_qty_match.group(1))
            continue

        # 3. 匹配交易数据表头（必须包含“科目编码”和“结算类型”，确认数据区域）
        if not header_matched and len(line_cols) >= 5:
            if "科目编码" in line_cols and "结算类型" in line_cols and "电量" in line_cols:
                header_matched = True
                trade_data_start_flag = True  # 表头后才开始提取数据
                continue

        # 4. 提取交易科目（严格满足：表头后、5列结构、科目名合法）
        if trade_data_start_flag and header_matched and current_station and len(line_cols) == 5:
            trade_code = line_cols[0]  # 第1列：科目编码（必须为数字或特定编码格式）
            trade_name = line_cols[1]  # 第2列：结算类型（科目名称）
            
            # 过滤无效科目（关键！解决多余列问题）
            invalid_keywords = ['_', '县', '镇', '乡', '村', 'hf', 'HF', '合计', '小计', '电量', '电价', '电费']
            if (len(trade_name) >= 2 and len(trade_name) <= 20  # 科目名长度2-20字
                and not any(kw in trade_name for kw in invalid_keywords)
                and (trade_code.isdigit() or trade_code.startswith(('10', '20')))):  # 科目编码为数字或特定前缀
                
                if trade_name not in all_trade_names:  # 去重但保留顺序
                    all_trade_names.append(trade_name)

    # 第二步：按顺序提取每个场站的交易数据
    current_station = None
    current_station_meter_qty = None
    trade_data_start_flag = False
    header_matched = False
    station_trade_data = {}

    for line_idx, line in enumerate(pdf_lines):
        line = line.replace('\\t', ' ').strip()
        line_cols = [col.strip() for col in re.split(r'\s{2,}', line) if col.strip()]

        # 场站切换：保存上一个场站数据
        station_match = re.search(station_pattern, line)
        if station_match:
            if current_station and station_trade_data and all_trade_names:
                # 构建场站完整数据（按科目顺序）
                station_row = {
                    '公司名称': company_name,
                    '场站名称': current_station,
                    '清分日期': clear_date,
                    '文件合计电量(兆瓦时)': total_quantity,
                    '文件合计电费(元)': total_amount,
                    '场站计量电量(兆瓦时)': current_station_meter_qty
                }
                # 按原始顺序补充科目数据
                for trade in all_trade_names:
                    station_row[f'{trade}_电量'] = station_trade_data.get(trade, {}).get('电量')
                    station_row[f'{trade}_电价'] = station_trade_data.get(trade, {}).get('电价')
                    station_row[f'{trade}_电费'] = station_trade_data.get(trade, {}).get('电费')
                all_station_data.append(station_row)
            
            # 初始化新场站
            current_station = station_match.group(1)
            station_trade_data = {}
            trade_data_start_flag = False
            header_matched = False
            continue

        # 识别计量电量
        if current_station and "计量电量" in line and "：" in line:
            meter_qty_match = re.search(r'计量电量\s*[:：]\s*(\S+)', line)
            if meter_qty_match:
                current_station_meter_qty = safe_convert_to_numeric(meter_qty_match.group(1))
            continue

        # 匹配交易数据表头
        if not header_matched and len(line_cols) >= 5:
            if "科目编码" in line_cols and "结算类型" in line_cols and "电量" in line_cols:
                header_matched = True
                trade_data_start_flag = True
                continue

        # 提取交易数据（严格5列结构，对应编码、名称、电量、电价、电费）
        if trade_data_start_flag and header_matched and current_station and len(line_cols) == 5:
            trade_code = line_cols[0]
            trade_name = line_cols[1]
            # 只处理已识别的有效科目
            if trade_name in all_trade_names:
                quantity = safe_convert_to_numeric(line_cols[2])  # 第3列：电量
                price = safe_convert_to_numeric(line_cols[3])     # 第4列：电价
                fee = safe_convert_to_numeric(line_cols[4])      # 第5列：电费
                station_trade_data[trade_name] = {
                    '电量': quantity,
                    '电价': price,
                    '电费': fee
                }

    # 保存最后一个场站数据
    if current_station and station_trade_data and all_trade_names:
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
    
    return all_station_data, all_trade_names

def extract_data_from_pdf(file_obj, file_name):
    """从PDF文件对象提取数据（精准+顺序）"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            if not pdf.pages:
                raise ValueError("PDF无有效页面")
            
            # 读取所有页面文本（保留原始换行，避免数据错乱）
            all_text = ""
            for page in pdf.pages:
                page_text = page.extract_text_simple()  # 用simple提取，减少格式干扰
                if page_text:
                    all_text += page_text + "\n"
            pdf_lines = [line.strip() for line in all_text.split('\n') if line.strip() and len(line.strip()) >= 3]  # 过滤过短行
            if not pdf_lines:
                raise ValueError("PDF为扫描件，无可用文本")

        # 提取基础信息
        company_name = extract_company_name(pdf_lines)
        clear_date = extract_clear_date(pdf_lines)
        total_quantity, total_amount = extract_total_data(pdf_lines)
        
        # 提取所有场站数据和科目（按顺序）
        station_data_list, all_trade_names = extract_station_data(
            pdf_lines, company_name, clear_date, total_quantity, total_amount
        )
        
        return station_data_list, all_trade_names

    except Exception as e:
        st.warning(f"处理PDF {file_name} 出错: {str(e)}")
        return [], []

def extract_data_from_excel(file_obj, file_name):
    """Excel文件处理（保持兼容，按顺序提取）"""
    try:
        df = pd.read_excel(file_obj, dtype=object)
        company_name = "未知公司"
        # 从文件名提取公司名（过滤无效字符）
        name_without_ext = file_name.split('.')[0]
        if "晶盛" in name_without_ext:
            company_name = "大庆晶盛光伏电站"
        
        # 提取日期
        date_match = re.search(r'\d{4}-\d{2}-\d{2}', name_without_ext)
        clear_date = date_match.group() if date_match else None

        # 提取合计数据
        total_quantity = safe_convert_to_numeric(df.iloc[0, 3] if len(df) > 0 else None)
        total_amount = safe_convert_to_numeric(df.iloc[0, 5] if len(df) > 0 else None)

        # 固定Excel科目（按常见顺序）
        excel_trade_names = [
            '优先发电交易', '电网企业代理购电交易', '省内电力直接交易'
        ]
        
        # 构建场站数据
        station_data = [{
            '公司名称': company_name,
            '场站名称': company_name.replace('有限公司', '场站'),
            '清分日期': clear_date,
            '文件合计电量(兆瓦时)': total_quantity,
            '文件合计电费(元)': total_amount,
            '场站计量电量(兆瓦时)': total_quantity
        }]
        
        return station_data, excel_trade_names

    except Exception as e:
        st.warning(f"处理Excel {file_name} 出错: {str(e)}")
        return [], []

# ---------------------- 数据汇总与导出（按科目顺序） ----------------------
def calculate_summary_row(data_df, all_trade_names):
    """计算汇总行（按科目顺序）"""
    summary_row = {
        '公司名称': '总计',
        '场站名称': '总计',
        '清分日期': '',
        '文件合计电量(兆瓦时)': data_df['文件合计电量(兆瓦时)'].dropna().sum(),
        '文件合计电费(元)': data_df['文件合计电费(元)'].dropna().sum(),
        '场站计量电量(兆瓦时)': data_df['场站计量电量(兆瓦时)'].dropna().sum()
    }
    
    # 按原始顺序汇总科目数据
    for trade in all_trade_names:
        summary_row[f'{trade}_电量'] = data_df[f'{trade}_电量'].dropna().sum()
        summary_row[f'{trade}_电费'] = data_df[f'{trade}_电费'].dropna().sum()
        # 电价取有效平均值（排除0和空值）
        price_vals = data_df[f'{trade}_电价'].dropna()
        price_vals = price_vals[price_vals > 0.01]  # 排除极小值
        summary_row[f'{trade}_电价'] = round(price_vals.mean(), 3) if not price_vals.empty else None
    
    return pd.DataFrame([summary_row])

def to_excel_bytes(df, report_df):
    """Excel导出（保持列顺序）"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='结算数据明细', index=False)
        report_df.to_excel(writer, sheet_name='处理报告', index=False)
    output.seek(0)
    return output

# ---------------------- Streamlit 页面布局 ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取（精准版）", layout="wide")
    
    # 页面标题
    st.title("📊 黑龙江日清分结算单数据提取工具（精准科目+顺序保留）")
    st.divider()

    # 1. 文件上传区域（提示PDF格式要求）
    st.subheader("📁 上传文件")
    st.caption("支持PDF/Excel，PDF需满足：① 可复制文本 ② 含“机组 某某风电场”标识 ③ 交易表头含“科目编码+结算类型”")
    uploaded_files = st.file_uploader(
        "可批量上传（PDF将按原始顺序提取科目，过滤无效列）",
        type=['pdf', 'xlsx'],
        accept_multiple_files=True
    )

    # 2. 数据处理逻辑
    if uploaded_files and st.button("🚀 开始处理", type="primary"):
        st.divider()
        st.subheader("⚙️ 处理进度")
        
        all_station_data = []
        all_trade_names = []  # 列表存储，保留全局科目顺序
        total_files = len(uploaded_files)
        processed_files = 0

        # 批量处理
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, file in enumerate(uploaded_files):
            file_name = file.name
            status_text.text(f"正在处理：{file_name}（{idx+1}/{total_files}）")
            
            # 按文件类型提取
            if file_name.lower().endswith('.pdf'):
                station_data, trade_names = extract_data_from_pdf(file, file_name)
            else:
                station_data, trade_names = extract_data_from_excel(file, file_name)
            
            # 累积数据（合并科目顺序，去重但保留首次出现顺序）
            if station_data:
                all_station_data.extend(station_data)
                for trade in trade_names:
                    if trade not in all_trade_names:
                        all_trade_names.append(trade)
                processed_files += 1
            
            # 更新进度
            progress_bar.progress((idx + 1) / total_files)

        progress_bar.empty()
        status_text.text("处理完成！")

        # 3. 结果展示与导出
        if all_station_data and all_trade_names:
            st.divider()
            st.subheader("📈 提取结果（按PDF原始顺序）")
            
            # 构建结果列名（基础列+按顺序的科目列）
            base_columns = [
                '公司名称', '场站名称', '清分日期',
                '文件合计电量(兆瓦时)', '文件合计电费(元)', '场站计量电量(兆瓦时)'
            ]
            trade_columns = []
            for trade in all_trade_names:
                trade_columns.extend([f'{trade}_电量', f'{trade}_电价', f'{trade}_电费'])
            result_columns = base_columns + trade_columns

            # 构建DataFrame（确保列顺序正确）
            result_df = pd.DataFrame(all_station_data)
            # 补充缺失列（不同文件科目差异）
            for col in result_columns:
                if col not in result_df.columns:
                    result_df[col] = None
            # 严格按目标列顺序排序
            result_df = result_df[result_columns]
            # 数值列格式化
            numeric_cols = [col for col in result_columns if any(key in col for key in ['电量', '电价', '电费'])]
            result_df[numeric_cols] = result_df[numeric_cols].apply(pd.to_numeric, errors='coerce')

            # 按公司、场站、日期排序
            result_df['清分日期'] = pd.to_datetime(result_df['清分日期'], errors='coerce')
            result_df = result_df.sort_values(['公司名称', '场站名称', '清分日期']).reset_index(drop=True)
            result_df['清分日期'] = result_df['清分日期'].dt.strftime('%Y-%m-%d').fillna('')

            # 添加汇总行（按科目顺序）
            summary_row = calculate_summary_row(result_df, all_trade_names)
            result_df = pd.concat([result_df, summary_row], ignore_index=True)

            # 生成处理报告
            failed_files = total_files - processed_files
            success_rate = f"{processed_files / total_files:.2%}" if total_files > 0 else "0%"
            stations = result_df['场站名称'].unique()
            station_count = len(stations) - 1 if '总计' in stations else len(stations)
            valid_rows = len(result_df) - 1

            report_df = pd.DataFrame({
                '统计项': ['文件总数', '成功处理数', '失败数', '处理成功率', '涉及场站数', '有效数据行数', '提取科目数'],
                '数值': [total_files, processed_files, failed_files,
                         success_rate, station_count, valid_rows, len(all_trade_names)]
            })

            # 展示结果（分标签页）
            tab1, tab2 = st.tabs(["结算数据明细（按顺序）", "处理报告"])
            with tab1:
                st.dataframe(result_df, use_container_width=True, height=500)
            with tab2:
                st.dataframe(report_df, use_container_width=True)

            # 导出Excel（保留顺序）
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            download_filename = f"黑龙江结算数据提取_精准版_{current_time}.xlsx"
            excel_bytes = to_excel_bytes(result_df, report_df)

            st.divider()
            st.download_button(
                label="📥 导出精准版Excel",
                data=excel_bytes,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            # 显示统计信息
            st.info(
                f"""处理完成！
                - 总计上传 {total_files} 个文件，成功处理 {processed_files} 个（成功率 {success_rate}）
                - 提取 {len(all_trade_names)} 个科目（按PDF原始顺序），涉及 {station_count} 个场站，{valid_rows} 行有效数据
                - 已过滤“hf_”“县_”等无效列，科目数据精准匹配
                """
            )
        else:
            st.warning("⚠️ 未提取到有效数据！请检查：")
            st.markdown("""
                1. PDF是否为可复制文本（非扫描件）；
                2. PDF是否包含“机组 某某风电场”的场站标识；
                3. 交易数据区域是否有“科目编码+结算类型+电量+电价+电费”表头。
            """)

    # 无文件上传时的提示
    elif not uploaded_files and st.button("🚀 开始处理", disabled=True):
        st.warning("请先上传PDF/Excel文件！")

if __name__ == "__main__":
    main()
