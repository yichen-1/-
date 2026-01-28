import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心配置（可根据实际需求调整） ----------------------
# 特殊科目：无电量/电价，仅需提取电费（不生成电量/电价列）
SPECIAL_TRADE_KEYWORDS = ['阻塞费用', '价差费用']
# 无效行过滤关键词（彻底排除含这些字符的行）
INVALID_LINE_KEYWORDS = ['hf', 'HF', '县', '镇', '乡', '村', '_', '—']
# 交易表头关键词（放宽匹配，只要包含编码和类型即可）
TRADE_HEADER_KEYWORDS = ['科目编码', '结算类型']

# ---------------------- 核心工具函数 ----------------------
def safe_convert_to_numeric(value, default=None):
    """安全转换为数值，兼容逗号分隔的金额和空值"""
    try:
        if pd.notna(value) and value is not None:
            str_val = str(value).strip()
            if str_val in ['/', 'NA', 'None', '', '无', '——', '0.00', '-']:
                return default
            cleaned_value = str_val.replace(',', '').replace(' ', '').strip()
            return pd.to_numeric(cleaned_value)
        return default
    except (ValueError, TypeError):
        return default

def filter_invalid_lines(pdf_lines):
    """过滤含无效关键词的行，减少干扰"""
    valid_lines = []
    for line in pdf_lines:
        line = line.replace('\\t', ' ').strip()
        # 过滤过短行（少于2字符）和含无效关键词的行
        if len(line) >= 2 and not any(kw in line for kw in INVALID_LINE_KEYWORDS):
            valid_lines.append(line)
    return valid_lines

def extract_company_name(pdf_lines):
    """精准提取公司名称"""
    for line in pdf_lines:
        if "公司名称:" in line or "公司名称：" in line:
            # 只保留中文、数字、括号和“有限公司”后缀
            company_match = re.search(r'[\u4e00-\u9fa5a-zA-Z0-9()（）]+有限公司', line)
            if company_match:
                return company_match.group().strip()
    return "未知公司"

def extract_clear_date(pdf_lines):
    """精准提取清分日期"""
    date_pattern = r'清分日期\s*[:：]?\s*(\d{4}-\d{2}-\d{2})'
    for line in pdf_lines:
        date_match = re.search(date_pattern, line)
        if date_match:
            return date_match.group(1)
    return None

# ---------------------- 核心提取逻辑（全科目覆盖+顺序保留） ----------------------
def classify_trade_type(trade_name):
    """判断科目类型：常规科目（3列）/特殊科目（1列）"""
    return 'special' if any(kw in trade_name for kw in SPECIAL_TRADE_KEYWORDS) else 'normal'

def extract_trade_info(line_cols):
    """根据科目类型提取数据：常规科目（电量+电价+电费）/特殊科目（仅电费）"""
    trade_code = line_cols[0].strip()
    trade_name = line_cols[1].strip()
    trade_type = classify_trade_type(trade_name)
    
    if trade_type == 'normal':
        # 常规科目：第3列电量，第4列电价，第5列电费（允许部分为空）
        quantity = safe_convert_to_numeric(line_cols[2] if len(line_cols)>=3 else None)
        price = safe_convert_to_numeric(line_cols[3] if len(line_cols)>=4 else None)
        fee = safe_convert_to_numeric(line_cols[4] if len(line_cols)>=5 else None)
        return trade_code, trade_name, trade_type, {'电量': quantity, '电价': price, '电费': fee}
    else:
        # 特殊科目：第2列或第3列电费（适配3列结构）
        fee_col_idx = 2 if len(line_cols)>=3 else 1
        fee = safe_convert_to_numeric(line_cols[fee_col_idx])
        return trade_code, trade_name, trade_type, {'电费': fee}

def extract_station_and_trades(pdf_lines):
    """
    提取所有场站和科目数据
    1. 按PDF从上到下顺序保留科目
    2. 适配常规/特殊科目结构
    3. 严格保留单个文件内的科目顺序
    """
    # 初始化变量
    station_pattern = r'机组\s+([^:：\s]{2,15}风电场)'  # 放宽场站名长度（2-15字）
    current_station = None
    current_station_meter_qty = None
    in_trade_area = False  # 是否进入交易数据区域
    all_stations = []  # 存储所有场站数据（含科目）
    file_total_quantity = None  # 文件级合计电量
    file_total_fee = None  # 文件级合计电费

    # 第一步：先提取文件级合计数据和过滤无效行
    filtered_lines = filter_invalid_lines(pdf_lines)
    for line in filtered_lines:
        line_cols = [col.strip() for col in re.split(r'\s{1,}', line) if col.strip()]
        # 提取文件级合计电量和电费
        if len(line_cols) >=4 and "合计电量" in line_cols and "合计电费" in line_cols:
            qty_idx = line_cols.index("合计电量") + 1 if "合计电量" in line_cols else -1
            fee_idx = line_cols.index("合计电费") + 1 if "合计电费" in line_cols else -1
            if qty_idx != -1 and qty_idx < len(line_cols):
                file_total_quantity = safe_convert_to_numeric(line_cols[qty_idx])
            if fee_idx != -1 and fee_idx < len(line_cols):
                file_total_fee = safe_convert_to_numeric(line_cols[fee_idx])

    # 第二步：提取场站和科目数据（按顺序）
    current_station_trades = []  # 当前场站的科目列表（保留顺序）
    for line in filtered_lines:
        line_cols = [col.strip() for col in re.split(r'\s{1,}', line) if col.strip()]
        line_len = len(line_cols)

        # 1. 识别场站切换（保存上一个场站数据）
        station_match = re.search(station_pattern, line)
        if station_match:
            if current_station and current_station_trades:
                # 保存当前场站数据
                all_stations.append({
                    '场站名称': current_station,
                    '计量电量': current_station_meter_qty,
                    '科目列表': current_station_trades.copy()  # 深拷贝，避免引用问题
                })
            # 初始化新场站
            current_station = station_match.group(1)
            current_station_meter_qty = None
            current_station_trades = []
            in_trade_area = False
            continue

        # 2. 提取当前场站的计量电量（适配“计量电量：XXX”或“计量电量 XXX”）
        if current_station and "计量电量" in line:
            meter_match = re.search(r'计量电量\s*[:：]?\s*(\S+)', line)
            if meter_match:
                current_station_meter_qty = safe_convert_to_numeric(meter_match.group(1))
            continue

        # 3. 识别交易数据区域（只要包含表头关键词即进入）
        if not in_trade_area and all(kw in line_cols for kw in TRADE_HEADER_KEYWORDS):
            in_trade_area = True
            continue

        # 4. 提取科目数据（区分常规/特殊科目）
        if in_trade_area and current_station and line_len >=2:
            # 科目编码必须为数字或特定格式（排除纯文本行）
            if line_cols[0].isdigit() or line_cols[0].startswith(('10', '20', '30')):
                try:
                    trade_code, trade_name, trade_type, trade_data = extract_trade_info(line_cols)
                    # 过滤无效科目名（长度1-20字，排除纯数字/符号）
                    if 1 <= len(trade_name) <=20 and not trade_name.isdigit():
                        current_station_trades.append({
                            '编码': trade_code,
                            '名称': trade_name,
                            '类型': trade_type,
                            '数据': trade_data
                        })
                except Exception:
                    continue

    # 保存最后一个场站数据
    if current_station and current_station_trades:
        all_stations.append({
            '场站名称': current_station,
            '计量电量': current_station_meter_qty,
            '科目列表': current_station_trades.copy()
        })

    return all_stations, file_total_quantity, file_total_fee

def extract_data_from_pdf(file_obj, file_name):
    """从PDF提取完整数据（场站+科目+顺序保留）"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            if not pdf.pages:
                raise ValueError("PDF无有效页面")
            
            # 读取所有页面文本（保留原始顺序）
            all_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()  # 不用simple，避免丢失结构
                if page_text:
                    all_text += page_text + "\n"
            pdf_lines = [line.strip() for line in all_text.split('\n') if line.strip()]
            if not pdf_lines:
                raise ValueError("PDF为扫描件，无可用文本")

        # 提取基础信息
        company_name = extract_company_name(pdf_lines)
        clear_date = extract_clear_date(pdf_lines)
        # 提取场站和科目数据
        all_stations, file_total_quantity, file_total_fee = extract_station_and_trades(pdf_lines)

        # 整理输出格式
        output_data = []
        global_trade_order = []  # 单个文件内的科目顺序（全局复用）
        # 先收集单个文件内的科目顺序（按提取顺序）
        for station in all_stations:
            for trade in station['科目列表']:
                if trade['名称'] not in global_trade_order:
                    global_trade_order.append(trade['名称'])
        
        # 构建每个场站的数据行
        for station in all_stations:
            station_row = {
                '公司名称': company_name,
                '场站名称': station['场站名称'],
                '清分日期': clear_date,
                '文件合计电量(兆瓦时)': file_total_quantity,
                '文件合计电费(元)': file_total_fee,
                '场站计量电量(兆瓦时)': station['计量电量']
            }
            # 按顺序添加科目数据
            for trade_name in global_trade_order:
                # 找到当前科目的数据
                trade_data = next((t['数据'] for t in station['科目列表'] if t['名称'] == trade_name), None)
                if trade_data:
                    trade_type = next((t['类型'] for t in station['科目列表'] if t['名称'] == trade_name), 'normal')
                    if trade_type == 'normal':
                        # 常规科目：添加电量、电价、电费
                        station_row[f'{trade_name}_电量'] = trade_data.get('电量')
                        station_row[f'{trade_name}_电价'] = trade_data.get('电价')
                        station_row[f'{trade_name}_电费'] = trade_data.get('电费')
                    else:
                        # 特殊科目：仅添加电费
                        station_row[f'{trade_name}_电费'] = trade_data.get('电费')
            output_data.append(station_row)
        
        return output_data, global_trade_order

    except Exception as e:
        st.warning(f"处理PDF {file_name} 出错: {str(e)}")
        return [], []

# ---------------------- Excel处理（保持兼容） ----------------------
def extract_data_from_excel(file_obj, file_name):
    """Excel文件处理（适配动态科目）"""
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
        total_fee = safe_convert_to_numeric(df.iloc[0, 5] if len(df) > 0 else None)

        # 常规Excel科目（按顺序）
        excel_trades = [
            '优先发电交易', '电网企业代理购电交易', '省内电力直接交易'
        ]
        # 构建数据行
        station_row = {
            '公司名称': company_name,
            '场站名称': company_name.replace('有限公司', '场站'),
            '清分日期': clear_date,
            '文件合计电量(兆瓦时)': total_quantity,
            '文件合计电费(元)': total_fee,
            '场站计量电量(兆瓦时)': total_quantity
        }
        # 添加科目数据
        for trade in excel_trades:
            station_row[f'{trade}_电量'] = None
            station_row[f'{trade}_电价'] = None
            station_row[f'{trade}_电费'] = None
        
        return [station_row], excel_trades

    except Exception as e:
        st.warning(f"处理Excel {file_name} 出错: {str(e)}")
        return [], []

# ---------------------- 数据汇总与导出（动态列适配） ----------------------
def build_result_columns(global_trade_order):
    """根据科目顺序和类型，动态构建结果列（特殊科目仅电费列）"""
    base_cols = [
        '公司名称', '场站名称', '清分日期',
        '文件合计电量(兆瓦时)', '文件合计电费(元)', '场站计量电量(兆瓦时)'
    ]
    trade_cols = []
    for trade_name in global_trade_order:
        trade_type = 'special' if any(kw in trade_name for kw in SPECIAL_TRADE_KEYWORDS) else 'normal'
        if trade_type == 'normal':
            trade_cols.extend([f'{trade_name}_电量', f'{trade_name}_电价', f'{trade_name}_电费'])
        else:
            trade_cols.append(f'{trade_name}_电费')
    return base_cols + trade_cols

def calculate_summary_row(result_df, global_trade_order):
    """计算汇总行（适配常规/特殊科目）"""
    summary_row = {
        '公司名称': '总计',
        '场站名称': '总计',
        '清分日期': '',
        '文件合计电量(兆瓦时)': result_df['文件合计电量(兆瓦时)'].dropna().sum(),
        '文件合计电费(元)': result_df['文件合计电费(元)'].dropna().sum(),
        '场站计量电量(兆瓦时)': result_df['场站计量电量(兆瓦时)'].dropna().sum()
    }
    # 按顺序汇总科目数据
    for trade_name in global_trade_order:
        trade_type = 'special' if any(kw in trade_name for kw in SPECIAL_TRADE_KEYWORDS) else 'normal'
        if trade_type == 'normal':
            # 常规科目：电量/电费求和，电价求平均
            summary_row[f'{trade_name}_电量'] = result_df[f'{trade_name}_电量'].dropna().sum()
            summary_row[f'{trade_name}_电费'] = result_df[f'{trade_name}_电费'].dropna().sum()
            price_vals = result_df[f'{trade_name}_电价'].dropna()
            price_vals = price_vals[price_vals > 0.01]
            summary_row[f'{trade_name}_电价'] = round(price_vals.mean(), 3) if not price_vals.empty else None
        else:
            # 特殊科目：仅电费求和
            summary_row[f'{trade_name}_电费'] = result_df[f'{trade_name}_电费'].dropna().sum()
    return pd.DataFrame([summary_row])

def to_excel_bytes(df, report_df):
    """Excel导出（保留列顺序）"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='结算数据明细', index=False)
        report_df.to_excel(writer, sheet_name='处理报告', index=False)
    output.seek(0)
    return output

# ---------------------- Streamlit 页面布局 ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取（全科目版）", layout="wide")
    
    # 页面标题与说明
    st.title("📊 黑龙江日清分结算单数据提取工具（全科目+特殊场景适配）")
    st.divider()
    st.subheader("🔍 功能说明")
    st.markdown("""
    - **全科目覆盖**：自动提取“送江苏/浙江绿电交易”“阻塞费用”“价差费用”等所有科目
    - **特殊科目适配**：“阻塞费用”“价差费用”仅生成电费列，无多余电量/电价列
    - **顺序严格保留**：按PDF从上到下的原始顺序提取科目，不打乱
    - **无效列过滤**：彻底排除“hf”“县”等无关字符，无多余列
    """)

    # 1. 文件上传区域
    st.subheader("📁 上传文件")
    uploaded_files = st.file_uploader(
        "支持PDF/Excel批量上传（PDF优先，自动适配所有科目）",
        type=['pdf', 'xlsx'],
        accept_multiple_files=True
    )

    # 2. 数据处理逻辑
    if uploaded_files and st.button("🚀 开始处理", type="primary"):
        st.divider()
        st.subheader("⚙️ 处理进度")
        
        all_output_data = []
        global_trade_order = []  # 全局科目顺序（按文件处理顺序追加，保留单个文件内顺序）
        total_files = len(uploaded_files)
        processed_files = 0

        # 批量处理文件（按上传顺序）
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, file in enumerate(uploaded_files):
            file_name = file.name
            status_text.text(f"正在处理：{file_name}（{idx+1}/{total_files}）")
            
            # 按文件类型提取
            if file_name.lower().endswith('.pdf'):
                file_output, file_trade_order = extract_data_from_pdf(file, file_name)
            else:
                file_output, file_trade_order = extract_data_from_excel(file, file_name)
            
            # 累积数据（关键：保留单个文件内的科目顺序，全局按文件顺序追加）
            if file_output:
                all_output_data.extend(file_output)
                # 追加科目顺序（不重复）
                for trade in file_trade_order:
                    if trade not in global_trade_order:
                        global_trade_order.append(trade)
                processed_files += 1
            
            # 更新进度
            progress_bar.progress((idx + 1) / total_files)

        progress_bar.empty()
        status_text.text("处理完成！")

        # 3. 结果展示与导出
        if all_output_data and global_trade_order:
            st.divider()
            st.subheader("📈 提取结果（按PDF原始顺序）")
            
            # 动态构建结果列（适配常规/特殊科目）
            result_columns = build_result_columns(global_trade_order)
            # 构建DataFrame
            result_df = pd.DataFrame(all_output_data)
            # 补充缺失列（不同文件科目差异）
            for col in result_columns:
                if col not in result_df.columns:
                    result_df[col] = None
            # 严格按顺序排列列
            result_df = result_df[result_columns]
            # 数值列格式化
            numeric_cols = [col for col in result_columns if any(key in col for key in ['电量', '电价', '电费'])]
            result_df[numeric_cols] = result_df[numeric_cols].apply(pd.to_numeric, errors='coerce')

            # 排序（公司→场站→日期）
            result_df['清分日期'] = pd.to_datetime(result_df['清分日期'], errors='coerce')
            result_df = result_df.sort_values(['公司名称', '场站名称', '清分日期']).reset_index(drop=True)
            result_df['清分日期'] = result_df['清分日期'].dt.strftime('%Y-%m-%d').fillna('')

            # 添加汇总行
            summary_row = calculate_summary_row(result_df, global_trade_order)
            result_df = pd.concat([result_df, summary_row], ignore_index=True)

            # 生成处理报告
            failed_files = total_files - processed_files
            success_rate = f"{processed_files / total_files:.2%}" if total_files > 0 else "0%"
            stations = result_df['场站名称'].unique()
            station_count = len(stations) - 1 if '总计' in stations else len(stations)
            valid_rows = len(result_df) - 1
            trade_count = len(global_trade_order)

            report_df = pd.DataFrame({
                '统计项': ['文件总数', '成功处理数', '失败数', '处理成功率', '涉及场站数', '有效数据行数', '提取科目数'],
                '数值': [total_files, processed_files, failed_files,
                         success_rate, station_count, valid_rows, trade_count]
            })

            # 展示结果（分标签页）
            tab1, tab2 = st.tabs(["结算数据明细", "处理报告"])
            with tab1:
                st.dataframe(result_df, use_container_width=True, height=600)
                # 显示科目顺序说明
                st.caption(f"科目提取顺序（按PDF原始顺序）：{', '.join(global_trade_order)}")
            with tab2:
                st.dataframe(report_df, use_container_width=True)

            # 导出Excel
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            download_filename = f"黑龙江结算数据提取_全科目版_{current_time}.xlsx"
            excel_bytes = to_excel_bytes(result_df, report_df)

            st.divider()
            st.download_button(
                label="📥 导出全科目版Excel",
                data=excel_bytes,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            # 显示统计信息
            st.success(
                f"""✅ 处理完成！
                - 共处理 {total_files} 个文件，成功 {processed_files} 个（成功率 {success_rate}）
                - 提取 {trade_count} 个科目（含特殊科目），涉及 {station_count} 个场站，{valid_rows} 行有效数据
                - 特殊科目“阻塞费用”“价差费用”仅保留电费列，无多余列
                """
            )
        else:
            st.warning("⚠️ 未提取到有效数据！请检查PDF是否为可复制文本，且包含“机组 某某风电场”标识。")

    # 无文件上传时的提示
    elif not uploaded_files and st.button("🚀 开始处理", disabled=True):
        st.warning("请先上传PDF/Excel文件！")

if __name__ == "__main__":
    main()
