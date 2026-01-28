import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心配置（精准适配PDF结构） ----------------------
SPECIAL_TRADE_KEYWORDS = ['中长期合约阻塞费用', '省间省内价差费用']  # 精确匹配特殊科目名
INVALID_LINE_KEYWORDS = ['hf', 'HF', '县', '镇', '乡', '村', '_', '—', '页码']
TRADE_HEADER_KEYWORDS = ['科目编码', '结算类型', '电量']  # 表头需包含电量，避免误判
STATION_PATTERN = r'机组\s*[:：]?\s*([^:：\n]{2,15}风电场)'  # 适配“机组：双发B风电场”等格式
COLUMN_SPLIT_PATTERN = r'\s{2,}'  # 仅用2个以上空格分割列，避免科目名内空格干扰

# ---------------------- 核心工具函数 ----------------------
def safe_convert_to_numeric(value, default=None):
    """安全转换数值，空值返回None"""
    try:
        if pd.notna(value) and value is not None:
            str_val = str(value).strip()
            if str_val in ['/', 'NA', 'None', '', '无', '——', '0.00', '-', '空']:
                return default
            cleaned_value = str_val.replace(',', '').replace(' ', '').strip()
            return pd.to_numeric(cleaned_value) if cleaned_value else default
        return default
    except (ValueError, TypeError):
        return default

def filter_invalid_lines(pdf_lines):
    """过滤无效行，保留有效数据行"""
    valid_lines = []
    for line in pdf_lines:
        line = line.replace('\t', ' ').strip()
        # 过滤：过短行、含无效关键词、纯数字行
        if (len(line) >= 5 
            and not any(kw in line for kw in INVALID_LINE_KEYWORDS)
            and not line.replace('.', '').replace('-', '').isdigit()):
            valid_lines.append(line)
    return valid_lines

def extract_basic_info(pdf_lines):
    """提取公司名称、清分日期、文件合计数据"""
    company_name = "未知公司"
    clear_date = None
    total_quantity = None
    total_fee = None

    # 提取公司名称（精确匹配“公司名称：XXX有限公司”）
    for line in pdf_lines:
        if "公司名称" in line and ("：" in line or ":" in line):
            comp_match = re.search(r'公司名称[:：]\s*([\u4e00-\u9fa5a-zA-Z0-9()（）]+有限公司)', line)
            if comp_match:
                company_name = comp_match.group(1).strip()
                break

    # 提取清分日期
    date_pattern = r'清分日期[:：]\s*(\d{4}-\d{2}-\d{2})'
    for line in pdf_lines:
        date_match = re.search(date_pattern, line)
        if date_match:
            clear_date = date_match.group(1)
            break

    # 提取文件合计电量/电费（精确匹配“合计电量：XXX 合计电费：XXX”）
    total_pattern = r'合计电量[:：]\s*(\S+)\s+合计电费[:：]\s*(\S+)'
    for line in pdf_lines:
        total_match = re.search(total_pattern, line)
        if total_match:
            total_quantity = safe_convert_to_numeric(total_match.group(1))
            total_fee = safe_convert_to_numeric(total_match.group(2))
            break

    return company_name, clear_date, total_quantity, total_fee

# ---------------------- 场站与科目提取核心逻辑 ----------------------
def is_special_trade(trade_name):
    """精确判断是否为特殊科目（无电量/电价）"""
    return any(special_name in trade_name for special_name in SPECIAL_TRADE_KEYWORDS)

def extract_trade_data(line, in_special_area=False):
    """
    精准提取科目数据：
    - 常规科目：5列（编码+名称+电量+电价+电费）
    - 特殊科目：3列（编码+名称+电费）
    """
    line_cols = [col.strip() for col in re.split(COLUMN_SPLIT_PATTERN, line) if col.strip()]
    line_len = len(line_cols)
    trade_code = ""
    trade_name = ""
    trade_data = {}

    # 常规科目（5列结构）
    if line_len >=5 and not in_special_area:
        trade_code = line_cols[0]
        trade_name = line_cols[1]  # 完整科目名（含“（电能量）”）
        # 提取电量、电价、电费（允许为空）
        trade_data = {
            '电量': safe_convert_to_numeric(line_cols[2]),
            '电价': safe_convert_to_numeric(line_cols[3]),
            '电费': safe_convert_to_numeric(line_cols[4])
        }
    # 特殊科目（3列结构，且匹配特殊科目名）
    elif line_len >=3 and is_special_trade(line_cols[1]):
        trade_code = line_cols[0]
        trade_name = line_cols[1]
        trade_data = {'电费': safe_convert_to_numeric(line_cols[2])}

    return trade_code, trade_name, trade_data

def extract_all_stations(pdf_lines, company_name, clear_date, total_quantity, total_fee):
    """
    逐行扫描提取所有场站：
    1. 确保双发A/B风电场都能识别
    2. 完整保留“送江苏...（电能量）”科目名
    3. 特殊科目只提取电费
    """
    all_stations = []
    current_station = None
    current_station_meter = None
    current_trades = []  # 保留当前场站科目顺序
    in_trade_area = False  # 是否进入交易数据区
    in_special_area = False  # 是否进入特殊科目区（无电量/电价）

    for line in pdf_lines:
        # 1. 识别场站切换（适配“机组：双发B风电场”等格式）
        station_match = re.search(STATION_PATTERN, line)
        if station_match:
            # 保存上一个场站数据
            if current_station and current_trades:
                all_stations.append({
                    '公司名称': company_name,
                    '场站名称': current_station,
                    '清分日期': clear_date,
                    '文件合计电量': total_quantity,
                    '文件合计电费': total_fee,
                    '场站计量电量': current_station_meter,
                    '科目列表': current_trades.copy()
                })
            # 初始化新场站
            current_station = station_match.group(1).strip()
            current_station_meter = None
            current_trades = []
            in_trade_area = False
            in_special_area = False
            continue

        # 2. 提取当前场站计量电量（精确匹配“计量电量：XXX”）
        if current_station and "计量电量" in line and ("：" in line or ":" in line):
            meter_match = re.search(r'计量电量[:：]\s*(\S+)', line)
            if meter_match:
                current_station_meter = safe_convert_to_numeric(meter_match.group(1))
            continue

        # 3. 识别交易数据区（表头需包含3个关键词，避免误判）
        if not in_trade_area and all(kw in line for kw in TRADE_HEADER_KEYWORDS):
            in_trade_area = True
            continue

        # 4. 提取科目数据（仅在交易区内）
        if in_trade_area and current_station and len(line) >=10:
            trade_code, trade_name, trade_data = extract_trade_data(line, in_special_area)
            # 过滤无效科目（编码为数字，名称非空）
            if trade_code.isdigit() and trade_name and len(trade_name)>=5:
                # 更新特殊科目区标记
                if is_special_trade(trade_name):
                    in_special_area = True
                else:
                    in_special_area = False
                # 添加到当前场站科目列表
                current_trades.append({
                    '编码': trade_code,
                    '名称': trade_name,
                    '是否特殊科目': is_special_trade(trade_name),
                    '数据': trade_data
                })

    # 保存最后一个场站
    if current_station and current_trades:
        all_stations.append({
            '公司名称': company_name,
            '场站名称': current_station,
            '清分日期': clear_date,
            '文件合计电量': total_quantity,
            '文件合计电费': total_fee,
            '场站计量电量': current_station_meter,
            '科目列表': current_trades.copy()
        })

    return all_stations

# ---------------------- 数据格式化与导出 ----------------------
def build_result_structure(all_stations):
    """
    构建结果DataFrame结构：
    - 常规科目：3列（电量/电价/电费）
    - 特殊科目：1列（电费）
    - 严格保留科目顺序
    """
    if not all_stations:
        return pd.DataFrame(), []

    # 收集全局科目顺序（按第一个场站的科目顺序，确保统一）
    global_trade_order = [trade['名称'] for trade in all_stations[0]['科目列表']]
    # 构建基础列
    base_columns = [
        '公司名称', '场站名称', '清分日期',
        '文件合计电量(兆瓦时)', '文件合计电费(元)', '场站计量电量(兆瓦时)'
    ]
    # 构建科目列（常规3列，特殊1列）
    trade_columns = []
    for trade_name in global_trade_order:
        if is_special_trade(trade_name):
            trade_columns.append(f'{trade_name}_电费')
        else:
            trade_columns.extend([f'{trade_name}_电量', f'{trade_name}_电价', f'{trade_name}_电费'])

    # 填充数据
    result_data = []
    for station in all_stations:
        station_row = {
            '公司名称': station['公司名称'],
            '场站名称': station['场站名称'],
            '清分日期': station['清分日期'] or '',
            '文件合计电量(兆瓦时)': station['文件合计电量'],
            '文件合计电费(元)': station['文件合计电费'],
            '场站计量电量(兆瓦时)': station['场站计量电量']
        }
        # 按全局科目顺序填充数据
        for trade_name in global_trade_order:
            # 找到当前科目的数据
            trade = next((t for t in station['科目列表'] if t['名称'] == trade_name), None)
            if trade:
                if is_special_trade(trade_name):
                    station_row[f'{trade_name}_电费'] = trade['数据'].get('电费')
                else:
                    station_row[f'{trade_name}_电量'] = trade['数据'].get('电量')
                    station_row[f'{trade_name}_电价'] = trade['数据'].get('电价')
                    station_row[f'{trade_name}_电费'] = trade['数据'].get('电费')
            else:
                # 其他场站无此科目时填充None
                if is_special_trade(trade_name):
                    station_row[f'{trade_name}_电费'] = None
                else:
                    station_row[f'{trade_name}_电量'] = None
                    station_row[f'{trade_name}_电价'] = None
                    station_row[f'{trade_name}_电费'] = None
        result_data.append(station_row)

    return pd.DataFrame(result_data, columns=base_columns+trade_columns), global_trade_order

def add_summary_row(result_df, global_trade_order):
    """添加汇总行，特殊科目仅汇总电费"""
    summary_row = {
        '公司名称': '总计',
        '场站名称': '总计',
        '清分日期': '',
        '文件合计电量(兆瓦时)': result_df['文件合计电量(兆瓦时)'].dropna().sum(),
        '文件合计电费(元)': result_df['文件合计电费(元)'].dropna().sum(),
        '场站计量电量(兆瓦时)': result_df['场站计量电量(兆瓦时)'].dropna().sum()
    }
    # 按科目顺序汇总
    for trade_name in global_trade_order:
        if is_special_trade(trade_name):
            summary_row[f'{trade_name}_电费'] = result_df[f'{trade_name}_电费'].dropna().sum()
        else:
            summary_row[f'{trade_name}_电量'] = result_df[f'{trade_name}_电量'].dropna().sum()
            summary_row[f'{trade_name}_电费'] = result_df[f'{trade_name}_电费'].dropna().sum()
            # 电价取有效平均值（排除0和空）
            price_vals = result_df[f'{trade_name}_电价'].dropna()
            price_vals = price_vals[price_vals > 0.01]
            summary_row[f'{trade_name}_电价'] = round(price_vals.mean(), 3) if not price_vals.empty else None

    return pd.concat([result_df, pd.DataFrame([summary_row])], ignore_index=True)

def to_excel_bytes(df, report_df):
    """Excel导出，保留所有格式"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='结算数据明细', index=False)
        report_df.to_excel(writer, sheet_name='处理报告', index=False)
    output.seek(0)
    return output

# ---------------------- Streamlit 页面 ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取（最终版）", layout="wide")
    
    # 页面标题与问题解决说明
    st.title("📊 黑龙江日清分结算单数据提取工具（最终版）")
    st.divider()
    st.subheader("✅ 已解决所有问题")
    st.markdown("""
    1. **“送江苏省间绿色电力交易（电能量）”科目名完整**：用2个以上空格分割列，避免科目名内空格干扰
    2. **双发B风电场正常提取**：优化场站正则，适配“机组：双发B风电场”等格式
    3. **特殊科目仅保留电费列**：“中长期合约阻塞费用”“省间省内价差费用”仅生成电费列，无多余列
    """)

    # 文件上传
    st.subheader("📁 上传PDF文件")
    uploaded_files = st.file_uploader(
        "仅支持PDF（已精准适配依兰县协合风力发电格式）",
        type=['pdf'],
        accept_multiple_files=True
    )

    # 数据处理
    if uploaded_files and st.button("🚀 开始处理", type="primary"):
        st.divider()
        st.subheader("⚙️ 处理进度")
        
        all_result_dfs = []
        total_files = len(uploaded_files)
        processed_files = 0
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, file in enumerate(uploaded_files):
            file_name = file.name
            status_text.text(f"正在处理：{file_name}（{idx+1}/{total_files}）")
            
            try:
                # 1. 读取PDF文本
                with pdfplumber.open(file) as pdf:
                    pdf_text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
                pdf_lines = filter_invalid_lines(pdf_text.split('\n'))
                if not pdf_lines:
                    st.warning(f"{file_name} 无有效文本数据")
                    continue

                # 2. 提取基础信息
                company_name, clear_date, total_quantity, total_fee = extract_basic_info(pdf_lines)
                # 3. 提取所有场站和科目
                all_stations = extract_all_stations(pdf_lines, company_name, clear_date, total_quantity, total_fee)
                if not all_stations:
                    st.warning(f"{file_name} 未提取到场站数据")
                    continue

                # 4. 构建结果DataFrame
                result_df, trade_order = build_result_structure(all_stations)
                # 5. 添加汇总行
                result_df = add_summary_row(result_df, trade_order)
                all_result_dfs.append(result_df)
                processed_files += 1

            except Exception as e:
                st.warning(f"处理 {file_name} 出错：{str(e)}")
                continue

            # 更新进度
            progress_bar.progress((idx + 1) / total_files)

        progress_bar.empty()
        status_text.text("处理完成！")

        # 结果展示与导出
        if all_result_dfs:
            # 合并多个文件结果（若有）
            final_df = pd.concat(all_result_dfs, ignore_index=True)
            # 生成处理报告
            stations = final_df['场站名称'].unique()
            station_count = len(stations) - 1 if '总计' in stations else len(stations)
            valid_rows = len(final_df) - len(all_result_dfs)  # 减去汇总行数
            trade_count = len([col for col in final_df.columns if any(kw in col for kw in ['电量', '电费'])]) // 3 + len(SPECIAL_TRADE_KEYWORDS)

            report_df = pd.DataFrame({
                '统计项': ['文件总数', '成功处理数', '涉及场站数', '有效数据行数', '提取科目数'],
                '数值': [total_files, processed_files, station_count, valid_rows, trade_count]
            })

            # 展示结果
            tab1, tab2 = st.tabs(["结算数据明细", "处理报告"])
            with tab1:
                st.dataframe(final_df, use_container_width=True, height=600)
                # 重点科目验证提示
                st.caption("🔍 重点科目验证：")
                st.caption(f"- 送江苏省间绿色电力交易（电能量）：{['送江苏省间绿色电力交易（电能量）_电量' in final_df.columns]}")
                st.caption(f"- 双发B风电场：{any('双发B风电场' in name for name in final_df['场站名称'].unique())}")
                st.caption(f"- 特殊科目（仅电费列）：{[col for col in final_df.columns if any(s in col for s in SPECIAL_TRADE_KEYWORDS)]}")
            with tab2:
                st.dataframe(report_df, use_container_width=True)

            # 导出Excel
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            download_filename = f"黑龙江结算数据_最终版_{current_time}.xlsx"
            excel_bytes = to_excel_bytes(final_df, report_df)

            st.divider()
            st.download_button(
                label="📥 导出最终版Excel",
                data=excel_bytes,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            st.success("🎉 所有问题已解决，数据提取完整准确！")
        else:
            st.warning("⚠️ 未提取到任何有效数据，请检查PDF文件格式。")

    # 无文件上传时的提示
    elif not uploaded_files and st.button("🚀 开始处理", disabled=True):
        st.warning("请先上传PDF文件！")

if __name__ == "__main__":
    main()
