import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心配置（新增：解决科目/场站/无效列问题） ----------------------
# 新增：扩展目标科目（包含缺失科目，区分常规/特殊科目）
NORMAL_TRADES = [
    '优先发电交易',
    '电网企业代理购电交易',
    '省内电力直接交易',
    '送上海省间绿色电力交易(电能量 )',
    '送辽宁交易',
    '送华北交易',
    '送山东交易',
    '送浙江交易',
    '送江苏省间绿色电力交易（电能量）',  # 新增缺失科目
    '送浙江省间绿色电力交易（电能量）',  # 新增缺失科目
    '省内现货日前交易',
    '省内现货实时交易',
    '省间现货日前交易',
    '省间现货日内交易'
]
# 新增：特殊科目（仅提取电费）
SPECIAL_TRADES = [
    '中长期合约阻塞费用',
    '省间省内价差费用'
]
# 新增：无效关键词过滤（消除hf_、县_等多余列）
INVALID_KEYWORDS = ['hf', 'HF', '县', '镇', '乡', '村', '_', '—']
# 新增：优化场站识别（支持双发A/B风电场）
STATION_PATTERNS = [
    r'公司名称:\s*([^\s]+风电场)',  # 优先匹配风电场名称
    r'机组\s*[:：]?\s*([^\s]+风电场)',  # 适配双发A/B风电场格式
    r'公司名称:\s*([^\s]+有限公司)'  # 原有规则兜底
]

# ---------------------- 核心提取函数（保留原逻辑，新增优化） ----------------------
def extract_station_name(pdf_lines):
    """优化：适配双发A/B风电场，精准提取场站名称"""
    # 优先匹配风电场名称（解决双发B风电场提取问题）
    for pattern in STATION_PATTERNS:
        for line in pdf_lines:
            match = re.search(pattern, line)
            if match:
                station_name = match.group(1).strip()
                # 格式统一
                station_name = re.sub(r'太阳能发电有限公司$', '光伏电站', station_name)
                return station_name
    return "未知场站"

def safe_convert_to_numeric(value, default=None):
    """保留原逻辑，兼容更多空值场景"""
    try:
        if pd.notna(value) and value is not None:
            str_val = str(value).strip()
            # 新增：兼容更多空值标识
            if str_val in ['/', 'NA', 'None', '', '无', '——', '0.00', '-', '空']:
                return default
            cleaned_value = str_val.replace(',', '').replace(' ', '').strip()
            return pd.to_numeric(cleaned_value)
        return default
    except (ValueError, TypeError):
        return default

def filter_invalid_lines(pdf_lines):
    """新增：过滤含无效关键词的行，消除多余列"""
    valid_lines = []
    for line in pdf_lines:
        line = line.strip()
        # 过滤：过短行、含无效关键词、纯数字行
        if (len(line) >= 5 
            and not any(kw in line for kw in INVALID_KEYWORDS)
            and not line.replace('.', '').replace('-', '').isdigit()):
            valid_lines.append(line)
    return valid_lines

def extract_trade_data_by_column(trade_name, pdf_lines, is_special=False):
    """优化：适配常规/特殊科目提取，解决科目拆分问题"""
    quantity = None
    price = None
    fee = None

    # 新增：用2个以上空格分割列（避免科目名内空格拆分）
    for idx, line in enumerate(pdf_lines):
        line_cols = [col.strip() for col in re.split(r'\s{2,}', line) if col.strip()]
        
        # 常规科目：5列结构（编码+名称+电量+电价+电费）
        if not is_special and len(line_cols) >= 5 and trade_name in line_cols[1]:
            quantity = safe_convert_to_numeric(line_cols[2])
            price = safe_convert_to_numeric(line_cols[3])
            fee = safe_convert_to_numeric(line_cols[4])
            break
        # 特殊科目：3列结构（编码+名称+电费）
        elif is_special and len(line_cols) >= 3 and trade_name in line_cols[1]:
            fee = safe_convert_to_numeric(line_cols[2])
            break
    return [quantity, price, fee] if not is_special else [fee]

def extract_data_from_pdf(file_obj, file_name):
    """保留原结构，新增：支持特殊科目、过滤无效行、完整科目提取"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            if not pdf.pages:
                raise ValueError("PDF无有效页面")
            
            all_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    all_text += page_text + "\n"
        # 新增：过滤无效行（解决多余列问题）
        pdf_lines = filter_invalid_lines(all_text.split('\n'))
        if not pdf_lines:
            raise ValueError("PDF为扫描件，无可用文本")

        # 1. 提取场站名称（优化后）
        station_name = extract_station_name(pdf_lines)

        # 2. 提取清分日期（保留原逻辑）
        date = None
        date_pattern = r'清分日期\s*(\d{4}-\d{2}-\d{2})'
        for line in pdf_lines:
            date_match = re.search(date_pattern, line)
            if date_match:
                date = date_match.group(1)
                break

        # 3. 提取合计电量和合计电费（保留原逻辑）
        total_quantity = None
        total_amount = None
        for line in pdf_lines:
            line_cols = [col.strip() for col in re.split(r'\s{2,}', line) if col.strip()]  # 优化列分割
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

        # 4. 提取所有目标科目的数据（新增：区分常规/特殊科目）
        all_trade_data = []
        # 提取常规科目（3列：电量/电价/电费）
        for trade in NORMAL_TRADES:
            trade_data = extract_trade_data_by_column(trade, pdf_lines, is_special=False)
            all_trade_data.extend(trade_data)
        # 提取特殊科目（仅1列：电费）
        for trade in SPECIAL_TRADES:
            trade_data = extract_trade_data_by_column(trade, pdf_lines, is_special=True)
            all_trade_data.extend(trade_data)

        return [station_name, date, total_quantity, total_amount] + all_trade_data

    except Exception as e:
        st.warning(f"处理PDF {file_name} 出错: {str(e)}")
        # 新增：适配特殊科目后的返回值长度
        return ["未知场站", None, None, None] + [None] * (len(NORMAL_TRADES)*3 + len(SPECIAL_TRADES))

def extract_data_from_excel(file_obj, file_name):
    """保留原Excel处理逻辑，新增：适配新科目"""
    try:
        df = pd.read_excel(file_obj, dtype=object)
        station_name = "未知场站"
        # 从文件名提取场站名（保留原逻辑）
        name_without_ext = file_name.split('.')[0]
        if "晶盛" in name_without_ext:
            station_name = "大庆晶盛光伏电站"
        
        # 提取日期（保留原逻辑）
        date_match = re.search(r'\d{4}-\d{2}-\d{2}', name_without_ext)
        date = date_match.group() if date_match else None

        # 提取合计数据（保留原逻辑）
        total_quantity = safe_convert_to_numeric(df.iloc[0, 3] if len(df) > 0 else None)
        total_amount = safe_convert_to_numeric(df.iloc[0, 5] if len(df) > 0 else None)

        # 提取科目数据（新增：适配常规/特殊科目）
        all_trade_data = []
        # 常规科目：填充None（3列）
        for _ in NORMAL_TRADES:
            all_trade_data.extend([None, None, None])
        # 特殊科目：填充None（1列）
        for _ in SPECIAL_TRADES:
            all_trade_data.append(None)

        return [station_name, date, total_quantity, total_amount] + all_trade_data

    except Exception as e:
        st.warning(f"处理Excel {file_name} 出错: {str(e)}")
        return ["未知场站", None, None, None] + [None] * (len(NORMAL_TRADES)*3 + len(SPECIAL_TRADES))

def calculate_summary_row(data_df):
    """优化：适配特殊科目汇总（仅汇总电费）"""
    # 常规科目：求和电量/电费，平均电价
    sum_cols = [col for col in data_df.columns if any(key in col for key in ['电量', '电费']) and not any(s in col for s in SPECIAL_TRADES)]
    avg_cols = [col for col in data_df.columns if '电价' in col]
    # 特殊科目：仅求和电费
    special_fee_cols = [col for col in data_df.columns if any(s in col for s in SPECIAL_TRADES) and '电费' in col]

    summary_row = {'场站名称': '总计', '清分日期': ''}
    # 常规科目汇总
    for col in sum_cols:
        valid_vals = data_df[col].dropna()
        summary_row[col] = valid_vals.sum() if not valid_vals.empty else 0
    for col in avg_cols:
        valid_vals = data_df[col].dropna()
        summary_row[col] = round(valid_vals.mean(), 3) if not valid_vals.empty else None
    # 特殊科目汇总
    for col in special_fee_cols:
        valid_vals = data_df[col].dropna()
        summary_row[col] = valid_vals.sum() if not valid_vals.empty else 0

    return pd.DataFrame([summary_row])

def to_excel_bytes(df, report_df):
    """保留原逻辑：转为Excel字节流"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='结算数据明细', index=False)
        report_df.to_excel(writer, sheet_name='处理报告', index=False)
    output.seek(0)
    return output

# ---------------------- Streamlit 页面布局与交互（保留所有原功能） ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取", layout="wide")
    
    # 页面标题（保留原样式）
    st.title("📊 黑龙江日清分结算单数据提取工具")
    st.divider()

    # 1. 文件上传区域（保留原逻辑）
    st.subheader("📁 上传文件")
    uploaded_files = st.file_uploader(
        "支持PDF/Excel格式，可批量上传",
        type=['pdf', 'xlsx'],
        accept_multiple_files=True
    )

    # 2. 数据处理逻辑（保留原流程，适配新科目）
    if uploaded_files and st.button("🚀 开始处理", type="primary"):
        st.divider()
        st.subheader("⚙️ 处理进度")
        
        all_data = []
        total_files = len(uploaded_files)
        processed_files = 0

        # 批量处理上传的文件（保留原进度条）
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, file in enumerate(uploaded_files):
            file_name = file.name
            status_text.text(f"正在处理：{file_name}")
            
            # 根据文件类型调用对应提取函数（保留原逻辑）
            if file_name.lower().endswith('.pdf'):
                data = extract_data_from_pdf(file, file_name)
            else:
                data = extract_data_from_excel(file, file_name)
            
            # 验证数据有效性（保留原逻辑）
            if data[1] is not None and any(isinstance(val, (float, int)) for val in data[2:] if val is not None):
                all_data.append(data)
                processed_files += 1
            
            # 更新进度（保留原逻辑）
            progress_bar.progress((idx + 1) / total_files)

        progress_bar.empty()
        status_text.text("处理完成！")

        # 3. 结果展示与导出（适配新科目列）
        if all_data:
            st.divider()
            st.subheader("📈 提取结果")
            
            # 构建结果列（新增：包含所有新科目，特殊科目仅电费列）
            result_columns = [
                '场站名称', '清分日期', '合计电量(兆瓦时)', '合计电费(元)'
            ]
            # 常规科目列（3列：电量/电价/电费）
            for trade in NORMAL_TRADES:
                # 列名简化（避免过长）
                trade_col_name = trade.replace('（电能量）', '').replace('(电能量 )', '').replace('省间绿色电力交易', '省间绿电交易')
                result_columns.extend([
                    f'{trade_col_name}_电量',
                    f'{trade_col_name}_电价',
                    f'{trade_col_name}_电费'
                ])
            # 特殊科目列（仅电费列）
            for trade in SPECIAL_TRADES:
                result_columns.append(f'{trade}_电费')

            # 构建DataFrame（保留原逻辑）
            result_df = pd.DataFrame(all_data, columns=result_columns)
            num_cols = result_df.columns[2:]
            result_df[num_cols] = result_df[num_cols].apply(pd.to_numeric, errors='coerce')

            # 排序并格式化日期（保留原逻辑）
            result_df['清分日期'] = pd.to_datetime(result_df['清分日期'], errors='coerce')
            result_df = result_df.sort_values(['场站名称', '清分日期']).reset_index(drop=True)
            result_df['清分日期'] = result_df['清分日期'].dt.strftime('%Y-%m-%d').fillna('')

            # 添加汇总行（优化后）
            summary_row = calculate_summary_row(result_df)
            result_df = pd.concat([result_df, summary_row], ignore_index=True)

            # 生成处理报告（保留原逻辑）
            failed_files = total_files - processed_files
            success_rate = f"{processed_files / total_files:.2%}" if total_files > 0 else "0%"
            stations = result_df['场站名称'].unique()
            station_count = len(stations) - 1 if '总计' in stations else len(stations)
            valid_rows = len(result_df) - 1

            report_df = pd.DataFrame({
                '统计项': ['文件总数', '成功处理数', '失败数', '处理成功率', '涉及场站数', '有效数据行数'],
                '数值': [total_files, processed_files, failed_files,
                         success_rate, station_count, valid_rows]
            })

            # 展示结果表格（保留原标签页）
            tab1, tab2 = st.tabs(["结算数据明细", "处理报告"])
            with tab1:
                st.dataframe(result_df, use_container_width=True)
            with tab2:
                st.dataframe(report_df, use_container_width=True)

            # 生成下载文件（保留原逻辑）
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            download_filename = f"黑龙江结算数据提取_{current_time}.xlsx"
            excel_bytes = to_excel_bytes(result_df, report_df)

            # 下载按钮（保留原样式）
            st.divider()
            st.download_button(
                label="📥 导出Excel文件",
                data=excel_bytes,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            # 显示统计信息（保留原逻辑）
            st.info(
                f"""处理完成！
                - 总计上传 {total_files} 个文件，成功处理 {processed_files} 个（成功率 {success_rate}）
                - 涉及 {station_count} 个场站，{valid_rows} 行有效数据
                - 已提取所有科目（含送江苏/浙江绿电交易、阻塞费用、价差费用）
                """
            )
        else:
            st.warning("⚠️ 未提取到有效数据！请检查：")
            st.markdown("""
                1. PDF是否为可复制文本（非扫描件）；
                2. 文件是否为黑龙江日清分格式；
                3. Excel文件格式是否匹配。
            """)

    # 无文件上传时的提示（保留原逻辑）
    elif not uploaded_files and st.button("🚀 开始处理", disabled=True):
        st.warning("请先上传PDF/Excel文件！")

if __name__ == "__main__":
    main()
