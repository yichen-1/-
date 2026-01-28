import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心配置（适配全角符号+扩展科目） ----------------------
# 常规科目（3列：电量/电价/电费）
NORMAL_TRADES = [
    '优先发电交易',
    '电网企业代理购电交易',
    '省内电力直接交易',
    '送上海省间绿色电力交易(电能量 )',
    '送辽宁交易',
    '送华北交易',
    '送山东交易',
    '送浙江交易',
    '送江苏省间绿色电力交易（电能量）',
    '送浙江省间绿色电力交易（电能量）',
    '省内现货日前交易',
    '省内现货实时交易',
    '省间现货日前交易',
    '省间现货日内交易'
]
# 特殊科目（仅1列：电费）
SPECIAL_TRADES = [
    '中长期合约阻塞费用',
    '省间省内价差费用'
]
# 无效关键词（过滤hf、县等）
INVALID_KEYWORDS = ['hf', 'HF', '县', '镇', '乡', '村', '_', '—']
# 场站识别规则（支持全角/半角冒号，适配双发A/B风电场）
STATION_PATTERNS = [
    r'公司名称[:：]\s*([^\s]+风电场)',
    r'机组\s*[:：]?\s*([^\s]+风电场)',
    r'公司名称[:：]\s*([^\s]+有限公司)'
]
# 清分日期识别规则（支持全角/半角冒号）
DATE_PATTERN = r'清分日期[:：]\s*(\d{4}-\d{2}-\d{2})'

# ---------------------- 版本检查（调试用） ----------------------
def check_dependency_versions():
    """检查关键库版本，方便排查问题"""
    st.sidebar.subheader("🔧 环境版本信息")
    st.sidebar.write(f"pdfplumber版本：{pdfplumber.__version__ if hasattr(pdfplumber, '__version__') else '未知'}")
    st.sidebar.write(f"pandas版本：{pd.__version__}")
    st.sidebar.write(f"streamlit版本：{st.__version__}")
    st.sidebar.divider()

# ---------------------- 核心工具函数 ----------------------
def safe_convert_to_numeric(value, default=None):
    """安全转换数值，兼容更多空值场景"""
    try:
        if pd.notna(value) and value is not None:
            str_val = str(value).strip()
            if str_val in ['/', 'NA', 'None', '', '无', '——', '0.00', '-', '空']:
                return default
            cleaned_value = str_val.replace(',', '').replace(' ', '').strip()
            return pd.to_numeric(cleaned_value)
        return default
    except (ValueError, TypeError):
        return default

def filter_invalid_lines(pdf_lines):
    """放宽过滤条件：仅过滤无效关键词，保留所有长度≥2的行"""
    valid_lines = []
    for line in pdf_lines:
        line = line.strip()
        if len(line) >= 2 and not any(kw in line for kw in INVALID_KEYWORDS):
            valid_lines.append(line)
    return valid_lines

def extract_station_name(pdf_lines):
    """适配全角符号，精准提取场站名称"""
    for pattern in STATION_PATTERNS:
        for line in pdf_lines:
            match = re.search(pattern, line)
            if match:
                station_name = match.group(1).strip()
                station_name = re.sub(r'太阳能发电有限公司$', '光伏电站', station_name)
                return station_name
    return "未知场站"

def extract_trade_data_by_column(trade_name, pdf_lines, is_special=False):
    """适配常规/特殊科目，用2个以上空格分割列"""
    quantity = None
    price = None
    fee = None

    for line in pdf_lines:
        line_cols = [col.strip() for col in re.split(r'\s{2,}', line) if col.strip()]
        # 常规科目（5列）
        if not is_special and len(line_cols) >= 5 and trade_name in line_cols[1]:
            quantity = safe_convert_to_numeric(line_cols[2])
            price = safe_convert_to_numeric(line_cols[3])
            fee = safe_convert_to_numeric(line_cols[4])
            break
        # 特殊科目（3列）
        elif is_special and len(line_cols) >= 3 and trade_name in line_cols[1]:
            fee = safe_convert_to_numeric(line_cols[2])
            break
    return [quantity, price, fee] if not is_special else [fee]

# ---------------------- PDF/Excel提取核心函数（带调试） ----------------------
def extract_data_from_pdf(file_obj, file_name):
    """PDF提取：带调试信息，适配全角符号"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            if not pdf.pages:
                raise ValueError("PDF无有效页面")
            
            all_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    all_text += page_text + "\n"
        
        # 调试：显示提取的前1000字符（方便排查）
        st.subheader(f"📝 {file_name} 文本提取结果（前1000字符）")
        st.text(all_text[:1000] if all_text else "无提取到文本（可能是扫描件）")
        
        # 过滤无效行
        pdf_lines = filter_invalid_lines(all_text.split('\n'))
        if not pdf_lines:
            raise ValueError("PDF为扫描件/无有效文本")

        # 1. 提取场站名称
        station_name = extract_station_name(pdf_lines)
        st.write(f"📍 提取的场站名称：{station_name}")

        # 2. 提取清分日期（支持全角冒号）
        date = None
        for line in pdf_lines:
            date_match = re.search(DATE_PATTERN, line)
            if date_match:
                date = date_match.group(1)
                break
        st.write(f"📅 提取的清分日期：{date if date else '未识别到'}")

        # 3. 提取合计电量/电费
        total_quantity = None
        total_amount = None
        for line in pdf_lines:
            line_cols = [col.strip() for col in re.split(r'\s{2,}', line) if col.strip()]
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
        st.write(f"📊 合计电量：{total_quantity} 兆瓦时 | 合计电费：{total_amount} 元")

        # 4. 提取科目数据
        all_trade_data = []
        # 常规科目
        for trade in NORMAL_TRADES:
            trade_data = extract_trade_data_by_column(trade, pdf_lines, is_special=False)
            all_trade_data.extend(trade_data)
        # 特殊科目
        for trade in SPECIAL_TRADES:
            trade_data = extract_trade_data_by_column(trade, pdf_lines, is_special=True)
            all_trade_data.extend(trade_data)
        
        st.success(f"✅ {file_name} 数据提取完成！")
        st.divider()
        return [station_name, date, total_quantity, total_amount] + all_trade_data

    except Exception as e:
        st.error(f"❌ 处理PDF {file_name} 出错: {str(e)}")
        return ["未知场站", None, None, None] + [None] * (len(NORMAL_TRADES)*3 + len(SPECIAL_TRADES))

def extract_data_from_excel(file_obj, file_name):
    """Excel提取：保留原有逻辑，适配新科目"""
    try:
        df = pd.read_excel(file_obj, dtype=object)
        station_name = "未知场站"
        # 从文件名提取场站名
        name_without_ext = file_name.split('.')[0]
        if "晶盛" in name_without_ext:
            station_name = "大庆晶盛光伏电站"
        
        # 提取日期
        date_match = re.search(r'\d{4}-\d{2}-\d{2}', name_without_ext)
        date = date_match.group() if date_match else None

        # 提取合计数据
        total_quantity = safe_convert_to_numeric(df.iloc[0, 3] if len(df) > 0 else None)
        total_amount = safe_convert_to_numeric(df.iloc[0, 5] if len(df) > 0 else None)

        # 提取科目数据
        all_trade_data = []
        for _ in NORMAL_TRADES:
            all_trade_data.extend([None, None, None])
        for _ in SPECIAL_TRADES:
            all_trade_data.append(None)

        st.success(f"✅ {file_name} Excel数据提取完成！")
        return [station_name, date, total_quantity, total_amount] + all_trade_data

    except Exception as e:
        st.error(f"❌ 处理Excel {file_name} 出错: {str(e)}")
        return ["未知场站", None, None, None] + [None] * (len(NORMAL_TRADES)*3 + len(SPECIAL_TRADES))

# ---------------------- 汇总与导出函数 ----------------------
def calculate_summary_row(data_df):
    """汇总行：适配特殊科目"""
    sum_cols = [col for col in data_df.columns if any(key in col for key in ['电量', '电费']) and not any(s in col for s in SPECIAL_TRADES)]
    avg_cols = [col for col in data_df.columns if '电价' in col]
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
    """转为Excel字节流"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='结算数据明细', index=False)
        report_df.to_excel(writer, sheet_name='处理报告', index=False)
    output.seek(0)
    return output

# ---------------------- Streamlit主界面 ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取", layout="wide")
    
    # 版本检查（侧边栏）
    check_dependency_versions()

    # 页面标题
    st.title("📊 黑龙江日清分结算单数据提取工具（最终版）")
    st.divider()

    # 1. 文件上传
    st.subheader("📁 上传文件")
    uploaded_files = st.file_uploader(
        "支持PDF/Excel格式，可批量上传",
        type=['pdf', 'xlsx'],
        accept_multiple_files=True
    )

    # 2. 数据处理
    if uploaded_files and st.button("🚀 开始处理", type="primary"):
        st.divider()
        st.subheader("⚙️ 处理进度与调试信息")
        
        all_data = []
        total_files = len(uploaded_files)
        processed_files = 0

        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, file in enumerate(uploaded_files):
            file_name = file.name
            status_text.text(f"正在处理：{file_name}（{idx+1}/{total_files}）")
            
            # 按文件类型提取
            if file_name.lower().endswith('.pdf'):
                data = extract_data_from_pdf(file, file_name)
            else:
                data = extract_data_from_excel(file, file_name)
            
            # 放宽有效数据判断：只要有数值就保留（即使日期为None）
            if any(isinstance(val, (float, int)) for val in data[2:] if val is not None):
                all_data.append(data)
                processed_files += 1
            
            progress_bar.progress((idx + 1) / total_files)

        progress_bar.empty()
        status_text.text("处理完成！")

        # 3. 结果展示
        if all_data:
            st.divider()
            st.subheader("📈 提取结果")
            
            # 构建列名
            result_columns = ['场站名称', '清分日期', '合计电量(兆瓦时)', '合计电费(元)']
            # 常规科目列
            for trade in NORMAL_TRADES:
                trade_col_name = trade.replace('（电能量）', '').replace('(电能量 )', '').replace('省间绿色电力交易', '省间绿电交易')
                result_columns.extend([f'{trade_col_name}_电量', f'{trade_col_name}_电价', f'{trade_col_name}_电费'])
            # 特殊科目列
            for trade in SPECIAL_TRADES:
                result_columns.append(f'{trade}_电费')

            # 构建DataFrame
            result_df = pd.DataFrame(all_data, columns=result_columns)
            num_cols = result_df.columns[2:]
            result_df[num_cols] = result_df[num_cols].apply(pd.to_numeric, errors='coerce')

            # 排序格式化
            result_df['清分日期'] = pd.to_datetime(result_df['清分日期'], errors='coerce')
            result_df = result_df.sort_values(['场站名称', '清分日期']).reset_index(drop=True)
            result_df['清分日期'] = result_df['清分日期'].dt.strftime('%Y-%m-%d').fillna('')

            # 添加汇总行
            summary_row = calculate_summary_row(result_df)
            result_df = pd.concat([result_df, summary_row], ignore_index=True)

            # 处理报告
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

            # 展示结果
            tab1, tab2 = st.tabs(["结算数据明细", "处理报告"])
            with tab1:
                st.dataframe(result_df, use_container_width=True, height=600)
            with tab2:
                st.dataframe(report_df, use_container_width=True)

            # 导出Excel
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            download_filename = f"黑龙江结算数据提取_{current_time}.xlsx"
            excel_bytes = to_excel_bytes(result_df, report_df)

            st.divider()
            st.download_button(
                label="📥 导出Excel文件",
                data=excel_bytes,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            # 统计信息
            st.info(
                f"""✅ 处理完成！
                - 总计上传 {total_files} 个文件，成功处理 {processed_files} 个（成功率 {success_rate}）
                - 涉及 {station_count} 个场站，{valid_rows} 行有效数据
                - 已提取所有科目（含送江苏/浙江绿电交易、阻塞费用、价差费用）
                """
            )
        else:
            st.warning("⚠️ 未提取到有效数据！请检查：")
            st.markdown("""
                1. PDF是否为可复制文本（非扫描件）；
                2. PDF中是否包含“清分日期”“机组 双发A/B风电场”等关键信息；
                3. 已开启调试模式，可查看文本提取结果定位问题。
            """)

    # 无文件上传时的提示
    elif not uploaded_files and st.button("🚀 开始处理", disabled=True):
        st.warning("请先上传PDF/Excel文件！")

if __name__ == "__main__":
    main()
