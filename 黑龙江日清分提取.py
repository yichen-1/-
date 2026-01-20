import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心提取函数（保留原逻辑，适配Streamlit文件对象） ----------------------
def extract_station_name(pdf_lines):
    """从PDF内容提取场站名称（优先取公司名称，更精准）"""
    for line in pdf_lines:
        if "公司名称:" in line:
            station_name = re.sub(r'公司名称:\s*', '', line).strip()
            station_name = re.sub(r'太阳能发电有限公司$', '光伏电站', station_name)
            return station_name
    return "未知场站"

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

def extract_trade_data_by_column(trade_name, pdf_lines):
    """适配黑龙江PDF格式：按"结算类型"匹配，提取电量/电价/电费"""
    quantity = None
    price = None
    fee = None

    for idx, line in enumerate(pdf_lines):
        line_cols = [col.strip() for col in re.split(r'\s+', line) if col.strip()]
        if len(line_cols) >= 5 and trade_name in line_cols[1]:
            quantity = safe_convert_to_numeric(line_cols[2])
            price = safe_convert_to_numeric(line_cols[3])
            fee = safe_convert_to_numeric(line_cols[4])
            break
    return [quantity, price, fee]

def extract_data_from_pdf(file_obj, file_name):
    """适配Streamlit：接收文件对象而非路径"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            if not pdf.pages:
                raise ValueError("PDF无有效页面")
            
            all_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    all_text += page_text + "\n"
            pdf_lines = [line.strip() for line in all_text.split('\n') if line.strip()]
            if not pdf_lines:
                raise ValueError("PDF为扫描件，无可用文本")

        # 1. 提取场站名称
        station_name = extract_station_name(pdf_lines)

        # 2. 提取清分日期
        date = None
        date_pattern = r'清分日期\s*(\d{4}-\d{2}-\d{2})'
        for line in pdf_lines:
            date_match = re.search(date_pattern, line)
            if date_match:
                date = date_match.group(1)
                break

        # 3. 提取合计电量和合计电费
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

        # 4. 目标结算科目
        target_trades = [
            '优先发电交易',
            '电网企业代理购电交易',
            '省内电力直接交易',
            '送上海省间绿色电力交易(电能量 )',
            '送辽宁交易',
            '送华北交易',
            '送山东交易',
            '送浙江交易',
            '省内现货日前交易',
            '省内现货实时交易',
            '省间现货日前交易',
            '省间现货日内交易'
        ]

        # 5. 提取所有目标科目的数据
        all_trade_data = []
        for trade in target_trades:
            trade_data = extract_trade_data_by_column(trade, pdf_lines)
            all_trade_data.extend(trade_data)

        return [station_name, date, total_quantity, total_amount] + all_trade_data

    except Exception as e:
        st.warning(f"处理PDF {file_name} 出错: {str(e)}")
        return ["未知场站", None, None, None] + [None] * 36

def extract_data_from_excel(file_obj, file_name):
    """适配Streamlit：接收文件对象而非路径"""
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

        # 目标科目
        target_trades = [
            '优先发电交易', '电网企业代理购电交易', '省内电力直接交易',
            '送上海省间绿色电力交易(电能量 )', '送辽宁交易', '送华北交易',
            '送山东交易', '送浙江交易', '省内现货日前交易',
            '省内现货实时交易', '省间现货日前交易', '省间现货日内交易'
        ]

        all_trade_data = []
        for _ in target_trades:
            all_trade_data.extend([None, None, None])

        return [station_name, date, total_quantity, total_amount] + all_trade_data

    except Exception as e:
        st.warning(f"处理Excel {file_name} 出错: {str(e)}")
        return ["未知场站", None, None, None] + [None] * 36

def calculate_summary_row(data_df):
    """计算汇总行（求和电量/电费，平均电价）"""
    sum_cols = [col for col in data_df.columns if any(key in col for key in ['电量', '电费'])]
    avg_cols = [col for col in data_df.columns if '电价' in col]

    summary_row = {'场站名称': '总计', '清分日期': ''}
    for col in sum_cols:
        valid_vals = data_df[col].dropna()
        summary_row[col] = valid_vals.sum() if not valid_vals.empty else 0
    for col in avg_cols:
        valid_vals = data_df[col].dropna()
        summary_row[col] = round(valid_vals.mean(), 3) if not valid_vals.empty else None

    return pd.DataFrame([summary_row])

def to_excel_bytes(df, report_df):
    """将DataFrame转为Excel字节流，用于下载"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='结算数据明细', index=False)
        report_df.to_excel(writer, sheet_name='处理报告', index=False)
    output.seek(0)
    return output

# ---------------------- Streamlit 页面布局与交互 ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取", layout="wide")
    
    # 页面标题
    st.title("📊 黑龙江日清分结算单数据提取工具")
    st.divider()

    # 1. 文件上传区域
    st.subheader("📁 上传文件")
    uploaded_files = st.file_uploader(
        "支持PDF/Excel格式，可批量上传",
        type=['pdf', 'xlsx'],
        accept_multiple_files=True
    )

    # 2. 数据处理逻辑
    if uploaded_files and st.button("🚀 开始处理", type="primary"):
        st.divider()
        st.subheader("⚙️ 处理进度")
        
        all_data = []
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
                data = extract_data_from_pdf(file, file_name)
            else:
                data = extract_data_from_excel(file, file_name)
            
            # 验证数据有效性
            if data[1] is not None and any(isinstance(val, (float, int)) for val in data[2:] if val is not None):
                all_data.append(data)
                processed_files += 1
            
            # 更新进度
            progress_bar.progress((idx + 1) / total_files)

        progress_bar.empty()
        status_text.text("处理完成！")

        # 3. 结果展示与导出
        if all_data:
            st.divider()
            st.subheader("📈 提取结果")
            
            # 构建结果DataFrame
            result_columns = [
                '场站名称', '清分日期', '合计电量(兆瓦时)', '合计电费(元)',
                '优先发电交易_电量', '优先发电交易_电价', '优先发电交易_电费',
                '电网企业代理购电_电量', '电网企业代理购电_电价', '电网企业代理购电_电费',
                '省内电力直接交易_电量', '省内电力直接交易_电价', '省内电力直接交易_电费',
                '送上海省间绿电交易_电量', '送上海省间绿电交易_电价', '送上海省间绿电交易_电费',
                '送辽宁交易_电量', '送辽宁交易_电价', '送辽宁交易_电费',
                '送华北交易_电量', '送华北交易_电价', '送华北交易_电费',
                '送山东交易_电量', '送山东交易_电价', '送山东交易_电费',
                '送浙江交易_电量', '送浙江交易_电价', '送浙江交易_电费',
                '省内现货日前交易_电量', '省内现货日前交易_电价', '省内现货日前交易_电费',
                '省内现货实时交易_电量', '省内现货实时交易_电价', '省内现货实时交易_电费',
                '省间现货日前交易_电量', '省间现货日前交易_电价', '省间现货日前交易_电费',
                '省间现货日内交易_电量', '省间现货日内交易_电价', '省间现货日内交易_电费'
            ]

            result_df = pd.DataFrame(all_data, columns=result_columns)
            num_cols = result_df.columns[2:]
            result_df[num_cols] = result_df[num_cols].apply(pd.to_numeric, errors='coerce')

            # 排序并格式化日期
            result_df['清分日期'] = pd.to_datetime(result_df['清分日期'], errors='coerce')
            result_df = result_df.sort_values(['场站名称', '清分日期']).reset_index(drop=True)
            result_df['清分日期'] = result_df['清分日期'].dt.strftime('%Y-%m-%d').fillna('')

            # 添加汇总行
            summary_row = calculate_summary_row(result_df)
            result_df = pd.concat([result_df, summary_row], ignore_index=True)

            # 生成处理报告
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
                - 涉及 {station_count} 个场站，{valid_rows} 行有效数据
                """
            )
        else:
            st.warning("⚠️ 未提取到有效数据！请检查：")
            st.markdown("""
                1. PDF是否为可复制文本（非扫描件）；
                2. 文件是否为黑龙江日清分格式；
                3. Excel文件格式是否匹配。
            """)

    # 无文件上传时的提示
    elif not uploaded_files and st.button("🚀 开始处理", disabled=True):
        st.warning("请先上传PDF/Excel文件！")

if __name__ == "__main__":
    main()
