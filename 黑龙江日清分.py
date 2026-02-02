import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO
import sys
import os
from openpyxl.styles import PatternFill

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心配置（保留通用化） ----------------------
REDUNDANT_KEYWORDS = [
    "内部使用", "CONFIDENTIAL", "草稿", "现货试结算期间", "日清分单",
    "公司名称", "编号：", "单位：", "清分日期", "合计电量", "合计电费",
    "计量电量", "电能量电费", "科目编码", "审批：", "审核：", "编制：", "加盖电子签章", "dqjs"
]
TRADE_CODE_MAP = {
    "0101010101": "优先发电交易",
    "0101020101": "电网企业代理购电交易", 
    "0101020301": "省内电力直接交易",
    "0101040203": "送上海省间绿色电力交易(电能量)",
    "0101040301": "送辽宁交易",
    "0101040321": "送华北交易", 
    "0101040322": "送山东交易",
    "0101040330": "送浙江交易",
    "0102020101": "省内现货日前交易",
    "0102020301": "省内现货实时交易",
    "0102010101": "省间现货日前交易",
    "0102010201": "省间现货日内交易",
    "0202030001": "中长期合约阻塞费用",
    "0202030002": "省间省内价差费用",
    "0101070101": "现货结算价差调整",
    "0101090101": "辅助服务费用分摊",
    "0101100101": "偏差考核费用",
    "101070101": "现货结算价差调整",
    "101090101": "辅助服务费用分摊",
    "101100101": "偏差考核费用"
}
DATA_RULES = {
    "电量(兆瓦时)": {"min": 0, "max": 5000},
    "电价(元/兆瓦时)": {"min": 0, "max": 2000},
    "电费(元)": {"min": 0, "max": 10000000}
}

# ---------------------- 核心工具函数（精准优化） ----------------------
def remove_redundant_text(text):
    if not text:
        return ""
    cleaned = str(text).strip()
    for keyword in REDUNDANT_KEYWORDS:
        cleaned = cleaned.replace(keyword, "")
    cleaned = re.sub(r'\s+', ' ', cleaned)
    cleaned = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9\.\-\: ]', '', cleaned)
    return cleaned.strip()

def safe_convert_to_numeric(value, data_type=""):
    if value is None or pd.isna(value) or value == '':
        return None
    val_str = remove_redundant_text(value)
    if re.match(r'^\d{10}$', val_str) or val_str in ['-', '.', '', '—', '——']:
        return None
    try:
        cleaned = re.sub(r'[^\d.-]', '', val_str.replace('，', ',').replace('。', '.'))
        if not cleaned or cleaned in ['-', '.']:
            return None
        num = float(cleaned)
        if data_type in DATA_RULES:
            rule = DATA_RULES[data_type]
            if num < rule["min"] or num > rule["max"]:
                return None
        return num
    except (ValueError, TypeError):
        return None

def extract_general_info(pdf_text, file_name):
    """优化1：场站名称优先取“机组”对应的名称"""
    clean_text = remove_redundant_text(pdf_text)
    lines = clean_text.split('\n')
    
    # 1. 公司名称（通用提取）
    company_name = "未知发电公司"
    company_match = re.search(r'公司名称[:：]\s*([^，。\n]+公司)', clean_text)
    if company_match:
        company_name = company_match.group(1).strip()
    else:
        company_match = re.search(r'([^_]+公司|[^_]+发电)', file_name)
        if company_match:
            company_name = company_match.group(1).strip()
    
    # 2. 场站名称：优先从“机组：XXX”提取（用户核心需求）
    station_name = "未知场站"
   机组_match = re.search(r'机组[:：]\s*([^，。\n]+)', clean_text)  # 匹配“机组：晶盛光伏电站”
    if 机组_match:
        station_name = 机组_match.group(1).strip()
    else:
        # 备选：从文本中找场站类型关键词
        for line in lines:
            for type_key in ["风电场", "光伏电站", "储能电站", "电站", "场站"]:
                if type_key in line:
                    match = re.search(r'([^，。\n]+' + type_key + ')', line)
                    if match:
                        station_name = match.group(1).strip()
                        break
            if station_name != "未知场站":
                break
        # 再备选：从文件名提取
        if station_name == "未知场站":
            for type_key in ["风电场", "光伏电站", "储能电站", "电站", "场站"]:
                if type_key in file_name:
                    match = re.search(r'([^_]+' + type_key + ')', file_name)
                    if match:
                        station_name = match.group(1).strip()
                        break
    
    # 3. 清分日期（通用提取）
    date = None
    date_patterns = [
        r'清分日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2}|\d{4}年\d{1,2}月\d{1,2}日)',
        r'结算日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2}|\d{4}年\d{1,2}月\d{1,2}日)',
        r'日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2}|\d{4}年\d{1,2}月\d{1,2}日)',
        r'(\d{4}-\d{1,2}-\d{1,2})'
    ]
    for pattern in date_patterns:
        match = re.search(pattern, clean_text)
        if match:
            date_str = match.group(1).strip()
            if "年" in date_str:
                date_str = date_str.replace("年", "-").replace("月", "-").replace("日", "")
            date = date_str
            break
    if not date:
        date_match = re.search(r'(\d{4}-\d{1,2}-\d{1,2}|\d{8})', file_name)
        if date_match:
            date_str = date_match.group(1)
            if len(date_str) == 8:
                date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
            else:
                date = date_str
    
    # 4. 小计数据（通用提取）
    subtotal_qty = None
    subtotal_fee = None
    subtotal_match = re.search(
        r'小计[:：]?\s*电量[:：]?\s*([\d\.]+)\s*.*?电价[:：]?\s*([\d\.]+)\s*.*?电费[:：]?\s*([\d\.]+)',
        clean_text, re.DOTALL
    )
    if subtotal_match:
        subtotal_qty = safe_convert_to_numeric(subtotal_match.group(1), "电量(兆瓦时)")
        subtotal_fee = safe_convert_to_numeric(subtotal_match.group(3), "电费(元)")
    
    return company_name, station_name, date, subtotal_qty, subtotal_fee

def filter_valid_table_rows(table):
    """优化2：保留含“结算类型”的表头行，便于后续定位"""
    valid_rows = []
    for row in table:
        row_clean = [remove_redundant_text(cell) for cell in row]
        row_str = ''.join(row_clean).replace(" ", "")
        is_empty = all(cell == '' for cell in row_clean)
        # 保留：1. 含“结算类型”的表头行；2. 含编码/科目/数据的行；3. 非空
        has_settlement_type = "结算类型" in row_str
        has_code = any(re.match(r'^\d{9,10}$', cell.replace(" ", "")) for cell in row_clean)
        has_trade = any(len(cell) >= 4 for cell in row_clean if "科目" not in cell)
        has_data = any(safe_convert_to_numeric(cell) is not None for cell in row_clean)
        
        if (has_settlement_type or has_code or has_trade or has_data) and not is_empty:
            valid_rows.append(row_clean)
    return valid_rows

def extract_valid_trade_data(table, company_name, station_name, date):
    """优化3：从“结算类型”表头行的下一行开始提取科目（避免未识别开头）"""
    trade_records = []
    valid_rows = filter_valid_table_rows(table)
    if len(valid_rows) < 2:  # 至少需要“结算类型”表头行 + 1行数据
        return trade_records
    
    # 步骤1：找到“结算类型”所在的表头行（确定数据起始位置）
    settlement_header_idx = -1
    for idx, row in enumerate(valid_rows):
        row_str = ''.join(row).replace(" ", "")
        if "结算类型" in row_str:
            settlement_header_idx = idx
            break
    # 若未找到，默认第0行为表头
    if settlement_header_idx == -1:
        settlement_header_idx = 0
    # 数据从“结算类型”行的下一行开始（用户核心需求）
    data_start_idx = settlement_header_idx + 1
    if data_start_idx >= len(valid_rows):
        return trade_records  # 无数据行，返回空
    
    # 步骤2：定位列（基于“结算类型”表头行）
    header_row = valid_rows[settlement_header_idx]
    cols = {"code": -1, "name": -1, "qty": -1, "price": -1, "fee": -1}
    for col_idx, cell in enumerate(header_row):
        cell_clean = remove_redundant_text(cell).lower()
        if "编码" in cell_clean:
            cols["code"] = col_idx
        elif "结算类型" in cell_clean:  # 明确“结算类型”列为科目名称列
            cols["name"] = col_idx
        elif "电量" in cell_clean and "价" not in cell_clean:
            cols["qty"] = col_idx
        elif "电价" in cell_clean or "单价" in cell_clean:
            cols["price"] = col_idx
        elif "电费" in cell_clean or "金额" in cell_clean:
            cols["fee"] = col_idx
    # 兜底：按“编码→结算类型→电量→电价→电费”固定顺序
    if any(v == -1 for v in cols.values()) and len(header_row) >= 5:
        cols = {"code": 0, "name": 1, "qty": 2, "price": 3, "fee": 4}
    
    # 步骤3：从数据起始行开始解析（跳过表头）
    for row_idx in range(data_start_idx, len(valid_rows)):
        row = valid_rows[row_idx]
        row_str = ''.join([remove_redundant_text(cell) for cell in row]).replace(" ", "")
        
        # 跳过合计行
        if "合计" in row_str and "小计" not in row_str:
            continue
        
        # 提取编码和科目名称（从“结算类型”列取名称）
        trade_code = row[cols["code"]].strip().replace(" ", "") if (cols["code"] < len(row)) else ""
        trade_text = row[cols["name"]].strip() if (cols["name"] < len(row)) else ""
        # 精准匹配科目名称
        trade_name = TRADE_CODE_MAP.get(trade_code, "")
        if not trade_name:
            trade_keywords = {
                "优先发电": "优先发电交易",
                "代理购电": "电网企业代理购电交易",
                "直接交易": "省内电力直接交易",
                "现货日前": "省内现货日前交易",
                "现货实时": "省内现货实时交易",
                "阻塞费用": "中长期合约阻塞费用",
                "价差费用": "省间省内价差费用",
                "现货结算": "现货结算价差调整",
                "辅助服务": "辅助服务费用分摊",
                "偏差考核": "偏差考核费用"
            }
            for key, name in trade_keywords.items():
                if key in trade_text:
                    trade_name = name
                    break
        # 仍未匹配到则跳过（避免未识别科目开头）
        if not trade_name:
            continue
        
        # 提取数据
        quantity = safe_convert_to_numeric(row[cols["qty"]], "电量(兆瓦时)") if (cols["qty"] < len(row)) else None
        price = safe_convert_to_numeric(row[cols["price"]], "电价(元/兆瓦时)") if (cols["price"] < len(row)) else None
        fee = safe_convert_to_numeric(row[cols["fee"]], "电费(元)") if (cols["fee"] < len(row)) else None
        
        # 特殊科目处理
        if trade_name in ["中长期合约阻塞费用", "省间省内价差费用", "辅助服务费用分摊", "偏差考核费用"]:
            quantity = None
            price = None
        
        # 只保留有有效数据的科目
        if quantity is None and fee is None and trade_name not in ["中长期合约阻塞费用", "省间省内价差费用", "辅助服务费用分摊", "偏差考核费用"]:
            continue
        
        trade_records.append({
            "公司名称": company_name,
            "场站名称": station_name,
            "清分日期": date,
            "科目名称": trade_name,
            "原始科目编码": trade_code,
            "原始科目文本": trade_text,
            "是否小计行": "小计" in row_str,
            "电量(兆瓦时)": quantity,
            "电价(元/兆瓦时)": price,
            "电费(元)": fee,
            "提取状态": "成功"
        })
    
    return trade_records

# ---------------------- 通用PDF解析主函数 ----------------------
def parse_pdf_general(file_obj, file_name):
    try:
        file_obj.seek(0)
        file_bytes = BytesIO(file_obj.read())
        file_bytes.seek(0)
        
        all_text = ""
        all_tables = []
        with pdfplumber.open(file_bytes) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                all_text += text + "\n"
                tables = page.extract_tables({
                    "vertical_strategy": "lines_strict",
                    "horizontal_strategy": "lines_strict",
                    "snap_tolerance": 1,
                    "join_tolerance": 1,
                    "edge_min_length": 3
                })
                all_tables.extend(tables)
        
        company_name, station_name, date, subtotal_qty, subtotal_fee = extract_general_info(all_text, file_name)
        trade_records = []
        for table in all_tables:
            if len(table) < 2:
                continue
            table_data = extract_valid_trade_data(table, company_name, station_name, date)
            trade_records.extend(table_data)
        
        # 补充小计行
        if (subtotal_qty is not None or subtotal_fee is not None) and len(trade_records) > 0:
            trade_records.append({
                "公司名称": company_name,
                "场站名称": station_name,
                "清分日期": date,
                "科目名称": "当日小计",
                "原始科目编码": "SUBTOTAL",
                "原始科目文本": "当日小计",
                "是否小计行": True,
                "电量(兆瓦时)": subtotal_qty,
                "电价(元/兆瓦时)": None,
                "电费(元)": subtotal_fee,
                "提取状态": "成功"
            })
        
        # 去重
        unique_records = []
        seen_keys = set()
        for rec in trade_records:
            key = f"{rec['场站名称']}_{rec['科目名称']}_{rec['原始科目编码']}"
            if key not in seen_keys:
                seen_keys.add(key)
                unique_records.append(rec)
        
        return unique_records if len(unique_records) > 0 else [{
            "公司名称": "未知发电公司",
            "场站名称": "未知场站",
            "清分日期": None,
            "科目名称": "解析失败",
            "原始科目编码": "",
            "原始科目文本": "",
            "是否小计行": False,
            "电量(兆瓦时)": None,
            "电价(元/兆瓦时)": None,
            "电费(元)": None,
            "提取状态": "解析错误"
        }]
    
    except Exception as e:
        st.error(f"PDF解析错误: {str(e)}")
        return [{
            "公司名称": "未知发电公司",
            "场站名称": "未知场站",
            "清分日期": None,
            "科目名称": "解析失败",
            "原始科目编码": "",
            "原始科目文本": "",
            "是否小计行": False,
            "电量(兆瓦时)": None,
            "电价(元/兆瓦时)": None,
            "电费(元)": None,
            "提取状态": "解析错误"
        }]

# ---------------------- 通用Streamlit应用 ----------------------
def main():
    st.set_page_config(page_title="通用日清分数据提取工具（最终版）", layout="wide")
    
    st.title("📊 通用现货日清分结算单数据提取工具（精准版）")
    st.markdown("**核心特性：场站名取“机组”字段 | 科目从“结算类型”行开始 | 无未识别科目开头**")
    st.divider()
    
    uploaded_files = st.file_uploader(
        "上传PDF文件（支持多场站批量上传）",
        type=["pdf"],
        accept_multiple_files=True
    )
    
    if uploaded_files and st.button("🚀 开始批量提取", type="primary"):
        st.divider()
        st.subheader("⚙️ 处理进度")
        
        all_results = []
        progress_bar = st.progress(0)
        
        for idx, file in enumerate(uploaded_files):
            st.write(f"正在处理：{file.name}")
            file_results = parse_pdf_general(file, file.name)
            all_results.extend(file_results)
            progress_bar.progress((idx + 1) / len(uploaded_files))
            file.close()
        
        progress_bar.empty()
        
        df = pd.DataFrame(all_results).fillna("")
        col_order = [
            "公司名称", "场站名称", "清分日期", "科目名称", 
            "原始科目编码", "原始科目文本", "是否小计行",
            "电量(兆瓦时)", "电价(元/兆瓦时)", "电费(元)", "提取状态"
        ]
        df = df[col_order]
        
        st.subheader("📈 批量提取结果")
        styled_df = df.style.apply(
            lambda row: ["background-color: #e6f3ff" if row["是否小计行"] else "" for _ in row],
            axis=1
        )
        st.dataframe(styled_df, use_container_width=True)
        
        total_stations = df["场站名称"].nunique()
        total_trades = len(df[df["是否小计行"] == False])
        st.info(f"**统计：** 覆盖场站 {total_stations} 个 | 有效科目 {total_trades} 个 | 首科目均为有效科目")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="多场站日清分数据")
            ws = writer.sheets["多场站日清分数据"]
            light_blue = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
            for row in range(2, len(df) + 2):
                if df.iloc[row-2]["是否小计行"]:
                    for col in range(1, len(col_order) + 1):
                        ws.cell(row=row, column=col).fill = light_blue
        
        output.seek(0)
        st.download_button(
            label="📥 下载多场站Excel",
            data=output,
            file_name=f"多场站日清分数据_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.success("✅ 批量提取完成！场站名精准取自“机组”字段，科目从“结算类型”行开始，无未识别科目开头")
    
    else:
        st.info("👆 请上传任意场站的现货日清分结算单PDF（支持批量）")

if __name__ == "__main__":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    main()
