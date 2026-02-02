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

# ---------------------- 核心配置（强化边界限制） ----------------------
REDUNDANT_KEYWORDS = [
    "内部使用", "CONFIDENTIAL", "草稿", "现货试结算期间", "日清分单据",
    "公司名称", "编号：", "单位：", "清分日期", "合计电量", "合计电费",
    "电能量电费", "审批：", "审核：", "编制：", "加盖电子签章",
    "ylxxhfd", "yxxchfd", "依兰县协合风力发电有限公司", "依", "依兰", "协合",
    "县", "风力发电", "有限公司"
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
    "0101020201": "省内绿色电力交易(电能量)",
    "0101040202": "送江苏省间绿色电力交易(电能量)",
    "0101040204": "送浙江省间绿色电力交易(电能量)",
    "0101100001": "日融合交易",
    "101010101": "优先发电交易",
    "101020201": "省内绿色电力交易(电能量)",
    "101040202": "送江苏省间绿色电力交易(电能量)",
    "101040204": "送浙江省间绿色电力交易(电能量)",
    "101100001": "日融合交易",
    "101040330": "送浙江交易",
    "101070101": "现货结算价差调整",
    "101090101": "辅助服务费用分摊",
    "101100101": "偏差考核费用"
}
TRADE_KEYWORDS = {
    "优先发电": "优先发电交易",
    "代理购电": "电网企业代理购电交易",
    "直接交易": "省内电力直接交易",
    "绿色电力": "省内绿色电力交易(电能量)",
    "送江苏": "送江苏省间绿色电力交易(电能量)",
    "送浙江": "送浙江交易",
    "送浙江省间": "送浙江省间绿色电力交易(电能量)",
    "送上海": "送上海省间绿色电力交易(电能量)",
    "送辽宁": "送辽宁交易",
    "送华北": "送华北交易",
    "送山东": "送山东交易",
    "日融合": "日融合交易",
    "现货日前": "省内现货日前交易",
    "现货实时": "省内现货实时交易",
    "阻塞费用": "中长期合约阻塞费用",
    "价差费用": "省间省内价差费用",
    "现货结算": "现货结算价差调整",
    "辅助服务": "辅助服务费用分摊",
    "偏差考核": "偏差考核费用"
}
DATA_RULES = {
    "电量(兆瓦时)": {"min": -1000, "max": 5000},
    "电价(元/兆瓦时)": {"min": 0, "max": 2000},
    "电费(元)": {"min": -10000, "max": 10000000}
}
# 明确排除无关关键词（避免纳入场站名称）
EXCLUDE_KEYWORDS = ["计量电量", "电量", "电价", "电费", "合计", "小计", "编制", "审核"]
STATION_SPLIT_KEYWORDS = ["机组", "机组名称", "双发A风电场", "双发B风电场", "A风电场", "B风电场"]
STATION_CORE_NAMES = ["双发A风电场", "双发B风电场", "晶盛光伏电站"]  # 仅保留纯净核心名称
STATION_TYPE_KEYWORDS = ["风电场", "光伏电站", "储能电站", "电站", "场站"]

# ---------------------- 核心工具函数（精准截断修复） ----------------------
def remove_redundant_text(text):
    if not text:
        return ""
    cleaned = str(text).strip()
    # 保留“机组”关键词，清理其他冗余
    for keyword in REDUNDANT_KEYWORDS:
        if keyword != "机组":
            cleaned = cleaned.replace(keyword, "")
    # 清理单个字符和乱码
    single_watermarks = ["依", "兰", "协", "合", "县", "电", "力", "发", "限"]
    for char in single_watermarks:
        cleaned = cleaned.replace(char, "")
    cleaned = re.sub(r'\s+', ' ', cleaned)
    cleaned = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9\.\-\: ]', '', cleaned)
    return cleaned.strip()

def truncate_station_name(station_name):
    """新增：精准截断场站名称，排除无关数据（核心修复）"""
    if not station_name:
        return "未知场站"
    # 1. 优先匹配核心名称，直接返回（最精准）
    for core_name in STATION_CORE_NAMES:
        if core_name in station_name:
            return core_name
    # 2. 遇到排除关键词则截断（如“双发A风电场 计量电量”→“双发A风电场”）
    for exclude_key in EXCLUDE_KEYWORDS:
        if exclude_key in station_name:
            return station_name.split(exclude_key)[0].strip()
    # 3. 按标点符号截断
    for separator in ["，", "。", "；", "：", "\n", "\t"]:
        if separator in station_name:
            return station_name.split(separator)[0].strip()
    # 4. 保留含场站类型的部分（避免过长）
    for type_key in STATION_TYPE_KEYWORDS:
        if type_key in station_name:
            type_idx = station_name.find(type_key)
            return station_name[:type_idx + len(type_key)].strip()
    return station_name.strip()

def extract_station_from_text(pdf_text):
    """修复：正则增加边界限制，避免包含无关数据"""
    clean_text = remove_redundant_text(pdf_text)
    # 正则优化：匹配“机组：XXX风电场”且排除后续无关内容
    station_patterns = [
        # 匹配“机组：双发A风电场”，且后面不包含排除关键词
        r'机组[:：\s]*([^，。\n计量电量]+风电场)',
        r'机组名称[:：\s]*([^，。\n计量电量]+风电场)',
        # 精准匹配双发A/B风电场
        r'(双发[AB]风电场)',
        r'([^，。\n]+双发[AB]风电场)'
    ]
    for pattern in station_patterns:
        match = re.search(pattern, clean_text)
        if match:
            raw_name = match.group(1).strip()
            return truncate_station_name(raw_name)  # 截断处理
    # 兜底：提取含场站类型的名称
    for type_key in STATION_TYPE_KEYWORDS:
        match = re.search(r'([^，。\n]+' + type_key + ')', clean_text)
        if match:
            raw_name = match.group(1).strip()
            return truncate_station_name(raw_name)
    return "未知场站"

def extract_station_from_filename(file_name):
    """修复：从文件名提取时也做截断"""
    if not file_name:
        return "未知场站"
    name_patterns = [
        r'(双发[AB]风电场)',
        r'([^_]+双发[AB]风电场[^_]+)',
        r'([^_]+风电场)'
    ]
    for pattern in name_patterns:
        match = re.search(pattern, file_name)
        if match:
            raw_name = match.group(1).strip()
            return truncate_station_name(raw_name)
    return "未知场站"

def clean_station_name(station_name):
    """修复：增加截断步骤，确保纯净"""
    if not station_name or station_name == "未知场站":
        return "未知场站"
    # 1. 先清理冗余文本
    cleaned = remove_redundant_text(station_name)
    # 2. 关键步骤：精准截断
    truncated = truncate_station_name(cleaned)
    # 3. 最终验证是否为核心名称
    for core_name in STATION_CORE_NAMES:
        if core_name in truncated:
            return core_name
    return truncated if truncated else "未知场站"

def safe_convert_to_numeric(value, data_type=""):
    if value is None or pd.isna(value) or value == '':
        return None
    val_str = remove_redundant_text(value)
    if re.match(r'^\d{9,10}$', val_str) or val_str in ['-', '.', '', '—', '——']:
        return None
    try:
        cleaned = re.sub(r'[^\d\-\.]', '', val_str.replace('，', ',').replace('。', '.'))
        if not cleaned or cleaned in ['-', '.', '-.' , '-.']:
            return None
        num = float(cleaned)
        if data_type in DATA_RULES:
            rule = DATA_RULES[data_type]
            if num < rule["min"] or num > rule["max"]:
                return None
        return num
    except (ValueError, TypeError):
        return None

def extract_company_info(pdf_text, file_name):
    clean_text = remove_redundant_text(pdf_text)
    company_name = "未知发电公司"
    company_match = re.search(r'公司名称[:：]\s*([^，。\n]+公司)', clean_text)
    if company_match:
        company_name = company_match.group(1).strip()
    else:
        company_match = re.search(r'([^_]+公司|[^_]+发电)', file_name)
        if company_match:
            company_name = company_match.group(1).strip()
    return company_name

def extract_clear_date(pdf_text, file_name):
    clean_text = remove_redundant_text(pdf_text)
    date = None
    date_patterns = [
        r'清分日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2}|\d{4}年\d{1,2}月\d{1,2}日)',
        r'结算日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2}|\d{4}年\d{1,2}月\d{1,2}日)',
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
    return date

def split_double_station_tables(all_tables, pdf_text, file_name):
    """修复：拆分时对场站名称做截断处理"""
    clean_text = remove_redundant_text(pdf_text)
    merged_rows = []
    for table in all_tables:
        if not table:
            continue
        for row in table:
            cleaned_row = [remove_redundant_text(cell) for cell in row]
            if any(cell.strip() != "" for cell in cleaned_row):
                merged_rows.append(cleaned_row)
    
    if not merged_rows:
        text_station = extract_station_from_text(pdf_text)
        if text_station != "未知场站":
            return [(text_station, [])]
        file_station = extract_station_from_filename(file_name)
        return [(file_station, [])]
    
    station_segments = []
    current_segment = []
    # 初始场站从文本提取（已截断）
    current_station = extract_station_from_text(pdf_text)
    
    for row in merged_rows:
        row_str = ''.join(row).replace(" ", "")
        has_station_key = any(keyword in row_str for keyword in STATION_SPLIT_KEYWORDS)
        
        if has_station_key:
            # 保存上一段：场站名称已截断
            if current_segment:
                cleaned_station = clean_station_name(current_station)
                station_segments.append((cleaned_station, current_segment))
            # 提取当前场站并截断
            row_text = ' '.join([remove_redundant_text(cell) for cell in row])
            station_match = re.search(r'机组[:：\s]*([^，。\n]+)', row_text) or re.search(r'(双发[AB]风电场)', row_text)
            if station_match:
                current_station = truncate_station_name(station_match.group(1))
            current_segment = [row]
        else:
            current_segment.append(row)
    
    # 保存最后一段
    if current_segment:
        cleaned_station = clean_station_name(current_station)
        if cleaned_station == "未知场站":
            cleaned_station = extract_station_from_filename(file_name)
        station_segments.append((cleaned_station, current_segment))
    
    # 过滤无效段
    valid_segments = [(station, seg) for station, seg in station_segments if len(seg) >= 2 or station != "未知场站"]
    if not valid_segments:
        fallback_station = extract_station_from_filename(file_name)
        valid_segments = [(fallback_station, merged_rows)]
    
    return valid_segments

def get_trade_name(trade_code, trade_text):
    if trade_code in TRADE_CODE_MAP:
        return TRADE_CODE_MAP[trade_code]
    for key, name in TRADE_KEYWORDS.items():
        if key in trade_text:
            return name
    return "未识别科目"

def parse_single_station_data(station_name, table_segment, company_name, clear_date):
    trade_records = []
    valid_rows = []
    
    for row in table_segment:
        if not row:
            continue
        row_clean = [remove_redundant_text(cell) for cell in row]
        row_str = ''.join(row_clean).replace(" ", "")
        is_empty = all(cell == '' for cell in row_clean)
        is_header = any(keyword in row_str for keyword in ["科目编码", "结算类型", "电量", "电价", "电费"])
        
        has_code = any(re.match(r'^\d{9,10}$', cell.replace(" ", "")) for cell in row_clean)
        has_trade_key = any(key in row_str for key in TRADE_KEYWORDS.keys())
        has_valid_data = any(safe_convert_to_numeric(cell) is not None for cell in row_clean if cell not in ['', '-'])
        
        if (has_code or has_trade_key or has_valid_data) and not is_empty and not is_header:
            valid_rows.append(row_clean)
    
    if len(valid_rows) < 2:
        return trade_records
    
    cols = {"code": -1, "name": -1, "qty": -1, "price": -1, "fee": -1}
    header_idx = -1
    for idx, row in enumerate(valid_rows[:3]):
        row_str = ''.join(row).replace(" ", "")
        if "结算类型" in row_str:
            header_idx = idx
            break
    if header_idx == -1:
        header_idx = 0
    header_row = valid_rows[header_idx]
    
    for col_idx, cell in enumerate(header_row):
        cell_clean = remove_redundant_text(cell).lower()
        if "编码" in cell_clean:
            cols["code"] = col_idx
        elif "结算类型" in cell_clean:
            cols["name"] = col_idx
        elif "电量" in cell_clean and "价" not in cell_clean:
            cols["qty"] = col_idx
        elif "电价" in cell_clean or "单价" in cell_clean:
            cols["price"] = col_idx
        elif "电费" in cell_clean or "金额" in cell_clean:
            cols["fee"] = col_idx
    
    if any(v == -1 for v in cols.values()) and len(header_row) >= 5:
        cols = {"code": 0, "name": 1, "qty": 2, "price": 3, "fee": 4}
    
    data_start_idx = header_idx + 1
    for row_idx in range(data_start_idx, len(valid_rows)):
        row = valid_rows[row_idx]
        row_str = ''.join(row).replace(" ", "")
        is_subtotal = "小计" in row_str
        
        if "合计" in row_str and not is_subtotal:
            continue
        
        trade_code = row[cols["code"]].strip().replace(" ", "") if (cols["code"] < len(row)) else ""
        trade_text = row[cols["name"]].strip() if (cols["name"] < len(row)) else ""
        trade_name = get_trade_name(trade_code, trade_text)
        
        if trade_name == "未识别科目" and not is_subtotal:
            continue
        
        if is_subtotal:
            subtotal_qty = None
            subtotal_fee = None
            nums = [safe_convert_to_numeric(cell, "电量(兆瓦时)") for cell in row if safe_convert_to_numeric(cell) is not None]
            fee_nums = [safe_convert_to_numeric(cell, "电费(元)") for cell in row if safe_convert_to_numeric(cell) is not None]
            if nums:
                subtotal_qty = nums[0]
            if fee_nums:
                subtotal_fee = fee_nums[-1]
            trade_records.append({
                "公司名称": company_name,
                "场站名称": station_name,
                "清分日期": clear_date,
                "科目名称": "当日小计",
                "原始科目编码": trade_code,
                "原始科目文本": trade_text,
                "电量(兆瓦时)": subtotal_qty,
                "电价(元/兆瓦时)": None,
                "电费(元)": subtotal_fee
            })
            continue
        
        quantity = safe_convert_to_numeric(row[cols["qty"]], "电量(兆瓦时)") if (cols["qty"] < len(row)) else None
        price = safe_convert_to_numeric(row[cols["price"]], "电价(元/兆瓦时)") if (cols["price"] < len(row)) else None
        fee = safe_convert_to_numeric(row[cols["fee"]], "电费(元)") if (cols["fee"] < len(row)) else None
        
        if "阻塞费用" in trade_name or "价差费用" in trade_name or "辅助服务" in trade_name or "偏差考核" in trade_name:
            quantity = None
            price = None
        
        if quantity is None and fee is None:
            continue
        
        trade_records.append({
            "公司名称": company_name,
            "场站名称": station_name,
            "清分日期": clear_date,
            "科目名称": trade_name,
            "原始科目编码": trade_code,
            "原始科目文本": trade_text,
            "电量(兆瓦时)": quantity,
            "电价(元/兆瓦时)": price,
            "电费(元)": fee
        })
    
    return trade_records

# ---------------------- 主解析函数 ----------------------
def parse_pdf_final(file_obj, file_name):
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
        
        company_name = extract_company_info(all_text, file_name)
        clear_date = extract_clear_date(all_text, file_name)
        station_segments = split_double_station_tables(all_tables, all_text, file_name)
        
        all_records = []
        for station_name, table_segment in station_segments:
            # 最终确保场站名称已截断
            final_station = clean_station_name(station_name)
            station_data = parse_single_station_data(final_station, table_segment, company_name, clear_date)
            all_records.extend(station_data)
        
        # 去重
        unique_records = []
        seen_keys = set()
        for rec in all_records:
            key = f"{rec['场站名称']}_{rec['科目名称']}_{rec['原始科目编码']}"
            if key not in seen_keys:
                seen_keys.add(key)
                unique_records.append(rec)
        
        if not unique_records:
            fallback_station = extract_station_from_filename(file_name)
            unique_records.append({
                "公司名称": company_name,
                "场站名称": fallback_station,
                "清分日期": clear_date,
                "科目名称": "无有效数据",
                "原始科目编码": "",
                "原始科目文本": "",
                "电量(兆瓦时)": None,
                "电价(元/兆瓦时)": None,
                "电费(元)": None
            })
        
        return unique_records
    
    except Exception as e:
        st.error(f"PDF解析错误: {str(e)}")
        fallback_station = extract_station_from_filename(file_name)
        return [{
            "公司名称": "未知发电公司",
            "场站名称": fallback_station,
            "清分日期": None,
            "科目名称": "解析失败",
            "原始科目编码": "",
            "原始科目文本": "",
            "电量(兆瓦时)": None,
            "电价(元/兆瓦时)": None,
            "电费(元)": None
        }]

# ---------------------- Streamlit应用 ----------------------
def main():
    st.set_page_config(page_title="通用日清分数据提取工具（场站精准版）", layout="wide")
    
    st.title("📊 通用现货日清分结算单数据提取工具（双场站精准版）")
    st.markdown("**核心修复：场站名称精准截断 | 排除计量电量等无关数据 | 双发A/B纯净显示**")
    st.divider()
    
    uploaded_files = st.file_uploader(
        "上传PDF文件（支持双场站/多页面）",
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
            file_results = parse_pdf_final(file, file.name)
            all_results.extend(file_results)
            progress_bar.progress((idx + 1) / len(uploaded_files))
            file.close()
        
        progress_bar.empty()
        
        df = pd.DataFrame(all_results).fillna("")
        display_cols = [
            "公司名称", "场站名称", "清分日期", "科目名称", 
            "原始科目编码", "原始科目文本", "电量(兆瓦时)", 
            "电价(元/兆瓦时)", "电费(元)"
        ]
        df_display = df[[col for col in display_cols if col in df.columns]]
        
        st.subheader("📈 批量提取结果（场站名称纯净）")
        def highlight_rows(row):
            if row["科目名称"] == "当日小计":
                return ["background-color: #e6f3ff"] * len(row)
            elif row["场站名称"] == "双发A风电场":
                return ["background-color: #f0fff4"] * len(row)
            elif row["场站名称"] == "双发B风电场":
                return ["background-color: #fff8f0"] * len(row)
            else:
                return [""] * len(row)
        styled_df = df_display.style.apply(highlight_rows, axis=1)
        st.dataframe(styled_df, use_container_width=True)
        
        total_stations = df["场站名称"].nunique()
        total_trades = len(df[(df["科目名称"] != "当日小计") & (df["科目名称"] != "无有效数据") & (df["科目名称"] != "解析失败")])
        subtotal_count = len(df[df["科目名称"] == "当日小计"])
        st.info(f"**统计：** 覆盖场站 {total_stations} 个 | 有效科目 {total_trades} 个 | 小计行 {subtotal_count} 个")
        
        # 下载Excel
        download_cols = [
            "公司名称", "场站名称", "清分日期", "科目名称", 
            "电量(兆瓦时)", "电价(元/兆瓦时)", "电费(元)"
        ]
        df_download = df[[col for col in download_cols if col in df.columns]]
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_download.to_excel(writer, index=False, sheet_name="多场站日清分数据")
            ws = writer.sheets["多场站日清分数据"]
            light_blue = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
            for row in range(2, len(df_download) + 2):
                if df_download.iloc[row-2]["科目名称"] == "当日小计":
                    for col in range(1, len(df_download.columns) + 1):
                        ws.cell(row=row, column=col).fill = light_blue
        
        output.seek(0)
        st.download_button(
            label="📥 下载Excel（不含原始编码/文本）",
            data=output,
            file_name=f"多场站日清分数据_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.success("✅ 提取完成！场站名称已精准截断，仅保留“双发A风电场”“双发B风电场”，无无关数据")
    
    else:
        st.info("👆 请上传双场站（如双发A/B风电场）的现货日清分结算单PDF")

if __name__ == "__main__":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    main()
