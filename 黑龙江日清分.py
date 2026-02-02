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

# ---------------------- 核心配置（关键扩展+放宽规则） ----------------------
# 1. 冗余文本关键词（补充更多PDF非表格文本）
REDUNDANT_KEYWORDS = [
    "协合能源", "大庆晶盛", "太阳能发电", "内部使用", "CONFIDENTIAL", "草稿",
    "现货试结算期间", "日清分单", "公司名称", "编号：", "单位：", "清分日期",
    "合计电量", "合计电费", "机组", "计量电量", "电能量电费", "科目编码", "结算类型",
    "审批：", "审核：", "编制：", "加盖电子签章", "dqjs2627800", "2026年1月",
    "现货结算价差调整", "辅助服务费用分摊", "偏差考核费用"
]
# 2. 扩展科目编码-名称映射（覆盖所有可能科目）
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
    "0101070101": "现货结算价差调整",  # 新增遗漏科目
    "0101090101": "辅助服务费用分摊",  # 新增遗漏科目
    "0101100101": "偏差考核费用",      # 新增遗漏科目
    "101070101": "现货结算价差调整",   # 兼容9位编码（省略前导0）
    "101090101": "辅助服务费用分摊",
    "101100101": "偏差考核费用"
}
# 3. 放宽数据合理性规则（避免误过滤）
DATA_RULES = {
    "电量(兆瓦时)": {"min": 0, "max": 2000},  # 上限从1000→2000
    "电价(元/兆瓦时)": {"min": 0, "max": 1000}, # 上限从500→1000
    "电费(元)": {"min": 0, "max": 5000000}    # 上限从100万→500万
}
# 4. 允许保留的未识别科目关键词
ALLOWED_UNKNOWN_TRADES = ["现货结算价差调整", "辅助服务费用分摊", "偏差考核费用"]

# ---------------------- 核心工具函数（优化提取范围） ----------------------
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
    # 兼容9位/10位编码（不过滤9位数字，避免误判电量）
    if re.match(r'^\d{10}$', val_str) or val_str in ['-', '.', '', '—', '——']:
        return None
    
    try:
        cleaned = re.sub(r'[^\d.-]', '', val_str.replace('，', ',').replace('。', '.'))
        if not cleaned or cleaned in ['-', '.']:
            return None
        num = float(cleaned)
        
        # 放宽校验：仅过滤明显异常值（如负数、超极大值）
        if data_type in DATA_RULES:
            rule = DATA_RULES[data_type]
            if num < rule["min"] or num > rule["max"]:
                return None
        return num
    except (ValueError, TypeError):
        return None

def extract_fixed_info(pdf_text):
    # 1. 公司/场站名称
    company_match = re.search(r'公司名称[:：]\s*([^，。\n]+有限公司)', pdf_text)
    company_name = company_match.group(1).strip() if company_match else "大庆晶盛太阳能发电有限公司"
    station_name = f"{company_name}（晶盛光伏电站）"
    
    # 2. 清分日期（兼容“2026年01月01日”格式）
    date_match = re.search(r'清分日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2}|\d{4}年\d{1,2}月\d{1,2}日)', pdf_text)
    clear_date = ""
    if date_match:
        date_str = date_match.group(1).strip()
        if "年" in date_str:
            date_str = date_str.replace("年", "-").replace("月", "-").replace("日", "")
        clear_date = date_str
    clear_date = clear_date if clear_date else "2026-01-01"
    
    # 3. 小计数据（兼容“小计：电量 445.245 电费 35862.18”格式）
    subtotal_match = re.search(r'小计[:：]?\s*电量[:：]?\s*([\d\.]+)\s*电价[:：]?\s*([\d\.]+)\s*电费[:：]?\s*([\d\.]+)', pdf_text)
    subtotal_qty = safe_convert_to_numeric(subtotal_match.group(1), "电量(兆瓦时)") if subtotal_match else None
    subtotal_fee = safe_convert_to_numeric(subtotal_match.group(3), "电费(元)") if subtotal_match else None
    
    return station_name, clear_date, subtotal_qty, subtotal_fee

def filter_valid_table_rows(table):
    """放宽行过滤：保留未识别但关键的科目行"""
    valid_rows = []
    for row in table:
        row_clean = [remove_redundant_text(cell) for cell in row]
        row_str = ''.join(row_clean).replace(" ", "")
        
        # 保留条件：
        # 1. 含9/10位科目编码；2. 含关键科目关键词；3. 有有效数据（非空且非表头）
        has_code = any(re.match(r'^\d{9,10}$', cell.replace(" ", "")) for cell in row_clean)
        has_key_trade = any(trade in row_str for trade in ALLOWED_UNKNOWN_TRADES)
        has_data = any(safe_convert_to_numeric(cell) is not None for cell in row_clean if "元" not in cell)
        is_empty = all(cell == '' for cell in row_clean)
        is_header = any(keyword in row_str for keyword in ["科目编码", "结算类型", "电量", "电价", "电费", "合计"])
        
        if ((has_code or has_key_trade or has_data) and not is_empty and not is_header):
            valid_rows.append(row_clean)
    return valid_rows

def get_trade_name(trade_code, trade_text):
    """优化科目名称匹配：保留未识别但关键的科目"""
    if trade_code in TRADE_CODE_MAP:
        return TRADE_CODE_MAP[trade_code]
    # 匹配未识别但允许保留的科目
    for trade in ALLOWED_UNKNOWN_TRADES:
        if trade in trade_text:
            return trade
    return "未识别科目"

def extract_valid_trade_data(table, station_name, clear_date):
    """提取所有有效数据，包括未识别但关键的科目"""
    trade_records = []
    valid_rows = filter_valid_table_rows(table)
    if len(valid_rows) == 0:
        return trade_records
    
    # 定位列（兼容表头行位置偏移）
    cols = {"code": -1, "name": -1, "qty": -1, "price": -1, "fee": -1}
    # 检查前2行表头（避免表头行偏移导致列定位失败）
    for header_row in valid_rows[:2]:
        for col_idx, cell in enumerate(header_row):
            cell_clean = remove_redundant_text(cell).lower()
            if "编码" in cell_clean:
                cols["code"] = col_idx
            elif "类型" in cell_clean or "名称" in cell_clean:
                cols["name"] = col_idx
            elif "电量" in cell_clean and "价" not in cell_clean:
                cols["qty"] = col_idx
            elif "电价" in cell_clean or "单价" in cell_clean:
                cols["price"] = col_idx
            elif "电费" in cell_clean or "金额" in cell_clean:
                cols["fee"] = col_idx
        if all(v != -1 for v in cols.values()):
            break
    # 兜底：按固定顺序赋值（编码→类型→电量→电价→电费）
    if any(v == -1 for v in cols.values()) and len(valid_rows[0]) >= 5:
        cols = {"code": 0, "name": 1, "qty": 2, "price": 3, "fee": 4}
    
    # 解析所有有效行（不跳过未识别科目）
    for row_idx, row in enumerate(valid_rows):
        # 跳过表头行（前2行）
        if row_idx < 2 and ("编码" in ''.join(row) or "类型" in ''.join(row)):
            continue
        
        # 提取编码和名称
        trade_code = row[cols["code"]].strip().replace(" ", "") if (cols["code"] < len(row)) else ""
        trade_text = row[cols["name"]].strip() if (cols["name"] < len(row)) else ""
        trade_name = get_trade_name(trade_code, trade_text)
        
        # 提取数据（允许部分字段为空，如特殊科目无电量）
        quantity = safe_convert_to_numeric(row[cols["qty"]], "电量(兆瓦时)") if (cols["qty"] < len(row)) else None
        price = safe_convert_to_numeric(row[cols["price"]], "电价(元/兆瓦时)") if (cols["price"] < len(row)) else None
        fee = safe_convert_to_numeric(row[cols["fee"]], "电费(元)") if (cols["fee"] < len(row)) else None
        
        # 特殊科目处理（无电量/电价）
        if trade_name in ["中长期合约阻塞费用", "省间省内价差费用", "辅助服务费用分摊", "偏差考核费用"]:
            quantity = None
            price = None
        
        # 保留所有非空行（包括未识别科目）
        trade_records.append({
            "场站名称": station_name,
            "清分日期": clear_date,
            "科目名称": trade_name,
            "原始科目编码": trade_code,
            "原始科目文本": trade_text,
            "是否小计行": False,
            "电量(兆瓦时)": quantity,
            "电价(元/兆瓦时)": price,
            "电费(元)": fee,
            "提取状态": "成功" if (quantity is not None or fee is not None or trade_name in ALLOWED_UNKNOWN_TRADES) else "无有效数据"
        })
    
    return trade_records

# ---------------------- PDF解析主函数 ----------------------
def parse_pdf_final(file_obj):
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
                # 增强表格提取：保留更多可能的表格结构
                tables = page.extract_tables({
                    "vertical_strategy": "lines_strict",  # 严格按线条提取，减少文本干扰
                    "horizontal_strategy": "lines_strict",
                    "snap_tolerance": 0.5,
                    "join_tolerance": 0.5,
                    "edge_min_length": 3
                })
                all_tables.extend(tables)
        
        # 提取基础信息
        station_name, clear_date, subtotal_qty, subtotal_fee = extract_fixed_info(all_text)
        
        # 提取科目数据（合并所有表格）
        trade_records = []
        for table in all_tables:
            if len(table) < 2:  # 至少表头+1行数据
                continue
            table_data = extract_valid_trade_data(table, station_name, clear_date)
            trade_records.extend(table_data)
        
        # 补充小计行
        if subtotal_qty is not None or subtotal_fee is not None:
            trade_records.append({
                "场站名称": station_name,
                "清分日期": clear_date,
                "科目名称": "当日小计",
                "原始科目编码": "SUBTOTAL",
                "原始科目文本": "当日小计",
                "是否小计行": True,
                "电量(兆瓦时)": subtotal_qty,
                "电价(元/兆瓦时)": None,
                "电费(元)": subtotal_fee,
                "提取状态": "成功"
            })
        
        # 去重（按科目名称+原始编码）
        unique_records = []
        seen_keys = set()
        for rec in trade_records:
            key = f"{rec['科目名称']}_{rec['原始科目编码']}"
            if key not in seen_keys:
                seen_keys.add(key)
                unique_records.append(rec)
        
        return unique_records
    
    except Exception as e:
        st.error(f"PDF解析错误: {str(e)}")
        return [{
            "场站名称": "大庆晶盛太阳能发电有限公司（晶盛光伏电站）",
            "清分日期": "2026-01-01",
            "科目名称": "解析失败",
            "原始科目编码": "",
            "原始科目文本": "",
            "是否小计行": False,
            "电量(兆瓦时)": None,
            "电价(元/兆瓦时)": None,
            "电费(元)": None,
            "提取状态": "解析错误"
        }]

# ---------------------- Streamlit 应用 ----------------------
def main():
    st.set_page_config(page_title="日清分数据提取工具（全量提取版）", layout="wide")
    
    st.title("📊 日清分结算单数据提取工具（全量版）")
    st.markdown("**已优化：全科目提取 | 放宽过滤规则 | 保留未识别关键科目**")
    st.divider()
    
    uploaded_file = st.file_uploader("上传PDF文件（大庆晶盛光伏电站日清分单）", type=["pdf"], accept_multiple_files=False)
    
    if uploaded_file and st.button("🚀 开始全量提取", type="primary"):
        st.divider()
        st.write(f"正在处理：{uploaded_file.name}")
        
        # 解析PDF
        trade_data = parse_pdf_final(uploaded_file)
        uploaded_file.close()
        
        # 转换为DataFrame（保留原始信息便于核对）
        df = pd.DataFrame(trade_data).fillna("")
        col_order = [
            "场站名称", "清分日期", "科目名称", "原始科目编码", "原始科目文本",
            "是否小计行", "电量(兆瓦时)", "电价(元/兆瓦时)", "电费(元)", "提取状态"
        ]
        df = df[col_order]
        
        # 显示结果（高亮小计行和未识别科目）
        st.subheader("📈 全量提取结果")
        def highlight_rows(row):
            if row["是否小计行"]:
                return ["background-color: #e6f3ff"] * len(row)
            elif row["科目名称"] == "未识别科目" and row["提取状态"] == "成功":
                return ["background-color: #fff2e6"] * len(row)
            else:
                return [""] * len(row)
        styled_df = df.style.apply(highlight_rows, axis=1)
        st.dataframe(styled_df, use_container_width=True)
        
        # 统计信息（显示提取详情）
        total_trades = len(df[df["是否小计行"] == False])
        success_count = len(df[df["提取状态"] == "成功"])
        unknown_count = len(df[df["科目名称"] == "未识别科目"])
        st.info(f"**统计：** 总科目 {total_trades} 个 | 成功提取 {success_count} 个 | 未识别科目 {unknown_count} 个 | 清分日期：{df['清分日期'].iloc[0]}")
        
        # 下载Excel（包含原始编码和提取状态）
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="全量日清分数据")
            # 高亮特殊行
            ws = writer.sheets["全量日清分数据"]
            light_blue = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
            light_orange = PatternFill(start_color="FFF2E6", end_color="FFF2E6", fill_type="solid")
            for row in range(2, len(df) + 2):
                row_data = df.iloc[row-2]
                if row_data["是否小计行"]:
                    for col in range(1, len(col_order) + 1):
                        ws.cell(row=row, column=col).fill = light_blue
                elif row_data["科目名称"] == "未识别科目" and row_data["提取状态"] == "成功":
                    for col in range(1, len(col_order) + 1):
                        ws.cell(row=row, column=col).fill = light_orange
        
        output.seek(0)
        st.download_button(
            label="📥 下载全量Excel（含原始信息）",
            data=output,
            file_name=f"大庆晶盛光伏_全量日清分数据_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.success("✅ 全量提取完成！未识别科目已标记，可根据原始编码/文本补充映射")
    
    else:
        st.info("👆 请上传大庆晶盛光伏电站的现货日清分结算单PDF")

if __name__ == "__main__":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    main()
