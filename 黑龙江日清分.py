import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO
import sys
import os
# 新增：导入openpyxl样式模块（修复Excel样式错误）
from openpyxl.styles import PatternFill

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心配置（补充更多匹配项） ----------------------
WATERMARK_KEYWORDS = ["协合能源", "大庆晶盛", "太阳能发电", "内部使用", "CONFIDENTIAL", "草稿", "协合"]
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
    "0101050101": "省内绿色电力交易(电能量)",
    "0101060101": "日融合交易",
    "0101070101": "现货结算价差调整",
    "0101090101": "辅助服务费用分摊",
    "0101100101": "偏差考核费用"
}
TRADE_NAME_KEYWORDS = {
    "优先发电": "优先发电交易",
    "代理购电": "电网企业代理购电交易",
    "直接交易": "省内电力直接交易",
    "送上海": "送上海省间绿色电力交易(电能量)",
    "送辽宁": "送辽宁交易",
    "送华北": "送华北交易",
    "送山东": "送山东交易",
    "送浙江": "送浙江交易",
    "现货日前": "省内现货日前交易",
    "现货实时": "省内现货实时交易",
    "阻塞费用": "中长期合约阻塞费用",
    "价差费用": "省间省内价差费用",
    "辅助服务": "辅助服务费用分摊",
    "偏差考核": "偏差考核费用"
}
SPECIAL_TRADES = ["中长期合约阻塞费用", "省间省内价差费用", "辅助服务费用分摊", "偏差考核费用"]

# ---------------------- 核心工具函数（全修复） ----------------------
def remove_watermark(text):
    if not text:
        return ""
    cleaned_text = text
    for keyword in WATERMARK_KEYWORDS:
        cleaned_text = cleaned_text.replace(keyword, "")
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
    cleaned_text = re.sub(r'[\x00-\x1F\x7F]', '', cleaned_text)
    return cleaned_text.strip()

def safe_convert_to_numeric(value):
    if value is None or pd.isna(value) or value == '':
        return None
    val_str = remove_watermark(str(value)).strip()
    if re.match(r'^\d{9,10}$', val_str) or val_str in ['-', '.', '', '—', '——']:
        return None
    try:
        cleaned = re.sub(r'[^\d.-]', '', val_str.replace('，', ',').replace('。', '.'))
        if not cleaned or cleaned in ['-', '.']:
            return None
        num = float(cleaned)
        return num
    except (ValueError, TypeError):
        return None

def extract_clear_date(pdf_text):
    # 补充更多日期格式
    date_patterns = [
        r'清分日期[:：]\s*(\d{4}[年/-]\d{1,2}[月/-]\d{1,2}[日]?)',
        r'结算日期[:：]\s*(\d{4}[年/-]\d{1,2}[月/-]\d{1,2}[日]?)',
        r'(\d{4}年\d{1,2}月\d{1,2}日)\s*清分',
        r'(\d{4}-\d{1,2}-\d{1,2})\s*现货日清分',
        r'日期[:：]\s*(\d{4}[年/-]\d{1,2}[月/-]\d{1,2}[日]?)',  # 补充“日期：”格式
        r'(\d{4}\.\d{1,2}\.\d{1,2})'  # 补充“2026.01.01”格式
    ]
    for pattern in date_patterns:
        match = re.search(pattern, pdf_text)
        if match:
            date_str = match.group(1)
            date_str = re.sub(r'[年月日]', '-', date_str).rstrip('-')
            parts = date_str.split('-')
            if len(parts) == 3:
                year, month, day = parts
                return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            elif len(date_str.split('.')) == 3:
                year, month, day = date_str.split('.')
                return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
    return None

def extract_base_info(pdf_text):
    clean_text = remove_watermark(pdf_text)
    lines = clean_text.split('\n')
    
    # 修复1：放宽公司/场站名称匹配（从文本中找“场站”“电站”关键词）
    station_name = "未知场站"
    company_name = "未知公司"
    for line in lines:
        if "场站" in line or "电站" in line:
            match = re.search(r'([^，。\n]+[场站|电站])', line)
            if match:
                station_name = match.group(1).strip()
        if "公司" in line:
            match = re.search(r'([^，。\n]+公司)', line)
            if match:
                company_name = match.group(1).strip()
    
    # 修复2：提取日期
    date = extract_clear_date(clean_text)
    
    # 修复3：提取小计/合计
    subtotal_quantity = None
    subtotal_fee = None
    for line in lines:
        line_clean = remove_watermark(line).replace(' ', '').replace(',', '')
        if '小计' in line_clean:
            qty_match = re.search(r'小计电量[:：]([\d\.]+)|电量[:：]([\d\.]+)', line_clean)
            fee_match = re.search(r'小计电费[:：]([\d\.]+)|电费[:：]([\d\.]+)', line_clean)
            if qty_match:
                subtotal_quantity = safe_convert_to_numeric(next(g for g in qty_match.groups() if g))
            if fee_match:
                subtotal_fee = safe_convert_to_numeric(next(g for g in fee_match.groups() if g))
    
    return station_name, company_name, date, subtotal_quantity, subtotal_fee

def locate_table_columns(table_rows):
    # 修复4：补充更多表头关键词（匹配PDF实际表头）
    target_columns = {
        "科目编码": ["科目编码", "编码"],
        "科目名称": ["科目名称", "结算类型", "名称"],
        "交易电量": ["交易电量", "电量", "兆瓦时", "MWh"],
        "结算电价": ["结算电价", "电价", "元/兆瓦时", "元/MWh"],
        "结算电费": ["结算电费", "电费", "金额", "元"]
    }
    final_cols = {k: -1 for k in target_columns.keys()}
    used_cols = set()
    
    for row_idx, row in enumerate(table_rows[:3]):
        for col_idx, cell in enumerate(row):
            cell_clean = remove_watermark(str(cell)).lower().strip()
            for col_name, keywords in target_columns.items():
                if any(key in cell_clean for key in keywords) and col_idx not in used_cols:
                    final_cols[col_name] = col_idx
                    used_cols.add(col_idx)
                    break
    return final_cols

def correct_trade_name(trade_name):
    if not trade_name:
        return "未知科目"
    clean_name = remove_watermark(trade_name).strip()
    for keyword, correct_name in TRADE_NAME_KEYWORDS.items():
        if keyword in clean_name:
            return correct_name
    return clean_name if clean_name else "未知科目"

def extract_trade_data_from_tables(tables, station_name, clear_date):
    trade_records = []
    for table in tables:
        if len(table) < 3:
            continue
        clean_table = []
        for row in table:
            clean_row = [remove_watermark(str(cell)) for cell in row]
            if any(cell.strip() != '' for cell in clean_row):
                clean_table.append(clean_row)
        if len(clean_table) < 3:
            continue
        
        final_cols = locate_table_columns(clean_table)
        code_col = final_cols["科目编码"]
        name_col = final_cols["科目名称"]
        qty_col = final_cols["交易电量"]
        price_col = final_cols["结算电价"]
        fee_col = final_cols["结算电费"]
        
        if (code_col == -1 and name_col == -1) or (qty_col == -1 and fee_col == -1):
            continue
        
        for row_idx, row in enumerate(clean_table):
            if row_idx < 1:  # 仅跳过第1行表头
                continue
            row_clean = [cell.strip() for cell in row]
            if '合计' in ''.join(row_clean) and '小计' not in ''.join(row_clean):
                continue
            
            trade_code = row[code_col].strip() if (code_col != -1 and code_col < len(row)) else ''
            raw_name = row[name_col].strip() if (name_col != -1 and name_col < len(row)) else ''
            trade_name = TRADE_CODE_MAP.get(trade_code, correct_trade_name(raw_name))
            
            # 修复5：确保列索引不越界
            quantity = safe_convert_to_numeric(row[qty_col]) if (qty_col != -1 and qty_col < len(row)) else None
            price = safe_convert_to_numeric(row[price_col]) if (price_col != -1 and price_col < len(row)) else None
            fee = safe_convert_to_numeric(row[fee_col]) if (fee_col != -1 and fee_col < len(row)) else None
            
            if trade_name in SPECIAL_TRADES:
                quantity = None
                price = None
            
            is_subtotal = '小计' in ''.join(row_clean)
            if not is_subtotal and (quantity is None and fee is None):
                continue
            
            trade_records.append({
                "场站名称": station_name,
                "清分日期": clear_date,
                "科目名称": trade_name,
                "是否小计行": is_subtotal,
                "电量(兆瓦时)": quantity if not is_subtotal else quantity,
                "电价(元/兆瓦时)": price if not is_subtotal else None,
                "电费(元)": fee
            })
    return trade_records

def parse_pdf_file(file_obj, file_name):
    try:
        file_obj.seek(0)
        file_bytes = BytesIO(file_obj.read())
        file_bytes.seek(0)
        
        all_text = ""
        all_tables = []
        with pdfplumber.open(file_bytes) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                all_text += remove_watermark(text) + "\n"
                tables = page.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 2,
                    "join_tolerance": 2
                })
                all_tables.extend(tables)
        
        station_name, company_name, clear_date, subtotal_qty, subtotal_fee = extract_base_info(all_text)
        trade_records = extract_trade_data_from_tables(all_tables, station_name, clear_date)
        
        # 补充小计行
        if subtotal_qty or subtotal_fee:
            trade_records.append({
                "场站名称": station_name,
                "清分日期": clear_date,
                "科目名称": "当日小计",
                "是否小计行": True,
                "电量(兆瓦时)": subtotal_qty,
                "电价(元/兆瓦时)": None,
                "电费(元)": subtotal_fee
            })
        
        return trade_records
    except Exception as e:
        st.error(f"PDF解析失败（{file_name}）: {str(e)}")
        return [{
            "场站名称": "未知场站",
            "清分日期": None,
            "科目名称": "解析失败",
            "是否小计行": False,
            "电量(兆瓦时)": None,
            "电价(元/兆瓦时)": None,
            "电费(元)": None
        }]

# ---------------------- Streamlit 应用（修复Excel样式错误） ----------------------
def main():
    st.set_page_config(page_title="日清分数据提取工具（修复版）", layout="wide")
    
    st.title("📊 日清分结算单数据提取工具（稳定版）")
    st.markdown("**已修复：Excel样式错误 | 日期/场站提取 | 数据映射偏差**")
    st.divider()
    
    uploaded_files = st.file_uploader(
        "支持PDF格式（单文件上传）",
        type=['pdf'],
        accept_multiple_files=False
    )
    
    if uploaded_files and st.button("🚀 开始处理", type="primary"):
        st.divider()
        file = uploaded_files
        st.write(f"正在处理：{file.name}")
        trade_records = parse_pdf_file(file, file.name)
        file.close()
        
        result_df = pd.DataFrame(trade_records)
        col_order = ["场站名称", "清分日期", "科目名称", "是否小计行", "电量(兆瓦时)", "电价(元/兆瓦时)", "电费(元)"]
        result_df = result_df[col_order].fillna("")  # 空值显示为空字符串
        
        # 显示结果
        st.subheader("📈 提取结果")
        styled_df = result_df.style.apply(
            lambda row: ['background-color: #f0f8ff' if row["是否小计行"] else '' for _ in row],
            axis=1
        )
        st.dataframe(styled_df, use_container_width=True)
        
        # 统计
        subtotal_count = len(result_df[result_df["是否小计行"]])
        trade_count = len(result_df[~result_df["是否小计行"]])
        st.info(f"**统计：** {trade_count}个科目 + {subtotal_count}个小计行 | 场站：{result_df['场站名称'].iloc[0]} | 日期：{result_df['清分日期'].iloc[0] or '待识别'}")
        
        # 修复6：正确设置Excel样式（用openpyxl）
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name="数据")
            worksheet = writer.sheets["数据"]
            # 定义浅蓝色填充
            light_blue_fill = PatternFill(start_color="F0F8FF", end_color="F0F8FF", fill_type="solid")
            # 遍历行设置样式
            for row_idx in range(2, len(result_df) + 2):
                if result_df.iloc[row_idx - 2]["是否小计行"]:
                    for col_idx in range(1, len(col_order) + 1):
                        worksheet.cell(row=row_idx, column=col_idx).fill = light_blue_fill
        
        output.seek(0)
        st.download_button(
            label="📥 下载Excel",
            data=output,
            file_name=f"日清分数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.success("✅ 处理完成！")
    
    else:
        st.info("👆 请上传PDF文件开始处理")

if __name__ == "__main__":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    main()
