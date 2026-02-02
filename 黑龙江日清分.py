import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO
import sys
import os

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心配置（优化科目映射） ----------------------
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

# 特殊科目（仅含费用，无电量/电价）
SPECIAL_TRADES = ["中长期合约阻塞费用", "省间省内价差费用", "辅助服务费用分摊", "偏差考核费用"]

# ---------------------- 核心工具函数（重构提取逻辑） ----------------------
def safe_convert_to_numeric(value):
    """安全转换为数值 - 增强版"""
    if value is None or pd.isna(value) or value == '':
        return None
    
    val_str = str(value).strip().replace('\xa0', ' ')
    if re.match(r'^\d{9,10}$', val_str):  # 排除科目编码
        return None
    if val_str in ['-', '.', '', '—', '——', ' ', '\t', '\n']:
        return None
    
    try:
        cleaned = re.sub(r'[^\d.-]', '', val_str.replace('，', ',').replace('。', '.'))
        return float(cleaned) if cleaned and cleaned not in ['-', '.', ''] else None
    except (ValueError, TypeError):
        return None

def extract_base_info(pdf_text):
    """提取公司名称、清分日期、合计数据"""
    pdf_text = pdf_text.replace('\xa0', ' ').replace('\r', '\n').strip()
    lines = pdf_text.split('\n')
    
    # 提取公司名称
    company_name = "未知公司"
    for line in lines:
        if "公司名称" in line:
            match = re.search(r'公司名称[:：]\s*(.+?有限公司)', line)
            if match:
                company_name = match.group(1).strip()
                break
    
    # 提取清分日期
    date = None
    date_patterns = [r'清分日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2})', r'(\d{4}年\d{1,2}月\d{1,2}日)']
    for line in lines:
        for pattern in date_patterns:
            match = re.search(pattern, line)
            if match:
                date_str = match.group(1).replace('年', '-').replace('月', '-').replace('日', '')
                parts = date_str.split('-')
                if len(parts) == 3:
                    date = f"{parts[0]}-{parts[1].zfill(2)}-{parts[2].zfill(2)}"
                break
        if date:
            break
    
    # 提取合计电量、合计电费
    total_quantity, total_amount = None, None
    for line in lines:
        line_clean = line.replace(' ', '').replace(',', '').replace('，', '')
        qty_match = re.search(r'合计电量[:：]([\d\.]+)|总电量[:：]([\d\.]+)', line_clean)
        if qty_match:
            qty_val = next((g for g in qty_match.groups() if g), None)
            total_quantity = safe_convert_to_numeric(qty_val)
        
        fee_match = re.search(r'合计电费[:：]([\d\.]+)|总电费[:：]([\d\.]+)|合计金额[:：]([\d\.]+)', line_clean)
        if fee_match:
            fee_val = next((g for g in fee_match.groups() if g), None)
            total_amount = safe_convert_to_numeric(fee_val)
    
    return company_name, date, total_quantity, total_amount

def extract_trade_data_from_tables(tables):
    """按PDF格式提取科目数据（科目为行，电量/电价/电费为列）"""
    trade_records = []
    
    for table in tables:
        if len(table) < 3:  # 至少需要表头+数据行
            continue
        
        # 定位核心列索引（科目编码、结算类型、电量、电价、电费）
        code_col = -1
        name_col = -1
        qty_col = -1
        price_col = -1
        fee_col = -1
        
        # 遍历表头行（前3行）找列索引
        for i in range(min(3, len(table))):
            row = table[i]
            for j, cell in enumerate(row):
                cell_clean = str(cell).strip().lower().replace('\xa0', '')
                if any(key in cell_clean for key in ["科目编码", "编码"]):
                    code_col = j
                elif any(key in cell_clean for key in ["结算类型", "科目名称", "名称"]):
                    name_col = j
                elif any(key in cell_clean for key in ["电量", "兆瓦时"]):
                    qty_col = j
                elif any(key in cell_clean for key in ["电价", "单价"]):
                    price_col = j
                elif any(key in cell_clean for key in ["电费", "金额", "元"]):
                    fee_col = j
        
        # 必须包含核心列才继续
        if code_col == -1 or name_col == -1 or (qty_col == -1 and fee_col == -1):
            continue
        
        # 解析数据行（跳过表头和合计行）
        for i in range(len(table)):
            row = table[i]
            row_clean = [str(cell).strip().replace('\xa0', '') for cell in row]
            
            # 跳过空行和合计行
            if all(cell == '' for cell in row_clean) or any(key in ''.join(row_clean) for key in ["合计", "总计", "小计"]):
                continue
            
            # 提取科目编码和名称
            trade_code = row[code_col].strip() if code_col < len(row) else ''
            trade_name = row[name_col].strip() if name_col < len(row) else ''
            
            # 用编码映射标准名称
            if trade_code in TRADE_CODE_MAP:
                trade_name = TRADE_CODE_MAP[trade_code]
            elif not trade_name or trade_name in ['', '-']:
                continue
            
            # 提取电量、电价、电费
            quantity = safe_convert_to_numeric(row[qty_col]) if (qty_col < len(row) and qty_col != -1) else None
            price = safe_convert_to_numeric(row[price_col]) if (price_col < len(row) and price_col != -1) else None
            fee = safe_convert_to_numeric(row[fee_col]) if (fee_col < len(row) and fee_col != -1) else None
            
            # 特殊科目处理（仅保留费用）
            if trade_name in SPECIAL_TRADES:
                quantity = None
                price = None
            
            # 只保留有有效数据的记录
            if quantity is not None or price is not None or fee is not None or trade_name in SPECIAL_TRADES:
                trade_records.append({
                    "科目名称": trade_name,
                    "电量(兆瓦时)": quantity,
                    "电价(元/兆瓦时)": price,
                    "电费(元)": fee
                })
    
    return trade_records

def parse_pdf_file(file_obj):
    """解析PDF文件，返回结构化数据"""
    try:
        file_obj.seek(0)
        file_bytes = BytesIO(file_obj.read())
        file_bytes.seek(0)
        
        # 提取文本和表格
        all_text = ""
        all_tables = []
        with pdfplumber.open(file_bytes) as pdf:
            for page in pdf.pages:
                # 提取文本
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
                # 提取表格
                tables = page.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 3,
                    "join_tolerance": 3
                })
                for table in tables:
                    # 清理表格空行空列
                    cleaned_table = []
                    for row in table:
                        cleaned_row = [str(cell).strip() if cell is not None else '' for cell in row]
                        if any(cell != '' for cell in cleaned_row):
                            cleaned_table.append(cleaned_row)
                    if cleaned_table:
                        all_tables.append(cleaned_table)
        
        # 提取基础信息
        company_name, date, total_qty, total_fee = extract_base_info(all_text)
        
        # 提取科目交易数据
        trade_records = extract_trade_data_from_tables(all_tables)
        
        # 补充合计行
        if total_qty is not None or total_fee is not None:
            trade_records.append({
                "科目名称": "合计",
                "电量(兆瓦时)": total_qty,
                "电价(元/兆瓦时)": None,
                "电费(元)": total_fee
            })
        
        # 补充场站名称（默认公司+未知场站，可根据实际PDF调整）
        station_name = f"{company_name}（未知场站）"
        if "晶盛光伏电站" in all_text:
            station_name = f"{company_name}（晶盛光伏电站）"
        
        # 给每条记录添加场站和日期信息
        for record in trade_records:
            record["场站名称"] = station_name
            record["清分日期"] = date
        
        return trade_records
    
    except Exception as e:
        st.error(f"PDF解析失败: {str(e)}")
        return [{
            "场站名称": "未知场站",
            "清分日期": None,
            "科目名称": "解析失败",
            "电量(兆瓦时)": None,
            "电价(元/兆瓦时)": None,
            "电费(元)": None
        }]

# ---------------------- Streamlit 应用（适配新格式） ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取工具（按PDF格式）", layout="wide")
    
    st.title("📊 日清分结算单数据提取工具（科目行式）")
    st.markdown("**提取格式：科目为行，电量/电价/电费为列 | 精准匹配PDF原生结构**")
    st.divider()
    
    # 上传文件
    uploaded_files = st.file_uploader(
        "支持PDF格式，可批量上传",
        type=['pdf'],
        accept_multiple_files=True
    )
    
    if uploaded_files and st.button("🚀 开始处理", type="primary"):
        st.divider()
        st.subheader("⚙️ 处理进度")
        
        all_results = []
        progress_bar = st.progress(0)
        
        for idx, file in enumerate(uploaded_files):
            st.write(f"正在处理：{file.name}")
            records = parse_pdf_file(file)
            all_results.extend(records)
            progress_bar.progress((idx + 1) / len(uploaded_files))
            file.close()
        
        progress_bar.empty()
        
        # 转换为DataFrame并调整列顺序
        result_df = pd.DataFrame(all_results)
        col_order = ["场站名称", "清分日期", "科目名称", "电量(兆瓦时)", "电价(元/兆瓦时)", "电费(元)"]
        result_df = result_df[col_order]
        
        # 显示结果
        st.subheader("📈 提取结果")
        st.dataframe(result_df, use_container_width=True)
        
        # 统计信息
        st.info(f"**统计信息：** 共提取 {len(result_df)} 条科目记录，涉及 {result_df['场站名称'].nunique()} 个场站")
        
        # 数据完整性
        data_cols = ["电量(兆瓦时)", "电价(元/兆瓦时)", "电费(元)"]
        filled_cells = result_df[data_cols].notna().sum().sum()
        total_cells = len(result_df) * len(data_cols)
        st.info(f"**数据完整性：** 有值单元格 {filled_cells}/{total_cells} ({filled_cells/total_cells*100:.1f}%)")
        
        # 下载Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name="科目交易数据")
        output.seek(0)
        
        st.download_button(
            label="📥 下载Excel文件",
            data=output,
            file_name=f"日清分数据_科目行式_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.success("✅ 全部处理完成！")
    
    else:
        st.info("👆 请上传PDF文件开始处理")

if __name__ == "__main__":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    main()
