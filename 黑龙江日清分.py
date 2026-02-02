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

# ---------------------- 核心配置（新增水印过滤+扩展科目映射） ----------------------
# 1. 常见水印关键词（需根据实际PDF补充）
WATERMARK_KEYWORDS = ["协合能源", "大庆晶盛", "太阳能发电", "内部使用", "CONFIDENTIAL", "草稿"]
# 2. 完整科目映射（覆盖更多可能编码）
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
    "0101100101": "偏差考核费用",
    "0101030101": "居民农业用电交易"  # 新增可能科目
}
# 3. 科目名称关键词库（用于模糊修正）
TRADE_NAME_KEYWORDS = {
    "优先发电": "优先发电交易",
    "代理购电": "电网企业代理购电交易",
    "直接交易": "省内电力直接交易",
    "现货日前": "省内现货日前交易",
    "现货实时": "省内现货实时交易",
    "阻塞费用": "中长期合约阻塞费用",
    "价差费用": "省间省内价差费用",
    "辅助服务": "辅助服务费用分摊",
    "偏差考核": "偏差考核费用"
}
# 4. 特殊科目（仅含费用）
SPECIAL_TRADES = ["中长期合约阻塞费用", "省间省内价差费用", "辅助服务费用分摊", "偏差考核费用"]

# ---------------------- 核心工具函数（全链路优化） ----------------------
def remove_watermark(text):
    """第一步：移除水印干扰"""
    if not text:
        return ""
    # 1. 移除水印关键词
    cleaned_text = text
    for keyword in WATERMARK_KEYWORDS:
        cleaned_text = cleaned_text.replace(keyword, "")
    # 2. 移除连续空白符和特殊字符
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text)  # 多个空格转单个
    cleaned_text = re.sub(r'[\x00-\x1F\x7F]', '', cleaned_text)  # 控制字符
    return cleaned_text.strip()

def safe_convert_to_numeric(value):
    """第二步：安全转换数值（增加数据合理性校验）"""
    if value is None or pd.isna(value) or value == '':
        return None
    
    # 先移除水印残留
    val_str = remove_watermark(str(value)).strip()
    # 排除科目编码和纯符号
    if re.match(r'^\d{9,10}$', val_str) or val_str in ['-', '.', '', '—', '——']:
        return None
    
    try:
        # 清理千分位、全角符号
        cleaned = re.sub(r'[^\d.-]', '', val_str.replace('，', ',').replace('。', '.'))
        if not cleaned or cleaned in ['-', '.']:
            return None
        num = float(cleaned)
        # 合理性校验（示例：电量不会小于0，电价不会超过10000元/MWh）
        if '电量' in str(value) and num < 0:
            return None
        if '电价' in str(value) and (num < 0 or num > 10000):
            return None
        return num
    except (ValueError, TypeError):
        return None

def extract_clear_date(pdf_text):
    """第三步：精准提取清分日期（覆盖所有常见格式）"""
    date_patterns = [
        r'清分日期[:：]\s*(\d{4}[年/-]\d{1,2}[月/-]\d{1,2}[日]?)',  # 清分日期：2026-01-01/2026年01月01日
        r'结算日期[:：]\s*(\d{4}[年/-]\d{1,2}[月/-]\d{1,2}[日]?)',  # 结算日期：2026.01.01
        r'(\d{4}年\d{1,2}月\d{1,2}日)\s*清分',  # 2026年01月01日清分
        r'(\d{4}-\d{1,2}-\d{1,2})\s*现货日清分'  # 2026-01-01 现货日清分
    ]
    
    for pattern in date_patterns:
        match = re.search(pattern, pdf_text)
        if match:
            date_str = match.group(1)
            # 统一格式为YYYY-MM-DD
            date_str = re.sub(r'[年月日]', '-', date_str).rstrip('-')
            # 补全两位数月份/日期
            parts = date_str.split('-')
            if len(parts) == 3:
                year, month, day = parts
                return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
    return None

def extract_base_info(pdf_text):
    """第四步：提取基础信息（公司、日期、合计）"""
    # 先去水印
    clean_text = remove_watermark(pdf_text)
    lines = clean_text.split('\n')
    
    # 1. 公司名称（匹配“有限公司”结尾）
    company_name = "未知公司"
    for line in lines:
        if "公司名称" in line:
            match = re.search(r'公司名称[:：]\s*([^，。\n]+有限公司)', line)
            if match:
                company_name = match.group(1).strip()
                break
    
    # 2. 清分日期（调用精准提取函数）
    date = extract_clear_date(clean_text)
    
    # 3. 合计/小计数据（保留“小计”）
    total_quantity = None  # 总电量
    total_fee = None       # 总电费
    subtotal_quantity = None  # 小计电量
    subtotal_fee = None       # 小计电费
    
    for line in lines:
        line_clean = remove_watermark(line).replace(' ', '').replace(',', '')
        # 匹配“小计”（当日汇总）
        if '小计' in line_clean:
            qty_match = re.search(r'小计电量[:：]([\d\.]+)|电量小计[:：]([\d\.]+)', line_clean)
            fee_match = re.search(r'小计电费[:：]([\d\.]+)|电费小计[:：]([\d\.]+)', line_clean)
            if qty_match:
                subtotal_quantity = safe_convert_to_numeric(next(g for g in qty_match.groups() if g))
            if fee_match:
                subtotal_fee = safe_convert_to_numeric(next(g for g in fee_match.groups() if g))
        # 匹配“合计”（全局汇总）
        elif '合计' in line_clean and '小计' not in line_clean:
            qty_match = re.search(r'合计电量[:：]([\d\.]+)|总电量[:：]([\d\.]+)', line_clean)
            fee_match = re.search(r'合计电费[:：]([\d\.]+)|总电费[:：]([\d\.]+)', line_clean)
            if qty_match:
                total_quantity = safe_convert_to_numeric(next(g for g in qty_match.groups() if g))
            if fee_match:
                total_fee = safe_convert_to_numeric(next(g for g in fee_match.groups() if g))
    
    # 优先用“小计”（当日汇总更关键）
    final_qty = subtotal_quantity if subtotal_quantity is not None else total_quantity
    final_fee = subtotal_fee if subtotal_fee is not None else total_fee
    
    return company_name, date, final_qty, final_fee, (subtotal_quantity is not None or total_quantity is not None)

def locate_table_columns(table_rows):
    """第五步：精准定位表格列（基于完整表头文本）"""
    # 目标列的完整关键词（需与PDF表头完全匹配）
    target_columns = {
        "科目编码": [],
        "科目名称": [],
        "交易电量": [],  # 匹配“交易电量(兆瓦时)”“电量(MWh)”
        "结算电价": [],  # 匹配“结算电价(元/兆瓦时)”“电价(元)”
        "结算电费": []   # 匹配“结算电费(元)”“电费(元)”
    }
    
    # 遍历前3行表头，记录每列的匹配度
    for row_idx, row in enumerate(table_rows[:3]):
        for col_idx, cell in enumerate(row):
            cell_clean = remove_watermark(str(cell)).lower().strip()
            # 匹配科目编码
            if any(key in cell_clean for key in ["科目编码", "编码", "code"]):
                target_columns["科目编码"].append((row_idx, col_idx, 1.0))
            # 匹配科目名称
            elif any(key in cell_clean for key in ["科目名称", "结算类型", "名称", "type"]):
                target_columns["科目名称"].append((row_idx, col_idx, 1.0))
            # 匹配交易电量（必须包含“电量”+单位）
            elif any(key in cell_clean for key in ["电量", "mw", "兆瓦时"]) and not "价" in cell_clean:
                target_columns["交易电量"].append((row_idx, col_idx, 0.9))
            # 匹配结算电价（必须包含“价”+单位）
            elif any(key in cell_clean for key in ["电价", "单价", "price"]) and not "量" in cell_clean:
                target_columns["结算电价"].append((row_idx, col_idx, 0.9))
            # 匹配结算电费（必须包含“费”+单位）
            elif any(key in cell_clean for key in ["电费", "金额", "fee", "元"]) and not "量" in cell_clean and not "价" in cell_clean:
                target_columns["结算电费"].append((row_idx, col_idx, 0.9))
    
    # 确定最终列索引（取匹配度最高的列，避免重复）
    used_cols = set()
    final_cols = {}
    for col_name, matches in target_columns.items():
        if not matches:
            final_cols[col_name] = -1  # 未找到
            continue
        # 按行优先级（第1行表头 > 第2行）排序
        matches.sort(key=lambda x: (x[0], -x[2]))
        for row_idx, col_idx, score in matches:
            if col_idx not in used_cols:
                final_cols[col_name] = col_idx
                used_cols.add(col_idx)
                break
        else:
            final_cols[col_name] = -1
    
    return final_cols

def correct_trade_name(trade_name):
    """第六步：修正科目名称（水印污染后修复）"""
    if not trade_name:
        return "未知科目"
    # 先去水印
    clean_name = remove_watermark(trade_name).strip()
    # 1. 按关键词模糊匹配
    for keyword, correct_name in TRADE_NAME_KEYWORDS.items():
        if keyword in clean_name:
            return correct_name
    # 2. 若仍未知，返回清理后的名称
    return clean_name if clean_name else "未知科目"

def extract_trade_data_from_tables(tables, clear_date):
    """第七步：提取科目数据（保留小计行，修正名称）"""
    trade_records = []
    
    for table in tables:
        if len(table) < 4:  # 至少表头2行+数据1行+小计1行
            continue
        
        # 第一步：清理表格（去水印）
        clean_table = []
        for row in table:
            clean_row = [remove_watermark(str(cell)) for cell in row]
            if any(cell.strip() != '' for cell in clean_row):  # 跳过空行
                clean_table.append(clean_row)
        if len(clean_table) < 4:
            continue
        
        # 第二步：定位列索引
        final_cols = locate_table_columns(clean_table)
        code_col = final_cols["科目编码"]
        name_col = final_cols["科目名称"]
        qty_col = final_cols["交易电量"]
        price_col = final_cols["结算电价"]
        fee_col = final_cols["结算电费"]
        # 核心列必须存在（编码/名称 + 电量/电费）
        if (code_col == -1 and name_col == -1) or (qty_col == -1 and fee_col == -1):
            continue
        
        # 第三步：解析数据行（保留小计，跳过合计）
        for row_idx, row in enumerate(clean_table):
            row_clean = [cell.strip() for cell in row]
            # 跳过表头行（前2行）
            if row_idx < 2:
                continue
            # 跳过全局合计行（保留小计行）
            if '合计' in ''.join(row_clean) and '小计' not in ''.join(row_clean):
                continue
            
            # 提取基础信息
            trade_code = row[code_col].strip() if (code_col != -1 and code_col < len(row)) else ''
            raw_name = row[name_col].strip() if (name_col != -1 and name_col < len(row)) else ''
            # 修正科目名称
            trade_name = TRADE_CODE_MAP.get(trade_code, correct_trade_name(raw_name))
            
            # 提取数据（带合理性校验）
            quantity = safe_convert_to_numeric(row[qty_col]) if (qty_col != -1 and qty_col < len(row)) else None
            price = safe_convert_to_numeric(row[price_col]) if (price_col != -1 and price_col < len(row)) else None
            fee = safe_convert_to_numeric(row[fee_col]) if (fee_col != -1 and fee_col < len(row)) else None
            
            # 特殊科目处理（仅保留费用）
            if trade_name in SPECIAL_TRADES:
                quantity = None
                price = None
            
            # 标记是否为小计行
            is_subtotal = '小计' in ''.join(row_clean)
            
            # 新增：数据行必须关联日期
            if not is_subtotal and (quantity is None and fee is None):
                continue  # 非小计行无数据则跳过
            
            # 添加到结果
            trade_records.append({
                "科目名称": trade_name,
                "是否小计行": is_subtotal,
                "电量(兆瓦时)": quantity if not is_subtotal else quantity,
                "电价(元/兆瓦时)": price if not is_subtotal else None,  # 小计行无电价
                "电费(元)": fee,
                "原始科目编码": trade_code,
                "原始科目名称": raw_name
            })
    
    return trade_records

def parse_pdf_file(file_obj, file_name):
    """主解析函数（整合全链路优化）"""
    try:
        # 重置文件流
        file_obj.seek(0)
        file_bytes = BytesIO(file_obj.read())
        file_bytes.seek(0)
        
        # 提取文本和表格（去水印）
        all_text = ""
        all_tables = []
        with pdfplumber.open(file_bytes) as pdf:
            for page in pdf.pages:
                # 提取文本（去水印）
                text = page.extract_text() or ""
                all_text += remove_watermark(text) + "\n"
                # 提取表格（保留原始结构用于列定位）
                tables = page.extract_tables({
                    "vertical_strategy": "lines",  # 按表格线定位（精准度最高）
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 2,  # 缩小对齐公差
                    "join_tolerance": 2,
                    "edge_min_length": 8  # 过滤短线条干扰
                })
                all_tables.extend(tables)
        
        # 1. 提取基础信息（公司、日期、小计）
        company_name, clear_date, total_qty, total_fee, has_subtotal = extract_base_info(all_text)
        # 2. 提取科目数据（含小计行）
        trade_records = extract_trade_data_from_tables(all_tables, clear_date)
        # 3. 补充小计行（若未提取到，手动添加）
        if has_subtotal and not any(rec["是否小计行"] for rec in trade_records):
            trade_records.append({
                "科目名称": "当日小计",
                "是否小计行": True,
                "电量(兆瓦时)": total_qty,
                "电价(元/兆瓦时)": None,
                "电费(元)": total_fee,
                "原始科目编码": "SUBTOTAL",
                "原始科目名称": "当日小计"
            })
        
        # 4. 补充场站和日期信息
        station_name = f"{company_name}（晶盛光伏电站）" if "晶盛" in company_name else f"{company_name}（未知场站）"
        for record in trade_records:
            record["场站名称"] = station_name
            record["清分日期"] = clear_date
            # 删除原始字段（仅保留最终结果）
            record.pop("原始科目编码", None)
            record.pop("原始科目名称", None)
        
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

# ---------------------- Streamlit 应用（适配优化后逻辑） ----------------------
def main():
    st.set_page_config(page_title="日清分数据提取工具（水印优化版）", layout="wide")
    
    st.title("📊 日清分结算单数据提取工具（抗水印版）")
    st.markdown("**核心功能：水印过滤 | 日期精准提取 | 小计行保留 | 数据错位修复**")
    st.divider()
    
    # 上传文件
    uploaded_files = st.file_uploader(
        "支持PDF格式（推荐单文件上传，避免批量干扰）",
        type=['pdf'],
        accept_multiple_files=False  # 单文件上传更稳定
    )
    
    if uploaded_files and st.button("🚀 开始处理", type="primary"):
        st.divider()
        st.subheader("⚙️ 处理进度")
        
        # 处理单个文件（更稳定）
        file = uploaded_files
        st.write(f"正在处理：{file.name}")
        trade_records = parse_pdf_file(file, file.name)
        file.close()
        
        # 转换为DataFrame（调整列顺序）
        result_df = pd.DataFrame(trade_records)
        col_order = ["场站名称", "清分日期", "科目名称", "是否小计行", "电量(兆瓦时)", "电价(元/兆瓦时)", "电费(元)"]
        result_df = result_df[col_order]
        
        # 显示结果（高亮小计行）
        st.subheader("📈 提取结果（小计行已高亮）")
        # 高亮小计行
        styled_df = result_df.style.apply(
            lambda row: ['background-color: #f0f8ff' if row["是否小计行"] else '' for _ in row],
            axis=1
        )
        st.dataframe(styled_df, use_container_width=True)
        
        # 关键信息统计
        subtotal_count = result_df[result_df["是否小计行"]].shape[0]
        valid_trade_count = result_df[~result_df["是否小计行"]].shape[0]
        st.info(f"**统计信息：** 共提取 {valid_trade_count} 个科目 + {subtotal_count} 个小计行 | 清分日期：{result_df['清分日期'].iloc[0] or '未识别'}")
        
        # 数据完整性分析
        data_cols = ["电量(兆瓦时)", "电价(元/兆瓦时)", "电费(元)"]
        filled_cells = result_df[data_cols].notna().sum().sum()
        total_cells = len(result_df) * len(data_cols)
        st.info(f"**数据完整性：** 有效数据单元格 {filled_cells}/{total_cells} ({filled_cells/total_cells*100:.1f}%)")
        
        # 下载Excel（保留高亮格式）
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name="日清分数据")
            # 高亮小计行（Excel中）
            worksheet = writer.sheets["日清分数据"]
            for row_idx in range(2, len(result_df) + 2):  # 第1行是表头
                if result_df.iloc[row_idx - 2]["是否小计行"]:
                    for col_idx in range(1, len(col_order) + 1):
                        worksheet.cell(row=row_idx, column=col_idx).fill = pd.ExcelWriter._xlsx.styles.PatternFill(
                            start_color="F0F8FF", end_color="F0F8FF", fill_type="solid"
                        )
        output.seek(0)
        
        st.download_button(
            label="📥 下载Excel（小计行高亮）",
            data=output,
            file_name=f"日清分数据_抗水印版_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.success("✅ 处理完成！若仍有数据问题，请检查PDF水印是否已添加到WATERMARK_KEYWORDS中")
    
    else:
        st.info("👆 请上传单个PDF文件开始处理（建议先清理PDF水印再上传）")

if __name__ == "__main__":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    main()
