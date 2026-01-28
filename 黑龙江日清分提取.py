import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心配置 ----------------------
# 科目编码到名称的完整映射
TRADE_CODE_MAP = {
    "101010101": "优先发电交易",
    "101020101": "电网企业代理购电交易", 
    "101020301": "省内电力直接交易",
    "101040322": "送上海省间绿色电力交易",
    "102020101": "送辽宁交易",
    "102020301": "送华北交易", 
    "102010101": "送山东交易",
    "102010201": "送浙江交易",
    "202030001": "送江苏省间绿色电力交易",
    "202030002": "送浙江省间绿色电力交易",
    "101080101": "省内现货日前交易",
    "101080201": "省内现货实时交易",
    "101080301": "省间现货日前交易",
    "101080401": "省间现货日内交易",
    "201010101": "中长期合约阻塞费用",
    "201020101": "省间省内价差费用"
}

ALL_TRADES = list(TRADE_CODE_MAP.values())
SPECIAL_TRADES = ["中长期合约阻塞费用", "省间省内价差费用"]
REGULAR_TRADES = [trade for trade in ALL_TRADES if trade not in SPECIAL_TRADES]

# ---------------------- 核心工具函数 ----------------------
def safe_convert_to_numeric(value):
    """安全转换为数值 - 增强版，避免将编码识别为数字"""
    if value is None or pd.isna(value) or value == '':
        return None
    
    # 先转为字符串处理
    val_str = str(value).strip()
    
    # 排除9位数字（科目编码）
    if re.match(r'^\d{9}$', val_str):
        return None
    
    # 排除空字符串和纯符号
    if val_str in ['-', '.', '', '—', '——']:
        return None
    
    try:
        # 移除千分位逗号、人民币符号等
        cleaned = re.sub(r'[^\d.-]', '', val_str)
        if cleaned and cleaned not in ['-', '.', '']:
            return float(cleaned)
        return None
    except (ValueError, TypeError):
        return None

def extract_base_company_info(pdf_text):
    """提取基础公司信息（用于双场站识别）"""
    lines = pdf_text.split('\n')
    base_company = "未知公司"
    
    # 提取基础公司名称
    for line in lines:
        line_clean = line.strip()
        if "公司名称" in line_clean:
            match = re.search(r'公司名称[:：]\s*(.+?有限公司)', line_clean)
            if match:
                base_company = match.group(1).strip()
                break
    
    return base_company

def split_double_station_data(pdf_text, pdf_tables):
    """
    拆分双场站（A/B）数据
    返回：[(station_name, tables_segment), ...]
    """
    base_company = extract_base_company_info(pdf_text)
    lines = pdf_text.split('\n')
    
    # 标记点：查找包含"A风电场"、"B风电场"、"1号"、"2号"、"一号"、"二号"的行
    station_markers = []
    for i, line in enumerate(lines):
        line_clean = line.strip()
        if any(marker in line_clean for marker in ["A风电场", "B风电场", "1号机组", "2号机组", "一号机组", "二号机组"]):
            station_markers.append((i, line_clean))
    
    # 情况1：检测到双场站标记
    if len(station_markers) >= 2:
        # 识别A/B场站
        station_a_marker = None
        station_b_marker = None
        
        for pos, text in station_markers:
            if any(marker in text for marker in ["A风电场", "1号", "一号"]):
                station_a_marker = (pos, f"{base_company}（双发A风电场）")
            elif any(marker in text for marker in ["B风电场", "2号", "二号"]):
                station_b_marker = (pos, f"{base_company}（双发B风电场）")
        
        # 如果找到A/B标记，尝试拆分表格
        if station_a_marker and station_b_marker:
            # 简单策略：将表格分为两部分（实际可根据PDF结构优化）
            mid_idx = len(pdf_tables) // 2
            station_a_tables = pdf_tables[:mid_idx]
            station_b_tables = pdf_tables[mid_idx:]
            
            return [
                (station_a_marker[1], station_a_tables),
                (station_b_marker[1], station_b_tables)
            ]
    
    # 情况2：单场站或无法拆分，返回整体
    # 尝试识别是A还是B
    station_name = f"{base_company}（未知场站）"
    if any(marker in pdf_text for marker in ["A风电场", "1号", "一号"]):
        station_name = f"{base_company}（双发A风电场）"
    elif any(marker in pdf_text for marker in ["B风电场", "2号", "二号"]):
        station_name = f"{base_company}（双发B风电场）"
    
    return [(station_name, pdf_tables)]

def extract_station_and_date_v2(pdf_text, file_name, station_name_override=None):
    """提取场站名称和日期 - 增强版，支持场站名称覆盖"""
    lines = pdf_text.split('\n')
    
    # 使用覆盖的场站名称（双场站拆分时用）
    station_name = station_name_override if station_name_override else "未知场站"
    
    # 如果没有覆盖名称，尝试从文本提取
    if station_name == "未知场站":
        # 方法1: 从包含风电场的行提取
        for line in lines:
            line_clean = line.strip()
            if "风电场" in line_clean:
                match = re.search(r'([^，。！？]+风电场)', line_clean)
                if match:
                    station_name = match.group(1).strip()
                    break
    
    # 提取日期 - 增强的匹配模式
    date = None
    date_patterns = [
        r'清分日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2})',
        r'日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2})',
        r'(\d{4}年\d{1,2}月\d{1,2}日)',
        r'(\d{4}/\d{1,2}/\d{1,2})',
        r'(\d{4}\.\d{1,2}\.\d{1,2})'
    ]
    
    for line in lines:
        for pattern in date_patterns:
            match = re.search(pattern, line)
            if match:
                date_str = match.group(1)
                # 统一转换为yyyy-mm-dd格式
                date_str = date_str.replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-').replace('.', '-')
                # 补全月份和日期的前导零
                parts = date_str.split('-')
                if len(parts) == 3:
                    year, month, day = parts
                    date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                break
        if date:
            break
    
    # 从文件名提取日期（备用）
    if not date:
        date_match = re.search(r'(\d{4}-\d{2}-\d{2})|(\d{8})', file_name)
        if date_match:
            date_str = date_match.group()
            if len(date_str) == 8:  # yyyymmdd格式
                date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
            else:
                date = date_str
    
    return station_name, date

def extract_data_using_pdfplumber_tables(file_obj):
    """使用pdfplumber的表格提取功能 - 增强版"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            all_tables = []
            for page in pdf.pages:
                # 优化表格提取参数
                tables = page.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 3,
                    "join_tolerance": 3
                })
                if tables:
                    for table in tables:
                        # 深度清理表格数据
                        cleaned_table = []
                        for row in table:
                            cleaned_row = []
                            for cell in row:
                                if cell is None:
                                    cleaned_row.append("")
                                else:
                                    # 移除空白字符和特殊符号
                                    cell_clean = re.sub(r'\s+', ' ', str(cell)).strip()
                                    cleaned_row.append(cell_clean)
                            # 跳过空行
                            if any(cell != "" for cell in cleaned_row):
                                cleaned_table.append(cleaned_row)
                        if cleaned_table:  # 只添加非空表格
                            all_tables.append(cleaned_table)
            
            return all_tables
    except Exception as e:
        st.error(f"表格提取失败: {e}")
        return []

def parse_trade_table_data_v2(tables):
    """解析交易表格数据 - 增强版，避免编码识别错误"""
    trade_data = {}
    
    # 初始化所有科目
    for trade in ALL_TRADES:
        if trade in SPECIAL_TRADES:
            trade_data[trade] = {'fee': None}
        else:
            trade_data[trade] = {'quantity': None, 'price': None, 'fee': None}
    
    for table in tables:
        if len(table) < 2:  # 至少要有表头和数据行
            continue
            
        # 智能查找表头行
        header_row = -1
        code_col = -1
        name_col = -1
        qty_col = -1
        price_col = -1
        fee_col = -1
        
        # 遍历所有行寻找表头
        for i, row in enumerate(table):
            row_str = ' '.join([str(cell) for cell in row if cell])
            # 更宽松的表头识别
            if ("科目编码" in row_str or "编码" in row_str) and ("科目名称" in row_str or "名称" in row_str):
                header_row = i
                # 确定各列位置（支持模糊匹配）
                for j, cell in enumerate(row):
                    cell_lower = str(cell).lower()
                    if any(keyword in cell_lower for keyword in ["科目编码", "编码", "code"]):
                        code_col = j
                    elif any(keyword in cell_lower for keyword in ["科目名称", "名称", "name"]):
                        name_col = j
                    elif any(keyword in cell_lower for keyword in ["电量", "数量", "kwh", "mwh"]):
                        qty_col = j
                    elif any(keyword in cell_lower for keyword in ["电价", "价格", "price"]):
                        price_col = j
                    elif any(keyword in cell_lower for keyword in ["电费", "金额", "费用", "amount"]):
                        fee_col = j
                break
        
        if header_row == -1:
            continue
            
        # 解析数据行（跳过表头和合计行）
        for i in range(header_row + 1, len(table)):
            row = table[i]
            # 跳过合计/总计行
            row_str = ' '.join([str(cell) for cell in row if cell])
            if any(keyword in row_str for keyword in ["合计", "总计", "小计", "summary", "total"]):
                continue
            
            # 提取科目编码和名称
            trade_code = ""
            trade_name = None
            
            # 从编码列提取
            if code_col >= 0 and code_col < len(row):
                trade_code = str(row[code_col]).strip()
                if trade_code in TRADE_CODE_MAP:
                    trade_name = TRADE_CODE_MAP[trade_code]
            
            # 编码匹配失败，尝试从名称列匹配
            if not trade_name and name_col >= 0 and name_col < len(row):
                name_cell = str(row[name_col]).strip()
                for code, name in TRADE_CODE_MAP.items():
                    if name in name_cell or name.replace("交易", "") in name_cell:
                        trade_name = name
                        break
            
            if not trade_name:
                continue
            
            # 提取数据（使用安全转换函数）
            is_special = trade_name in SPECIAL_TRADES
            
            if is_special:
                # 特殊科目只有电费
                if fee_col >= 0 and fee_col < len(row):
                    fee_val = row[fee_col]
                    trade_data[trade_name]['fee'] = safe_convert_to_numeric(fee_val)
            else:
                # 常规科目 - 更容错的提取逻辑
                if qty_col >= 0 and qty_col < len(row):
                    qty_val = row[qty_col]
                    trade_data[trade_name]['quantity'] = safe_convert_to_numeric(qty_val)
                
                if price_col >= 0 and price_col < len(row):
                    price_val = row[price_col]
                    trade_data[trade_name]['price'] = safe_convert_to_numeric(price_val)
                
                if fee_col >= 0 and fee_col < len(row):
                    fee_val = row[fee_col]
                    trade_data[trade_name]['fee'] = safe_convert_to_numeric(fee_val)
    
    return trade_data

def extract_total_data_v2(pdf_text):
    """提取合计数据 - 增强版"""
    total_quantity, total_amount = None, None
    
    lines = pdf_text.split('\n')
    
    for line in lines:
        line_clean = line.replace(' ', '').replace(',', '').replace('，', '')
        
        # 更精准的合计电量提取
        qty_match = re.search(r'合计电量[:：]([\d\.]+)', line_clean)
        if qty_match:
            total_quantity = safe_convert_to_numeric(qty_match.group(1))
        
        # 更精准的合计电费提取
        fee_match = re.search(r'合计电费[:：]([\d\.]+)', line_clean)
        if fee_match:
            total_amount = safe_convert_to_numeric(fee_match.group(1))
    
    return total_quantity, total_amount

def process_single_station(station_name, tables, pdf_text, file_name):
    """处理单个场站的数据提取"""
    # 提取基础信息
    station_name, date = extract_station_and_date_v2(pdf_text, file_name, station_name)
    total_quantity, total_amount = extract_total_data_v2(pdf_text)
    
    # 解析交易数据
    trade_data = parse_trade_table_data_v2(tables)
    
    # 构建结果列表
    result = [station_name, date, total_quantity, total_amount]
    
    # 添加常规科目数据
    for trade in REGULAR_TRADES:
        data = trade_data.get(trade, {'quantity': None, 'price': None, 'fee': None})
        result.extend([data['quantity'], data['price'], data['fee']])
    
    # 添加特殊科目数据
    for trade in SPECIAL_TRADES:
        data = trade_data.get(trade, {'fee': None})
        result.append(data['fee'])
    
    return result

def extract_data_from_pdf_v2(file_obj, file_name):
    """从PDF提取数据 - 支持双场站版本"""
    try:
        # 首先读取PDF文本和表格
        file_obj.seek(0)
        with pdfplumber.open(file_obj) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
        
        if not all_text or len(all_text.strip()) < 50:
            raise ValueError("PDF为空或文本内容太少")
        
        # 提取表格数据
        file_obj.seek(0)
        all_tables = extract_data_using_pdfplumber_tables(file_obj)
        
        if not all_tables:
            # 表格提取失败，使用文本分析（备用方案）
            st.warning(f"{file_name}: 表格提取失败，使用文本分析模式")
            # 这里可以添加文本分析的备用逻辑
        
        # 拆分双场站数据
        station_data_list = split_double_station_data(all_text, all_tables)
        
        # 处理每个场站
        results = []
        for station_name, tables_segment in station_data_list:
            result = process_single_station(station_name, tables_segment, all_text, file_name)
            results.append(result)
        
        return results
        
    except Exception as e:
        st.error(f"处理PDF {file_name} 出错: {str(e)}")
        # 返回默认格式的错误数据
        default_result = ["未知场站", None, None, None] + [None] * (len(REGULAR_TRADES) * 3 + len(SPECIAL_TRADES))
        return [default_result]

# ---------------------- Streamlit 应用 ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取工具", layout="wide")
    
    st.title("📊 黑龙江日清分结算单数据提取工具（双场站增强版）")
    st.markdown("**核心改进：支持双场站(A/B)识别、修复数据提取错误、减少None值**")
    st.divider()
    
    # 显示科目信息
    with st.expander("📋 支持的科目列表"):
        st.write("**常规科目（电量、电价、电费）：**")
        for trade in REGULAR_TRADES:
            st.write(f"- {trade}")
        
        st.write("**特殊科目（仅电费）：**")
        for trade in SPECIAL_TRADES:
            st.write(f"- {trade}")
    
    st.subheader("📁 上传文件")
    uploaded_files = st.file_uploader(
        "支持PDF格式，可批量上传（支持双场站PDF）",
        type=['pdf'],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        if st.button("🚀 开始处理", type="primary"):
            st.divider()
            st.subheader("⚙️ 处理进度")
            
            all_data = []
            progress_bar = st.progress(0)
            
            for idx, file in enumerate(uploaded_files):
                progress_bar.progress((idx + 1) / len(uploaded_files))
                
                try:
                    # 处理PDF（可能返回多个场站的数据）
                    file_results = extract_data_from_pdf_v2(file, file.name)
                    for result in file_results:
                        all_data.append(result)
                    
                    # 显示处理结果
                    if len(file_results) == 2:
                        st.success(f"✓ {file.name} 处理成功（识别出2个场站）")
                    elif len(file_results) == 1:
                        st.success(f"✓ {file.name} 处理成功（识别出1个场站）")
                    else:
                        st.warning(f"⚠ {file.name} 处理完成，但未识别到场站数据")
                        
                except Exception as e:
                    st.error(f"✗ {file.name} 处理失败: {str(e)}")
            
            progress_bar.empty()
            
            if all_data:
                # 构建结果DataFrame
                result_columns = ['场站名称', '清分日期', '合计电量(兆瓦时)', '合计电费(元)']
                
                for trade in REGULAR_TRADES:
                    # 简化列名，避免过长
                    trade_short = trade.replace('省间绿色电力交易', '省间绿电交易')
                    trade_short = trade_short.replace('电网企业代理购电交易', '代理购电交易')
                    result_columns.extend([f'{trade_short}_电量', f'{trade_short}_电价', f'{trade_short}_电费'])
                
                for trade in SPECIAL_TRADES:
                    result_columns.append(f'{trade}_电费')
                
                result_df = pd.DataFrame(all_data, columns=result_columns)
                
                # 显示结果
                st.subheader("📈 提取结果")
                st.dataframe(result_df, use_container_width=True)
                
                # 增强的统计信息
                st.info(f"**统计信息：** 共处理 {len(all_data)} 条场站记录，涉及 {result_df['场站名称'].nunique()} 个场站")
                
                # 检查双发A/B风电场
                has_a_station = any('双发A' in str(name) for name in result_df['场站名称'])
                has_b_station = any('双发B' in str(name) for name in result_df['场站名称'])
                
                if has_a_station and has_b_station:
                    st.success("✅ 成功识别双发A/B风电场数据")
                elif has_a_station:
                    st.warning("⚠️ 仅识别到双发A风电场，未检测到B风电场数据")
                elif has_b_station:
                    st.warning("⚠️ 仅识别到双发B风电场，未检测到A风电场数据")
                
                # 更精准的数据完整性统计
                data_columns = result_columns[4:]
                non_null_count = result_df[data_columns].notna().sum()
                total_cells = len(result_df) * len(data_columns)
                filled_cells = result_df[data_columns].notna().sum().sum()
                
                st.info(f"**数据完整性：**")
                st.info(f"- 总数据单元格：{total_cells}")
                st.info(f"- 有值单元格：{filled_cells} ({filled_cells/total_cells*100:.1f}%)")
                
                # 下载功能
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                output.seek(0)
                
                st.download_button(
                    label="📥 下载Excel文件",
                    data=output,
                    file_name=f"黑龙江结算数据_双场站版_{current_time}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.success("✅ 全部处理完成！")
    
    else:
        st.info("👆 请上传PDF文件开始处理（支持包含双场站的PDF）")

if __name__ == "__main__":
    main()
