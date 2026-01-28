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

# ---------------------- 核心配置（保留原有） ----------------------
TRADE_CODE_MAP = {
    "0101010101": "优先发电交易",
    "0101020101": "电网企业代理购电交易", 
    "0101020301": "省内电力直接交易",
    "0101040322": "送上海省间绿色电力交易",
    "0102020101": "送辽宁交易",
    "0102020301": "送华北交易", 
    "0102010101": "送山东交易",
    "0102010201": "送浙江交易",
    "0202030001": "送江苏省间绿色电力交易",
    "0202030002": "送浙江省间绿色电力交易",
    "0101080101": "省内现货日前交易",
    "0101080201": "省内现货实时交易",
    "0101080301": "省间现货日前交易",
    "0101080401": "省间现货日内交易",
    "0201010101": "中长期合约阻塞费用",
    "0201020101": "省间省内价差费用",
    "0101050101": "省内绿色电力交易(电能量)",
    "0101060101": "日融合交易",
    "0101070101": "现货结算价差调整",
    "0101090101": "辅助服务费用分摊",
    "0101100101": "偏差考核费用"
}

ALL_TRADES = list(TRADE_CODE_MAP.values())
SPECIAL_TRADES = ["中长期合约阻塞费用", "省间省内价差费用", "辅助服务费用分摊", "偏差考核费用"]
REGULAR_TRADES = [trade for trade in ALL_TRADES if trade not in SPECIAL_TRADES]

# ---------------------- 核心工具函数（关键修改） ----------------------
def safe_convert_to_numeric(value):
    """安全转换为数值 - 增强版，兼容网页端特殊字符"""
    if value is None or pd.isna(value) or value == '':
        return None
    
    # 先转为字符串处理，移除网页端常见的非断行空格(\xa0)
    val_str = str(value).strip().replace('\xa0', ' ')
    
    # 排除9/10位数字（科目编码）
    if re.match(r'^\d{9,10}$', val_str):
        return None
    
    # 排除空字符串和纯符号（补充网页端常见符号）
    if val_str in ['-', '.', '', '—', '——', ' ', '\t', '\n']:
        return None
    
    try:
        # 移除千分位逗号、人民币符号、全角符号等（增强）
        cleaned = re.sub(r'[^\d.-]', '', val_str.replace('，', ',').replace('。', '.'))
        if cleaned and cleaned not in ['-', '.', '']:
            return float(cleaned)
        return None
    except (ValueError, TypeError):
        return None

def extract_base_company_info(pdf_text):
    """提取基础公司信息 - 增强网页端字符兼容"""
    # 统一字符格式，移除特殊空白符
    pdf_text = pdf_text.replace('\xa0', ' ').replace('\r', '\n').strip()
    lines = pdf_text.split('\n')
    base_company = "未知公司"
    
    for line in lines:
        line_clean = line.strip()
        if "公司名称" in line_clean:
            # 增强正则，兼容全角冒号/空格
            match = re.search(r'公司名称[:：]\s*(.+?有限公司)', line_clean)
            if match:
                base_company = match.group(1).strip()
                break
    
    return base_company

def split_double_station_data(pdf_text, pdf_tables):
    """拆分双场站数据 - 增强鲁棒性，适配网页端文本"""
    # 统一文本格式
    pdf_text = pdf_text.replace('\xa0', ' ').replace('\r', '\n').strip()
    base_company = extract_base_company_info(pdf_text)
    lines = pdf_text.split('\n')
    
    station_markers = []
    for i, line in enumerate(lines):
        line_clean = line.strip()
        # 增强标记匹配，兼容网页端字符差异
        if any(marker in line_clean for marker in ["A风电场", "B风电场", "1号机组", "2号机组", "一号机组", "二号机组", "A场", "B场"]):
            station_markers.append((i, line_clean))
    
    # 优化拆分逻辑：即使标记不足2个，也尝试按表格数量拆分
    if len(station_markers) >= 2:
        station_a_marker = None
        station_b_marker = None
        
        for pos, text in station_markers:
            if any(marker in text for marker in ["A风电场", "1号", "一号", "A场"]):
                station_a_marker = (pos, f"{base_company}（双发A风电场）")
            elif any(marker in text for marker in ["B风电场", "2号", "二号", "B场"]):
                station_b_marker = (pos, f"{base_company}（双发B风电场）")
        
        if station_a_marker and station_b_marker:
            mid_idx = len(pdf_tables) // 2
            station_a_tables = pdf_tables[:mid_idx] if mid_idx > 0 else pdf_tables
            station_b_tables = pdf_tables[mid_idx:] if mid_idx > 0 else []
            # 避免空表格
            station_a_tables = station_a_tables if station_a_tables else pdf_tables
            station_b_tables = station_b_tables if station_b_tables else pdf_tables
            
            return [
                (station_a_marker[1], station_a_tables),
                (station_b_marker[1], station_b_tables)
            ]
    
    # 单场站处理：增强名称识别
    station_name = f"{base_company}（未知场站）"
    if any(marker in pdf_text for marker in ["A风电场", "1号", "一号", "A场"]):
        station_name = f"{base_company}（双发A风电场）"
    elif any(marker in pdf_text for marker in ["B风电场", "2号", "二号", "B场"]):
        station_name = f"{base_company}（双发B风电场）"
    
    return [(station_name, pdf_tables)]

def extract_station_and_date_v2(pdf_text, file_name, station_name_override=None):
    """提取场站名称和日期 - 增强网页端日期匹配"""
    # 统一文本格式
    pdf_text = pdf_text.replace('\xa0', ' ').replace('\r', '\n').strip()
    lines = pdf_text.split('\n')
    
    station_name = station_name_override if station_name_override else "未知场站"
    
    if station_name == "未知场站":
        for line in lines:
            line_clean = line.strip()
            if "风电场" in line_clean:
                match = re.search(r'([^，。！？、；]+风电场)', line_clean)
                if match:
                    station_name = match.group(1).strip()
                    break
    
    # 增强日期匹配：兼容更多格式，处理网页端字符
    date = None
    date_patterns = [
        r'清分日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2})',
        r'日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2})',
        r'(\d{4}年\d{1,2}月\d{1,2}日)',
        r'(\d{4}/\d{1,2}/\d{1,2})',
        r'(\d{4}\.\d{1,2}\.\d{1,2})',
        r'(\d{8})'  # 补充纯数字日期
    ]
    
    for line in lines:
        line = line.replace('\xa0', ' ')
        for pattern in date_patterns:
            match = re.search(pattern, line)
            if match:
                date_str = match.group(1)
                date_str = date_str.replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-').replace('.', '-')
                parts = date_str.split('-')
                if len(parts) == 3:
                    year, month, day = parts
                    date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                elif len(date_str) == 8:  # 处理纯数字日期
                    date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
                break
        if date:
            break
    
    # 从文件名提取日期（增强：兼容更多文件名格式）
    if not date:
        date_match = re.search(r'(\d{4}-\d{2}-\d{2})|(\d{8})|(\d{4}_\d{2}_\d{2})', file_name)
        if date_match:
            date_str = date_match.group()
            date_str = date_str.replace('_', '-')
            if len(date_str) == 8:
                date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
            else:
                date = date_str
    
    return station_name, date

def extract_data_using_pdfplumber_tables(file_obj):
    """使用pdfplumber提取表格 - 适配网页端pdfplumber版本差异"""
    try:
        # 重新读取文件流，避免网页端指针偏移
        file_obj.seek(0)
        with pdfplumber.open(file_obj) as pdf:
            all_tables = []
            for page in pdf.pages:
                # 兼容不同pdfplumber版本的参数（关键修复）
                try:
                    # 新版本参数
                    tables = page.extract_tables({
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "lines",
                        "snap_tolerance": 3,
                        "join_tolerance": 3,
                        "edge_min_length": 10
                    })
                except TypeError:
                    # 旧版本参数（网页端常见）
                    tables = page.extract_tables(
                        vertical_strategy="lines",
                        horizontal_strategy="lines",
                        snap_tolerance=3,
                        join_tolerance=3,
                        edge_min_length=10
                    )
                
                if tables:
                    for table in tables:
                        cleaned_table = []
                        for row in table:
                            cleaned_row = []
                            for cell in row:
                                if cell is None:
                                    cleaned_row.append("")
                                else:
                                    # 移除网页端特殊字符
                                    cell_clean = re.sub(r'\s+', ' ', str(cell)).replace('\xa0', ' ').strip()
                                    cleaned_row.append(cell_clean)
                            if any(cell != "" for cell in cleaned_row):
                                cleaned_table.append(cleaned_row)
                        if cleaned_table:
                            all_tables.append(cleaned_table)
            
            return all_tables
    except Exception as e:
        st.error(f"表格提取失败: {e} (pdfplumber版本: {pdfplumber.__version__})")
        return []

def parse_trade_table_data_v2(tables):
    """解析交易表格数据 - 增强网页端表头匹配"""
    trade_data = {}
    for trade in ALL_TRADES:
        if trade in SPECIAL_TRADES:
            trade_data[trade] = {'fee': None}
        else:
            trade_data[trade] = {'quantity': None, 'price': None, 'fee': None}
    
    for table in tables:
        if len(table) < 2:  # 降低表头行数要求，兼容网页端表格提取差异
            continue
            
        code_col = -1
        name_col = -1
        qty_col = -1
        price_col = -1
        fee_col = -1
        
        # 优化表头查找：兼容网页端表头行偏移
        header_row1 = -1
        for i, row in enumerate(table[:5]):  # 只检查前5行，避免无效遍历
            row_str = ' '.join([str(cell) for cell in row if cell]).replace('\xa0', ' ')
            if ("科目编码" in row_str or "编码" in row_str) and ("结算类型" in row_str or "名称" in row_str):
                header_row1 = i
                break
        
        if header_row1 == -1:
            # 降级匹配：只找编码/名称列
            for i, row in enumerate(table[:3]):
                row_str = ' '.join([str(cell) for cell in row if cell]).replace('\xa0', ' ')
                if "科目编码" in row_str or "编码" in row_str:
                    header_row1 = i
                    break
        
        if header_row1 == -1:
            continue
        
        header_row2 = header_row1 + 1
        if header_row2 >= len(table):
            header_row2 = header_row1  # 兼容单行表头
        
        # 匹配列索引：增强容错
        for j, cell in enumerate(table[header_row1]):
            cell_lower = str(cell).lower().replace('\xa0', ' ')
            if any(keyword in cell_lower for keyword in ["科目编码", "编码", "code"]):
                code_col = j
            elif any(keyword in cell_lower for keyword in ["结算类型", "科目名称", "名称", "name"]):
                name_col = j
        
        for j, cell in enumerate(table[header_row2]):
            cell_lower = str(cell).lower().replace('\xa0', ' ')
            if any(keyword in cell_lower for keyword in ["电量", "数量", "kwh", "mwh", "兆瓦时"]):
                qty_col = j
            elif any(keyword in cell_lower for keyword in ["电价", "价格", "单价", "price"]):
                price_col = j
            elif any(keyword in cell_lower for keyword in ["电费", "金额", "费用", "合计", "amount", "元"]):
                fee_col = j
        
        # 解析数据行：增强容错
        for i in range(header_row2 + 1, len(table)):
            row = table[i]
            row_str = ' '.join([str(cell) for cell in row if cell]).replace('\xa0', ' ')
            if any(keyword in row_str for keyword in ["合计", "总计", "小计", "summary", "total", "汇总"]):
                continue
            
            trade_code = ""
            trade_name = None
            
            # 编码列提取：兼容网页端编码格式
            if code_col >= 0 and code_col < len(row):
                trade_code = str(row[code_col]).strip().replace('\xa0', ' ')
                if len(trade_code) == 9:
                    trade_code = "0" + trade_code
                if trade_code in TRADE_CODE_MAP:
                    trade_name = TRADE_CODE_MAP[trade_code]
            
            # 名称列模糊匹配：增强关键词匹配
            if not trade_name and name_col >= 0 and name_col < len(row):
                name_cell = str(row[name_col]).strip().replace('\xa0', ' ')
                # 拆分关键词，增强匹配
                for code, name in TRADE_CODE_MAP.items():
                    name_parts = re.split(r'[()（）、-]', name)
                    if any(part.strip() in name_cell for part in name_parts if part.strip()):
                        trade_name = name
                        break
            
            if not trade_name:
                continue
            
            is_special = trade_name in SPECIAL_TRADES
            if is_special:
                if fee_col >= 0 and fee_col < len(row):
                    fee_val = row[fee_col]
                    trade_data[trade_name]['fee'] = safe_convert_to_numeric(fee_val)
            else:
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
    """提取合计数据 - 增强网页端文本匹配"""
    pdf_text = pdf_text.replace('\xa0', ' ').replace('\r', '\n').strip()
    total_quantity, total_amount = None, None
    lines = pdf_text.split('\n')
    
    for line in lines:
        line_clean = line.replace(' ', '').replace(',', '').replace('，', '').replace('\xa0', '')
        # 增强合计匹配：兼容更多表述
        qty_match = re.search(r'合计电量[:：]([\d\.]+)|总电量[:：]([\d\.]+)|电量合计[:：]([\d\.]+)', line_clean)
        if qty_match:
            # 取第一个非空匹配组
            qty_val = next((g for g in qty_match.groups() if g), None)
            if qty_val:
                total_quantity = safe_convert_to_numeric(qty_val)
        
        fee_match = re.search(r'合计电费[:：]([\d\.]+)|总电费[:：]([\d\.]+)|电费合计[:：]([\d\.]+)|合计金额[:：]([\d\.]+)', line_clean)
        if fee_match:
            fee_val = next((g for g in fee_match.groups() if g), None)
            if fee_val:
                total_amount = safe_convert_to_numeric(fee_val)
    
    return total_quantity, total_amount

def process_single_station(station_name, tables, pdf_text, file_name):
    """处理单个场站 - 保留原有逻辑"""
    station_name, date = extract_station_and_date_v2(pdf_text, file_name, station_name)
    total_quantity, total_amount = extract_total_data_v2(pdf_text)
    trade_data = parse_trade_table_data_v2(tables)
    
    result = [station_name, date, total_quantity, total_amount]
    for trade in REGULAR_TRADES:
        data = trade_data.get(trade, {'quantity': None, 'price': None, 'fee': None})
        result.extend([data['quantity'], data['price'], data['fee']])
    for trade in SPECIAL_TRADES:
        data = trade_data.get(trade, {'fee': None})
        result.append(data['fee'])
    
    return result

def extract_data_from_pdf_v2(file_obj, file_name):
    """从PDF提取数据 - 关键修复：彻底重置文件流"""
    try:
        # 关键修改1：复制文件流到BytesIO，避免网页端文件对象限制
        file_obj.seek(0)
        file_bytes = BytesIO(file_obj.read())
        file_bytes.seek(0)
        
        # 读取PDF文本
        with pdfplumber.open(file_bytes) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
        
        if not all_text or len(all_text.strip()) < 50:
            raise ValueError("PDF为空或文本内容太少")
        
        # 重新重置字节流，提取表格
        file_bytes.seek(0)
        all_tables = extract_data_using_pdfplumber_tables(file_bytes)
        
        if not all_tables:
            st.warning(f"{file_name}: 表格提取失败，使用文本分析模式")
        
        # 拆分双场站数据
        station_data_list = split_double_station_data(all_text, all_tables)
        
        # 处理每个场站
        results = []
        for station_name, tables_segment in station_data_list:
            result = process_single_station(station_name, tables_segment, all_text, file_name)
            results.append(result)
        
        # 关闭字节流，释放资源
        file_bytes.close()
        return results
        
    except Exception as e:
        st.error(f"处理PDF {file_name} 出错: {str(e)}")
        default_result = ["未知场站", None, None, None] + [None] * (len(REGULAR_TRADES) * 3 + len(SPECIAL_TRADES))
        return [default_result]

# ---------------------- Streamlit 应用（关键修改） ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取工具", layout="wide")
    
    st.title("📊 黑龙江日清分结算单数据提取工具（网页适配版）")
    st.markdown("**核心修复：适配网页端运行环境、统一文件流处理、增强字符兼容**")
    st.divider()
    
    # 显示环境信息（调试用）
    with st.expander("🔧 运行环境信息（调试）"):
        st.write(f"Python版本: {sys.version}")
        st.write(f"pdfplumber版本: {pdfplumber.__version__}")
        st.write(f"pandas版本: {pd.__version__}")
    
    # 显示科目信息
    with st.expander("📋 支持的科目列表（含新增）"):
        st.write("**常规科目（电量、电价、电费）：**")
        for trade in REGULAR_TRADES:
            st.write(f"- {trade}")
        st.write("**特殊科目（仅电费）：**")
        for trade in SPECIAL_TRADES:
            st.write(f"- {trade}")
    
    st.subheader("📁 上传文件")
    uploaded_files = st.file_uploader(
        "支持PDF格式，可批量上传（适配依兰协合风电PDF）",
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
                    # 关键修改2：批量处理时每次重置文件流
                    file.seek(0)
                    file_results = extract_data_from_pdf_v2(file, file.name)
                    for result in file_results:
                        all_data.append(result)
                    
                    if len(file_results) == 2:
                        st.success(f"✓ {file.name} 处理成功（识别出2个场站）")
                    elif len(file_results) == 1:
                        st.success(f"✓ {file.name} 处理成功（识别出1个场站）")
                    else:
                        st.warning(f"⚠ {file.name} 处理完成，但未识别到场站数据")
                        
                except Exception as e:
                    st.error(f"✗ {file.name} 处理失败: {str(e)}")
                finally:
                    # 关键修改3：释放文件资源
                    file.close()
            
            progress_bar.empty()
            
            if all_data:
                # 构建结果DataFrame
                result_columns = ['场站名称', '清分日期', '合计电量(兆瓦时)', '合计电费(元)']
                for trade in REGULAR_TRADES:
                    trade_short = trade.replace('省间绿色电力交易', '省间绿电交易')
                    trade_short = trade_short.replace('电网企业代理购电交易', '代理购电交易')
                    trade_short = trade_short.replace('(电能量)', '')
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
                
                # 数据完整性统计
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
                    file_name=f"黑龙江结算数据_网页适配版_{current_time}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.success("✅ 全部处理完成！")
    
    else:
        st.info("👆 请上传PDF文件开始处理（已适配依兰协合风电PDF）")

if __name__ == "__main__":
    # 关键修改4：设置环境变量，避免Streamlit网页端编码问题
    os.environ["PYTHONIOENCODING"] = "utf-8"
    main()
