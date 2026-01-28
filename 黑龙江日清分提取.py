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

# ---------------------- 核心提取函数（完全重写） ----------------------
def safe_convert_to_numeric(value):
    """安全转换为数值"""
    if value is None or pd.isna(value) or value == '':
        return None
    try:
        if isinstance(value, str):
            # 移除千分位逗号和其他非数字字符
            cleaned = re.sub(r'[^\d.-]', '', value)
            if cleaned and cleaned not in ['-', '.', '']:
                return float(cleaned)
        return float(value)
    except (ValueError, TypeError):
        return None

def extract_station_and_date(pdf_text):
    """提取场站名称和日期 - 改进版"""
    lines = pdf_text.split('\n')
    
    station_name = "未知场站"
    date = None
    
    # 提取场站名称
    for i, line in enumerate(lines):
        line_clean = line.strip()
        
        # 方法1: 从公司名称提取
        if "公司名称" in line_clean:
            match = re.search(r'公司名称[:：]\s*(.+?有限公司)', line_clean)
            if match:
                base_company = match.group(1).strip()
                
                # 查找机组信息判断A/B场站
                station_type = "未知场站"
                for j in range(max(0, i-3), min(len(lines), i+10)):
                    if "机组" in lines[j]:
                        if "B" in lines[j].upper() or "2" in lines[j] or "二" in lines[j]:
                            station_type = "双发B风电场"
                        elif "A" in lines[j].upper() or "1" in lines[j] or "一" in lines[j]:
                            station_type = "双发A风电场"
                        break
                
                station_name = f"{base_company}（{station_type}）"
                break
        
        # 方法2: 从包含风电场的行提取
        if "风电场" in line_clean and station_name == "未知场站":
            match = re.search(r'([^，。！？]+风电场)', line_clean)
            if match:
                station_name = match.group(1).strip()
    
    # 提取日期
    for line in lines:
        patterns = [
            r'清分日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2})',
            r'日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2})',
            r'(\d{4}年\d{1,2}月\d{1,2}日)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, line)
            if match:
                date_str = match.group(1)
                date_str = date_str.replace('年', '-').replace('月', '-').replace('日', '')
                date = date_str
                break
        if date:
            break
    
    return station_name, date

def extract_data_using_pdfplumber_tables(file_obj):
    """使用pdfplumber的表格提取功能 - 这是关键修复"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            # 尝试提取所有页面的表格
            all_tables = []
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        # 清理表格数据
                        cleaned_table = []
                        for row in table:
                            cleaned_row = [cell.strip() if cell else "" for cell in row]
                            cleaned_table.append(cleaned_row)
                        all_tables.append(cleaned_table)
            
            return all_tables
    except Exception as e:
        st.error(f"表格提取失败: {e}")
        return []

def parse_trade_table_data(tables):
    """解析交易表格数据"""
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
            
        # 查找表头行，识别列索引
        header_row = -1
        code_col = -1
        name_col = -1
        qty_col = -1
        price_col = -1
        fee_col = -1
        
        for i, row in enumerate(table):
            row_str = ' '.join([str(cell) for cell in row if cell])
            if "科目编码" in row_str or "编码" in row_str:
                header_row = i
                # 确定各列位置
                for j, cell in enumerate(row):
                    if cell and ("科目编码" in str(cell) or "编码" in str(cell)):
                        code_col = j
                    elif cell and ("科目名称" in str(cell) or "名称" in str(cell)):
                        name_col = j
                    elif cell and ("电量" in str(cell)):
                        qty_col = j
                    elif cell and ("电价" in str(cell)):
                        price_col = j
                    elif cell and ("电费" in str(cell)):
                        fee_col = j
                break
        
        if header_row == -1:
            continue
            
        # 解析数据行
        for i in range(header_row + 1, len(table)):
            row = table[i]
            if len(row) <= max(code_col, name_col, qty_col, price_col, fee_col):
                continue
                
            # 提取科目编码
            trade_code = str(row[code_col]) if code_col < len(row) else ""
            trade_name = None
            
            # 通过编码获取科目名称
            if trade_code in TRADE_CODE_MAP:
                trade_name = TRADE_CODE_MAP[trade_code]
            else:
                # 尝试通过名称匹配
                name_cell = str(row[name_col]) if name_col < len(row) else ""
                for code, name in TRADE_CODE_MAP.items():
                    if name in name_cell:
                        trade_name = name
                        break
            
            if not trade_name:
                continue
                
            # 提取数据
            is_special = trade_name in SPECIAL_TRADES
            
            if is_special:
                # 特殊科目只有电费
                fee_val = str(row[fee_col]) if fee_col < len(row) else ""
                trade_data[trade_name]['fee'] = safe_convert_to_numeric(fee_val)
            else:
                # 常规科目有电量、电价、电费
                qty_val = str(row[qty_col]) if qty_col < len(row) else ""
                price_val = str(row[price_col]) if price_col < len(row) else ""
                fee_val = str(row[fee_col]) if fee_col < len(row) else ""
                
                trade_data[trade_name]['quantity'] = safe_convert_to_numeric(qty_val)
                trade_data[trade_name]['price'] = safe_convert_to_numeric(price_val)
                trade_data[trade_name]['fee'] = safe_convert_to_numeric(fee_val)
    
    return trade_data

def extract_data_using_text_analysis(pdf_text):
    """备用方法：通过文本分析提取数据"""
    trade_data = {}
    
    # 初始化所有科目
    for trade in ALL_TRADES:
        if trade in SPECIAL_TRADES:
            trade_data[trade] = {'fee': None}
        else:
            trade_data[trade] = {'quantity': None, 'price': None, 'fee': None}
    
    lines = pdf_text.split('\n')
    
    # 查找交易数据区域
    data_start = -1
    for i, line in enumerate(lines):
        if "科目编码" in line and "科目名称" in line:
            data_start = i
            break
    
    if data_start == -1:
        return trade_data
    
    # 解析数据行
    for i in range(data_start + 1, len(lines)):
        line = lines[i].strip()
        if not line or "合计" in line or "总计" in line:
            continue
            
        # 查找科目编码
        code_match = re.search(r'\b(\d{9})\b', line)
        if not code_match:
            continue
            
        code = code_match.group(1)
        if code not in TRADE_CODE_MAP:
            continue
            
        trade_name = TRADE_CODE_MAP[code]
        is_special = trade_name in SPECIAL_TRADES
        
        # 提取数字（跳过科目编码）
        numbers = []
        line_parts = line.split()
        
        # 找到编码位置，从后面开始提取数字
        code_index = -1
        for idx, part in enumerate(line_parts):
            if code in part:
                code_index = idx
                break
        
        if code_index >= 0:
            # 提取编码后面的数字
            for j in range(code_index + 1, len(line_parts)):
                part = line_parts[j]
                num_match = re.search(r'-?[\d,]+\.?\d*', part)
                if num_match:
                    numbers.append(safe_convert_to_numeric(num_match.group()))
        
        # 分配数据
        if is_special:
            if numbers:
                trade_data[trade_name]['fee'] = numbers[0]
        else:
            if len(numbers) >= 3:
                trade_data[trade_name]['quantity'] = numbers[0]
                trade_data[trade_name]['price'] = numbers[1]
                trade_data[trade_name]['fee'] = numbers[2]
            elif len(numbers) == 2:
                trade_data[trade_name]['quantity'] = numbers[0]
                trade_data[trade_name]['fee'] = numbers[1]
            elif len(numbers) == 1:
                trade_data[trade_name]['fee'] = numbers[0]
    
    return trade_data

def extract_total_data(pdf_text):
    """提取合计数据"""
    total_quantity, total_amount = None, None
    
    lines = pdf_text.split('\n')
    
    for line in lines:
        line_clean = line.replace(' ', '')
        
        # 查找合计电量
        if "合计电量" in line_clean:
            match = re.search(r'合计电量[^\d]*([\d,]+\.?\d*)', line_clean)
            if match:
                total_quantity = safe_convert_to_numeric(match.group(1))
        
        # 查找合计电费
        if "合计电费" in line_clean:
            match = re.search(r'合计电费[^\d]*([\d,]+\.?\d*)', line_clean)
            if match:
                total_amount = safe_convert_to_numeric(match.group(1))
    
    return total_quantity, total_amount

def extract_data_from_pdf(file_obj, file_name):
    """从PDF提取数据 - 综合方法"""
    try:
        # 首先提取文本内容用于基本信息提取
        with pdfplumber.open(file_obj) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
        
        if not all_text or len(all_text.strip()) < 50:
            raise ValueError("PDF为空或文本内容太少")
        
        # 提取基本信息和合计数据
        station_name, date = extract_station_and_date(all_text)
        total_quantity, total_amount = extract_total_data(all_text)
        
        # 从文件名提取日期（备用）
        if not date:
            date_match = re.search(r'(\d{4}-\d{2}-\d{2})', file_name)
            if date_match:
                date = date_match.group(1)
        
        # 重置文件指针，重新读取用于表格提取
        file_obj.seek(0)
        
        # 方法1: 使用表格提取（优先）
        tables = extract_data_using_pdfplumber_tables(file_obj)
        trade_data = parse_trade_table_data(tables)
        
        # 方法2: 如果表格提取失败，使用文本分析
        if not any(trade_data[trade].get('fee') for trade in ALL_TRADES):
            trade_data = extract_data_using_text_analysis(all_text)
        
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
        
    except Exception as e:
        st.error(f"处理PDF {file_name} 出错: {str(e)}")
        # 返回正确长度的空数据
        return ["未知场站", None, None, None] + [None] * (len(REGULAR_TRADES) * 3 + len(SPECIAL_TRADES))

# ---------------------- Streamlit 应用 ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取工具", layout="wide")
    
    st.title("📊 黑龙江日清分结算单数据提取工具（终极修复版）")
    st.markdown("**修复重点：表格结构识别、双发B风电场识别、科目数据提取**")
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
        "支持PDF格式，可批量上传",
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
                    data = extract_data_from_pdf(file, file.name)
                    if data[1] is not None:  # 有日期视为成功
                        all_data.append(data)
                        st.success(f"✓ {file.name} 处理成功")
                    else:
                        st.warning(f"⚠ {file.name} 缺少日期信息")
                except Exception as e:
                    st.error(f"✗ {file.name} 处理失败: {str(e)}")
            
            progress_bar.empty()
            
            if all_data:
                # 构建结果DataFrame
                result_columns = ['场站名称', '清分日期', '合计电量(兆瓦时)', '合计电费(元)']
                
                for trade in REGULAR_TRADES:
                    trade_short = trade.replace('省间绿色电力交易', '省间绿电交易')
                    result_columns.extend([f'{trade_short}_电量', f'{trade_short}_电价', f'{trade_short}_电费'])
                
                for trade in SPECIAL_TRADES:
                    result_columns.append(f'{trade}_电费')
                
                result_df = pd.DataFrame(all_data, columns=result_columns)
                
                # 显示结果
                st.subheader("📈 提取结果")
                st.dataframe(result_df, use_container_width=True)
                
                # 统计信息
                st.info(f"**统计信息：** 共处理 {len(all_data)} 个文件，涉及 {result_df['场站名称'].nunique()} 个场站")
                
                # 检查双发B风电场是否存在
                has_b_station = any('双发B' in str(name) for name in result_df['场站名称'])
                if not has_b_station:
                    st.warning("⚠️ 未检测到双发B风电场数据，请检查PDF中机组信息")
                
                # 检查数据完整性
                data_columns = result_columns[4:]  # 跳过前4列基本信息
                non_null_count = result_df[data_columns].notna().sum().sum()
                total_cells = len(result_df) * len(data_columns)
                st.info(f"**数据完整性：** {non_null_count}/{total_cells} 个数据单元格有值 ({non_null_count/total_cells*100:.1f}%)")
                
                # 下载功能
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                output.seek(0)
                
                st.download_button(
                    label="📥 下载Excel文件",
                    data=output,
                    file_name=f"黑龙江结算数据_终极修复版_{current_time}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.success("✅ 处理完成！")
    
    else:
        st.info("👆 请上传PDF文件开始处理")

if __name__ == "__main__":
    main()
