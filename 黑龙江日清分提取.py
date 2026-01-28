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
# 科目编码到名称的映射（完整列表）
TRADE_CODE_MAP = {
    # 常规交易科目（9位数字编码）
    "101010101": ("优先发电交易", False),
    "101020101": ("电网企业代理购电交易", False), 
    "101020301": ("省内电力直接交易", False),
    "101040322": ("送上海省间绿色电力交易(电能量)", False),
    "102020101": ("送辽宁交易", False),
    "102020301": ("送华北交易", False), 
    "102010101": ("送山东交易", False),
    "102010201": ("送浙江交易", False),
    "202030001": ("送江苏省间绿色电力交易（电能量）", False),
    "202030002": ("送浙江省间绿色电力交易（电能量）", False),
    "101080101": ("省内现货日前交易", False),
    "101080201": ("省内现货实时交易", False),
    "101080301": ("省间现货日前交易", False),
    "101080401": ("省间现货日内交易", False),
    # 特殊科目（只有电费）
    "201010101": ("中长期合约阻塞费用", True),
    "201020101": ("省间省内价差费用", True)
}

# 所有科目名称列表（用于列顺序）
ALL_TRADES = [
    "优先发电交易",
    "电网企业代理购电交易", 
    "省内电力直接交易",
    "送上海省间绿色电力交易(电能量)",
    "送辽宁交易",
    "送华北交易", 
    "送山东交易",
    "送浙江交易",
    "送江苏省间绿色电力交易（电能量）",
    "送浙江省间绿色电力交易（电能量）",
    "省内现货日前交易",
    "省内现货实时交易",
    "省间现货日前交易",
    "省间现货日内交易",
    "中长期合约阻塞费用",
    "省间省内价差费用"
]

# 特殊科目列表（只有电费）
SPECIAL_TRADES = ["中长期合约阻塞费用", "省间省内价差费用"]

# ---------------------- 核心提取函数 ----------------------
def safe_convert_to_numeric(value):
    """安全转换为数值"""
    if value is None or pd.isna(value) or value == '':
        return None
    try:
        if isinstance(value, str):
            # 移除千分位逗号和其他非数字字符（保留负号、小数点和数字）
            cleaned = re.sub(r'[^\d.-]', '', value)
            if cleaned and cleaned != '-' and cleaned != '.':
                return float(cleaned)
        return float(value)
    except (ValueError, TypeError):
        return None

def extract_station_name(pdf_lines):
    """通用提取场站名称，支持多公司和场站拆分"""
    for i, line in enumerate(pdf_lines):
        line_clean = line.strip()
        
        # 匹配公司名称模式
        if "公司名称" in line_clean or "有限公司" in line_clean:
            # 提取公司名称
            company_match = re.search(r'公司名称[:：]\s*(.+?有限公司)', line_clean)
            if not company_match:
                # 尝试其他格式
                company_match = re.search(r'[:：]\s*(.+?有限公司)', line_clean)
            
            if company_match:
                company_name = company_match.group(1).strip()
                
                # 检查是否有场站信息（双发A/B风电场等）
                station_info = ""
                
                # 查找"机组"信息
                for j in range(i, min(i+5, len(pdf_lines))):
                    next_line = pdf_lines[j].strip()
                    if "机组" in next_line:
                        # 提取机组信息
                        if "B" in next_line.upper() or "双发B" in next_line:
                            station_info = "（双发B风电场）"
                        elif "A" in next_line.upper() or "双发A" in next_line:
                            station_info = "（双发A风电场）"
                        elif "风电场" in next_line:
                            # 提取具体的风电场名称
                            station_match = re.search(r'风电场[:：]\s*([^\s]+)', next_line)
                            if station_match:
                                station_info = f"（{station_match.group(1)}）"
                        break
                
                return f"{company_name}{station_info}"
    
    # 如果没找到，从包含"风电场"或"光伏电站"的行中提取
    for line in pdf_lines:
        if "风电场" in line or "光伏电站" in line:
            parts = line.split()
            for part in parts:
                if "风电场" in part or "光伏电站" in part:
                    return part.strip()
    
    return "未知场站"

def extract_date_from_pdf(pdf_lines):
    """提取清分日期"""
    for line in pdf_lines:
        # 尝试多种日期模式
        patterns = [
            r'清分日期[:：]\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})',
            r'日期[:：]\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})',
            r'(\d{4}年\d{1,2}月\d{1,2}日)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, line)
            if match:
                date_str = match.group(1)
                # 统一格式化为YYYY-MM-DD
                date_str = date_str.replace('年', '-').replace('月', '-').replace('日', '')
                date_str = date_str.replace('/', '-')
                return date_str
    
    return None

def extract_total_data(pdf_text):
    """提取合计电量、合计电费"""
    total_quantity = None
    total_amount = None
    
    lines = pdf_text.split('\n')
    
    for i, line in enumerate(lines):
        line_clean = line.replace(' ', '')
        
        # 查找"合计电量"和"合计电费"
        if "合计电量" in line_clean or "合计电费" in line_clean:
            # 尝试提取合计电量
            if "合计电量" in line_clean:
                qty_match = re.search(r'合计电量[^\d]*([\d,]+\.?\d*)', line_clean)
                if qty_match:
                    total_quantity = safe_convert_to_numeric(qty_match.group(1))
            
            # 尝试提取合计电费
            if "合计电费" in line_clean:
                amount_match = re.search(r'合计电费[^\d]*([\d,]+\.?\d*)', line_clean)
                if amount_match:
                    total_amount = safe_convert_to_numeric(amount_match.group(1))
            
            # 如果当前行没找到，检查下一行
            if (total_quantity is None or total_amount is None) and i+1 < len(lines):
                next_line = lines[i+1].replace(' ', '')
                numbers = re.findall(r'[\d,]+\.?\d*', next_line)
                if numbers:
                    if total_quantity is None and len(numbers) > 0:
                        total_quantity = safe_convert_to_numeric(numbers[0])
                    if total_amount is None and len(numbers) > 1:
                        total_amount = safe_convert_to_numeric(numbers[1])
    
    return total_quantity, total_amount

def extract_trade_data_from_pdf(pdf_text):
    """
    从PDF文本中提取所有交易数据
    关键改进：避免将科目编码误认为电量数据
    """
    # 初始化结果字典
    trade_data = {}
    for trade_name in ALL_TRADES:
        is_special = trade_name in SPECIAL_TRADES
        trade_data[trade_name] = {
            'quantity': None,
            'price': None,
            'fee': None,
            'is_special': is_special
        }
    
    lines = pdf_text.split('\n')
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        # 跳过空行
        if not line:
            i += 1
            continue
        
        # 检查是否包含科目编码
        code_match = re.search(r'\b(\d{9})\b', line)
        trade_name = None
        is_special = False
        
        if code_match:
            code = code_match.group(1)
            if code in TRADE_CODE_MAP:
                trade_name, is_special = TRADE_CODE_MAP[code]
        
        # 如果没有找到编码，尝试通过科目名称查找
        if not trade_name:
            for code, (name, special) in TRADE_CODE_MAP.items():
                # 检查科目名称是否在行中
                if name in line:
                    trade_name = name
                    is_special = special
                    break
        
        # 如果找到了科目，提取数据
        if trade_name:
            # 对于特殊科目（只有电费）
            if is_special:
                # 提取电费（跳过科目编码）
                fee_match = re.search(r'(?<!\d{9})\b(-?\d[\d,]*\.?\d*)\b', line[code_match.end() if code_match else 0:] if code_match else line)
                if not fee_match and i+1 < len(lines):
                    # 尝试下一行
                    next_line = lines[i+1].strip()
                    fee_match = re.search(r'\b(-?\d[\d,]*\.?\d*)\b', next_line)
                
                if fee_match:
                    trade_data[trade_name]['fee'] = safe_convert_to_numeric(fee_match.group(1))
            
            # 对于常规科目（电量、电价、电费）
            else:
                # 从行中提取所有数字（跳过科目编码）
                line_for_numbers = line
                if code_match:
                    # 移除科目编码部分
                    line_for_numbers = line[code_match.end():]
                
                numbers = re.findall(r'\b(-?\d[\d,]*\.?\d*)\b', line_for_numbers)
                
                # 如果当前行数字不够，检查下一行
                if len(numbers) < 3 and i+1 < len(lines):
                    next_line = lines[i+1].strip()
                    next_numbers = re.findall(r'\b(-?\d[\d,]*\.?\d*)\b', next_line)
                    numbers.extend(next_numbers)
                
                # 分配数据
                if len(numbers) >= 3:
                    trade_data[trade_name]['quantity'] = safe_convert_to_numeric(numbers[0])
                    trade_data[trade_name]['price'] = safe_convert_to_numeric(numbers[1])
                    trade_data[trade_name]['fee'] = safe_convert_to_numeric(numbers[2])
                elif len(numbers) == 2:
                    trade_data[trade_name]['quantity'] = safe_convert_to_numeric(numbers[0])
                    trade_data[trade_name]['fee'] = safe_convert_to_numeric(numbers[1])
                elif len(numbers) == 1:
                    trade_data[trade_name]['fee'] = safe_convert_to_numeric(numbers[0])
        
        i += 1
    
    return trade_data

def extract_data_from_pdf(file_obj, file_name):
    """从PDF提取数据 - 通用版本"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
        
        if not all_text or len(all_text.strip()) < 50:
            raise ValueError("PDF为空或文本内容太少，可能是扫描件")
        
        pdf_lines = [line.strip() for line in all_text.split('\n') if line.strip()]
        
        # 提取基本信息
        station_name = extract_station_name(pdf_lines)
        date = extract_date_from_pdf(pdf_lines)
        total_quantity, total_amount = extract_total_data(all_text)
        
        # 从文件名提取日期（备用）
        if not date:
            date_match = re.search(r'(\d{4}-\d{2}-\d{2})', file_name)
            if date_match:
                date = date_match.group(1)
        
        # 提取所有交易数据
        trade_data = extract_trade_data_from_pdf(all_text)
        
        # 构建结果列表（按ALL_TRADES顺序）
        result = [station_name, date, total_quantity, total_amount]
        
        for trade in ALL_TRADES:
            data = trade_data.get(trade, {'quantity': None, 'price': None, 'fee': None, 'is_special': False})
            is_special = data['is_special']
            
            if is_special:
                # 特殊科目：只有电费
                result.append(data['fee'])
            else:
                # 常规科目：电量、电价、电费
                result.extend([data['quantity'], data['price'], data['fee']])
        
        return result
        
    except Exception as e:
        st.error(f"处理PDF {file_name} 出错: {str(e)}")
        # 返回正确长度的空数据
        regular_count = len([t for t in ALL_TRADES if t not in SPECIAL_TRADES])
        special_count = len(SPECIAL_TRADES)
        total_columns = 4 + (regular_count * 3) + special_count
        return ["未知场站", None, None, None] + [None] * (total_columns - 4)

def calculate_summary_row(data_df):
    """计算汇总行"""
    if data_df.empty:
        return pd.DataFrame()
    
    summary_row = {'场站名称': '总计', '清分日期': ''}
    
    for col in data_df.columns:
        if col in ['场站名称', '清分日期']:
            continue
        
        # 电价列计算平均值
        if '电价' in col and '电费' not in col:
            valid_vals = data_df[col].dropna()
            if not valid_vals.empty:
                summary_row[col] = round(valid_vals.mean(), 4)
        else:
            # 其他列计算总和
            valid_vals = data_df[col].dropna()
            if not valid_vals.empty:
                summary_row[col] = valid_vals.sum()
    
    return pd.DataFrame([summary_row])

def to_excel_bytes(df, report_df):
    """转换为Excel字节流"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='结算数据明细', index=False)
        report_df.to_excel(writer, sheet_name='处理报告', index=False)
    output.seek(0)
    return output

# ---------------------- Streamlit 应用 ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取工具", layout="wide")
    
    st.title("📊 黑龙江日清分结算单数据提取工具（通用版）")
    st.markdown("**修复问题：科目编码识别、场站拆分、数据错位**")
    st.divider()
    
    # 显示科目信息
    with st.expander("📋 支持的科目列表"):
        regular_trades = [t for t in ALL_TRADES if t not in SPECIAL_TRADES]
        special_trades = SPECIAL_TRADES
        
        st.write("**常规科目（电量、电价、电费）：**")
        for trade in regular_trades:
            st.write(f"- {trade}")
        
        st.write("**特殊科目（仅电费）：**")
        for trade in special_trades:
            st.write(f"- {trade}")
    
    st.subheader("📁 上传文件")
    uploaded_files = st.file_uploader(
        "支持PDF格式，可批量上传",
        type=['pdf'],
        accept_multiple_files=True,
        help="请上传黑龙江日清分结算单PDF文件"
    )
    
    if uploaded_files:
        if st.button("🚀 开始处理", type="primary"):
            st.divider()
            st.subheader("⚙️ 处理进度")
            
            all_data = []
            total_files = len(uploaded_files)
            processed_files = 0
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, file in enumerate(uploaded_files):
                file_name = file.name
                status_text.text(f"正在处理：{file_name} ({idx+1}/{total_files})")
                
                try:
                    data = extract_data_from_pdf(file, file_name)
                    
                    # 检查数据有效性
                    if data[1] is not None:  # 有日期
                        all_data.append(data)
                        processed_files += 1
                        st.success(f"✓ {file_name} 处理成功")
                    else:
                        st.warning(f"⚠ {file_name} 缺少日期信息")
                    
                except Exception as e:
                    st.error(f"✗ {file_name} 处理失败: {str(e)[:100]}")
                
                progress_bar.progress((idx + 1) / total_files)
            
            progress_bar.empty()
            status_text.text("处理完成！")
            
            if all_data:
                st.divider()
                st.subheader("📈 提取结果")
                
                # 构建结果列
                result_columns = ['场站名称', '清分日期', '合计电量(兆瓦时)', '合计电费(元)']
                
                # 添加常规科目列
                regular_trades = [t for t in ALL_TRADES if t not in SPECIAL_TRADES]
                for trade in regular_trades:
                    trade_short = trade.replace('（电能量）', '').replace('(电能量)', '').replace('省间绿色电力交易', '省间绿电交易')
                    result_columns.extend([
                        f'{trade_short}_电量',
                        f'{trade_short}_电价',
                        f'{trade_short}_电费'
                    ])
                
                # 添加特殊科目列
                for trade in SPECIAL_TRADES:
                    result_columns.append(f'{trade}_电费')
                
                # 创建结果DataFrame
                result_df = pd.DataFrame(all_data, columns=result_columns)
                
                # 转换数值列
                for col in result_columns[2:]:
                    result_df[col] = pd.to_numeric(result_df[col], errors='coerce')
                
                # 添加汇总行
                if len(result_df) > 0:
                    summary_df = calculate_summary_row(result_df)
                    result_df = pd.concat([result_df, summary_df], ignore_index=True)
                
                # 显示结果
                tab1, tab2 = st.tabs(["结算数据明细", "处理报告"])
                
                with tab1:
                    st.dataframe(result_df, use_container_width=True)
                    st.info(f"**数据统计：** 共 {len(result_df)-1} 行数据，{len(result_columns)-4} 个科目")
                
                with tab2:
                    report_data = {
                        '统计项': ['上传文件数', '成功处理数', '失败数', '成功率', '提取场站数'],
                        '数值': [
                            total_files,
                            processed_files,
                            total_files - processed_files,
                            f"{(processed_files/total_files)*100:.1f}%" if total_files > 0 else "0%",
                            result_df['场站名称'].nunique() - 1  # 减去总计行
                        ]
                    }
                    report_df = pd.DataFrame(report_data)
                    st.dataframe(report_df, use_container_width=True)
                
                # 下载功能
                st.divider()
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                excel_bytes = to_excel_bytes(result_df, report_df)
                
                st.download_button(
                    label="📥 下载Excel文件",
                    data=excel_bytes,
                    file_name=f"黑龙江结算数据_通用版_{current_time}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.success(f"✅ 处理完成！成功提取 {processed_files}/{total_files} 个文件")
                st.info("""
                **关键改进：**
                1. **通用场站识别**：不锁定特定公司，支持多公司和场站拆分
                2. **科目编码识别**：使用正则表达式精确识别9位科目编码，避免误认为数据
                3. **数据防错**：特殊科目只提取电费，跳过科目编码
                4. **灵活匹配**：支持编码和名称双重匹配，提高提取成功率
                """)
                
            else:
                st.error("⚠️ 未提取到有效数据！")
                st.markdown("""
                **调试建议：**
                1. 确认PDF包含可复制文本
                2. 检查文件是否为标准黑龙江日清分格式
                3. 联系技术支持获取帮助
                """)
    
    else:
        st.info("👆 请上传PDF文件开始处理")

if __name__ == "__main__":
    main()
