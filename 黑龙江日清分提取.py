import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心配置优化 ----------------------
# 优化目标科目列表，适配黑龙江日清分格式
NORMAL_TRADES = [
    '优先发电交易',
    '电网企业代理购电交易',
    '省内电力直接交易',
    '送上海省间绿色电力交易(电能量)',
    '送辽宁交易',
    '送华北交易',
    '送山东交易',
    '送浙江交易',
    '送江苏省间绿色电力交易(电能量)',
    '送浙江省间绿色电力交易(电能量)',
    '省内现货日前交易',
    '省内现货实时交易',
    '省间现货日前交易',
    '省间现货日内交易'
]

SPECIAL_TRADES = [
    '中长期合约阻塞费用',
    '省间省内价差费用'
]

# ---------------------- 核心提取函数优化 ----------------------
def extract_station_name(pdf_lines):
    """从PDF文本中提取场站名称"""
    for line in pdf_lines:
        # 尝试多种匹配模式
        patterns = [
            r'公司名称[:：]\s*([^\n\r]+?)(?:公司|风电场|光伏电站|电站)',
            r'场站名称[:：]\s*([^\n\r]+)',
            r'([^\n\r]+?风电场)',
            r'([^\n\r]+?光伏电站)',
            r'([^\n\r]+?电站)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, line)
            if match:
                station_name = match.group(1).strip()
                # 清理常见后缀
                station_name = re.sub(r'(有限公司|公司|有限责任公司|责任公司)$', '', station_name)
                return station_name
    
    # 如果没找到，尝试从包含"依兰县"的行中提取
    for line in pdf_lines:
        if '依兰县' in line:
            parts = line.split()
            for part in parts:
                if '依兰县' in part and '有限公司' in part:
                    return part.strip()
    
    return "未知场站"

def extract_date_from_pdf(pdf_lines):
    """提取清分日期"""
    date_patterns = [
        r'清分日期[:：]\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})',
        r'日期[:：]\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})',
        r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})\s*日清分',
        r'(\d{4}年\d{1,2}月\d{1,2}日)',
    ]
    
    for line in pdf_lines:
        for pattern in date_patterns:
            match = re.search(pattern, line)
            if match:
                date_str = match.group(1)
                # 统一格式
                date_str = date_str.replace('年', '-').replace('月', '-').replace('日', '')
                date_str = re.sub(r'[/]', '-', date_str)
                return date_str
    
    return None

def extract_total_data(pdf_lines):
    """提取合计电量、电费"""
    total_quantity = None
    total_amount = None
    
    for i, line in enumerate(pdf_lines):
        if '合计' in line or '总计' in line or '合计电量' in line or '合计电费' in line:
            # 使用正则表达式提取数字
            # 查找合计电量
            qty_match = re.search(r'合计电量[^\d]*([\d,]+\.?\d*)', line)
            if not qty_match:
                qty_match = re.search(r'电量合计[^\d]*([\d,]+\.?\d*)', line)
            
            # 查找合计电费
            amount_match = re.search(r'合计电费[^\d]*([\d,]+\.?\d*)', line)
            if not amount_match:
                amount_match = re.search(r'电费合计[^\d]*([\d,]+\.?\d*)', line)
            
            if qty_match:
                total_quantity = float(qty_match.group(1).replace(',', ''))
            if amount_match:
                total_amount = float(amount_match.group(1).replace(',', ''))
    
    return total_quantity, total_amount

def extract_trade_data_from_pdf(pdf_text, trade_name):
    """从PDF文本中提取特定交易的数据"""
    lines = pdf_text.split('\n')
    
    for i, line in enumerate(lines):
        if trade_name in line:
            # 提取该行中的数字
            numbers = re.findall(r'[-]?[\d,]+\.?\d*', line)
            if len(numbers) >= 3:
                try:
                    # 通常格式为：编码 名称 电量 电价 电费
                    quantity = float(numbers[0].replace(',', '')) if len(numbers) > 0 else None
                    price = float(numbers[1].replace(',', '')) if len(numbers) > 1 else None
                    fee = float(numbers[2].replace(',', '')) if len(numbers) > 2 else None
                    return quantity, price, fee
                except:
                    pass
            
            # 如果当前行没有完整数据，尝试检查下一行
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                next_numbers = re.findall(r'[-]?[\d,]+\.?\d*', next_line)
                if len(next_numbers) >= 3:
                    try:
                        quantity = float(next_numbers[0].replace(',', '')) if len(next_numbers) > 0 else None
                        price = float(next_numbers[1].replace(',', '')) if len(next_numbers) > 1 else None
                        fee = float(next_numbers[2].replace(',', '')) if len(next_numbers) > 2 else None
                        return quantity, price, fee
                    except:
                        pass
    
    return None, None, None

def extract_trade_data_from_special(pdf_text, trade_name):
    """提取特殊交易数据（只有电费）"""
    lines = pdf_text.split('\n')
    
    for i, line in enumerate(lines):
        if trade_name in line:
            # 提取电费金额
            numbers = re.findall(r'[-]?[\d,]+\.?\d*', line)
            if numbers:
                try:
                    fee = float(numbers[0].replace(',', ''))
                    return fee
                except:
                    pass
    
    return None

def extract_data_from_pdf(file_obj, file_name):
    """从PDF提取数据 - 核心优化版本"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
            
            if not all_text or len(all_text.strip()) < 50:
                raise ValueError("PDF为空或文本内容太少，可能是扫描件")
        
        # 调试：显示提取的文本前500字符
        st.session_state['debug_text'] = all_text[:1000]  # 保存前1000字符用于调试
        
        # 按行分割
        pdf_lines = all_text.split('\n')
        
        # 提取基本信息
        station_name = extract_station_name(pdf_lines)
        date = extract_date_from_pdf(pdf_lines)
        total_quantity, total_amount = extract_total_data(pdf_lines)
        
        # 如果日期从文件名中提取
        if not date:
            date_match = re.search(r'(\d{4}-\d{2}-\d{2})', file_name)
            if date_match:
                date = date_match.group(1)
        
        # 提取常规交易数据
        normal_data = []
        for trade in NORMAL_TRADES:
            quantity, price, fee = extract_trade_data_from_pdf(all_text, trade)
            normal_data.extend([quantity, price, fee])
        
        # 提取特殊交易数据
        special_data = []
        for trade in SPECIAL_TRADES:
            fee = extract_trade_data_from_special(all_text, trade)
            special_data.append(fee)
        
        return [station_name, date, total_quantity, total_amount] + normal_data + special_data
        
    except Exception as e:
        st.error(f"处理PDF {file_name} 出错: {str(e)[:200]}")
        return ["未知场站", None, None, None] + [None] * (len(NORMAL_TRADES) * 3 + len(SPECIAL_TRADES))

def extract_data_from_excel(file_obj, file_name):
    """从Excel提取数据（简化版，实际需要根据具体格式调整）"""
    try:
        df = pd.read_excel(file_obj, dtype=str, header=None)
        
        # 尝试提取场站名称
        station_name = "未知场站"
        for i in range(min(10, len(df))):
            for j in range(min(5, len(df.columns))):
                cell_val = str(df.iat[i, j])
                if '风电场' in cell_val or '光伏电站' in cell_val:
                    station_name = cell_val.strip()
                    break
        
        # 提取日期
        date = None
        date_pattern = r'\d{4}-\d{2}-\d{2}'
        date_match = re.search(date_pattern, file_name)
        if date_match:
            date = date_match.group(0)
        
        return [station_name, date, None, None] + [None] * (len(NORMAL_TRADES) * 3 + len(SPECIAL_TRADES))
        
    except Exception as e:
        st.error(f"处理Excel {file_name} 出错: {str(e)}")
        return ["未知场站", None, None, None] + [None] * (len(NORMAL_TRADES) * 3 + len(SPECIAL_TRADES))

def calculate_summary_row(data_df):
    """计算汇总行"""
    summary_row = {'场站名称': '总计', '清分日期': ''}
    
    for col in data_df.columns:
        if col in ['场站名称', '清分日期']:
            continue
        
        if '电价' in col:
            # 计算平均电价
            valid_vals = data_df[col].dropna()
            if not valid_vals.empty:
                summary_row[col] = round(valid_vals.mean(), 4)
        else:
            # 计算总和
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

# ---------------------- 增强的Streamlit界面 ----------------------
def main():
    st.set_page_config(page_title="黑龙江日清分数据提取工具", layout="wide")
    
    st.title("📊 黑龙江日清分结算单数据提取工具")
    st.divider()
    
    # 文件上传区域
    st.subheader("📁 上传文件")
    uploaded_files = st.file_uploader(
        "支持PDF/Excel格式，可批量上传",
        type=['pdf', 'xlsx'],
        accept_multiple_files=True,
        help="请上传黑龙江日清分结算单文件，支持PDF和Excel格式"
    )
    
    # 调试选项
    with st.expander("🔧 调试选项（遇到问题时启用）"):
        show_debug = st.checkbox("显示调试信息", value=False)
        debug_file_index = st.number_input("调试文件索引", min_value=0, value=0, 
                                          help="选择要调试的文件序号（从0开始）")
    
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
                    if file_name.lower().endswith('.pdf'):
                        data = extract_data_from_pdf(file, file_name)
                    else:
                        data = extract_data_from_excel(file, file_name)
                    
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
            
            # 显示处理结果
            if all_data:
                st.divider()
                st.subheader("📈 提取结果")
                
                # 构建结果列
                result_columns = ['场站名称', '清分日期', '合计电量(兆瓦时)', '合计电费(元)']
                
                for trade in NORMAL_TRADES:
                    trade_short = trade.replace('（电能量）', '').replace('(电能量)', '')
                    result_columns.extend([
                        f'{trade_short}_电量',
                        f'{trade_short}_电价',
                        f'{trade_short}_电费'
                    ])
                
                for trade in SPECIAL_TRADES:
                    result_columns.append(f'{trade}_电费')
                
                # 创建结果DataFrame
                result_df = pd.DataFrame(all_data, columns=result_columns)
                
                # 转换数值列
                for col in result_columns[2:]:
                    result_df[col] = pd.to_numeric(result_df[col], errors='coerce')
                
                # 添加汇总行
                summary_df = calculate_summary_row(result_df)
                result_df = pd.concat([result_df, summary_df], ignore_index=True)
                
                # 显示结果
                tab1, tab2 = st.tabs(["结算数据明细", "处理报告"])
                
                with tab1:
                    st.dataframe(result_df, use_container_width=True)
                
                with tab2:
                    report_data = {
                        '统计项': ['上传文件数', '成功处理数', '失败数', '成功率'],
                        '数值': [
                            total_files,
                            processed_files,
                            total_files - processed_files,
                            f"{processed_files/total_files:.1%}" if total_files > 0 else "0%"
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
                    file_name=f"黑龙江结算数据_{current_time}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.success(f"✅ 处理完成！成功提取 {processed_files}/{total_files} 个文件")
            
            else:
                st.error("⚠️ 未提取到有效数据！")
                st.markdown("""
                **可能原因及解决方法：**
                1. **PDF格式问题**：检查是否为可复制文本的PDF（非扫描件）
                2. **文件格式不匹配**：确认是否为黑龙江日清分结算单标准格式
                3. **字段名称不匹配**：检查PDF中的交易名称是否与程序预设匹配
                """)
                
                # 显示调试信息
                if show_debug and 'debug_text' in st.session_state and 0 <= debug_file_index < len(uploaded_files):
                    st.divider()
                    st.subheader("🔍 调试信息")
                    st.text_area("提取的PDF文本（前1000字符）：", 
                                st.session_state.get('debug_text', '无调试信息'),
                                height=300)
                    
                    # 显示文件信息
                    debug_file = uploaded_files[int(debug_file_index)]
                    st.info(f"调试文件：{debug_file.name}")
    
    else:
        st.info("👆 请上传PDF或Excel文件开始处理")

if __name__ == "__main__":
    main()
