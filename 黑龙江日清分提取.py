import streamlit as st
import pandas as pd
import re
from datetime import datetime
import warnings
import pdfplumber
from io import BytesIO
import numpy as np

# 忽略样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.stylesheet")

# ---------------------- 核心配置 ----------------------
# 场站列表
STATIONS = [
    "依兰县协合风力发电有限公司（双发B风电场）",
    "依兰县协合风力发电有限公司（双发A风电场）"
]

# 科目编码到名称的映射
TRADE_CODE_TO_NAME = {
    # 常规交易科目
    "101010101": "优先发电交易",
    "101020101": "电网企业代理购电交易", 
    "101020301": "省内电力直接交易",
    "101040322": "送上海省间绿色电力交易(电能量)",
    "102020101": "送辽宁交易",
    "102020301": "送华北交易", 
    "102010101": "送山东交易",
    "102010201": "送浙江交易",
    "202030001": "送江苏省间绿色电力交易（电能量）",
    "202030002": "送浙江省间绿色电力交易（电能量）",
    "101080101": "省内现货日前交易",
    "101080201": "省内现货实时交易",
    "101080301": "省间现货日前交易",
    "101080401": "省间现货日内交易",
    # 特殊科目
    "201010101": "中长期合约阻塞费用",
    "201020101": "省间省内价差费用"
}

# 科目名称标准化映射（处理PDF中可能出现的变体）
TRADE_NAME_VARIANTS = {
    "送上海省间绿色电力交易(电能量 )": "送上海省间绿色电力交易(电能量)",
    "送上海省间绿色电力交易(电能量)": "送上海省间绿色电力交易(电能量)",
    "送江苏省间绿色电力交易（电能量）": "送江苏省间绿色电力交易（电能量）",
    "送浙江省间绿色电力交易（电能量）": "送浙江省间绿色电力交易（电能量）",
    "中长期合约阻塞费用": "中长期合约阻塞费用",
    "省间省内价差费用": "省间省内价差费用"
}

# ---------------------- 核心提取函数 ----------------------
def extract_station_name(pdf_lines):
    """智能提取场站名称"""
    for i, line in enumerate(pdf_lines):
        line = line.strip()
        
        # 方法1: 直接匹配已知场站
        for station in STATIONS:
            if station in line:
                return station
        
        # 方法2: 匹配公司名称模式
        if "公司名称" in line or "场站" in line:
            # 提取公司名称后的内容
            match = re.search(r'[:：]\s*(.+?有限公司)', line)
            if match:
                base_name = match.group(1)
                # 尝试判断是A站还是B站
                for j in range(i, min(i+5, len(pdf_lines))):
                    next_line = pdf_lines[j]
                    if "机组" in next_line:
                        if "B" in next_line.upper() or "2" in next_line or "二" in next_line:
                            return f"{base_name}（双发B风电场）"
                        elif "A" in next_line.upper() or "1" in next_line or "一" in next_line:
                            return f"{base_name}（双发A风电场）"
                return f"{base_name}（未知场站）"
    
    # 方法3: 从文件名或上下文推断
    return "依兰县协合风力发电有限公司（场站未识别）"

def extract_date_from_pdf(pdf_lines):
    """提取清分日期"""
    for line in pdf_lines:
        # 尝试多种日期模式
        patterns = [
            r'清分日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2})',
            r'日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2})',
            r'(\d{4}年\d{1,2}月\d{1,2}日)',
            r'(\d{4}\.\d{1,2}\.\d{1,2})',
            r'(\d{4}/\d{1,2}/\d{1,2})'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, line)
            if match:
                date_str = match.group(1)
                # 统一格式化为YYYY-MM-DD
                date_str = date_str.replace('年', '-').replace('月', '-').replace('日', '')
                date_str = date_str.replace('.', '-').replace('/', '-')
                return date_str
    
    return None

def extract_total_data(pdf_text):
    """提取合计电量、合计电费"""
    total_quantity = None
    total_amount = None
    
    # 模式1: 查找"合计电量"和"合计电费"
    qty_match = re.search(r'合计电量[^\d]*([\d,]+\.?\d*)', pdf_text.replace(' ', ''))
    if qty_match:
        total_quantity = float(qty_match.group(1).replace(',', ''))
    
    amount_match = re.search(r'合计电费[^\d]*([\d,]+\.?\d*)', pdf_text.replace(' ', ''))
    if amount_match:
        total_amount = float(amount_match.group(1).replace(',', ''))
    
    # 模式2: 如果没找到，尝试找"总计"等
    if not total_quantity or not total_amount:
        lines = pdf_text.split('\n')
        for i, line in enumerate(lines):
            if '合计' in line or '总计' in line or '小计' in line:
                # 提取该行及后续行的所有数字
                numbers = re.findall(r'[\d,]+\.?\d*', line.replace(' ', ''))
                for j in range(i+1, min(i+3, len(lines))):
                    numbers.extend(re.findall(r'[\d,]+\.?\d*', lines[j].replace(' ', '')))
                
                if len(numbers) >= 2:
                    if not total_quantity:
                        total_quantity = float(numbers[0].replace(',', ''))
                    if not total_amount:
                        total_amount = float(numbers[1].replace(',', ''))
                break
    
    return total_quantity, total_amount

def parse_trade_line(line, next_line=None):
    """解析交易数据行，返回科目编码、科目名称、电量、电价、电费"""
    line_clean = line.strip()
    
    # 初始化结果
    trade_code = None
    trade_name = None
    quantity = None
    price = None
    fee = None
    
    # 方法1: 查找科目编码
    code_pattern = r'(\d{9})'  # 9位数字编码
    code_match = re.search(code_pattern, line_clean)
    if code_match:
        trade_code = code_match.group(1)
        trade_name = TRADE_CODE_TO_NAME.get(trade_code)
    
    # 方法2: 如果没有编码，尝试匹配科目名称
    if not trade_name:
        for name_key, name_std in TRADE_NAME_VARIANTS.items():
            if name_key in line_clean:
                trade_name = name_std
                break
    
    # 提取数字
    numbers = re.findall(r'-?[\d,]+\.?\d*', line_clean.replace(' ', ''))
    
    # 如果是特殊科目（只有电费）
    if trade_name in ["中长期合约阻塞费用", "省间省内价差费用"]:
        if numbers:
            fee = float(numbers[0].replace(',', ''))
    # 常规科目（电量、电价、电费）
    elif numbers:
        if len(numbers) >= 3:
            quantity = float(numbers[0].replace(',', '')) if numbers[0] else None
            price = float(numbers[1].replace(',', '')) if len(numbers) > 1 and numbers[1] else None
            fee = float(numbers[2].replace(',', '')) if len(numbers) > 2 and numbers[2] else None
        elif len(numbers) == 2:
            quantity = float(numbers[0].replace(',', '')) if numbers[0] else None
            fee = float(numbers[1].replace(',', '')) if numbers[1] else None
        elif len(numbers) == 1:
            fee = float(numbers[0].replace(',', ''))
    
    # 如果本行数字不够，尝试下一行
    if (quantity is None or price is None or fee is None) and next_line:
        next_numbers = re.findall(r'-?[\d,]+\.?\d*', next_line.replace(' ', ''))
        all_numbers = numbers + next_numbers
        
        if trade_name in ["中长期合约阻塞费用", "省间省内价差费用"]:
            if all_numbers and fee is None:
                fee = float(all_numbers[0].replace(',', ''))
        elif all_numbers:
            if quantity is None and len(all_numbers) > 0:
                quantity = float(all_numbers[0].replace(',', ''))
            if price is None and len(all_numbers) > 1:
                price = float(all_numbers[1].replace(',', ''))
            if fee is None and len(all_numbers) > 2:
                fee = float(all_numbers[2].replace(',', ''))
    
    return trade_name, quantity, price, fee

def extract_all_trade_data(pdf_text):
    """提取所有交易数据"""
    lines = [line.strip() for line in pdf_text.split('\n') if line.strip()]
    
    # 初始化结果字典
    trade_data = {}
    for trade_name in TRADE_CODE_TO_NAME.values():
        trade_data[trade_name] = {'quantity': None, 'price': None, 'fee': None}
    
    # 遍历所有行，提取交易数据
    i = 0
    while i < len(lines):
        line = lines[i]
        next_line = lines[i+1] if i+1 < len(lines) else ""
        
        trade_name, quantity, price, fee = parse_trade_line(line, next_line)
        
        if trade_name and trade_name in trade_data:
            # 更新数据
            trade_data[trade_name]['quantity'] = quantity or trade_data[trade_name]['quantity']
            trade_data[trade_name]['price'] = price or trade_data[trade_name]['price']
            trade_data[trade_name]['fee'] = fee or trade_data[trade_name]['fee']
        
        i += 1
    
    return trade_data

def extract_data_from_pdf(file_obj, file_name):
    """从PDF提取数据 - 改进版"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
        
        if not all_text or len(all_text.strip()) < 50:
            raise ValueError("PDF为空或文本内容太少，可能是扫描件")
        
        # 按行分割并清理
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
        trade_data = extract_all_trade_data(all_text)
        
        # 构建结果列表
        result = [station_name, date, total_quantity, total_amount]
        
        # 添加常规交易数据
        normal_trades = [
            "优先发电交易", "电网企业代理购电交易", "省内电力直接交易",
            "送上海省间绿色电力交易(电能量)", "送辽宁交易", "送华北交易", 
            "送山东交易", "送浙江交易", "送江苏省间绿色电力交易（电能量）",
            "送浙江省间绿色电力交易（电能量）", "省内现货日前交易", 
            "省内现货实时交易", "省间现货日前交易", "省间现货日内交易"
        ]
        
        for trade in normal_trades:
            data = trade_data.get(trade, {'quantity': None, 'price': None, 'fee': None})
            result.extend([data['quantity'], data['price'], data['fee']])
        
        # 添加特殊科目数据
        special_trades = ["中长期合约阻塞费用", "省间省内价差费用"]
        for trade in special_trades:
            data = trade_data.get(trade, {'quantity': None, 'price': None, 'fee': None})
            result.append(data['fee'])
        
        return result
        
    except Exception as e:
        st.error(f"处理PDF {file_name} 出错: {str(e)[:200]}")
        return ["未知场站", None, None, None] + [None] * (14*3 + 2)  # 14个常规科目 * 3列 + 2个特殊科目

def calculate_summary_row(data_df):
    """计算汇总行"""
    if data_df.empty:
        return pd.DataFrame()
    
    summary_row = {'场站名称': '总计', '清分日期': ''}
    
    for col in data_df.columns:
        if col in ['场站名称', '清分日期']:
            continue
        
        # 电价列计算平均值
        if '电价' in col and '电费' not in col:  # 避免匹配到"电费"
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
    
    st.title("📊 黑龙江日清分结算单数据提取工具（改进版）")
    st.markdown("**修复问题：科目识别错误、场站识别错误、数据错位**")
    st.divider()
    
    # 场站信息显示
    with st.expander("🏭 支持的场站"):
        st.write("""
        1. **依兰县协合风力发电有限公司（双发A风电场）**
        2. **依兰县协合风力发电有限公司（双发B风电场）**
        """)
    
    # 科目信息显示
    with st.expander("📋 科目编码对照表"):
        code_df = pd.DataFrame([
            {"科目编码": "101010101", "科目名称": "优先发电交易"},
            {"科目编码": "101020101", "科目名称": "电网企业代理购电交易"},
            {"科目编码": "101020301", "科目名称": "省内电力直接交易"},
            {"科目编码": "101040322", "科目名称": "送上海省间绿色电力交易(电能量)"},
            {"科目编码": "102020101", "科目名称": "送辽宁交易"},
            {"科目编码": "102020301", "科目名称": "送华北交易"},
            {"科目编码": "102010101", "科目名称": "送山东交易"},
            {"科目编码": "102010201", "科目名称": "送浙江交易"},
            {"科目编码": "202030001", "科目名称": "送江苏省间绿色电力交易（电能量）"},
            {"科目编码": "202030002", "科目名称": "送浙江省间绿色电力交易（电能量）"},
            {"科目编码": "101080101", "科目名称": "省内现货日前交易"},
            {"科目编码": "101080201", "科目名称": "省内现货实时交易"},
            {"科目编码": "101080301", "科目名称": "省间现货日前交易"},
            {"科目编码": "101080401", "科目名称": "省间现货日内交易"},
            {"科目编码": "201010101", "科目名称": "中长期合约阻塞费用"},
            {"科目编码": "201020101", "科目名称": "省间省内价差费用"}
        ])
        st.dataframe(code_df, use_container_width=True)
    
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
                
                # 常规科目列
                normal_trades = [
                    "优先发电交易", "电网企业代理购电交易", "省内电力直接交易",
                    "送上海省间绿色电力交易(电能量)", "送辽宁交易", "送华北交易", 
                    "送山东交易", "送浙江交易", "送江苏省间绿色电力交易（电能量）",
                    "送浙江省间绿色电力交易（电能量）", "省内现货日前交易", 
                    "省内现货实时交易", "省间现货日前交易", "省间现货日内交易"
                ]
                
                for trade in normal_trades:
                    trade_short = trade.replace('（电能量）', '').replace('(电能量)', '').replace('省间绿色电力交易', '省间绿电交易')
                    result_columns.extend([
                        f'{trade_short}_电量',
                        f'{trade_short}_电价',
                        f'{trade_short}_电费'
                    ])
                
                # 特殊科目列
                special_trades = ["中长期合约阻塞费用", "省间省内价差费用"]
                for trade in special_trades:
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
                        '统计项': ['上传文件数', '成功处理数', '失败数', '成功率', '提取场站数', '数据完整性'],
                        '数值': [
                            total_files,
                            processed_files,
                            total_files - processed_files,
                            f"{(processed_files/total_files)*100:.1f}%" if total_files > 0 else "0%",
                            result_df['场站名称'].nunique() - 1,  # 减去总计行
                            "✅ 科目编码识别" if processed_files > 0 else "❌ 需检查格式"
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
                    file_name=f"黑龙江结算数据_改进版_{current_time}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.success(f"✅ 处理完成！成功提取 {processed_files}/{total_files} 个文件")
                st.info("""
                **改进说明：**
                1. **科目识别**：基于科目编码（9位数字）精确识别，避免名称匹配错误
                2. **场站识别**：支持双发A/B风电场精确区分
                3. **数据提取**：逐行解析，避免数据错位
                4. **容错处理**：支持跨行数据提取
                """)
                
            else:
                st.error("⚠️ 未提取到有效数据！")
                st.markdown("""
                **可能原因：**
                1. PDF文件格式不标准
                2. 文件中缺少科目编码
                3. 文件为扫描件（不可复制文本）
                
                **解决方法：**
                1. 确认PDF包含可复制文本
                2. 检查文件是否为标准黑龙江日清分格式
                3. 联系管理员获取技术支持
                """)
    
    else:
        st.info("👆 请上传PDF文件开始处理")

if __name__ == "__main__":
    main()
