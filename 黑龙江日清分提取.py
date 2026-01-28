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
# 科目编码到名称的映射
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

# 所有科目列表
ALL_TRADES = list(TRADE_CODE_MAP.values())
SPECIAL_TRADES = ["中长期合约阻塞费用", "省间省内价差费用"]
REGULAR_TRADES = [trade for trade in ALL_TRADES if trade not in SPECIAL_TRADES]

# ---------------------- 核心提取函数 ----------------------
def safe_convert_to_numeric(value):
    """安全转换为数值"""
    if value is None or pd.isna(value) or value == '':
        return None
    try:
        if isinstance(value, str):
            # 移除千分位逗号和其他非数字字符
            cleaned = re.sub(r'[^\d.-]', '', value)
            if cleaned and cleaned != '-' and cleaned != '.':
                return float(cleaned)
        return float(value)
    except (ValueError, TypeError):
        return None

def extract_station_name(pdf_text):
    """提取场站名称，支持多场站识别"""
    lines = pdf_text.split('\n')
    
    # 方法1: 从公司名称和机组信息提取
    company_name = None
    station_type = None
    
    for i, line in enumerate(lines):
        line_clean = line.strip()
        
        # 提取公司名称
        if "公司名称" in line_clean:
            match = re.search(r'公司名称[:：]\s*(.+?有限公司)', line_clean)
            if match:
                company_name = match.group(1).strip()
        
        # 提取机组信息判断A/B场站
        if "机组" in line_clean:
            if "B" in line_clean.upper() or "2" in line_clean or "二" in line_clean:
                station_type = "双发B风电场"
            elif "A" in line_clean.upper() or "1" in line_clean or "一" in line_clean:
                station_type = "双发A风电场"
        
        # 如果已经找到足够信息，提前返回
        if company_name and station_type:
            return f"{company_name}（{station_type}）"
    
    # 方法2: 从包含风电场名称的行提取
    for line in lines:
        if "风电场" in line:
            # 提取具体的风电场名称
            match = re.search(r'([^，。！？]+风电场)', line)
            if match:
                return match.group(1).strip()
    
    return "未知场站"

def extract_date_from_pdf(pdf_text):
    """提取清分日期"""
    lines = pdf_text.split('\n')
    
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
                return date_str
    
    return None

def extract_trade_data_using_table_structure(pdf_text):
    """使用表格结构解析交易数据（关键修复）"""
    trade_data = {}
    
    # 初始化所有科目
    for trade in ALL_TRADES:
        if trade in SPECIAL_TRADES:
            trade_data[trade] = {'fee': None}
        else:
            trade_data[trade] = {'quantity': None, 'price': None, 'fee': None}
    
    lines = pdf_text.split('\n')
    
    # 查找交易数据表格的开始位置
    table_start = -1
    for i, line in enumerate(lines):
        if "科目编码" in line and "科目名称" in line:
            table_start = i
            break
    
    if table_start == -1:
        return trade_data
    
    # 解析表格数据
    i = table_start + 1
    while i < len(lines):
        line = lines[i].strip()
        
        # 跳过空行和表头行
        if not line or "科目编码" in line or "合计" in line or "总计" in line:
            i += 1
            continue
        
        # 检查是否包含科目编码
        code_match = re.search(r'\b(\d{9})\b', line)
        if code_match:
            code = code_match.group(1)
            if code in TRADE_CODE_MAP:
                trade_name = TRADE_CODE_MAP[code]
                is_special = trade_name in SPECIAL_TRADES
                
                # 提取该行中的所有数字（跳过科目编码）
                numbers = []
                
                # 方法1: 按列分割（假设列之间用空格分隔）
                parts = re.split(r'\s+', line)
                code_index = -1
                for idx, part in enumerate(parts):
                    if code in part:
                        code_index = idx
                        break
                
                if code_index >= 0:
                    # 从科目编码后的部分提取数字
                    data_parts = parts[code_index + 1:]
                    for part in data_parts:
                        # 清理并提取数字
                        clean_part = re.sub(r'[^\d.-]', '', part)
                        if clean_part and clean_part not in ['-', '.']:
                            try:
                                num = float(clean_part)
                                numbers.append(num)
                            except:
                                pass
                
                # 方法2: 如果方法1没提取到足够数字，使用正则匹配
                if len(numbers) < (1 if is_special else 3):
                    line_after_code = line[code_match.end():]
                    number_matches = re.findall(r'-?\d[\d,]*\.?\d*', line_after_code)
                    numbers = [safe_convert_to_numeric(num) for num in number_matches]
                
                # 分配数据
                if is_special:
                    # 特殊科目：只有电费
                    if numbers:
                        trade_data[trade_name]['fee'] = numbers[0]
                else:
                    # 常规科目：电量、电价、电费
                    if len(numbers) >= 3:
                        trade_data[trade_name]['quantity'] = numbers[0]
                        trade_data[trade_name]['price'] = numbers[1]
                        trade_data[trade_name]['fee'] = numbers[2]
                    elif len(numbers) == 2:
                        trade_data[trade_name]['quantity'] = numbers[0]
                        trade_data[trade_name]['fee'] = numbers[1]
                    elif len(numbers) == 1:
                        trade_data[trade_name]['fee'] = numbers[0]
        
        i += 1
    
    return trade_data

def extract_total_data(pdf_text):
    """提取合计电量、合计电费"""
    total_quantity, total_amount = None, None
    
    lines = pdf_text.split('\n')
    
    for i, line in enumerate(lines):
        line_clean = line.replace(' ', '')
        
        if "合计电量" in line_clean:
            match = re.search(r'合计电量[^\d]*([\d,]+\.?\d*)', line_clean)
            if match:
                total_quantity = safe_convert_to_numeric(match.group(1))
        
        if "合计电费" in line_clean:
            match = re.search(r'合计电费[^\d]*([\d,]+\.?\d*)', line_clean)
            if match:
                total_amount = safe_convert_to_numeric(match.group(1))
    
    return total_quantity, total_amount

def extract_data_from_pdf(file_obj, file_name):
    """从PDF提取数据 - 修复表格结构解析"""
    try:
        with pdfplumber.open(file_obj) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
        
        if not all_text or len(all_text.strip()) < 50:
            raise ValueError("PDF为空或文本内容太少")
        
        # 提取基本信息
        station_name = extract_station_name(all_text)
        date = extract_date_from_pdf(all_text)
        total_quantity, total_amount = extract_total_data(all_text)
        
        # 从文件名提取日期（备用）
        if not date:
            date_match = re.search(r'(\d{4}-\d{2}-\d{2})', file_name)
            if date_match:
                date = date_match.group(1)
        
        # 使用表格结构提取交易数据
        trade_data = extract_trade_data_using_table_structure(all_text)
        
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

def main():
    st.set_page_config(page_title="黑龙江日清分数据提取工具", layout="wide")
    
    st.title("📊 黑龙江日清分结算单数据提取工具（表格结构修复版）")
    st.markdown("**修复问题：表格结构解析、科目编码识别、多场站支持**")
    st.divider()
    
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
                
                # 下载功能
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                output.seek(0)
                
                st.download_button(
                    label="📥 下载Excel文件",
                    data=output,
                    file_name=f"黑龙江结算数据_修复版_{current_time}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success(f"✅ 处理完成！成功提取 {len(all_data)}/{len(uploaded_files)} 个文件")

if __name__ == "__main__":
    main()
