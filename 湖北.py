import streamlit as st
import pandas as pd
import os
import zipfile
import re
from datetime import datetime, date, time
from openpyxl.styles import Alignment, PatternFill
from io import BytesIO

# 设置页面配置
st.set_page_config(
    page_title="功率预测与电价电量分析",
    page_icon="📊",
    layout="wide"
)

# ---------------------- 核心工具函数 ----------------------
def clean_unit_name(unit_name):
    """清理交易单元名称：去除括号及括号内的内容"""
    if pd.isna(unit_name) or unit_name == '':
        return ""
    unit_str = str(unit_name).strip()
    cleaned_str = re.sub(r'(\(.*?\)|（.*?）)', '', unit_str).strip()
    return cleaned_str

def truncate_to_two_decimal(x):
    """将数值截断到两位小数（只舍不入）"""
    if pd.isna(x):
        return None
    try:
        return float(int(float(x) * 100)) / 100
    except:
        return None

def format_worksheet(worksheet):
    """设置工作表格式：内容居中，列宽30"""
    alignment = Alignment(horizontal='center', vertical='center')
    for row in worksheet.iter_rows():
        for cell in row:
            cell.alignment = alignment
    for col in worksheet.columns:
        worksheet.column_dimensions[col[0].column_letter].width = 30

def extract_key_columns(df):
    """提取关键列（日期、时段、电量、电价）"""
    key_columns = {
        '日期': None,
        '时段': None,
        '时段名称': None,
        '电量': None,
        '电价': None
    }
    for col in df.columns:
        col_str = str(col).strip().lower()
        if '日期' in col_str:
            key_columns['日期'] = col
        elif '时段' in col_str and '名称' not in col_str:
            key_columns['时段'] = col
        elif '时段名称' in col_str:
            key_columns['时段名称'] = col
        elif '电量' in col_str:
            key_columns['电量'] = col
        elif '电价' in col_str:
            key_columns['电价'] = col
    if key_columns['电量'] is None and key_columns['电价'] is None:
        return pd.DataFrame()
    selected_cols = [col for col in key_columns.values() if col is not None]
    return df[selected_cols].copy()

def is_valid_excel_bytes(excel_bytes):
    """校验Excel字节流有效性"""
    try:
        with zipfile.ZipFile(BytesIO(excel_bytes), 'r') as zf:
            return '[Content_Types].xml' in zf.namelist()
    except:
        return False

# ---------------------- 核心业务函数 ----------------------
def generate_integrated_file_streamlit(source_excel_files, unit_station_mapping):
    """
    Streamlit版本：生成电量电价整合文件
    :param source_excel_files: Streamlit上传的文件列表（BytesIO）
    :param unit_station_mapping: 交易单元映射字典
    :return: 整合后的Excel字节流
    """
    # 初始化数据存储
    unit_data = {unit: [] for unit in unit_station_mapping.keys()}
    
    # 处理每个上传的文件
    for file_idx, uploaded_file in enumerate(source_excel_files):
        st.write(f"🔍 处理文件：{uploaded_file.name}")
        try:
            xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
            # 遍历工作表
            for sheet in xls.sheet_names:
                df = xls.parse(sheet)
                if df.empty or df.shape[1] < 1:
                    st.write(f"  - 工作表'{sheet}'无数据，跳过")
                    continue
                
                key_df = extract_key_columns(df)
                if key_df.empty:
                    st.write(f"  - 工作表'{sheet}'无电量/电价列，跳过")
                    continue
                
                # 按交易单元拆分
                for idx, row in df.iterrows():
                    try:
                        raw_unit = row.iloc[0]
                        cleaned_unit = clean_unit_name(raw_unit)
                        if cleaned_unit not in unit_station_mapping:
                            continue
                        key_row = key_df.iloc[idx:idx+1].copy()
                        if not key_row.empty:
                            key_row['数据来源'] = f"文件：{uploaded_file.name} | 工作表：{sheet} | 原始单元：{raw_unit}"
                            unit_data[cleaned_unit].append(key_row)
                    except Exception as e:
                        continue
        except Exception as e:
            st.error(f"处理文件 {uploaded_file.name} 出错：{str(e)}")
            continue
    
    # 生成整合Excel
    output_io = BytesIO()
    with pd.ExcelWriter(output_io, engine='openpyxl', mode='w') as writer:
        for cleaned_unit, station_name in unit_station_mapping.items():
            st.write(f"📝 生成{station_name}工作表")
            data_list = unit_data.get(cleaned_unit, [])
            if not data_list:
                pd.DataFrame({"提示": [f"无有效数据：{cleaned_unit}"]}).to_excel(
                    writer, sheet_name=station_name, index=False
                )
                format_worksheet(writer.sheets[station_name])
                continue
            
            merged_df = pd.concat(data_list, ignore_index=True)
            if '日期' in merged_df.columns:
                merged_df['日期'] = pd.to_datetime(merged_df['日期'], errors='coerce')
                merged_df = merged_df.sort_values(by=['日期', '时段']).reset_index(drop=True)
            
            # 小数处理
            for col in merged_df.columns:
                if '电量' in col or '电价' in col:
                    merged_df[col] = merged_df[col].apply(truncate_to_two_decimal)
            
            merged_df.to_excel(writer, sheet_name=station_name, index=False)
            format_worksheet(writer.sheets[station_name])
            st.write(f"  ✅ {station_name}：{len(merged_df)}行数据")
    
    output_io.seek(0)
    return output_io

def process_power_forecast_streamlit(forecast_file):
    """Streamlit版本：处理功率预测数据"""
    output_io = BytesIO()
    try:
        xls = pd.ExcelFile(forecast_file, engine='openpyxl')
        sheet_names = xls.sheet_names
        today = date.today()
        
        with pd.ExcelWriter(output_io, engine='openpyxl') as writer:
            for sheet_name in sheet_names:
                if sheet_name == '填写说明':
                    continue
                try:
                    df = xls.parse(sheet_name)
                except Exception as e:
                    st.error(f"解析工作表 '{sheet_name}' 出错：{str(e)}")
                    continue
                
                if df.empty or df.shape[0] < 4 or df.shape[1] < 2:
                    st.write(f"工作表 '{sheet_name}' 数据结构异常，跳过")
                    continue
                
                # 处理时间列
                time_column = df.iloc[:, 0]
                data_columns = df.columns[1:]
                times = []
                for t in time_column:
                    if isinstance(t, str):
                        try:
                            times.append(pd.to_datetime(t).time())
                        except:
                            times.append(None)
                    elif isinstance(t, (datetime, pd.Timestamp)):
                        times.append(t.time())
                    elif isinstance(t, time):
                        times.append(t)
                    else:
                        times.append(None)
                
                valid_times_mask = [t is not None for t in times]
                times = [t for t in times if t is not None]
                df = df[valid_times_mask].reset_index(drop=True)
                
                if not times:
                    st.write(f"工作表 '{sheet_name}' 无有效时间数据，跳过")
                    continue
                
                # 处理预测数据
                processed_data = []
                for col in data_columns:
                    try:
                        col_date = pd.to_datetime(col).date()
                        if col_date >= today:
                            col_data = df[col]
                            averaged_data = []
                            for i in range(0, len(col_data), 4):
                                segment = col_data[i:i+4]
                                if not segment.isna().all():
                                    avg_val = segment.mean()
                                    averaged_data.append(truncate_to_two_decimal(avg_val))
                                else:
                                    averaged_data.append(None)
                            if any(pd.notna(averaged_data)):
                                row = [col_date] + averaged_data
                                processed_data.append(row)
                    except Exception as e:
                        st.write(f"处理列 '{col}' 出错：{str(e)}")
                        continue
                
                if not processed_data:
                    st.write(f"工作表 '{sheet_name}' 无有效预测数据，跳过")
                    continue
                
                # 构建输出数据
                time_points = [time(hour=i) for i in range(24)]
                columns = ['时间'] + [row[0] for row in processed_data]
                processed_df = pd.DataFrame(columns=columns)
                processed_df['时间'] = [t.strftime('%H:%M:%S') for t in time_points]
                
                for i, row in enumerate(processed_data):
                    col_name = row[0]
                    for j in range(min(24, len(row[1:]))):
                        processed_df.loc[j, col_name] = row[j+1]
                
                processed_df = processed_df.dropna(axis=1, how='all')
                processed_df.to_excel(writer, sheet_name=sheet_name, index=False)
                format_worksheet(writer.sheets[sheet_name])
                st.write(f"✅ 工作表 '{sheet_name}' 处理完成")
    
    except Exception as e:
        st.error(f"处理预测数据出错：{str(e)}")
    
    output_io.seek(0)
    return output_io

def process_price_quantity_streamlit(price_quantity_file, summary_file):
    """Streamlit版本：处理电价电量数据"""
    output_io = BytesIO()
    try:
        xls_input = pd.ExcelFile(price_quantity_file, engine='openpyxl')
        xls_summary = pd.ExcelFile(summary_file, engine='openpyxl') if summary_file else None
        
        sheet_names = xls_input.sheet_names
        with pd.ExcelWriter(output_io, engine='openpyxl') as writer:
            for sheet_name in sheet_names:
                try:
                    df = xls_input.parse(sheet_name)
                except Exception as e:
                    st.error(f"解析工作表 '{sheet_name}' 出错：{str(e)}")
                    continue
                
                if df.empty:
                    st.write(f"工作表 '{sheet_name}' 为空，跳过")
                    continue
                
                # 提取关键列
                date_col = next((col for col in df.columns if '日期' in str(col)), None)
                quantity_cols = [col for col in df.columns if '电量' in str(col)]
                price_cols = [col for col in df.columns if '电价' in str(col)]
                
                if not date_col or not quantity_cols:
                    st.write(f"工作表 '{sheet_name}' 缺少日期/电量列，跳过")
                    continue
                
                # 解析数据
                dates = []
                quantity_data = []
                price_data = []
                for idx, row in df.iterrows():
                    try:
                        current_date = pd.to_datetime(row[date_col]).date()
                        dates.append(current_date)
                        quantities = [truncate_to_two_decimal(row[col]) for col in quantity_cols]
                        quantity_data.append(quantities)
                        prices = [truncate_to_two_decimal(row[col]) for col in price_cols]
                        price_data.append(prices)
                    except Exception as e:
                        st.write(f"解析{sheet_name}第{idx+1}行出错：{str(e)}")
                        continue
                
                if not dates or not quantity_data:
                    st.write(f"工作表 '{sheet_name}' 无有效数据，跳过")
                    continue
                
                # 读取汇总数据
                date_to_summary = {}
                if xls_summary and sheet_name in xls_summary.sheet_names:
                    try:
                        summary_df = xls_summary.parse(sheet_name)
                        summary_date_col = summary_df.columns[0] if not summary_df.empty else None
                        if summary_date_col:
                            for idx, row in summary_df.iterrows():
                                try:
                                    s_date = pd.to_datetime(row[summary_date_col]).date()
                                    s_quantity = truncate_to_two_decimal(row[1]) if len(row) > 1 and pd.notna(row[1]) else None
                                    if s_date and s_quantity:
                                        date_to_summary[s_date] = s_quantity
                                except:
                                    continue
                    except Exception as e:
                        st.write(f"读取{sheet_name}汇总数据出错：{str(e)}")
                
                # 生成处理后数据
                processed_data = []
                for i, (date, quantities, prices) in enumerate(zip(dates, quantity_data, price_data)):
                    row_data = [date] + quantities + prices
                    processed_data.append(row_data)
                    
                    if date in date_to_summary:
                        diffs = []
                        for q in quantities:
                            if pd.notna(q):
                                diff = q - date_to_summary[date]
                                diffs.append(truncate_to_two_decimal(diff))
                            else:
                                diffs.append(None)
                        diff_row = [f"{date} (差额)"] + diffs + prices
                        processed_data.append(diff_row)
                
                output_cols = ['日期'] + quantity_cols + price_cols
                processed_df = pd.DataFrame(processed_data, columns=output_cols)
                processed_df.to_excel(writer, sheet_name=sheet_name, index=False)
                format_worksheet(writer.sheets[sheet_name])
                st.write(f"✅ 工作表 '{sheet_name}' 处理完成")
    
    except Exception as e:
        st.error(f"处理电价电量数据出错：{str(e)}")
    
    output_io.seek(0)
    return output_io

def calculate_difference_streamlit(forecast_file, price_quantity_file):
    """Streamlit版本：计算差值"""
    # 功率预测系数
    station_coefficient = {
        '风储一期': 0.8*0.725*0.7 ,   
        '风储二期': 0.8*0.725*0.7,
        '栗溪': 0.8*0.725*0.7,
        '峪山一期': 0.8*0.725*0.7 ,
        '圣境山': 0.8*0.725*0.7,
        '襄北农光': 0.8*0.775*0.8,
        '浠水渔光': 0.8*0.775*0.8
    }
    
    output_io = BytesIO()
    try:
        forecast_xls = pd.ExcelFile(forecast_file, engine='openpyxl')
        price_quantity_xls = pd.ExcelFile(price_quantity_file, engine='openpyxl')
        
        forecast_sheet_names = forecast_xls.sheet_names
        price_quantity_sheet_names = price_quantity_xls.sheet_names

        with pd.ExcelWriter(output_io, engine='openpyxl') as writer:
            for sheet_name in forecast_sheet_names:
                if sheet_name == '填写说明':
                    continue
                if sheet_name not in price_quantity_sheet_names:
                    st.write(f"工作表 '{sheet_name}' 在电价电量文件中不存在，跳过")
                    continue

                try:
                    forecast_df = forecast_xls.parse(sheet_name)
                    price_quantity_df = price_quantity_xls.parse(sheet_name)
                except Exception as e:
                    st.error(f"解析{sheet_name}出错：{str(e)}")
                    continue

                if forecast_df.empty or len(forecast_df.columns) < 2:
                    st.write(f"{sheet_name}预测数据为空，跳过")
                    continue
                
                current_coeff = station_coefficient.get(sheet_name, 1.0)
                st.write(f"🔧 处理{sheet_name}：功率预测系数 = {round(current_coeff, 4)}")
                
                # 提取列
                time_col = forecast_df.iloc[:, 0]
                forecast_cols = forecast_df.columns[1:]
                quantity_cols = [col for col in price_quantity_df.columns if '电量' in str(col)]
                price_cols = [col for col in price_quantity_df.columns if '电价' in str(col)]
                
                if not quantity_cols:
                    st.write(f"{sheet_name}无电量列，跳过")
                    continue
                quantity_col = quantity_cols[0]
                price_col = price_cols[0] if price_cols else None

                # 计算差值
                processed_data = []
                for idx, row in forecast_df.iterrows():
                    if idx >= len(price_quantity_df):
                        st.write(f"{sheet_name}数据行数不足，第{idx+1}行跳过")
                        continue

                    current_time = row[0]
                    row_data = [current_time]
                    current_price = truncate_to_two_decimal(price_quantity_df.iloc[idx][price_col]) if (price_col and pd.notna(price_quantity_df.iloc[idx][price_col])) else None
                    
                    for col in forecast_cols:
                        forecast_val = row[col]
                        row_data.append(forecast_val)
                        
                        try:
                            quantity_val = price_quantity_df.iloc[idx][quantity_col]
                            if pd.notna(forecast_val) and pd.notna(quantity_val):
                                corrected_forecast = float(forecast_val) * current_coeff
                                diff_val = truncate_to_two_decimal(corrected_forecast - float(quantity_val))
                                
                                if diff_val < 0:
                                    max_negative = -float(quantity_val)
                                    diff_val = max(diff_val, max_negative)
                                
                                row_data.append(diff_val)
                            else:
                                row_data.append(None)
                        except Exception as e:
                            st.write(f"  计算{sheet_name}第{idx+1}行{col}列差值出错：{str(e)}")
                            row_data.append(None)
                    
                    row_data.append(current_price)
                    processed_data.append(row_data)

                # 构建列名
                new_cols = ['时间']
                for col in forecast_cols:
                    new_cols.extend([col, f'{col} (修正后差额)'])
                new_cols.append('对应时段电价')
                
                processed_df = pd.DataFrame(processed_data, columns=new_cols)
                if '对应时段电价' in processed_df.columns:
                    processed_df['对应时段电价'] = processed_df['对应时段电价'].apply(truncate_to_two_decimal)
                processed_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 设置格式 + 负值标黄
                worksheet = writer.sheets[sheet_name]
                format_worksheet(worksheet)
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                for col_idx in range(2, len(new_cols)-1, 2):
                    col_letter = chr(65 + col_idx)
                    for row_idx in range(1, len(processed_df) + 1):
                        cell = worksheet[f'{col_letter}{row_idx + 1}']
                        try:
                            val = float(cell.value) if cell.value is not None else None
                            if val is not None and val < 0:
                                cell.fill = yellow_fill
                        except:
                            continue

                st.write(f"✅ 工作表 '{sheet_name}' 处理完成")
    
    except Exception as e:
        st.error(f"计算差值出错：{str(e)}")
    
    output_io.seek(0)
    return output_io

# ---------------------- Streamlit 页面交互 ----------------------
def main():
    st.title("📊 功率预测与电价电量分析系统")
    st.divider()
    
    # 侧边栏：文件上传
    with st.sidebar:
        st.header("📁 文件上传")
        # 1. 功率预测文件
        forecast_file = st.file_uploader(
            "1. 上传功率预测文件（2025功率预测.xlsx）",
            type=["xlsx", "xls"],
            key="forecast"
        )
        
        # 2. 机组净合约电量文件（支持多文件上传）
        contract_files = st.file_uploader(
            "2. 上传机组净合约电量文件（支持多个）",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="contract"
        )
        
        # 3. 汇总文件（可选）
        summary_file = st.file_uploader(
            "3. 上传汇总文件（汇总.xlsx，可选）",
            type=["xlsx", "xls"],
            key="summary"
        )
        
        # 映射关系（固定）
        st.header("⚙️ 映射配置")
        unit_to_station = {
            "襄阳协合峪山泉水风电": "峪山一期",
            "荆门协合圣境山风电": "圣境山",
            "襄阳聚合光伏": "襄北农光",
            "三王（协合襄北）风电": "风储一期",
            "荆门协合栗溪风电": "栗溪",
            "襄州协合三王风光储能电站风电二期": "风储二期",
            "浠水聚合关口光伏": "浠水渔光"
        }
        st.write("交易单元 → 场站映射：")
        for k, v in unit_to_station.items():
            st.write(f"• {k} → {v}")
    
    # 主页面：执行流程
    st.header("🚀 执行分析流程")
    if st.button("开始处理", type="primary", disabled=not (forecast_file and contract_files)):
        with st.spinner("正在处理数据，请稍候..."):
            # 步骤1：生成电量电价整合文件
            st.subheader("步骤1：生成电量电价整合文件")
            integrated_io = generate_integrated_file_streamlit(contract_files, unit_to_station)
            st.download_button(
                label="📥 下载电量电价整合文件",
                data=integrated_io,
                file_name="电量电价整合.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # 步骤2：处理预测数据
            st.subheader("步骤2：处理功率预测数据")
            forecast_processed_io = process_power_forecast_streamlit(forecast_file)
            st.download_button(
                label="📥 下载处理后功率预测文件",
                data=forecast_processed_io,
                file_name="2025功率预测_处理后.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # 步骤3：处理电价电量数据
            st.subheader("步骤3：处理电价电量数据")
            price_quantity_processed_io = process_price_quantity_streamlit(integrated_io, summary_file)
            st.download_button(
                label="📥 下载处理后电价电量文件",
                data=price_quantity_processed_io,
                file_name="电量电价整合_处理后.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # 步骤4：计算差值
            st.subheader("步骤4：计算功率预测与电量差值")
            difference_io = calculate_difference_streamlit(forecast_processed_io, integrated_io)
            st.download_button(
                label="📥 下载调整结果文件",
                data=difference_io,
                file_name="调整结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        st.success("✅ 所有处理已完成！请下载对应的结果文件。")
    
    # 提示信息
    if not (forecast_file and contract_files):
        st.warning("⚠️ 请先上传【功率预测文件】和【机组净合约电量文件】后再执行处理！")

if __name__ == "__main__":
    main()
