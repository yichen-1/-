import streamlit as st
import pandas as pd
import os
import zipfile
import re
from datetime import datetime, date, time
from openpyxl.styles import Alignment, PatternFill
from io import BytesIO
import shutil

# ---------------------- 全局配置 ----------------------
# 持久化存储目录（部署时可修改为绝对路径）
STORAGE_DIR = os.path.join(os.path.expanduser('~'), 'power_analysis_storage')
CONTRACT_DIR = os.path.join(STORAGE_DIR, 'monthly_contracts')
# 自动创建存储目录
os.makedirs(CONTRACT_DIR, exist_ok=True)

# 设置页面配置
st.set_page_config(
    page_title="连续竞价调整系统",
    page_icon="📊",
    layout="wide"
)

# ---------------------- 工具函数 ----------------------
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

# ---------------------- 持久化相关函数 ----------------------
def save_monthly_contract_file(uploaded_file, month):
    """保存月度合约文件到本地存储"""
    # 生成文件名：月份_原文件名
    safe_filename = re.sub(r'[^\w\.-]', '_', uploaded_file.name)
    save_path = os.path.join(CONTRACT_DIR, f"{month}_{safe_filename}")
    # 保存文件
    with open(save_path, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    return save_path

def get_uploaded_months():
    """获取已上传的月份列表"""
    months = set()
    for filename in os.listdir(CONTRACT_DIR):
        if filename.startswith(('2025-', '2024-')) and (filename.endswith('.xlsx') or filename.endswith('.xls')):
            # 提取月份（格式：2025-11）
            month_part = filename.split('_')[0]
            if len(month_part) == 7 and '-' in month_part:
                months.add(month_part)
    return sorted(list(months))

def get_files_by_month(month):
    """获取指定月份的合约文件列表"""
    files = []
    for filename in os.listdir(CONTRACT_DIR):
        if filename.startswith(f"{month}_") and (filename.endswith('.xlsx') or filename.endswith('.xls')):
            file_path = os.path.join(CONTRACT_DIR, filename)
            if os.path.isfile(file_path):
                files.append(file_path)
    return files

def load_contract_files(selected_months):
    """加载选中月份的所有合约文件"""
    contract_files = []
    for month in selected_months:
        month_files = get_files_by_month(month)
        for file_path in month_files:
            # 转换为BytesIO供pandas读取
            with open(file_path, 'rb') as f:
                bytes_io = BytesIO(f.read())
                bytes_io.name = os.path.basename(file_path)  # 保留文件名
                contract_files.append(bytes_io)
    return contract_files

# ---------------------- 核心业务函数 ----------------------
def generate_integrated_file_streamlit(source_excel_files, unit_station_mapping):
    """生成电量电价整合文件"""
    unit_data = {unit: [] for unit in unit_station_mapping.keys()}
    
    for file_idx, uploaded_file in enumerate(source_excel_files):
        try:
            xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
            for sheet in xls.sheet_names:
                df = xls.parse(sheet)
                if df.empty or df.shape[1] < 1:
                    continue
                
                key_df = extract_key_columns(df)
                if key_df.empty:
                    continue
                
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
            continue
    
    output_io = BytesIO()
    with pd.ExcelWriter(output_io, engine='openpyxl', mode='w') as writer:
        for cleaned_unit, station_name in unit_station_mapping.items():
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
            
            for col in merged_df.columns:
                if '电量' in col or '电价' in col:
                    merged_df[col] = merged_df[col].apply(truncate_to_two_decimal)
            
            merged_df.to_excel(writer, sheet_name=station_name, index=False)
            format_worksheet(writer.sheets[station_name])
    
    output_io.seek(0)
    return output_io

def process_power_forecast_streamlit(forecast_file):
    """处理功率预测数据"""
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
                    continue
                
                if df.empty or df.shape[0] < 4 or df.shape[1] < 2:
                    continue
                
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
                    continue
                
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
                        continue
                
                if not processed_data:
                    continue
                
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
    
    except Exception as e:
        pass
    
    output_io.seek(0)
    return output_io

def process_price_quantity_streamlit(price_quantity_file):
    """处理电价电量数据"""
    output_io = BytesIO()
    try:
        xls_input = pd.ExcelFile(price_quantity_file, engine='openpyxl')
        sheet_names = xls_input.sheet_names
        
        with pd.ExcelWriter(output_io, engine='openpyxl') as writer:
            for sheet_name in sheet_names:
                try:
                    df = xls_input.parse(sheet_name)
                except Exception as e:
                    continue
                
                if df.empty:
                    continue
                
                date_col = next((col for col in df.columns if '日期' in str(col)), None)
                quantity_cols = [col for col in df.columns if '电量' in str(col)]
                price_cols = [col for col in df.columns if '电价' in str(col)]
                
                if not date_col or not quantity_cols:
                    continue
                
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
                        continue
                
                if not dates or not quantity_data:
                    continue
                
                # 生成数据
                processed_data = []
                for i, (date, quantities, prices) in enumerate(zip(dates, quantity_data, price_data)):
                    row_data = [date] + quantities + prices
                    processed_data.append(row_data)
                
                output_cols = ['日期'] + quantity_cols + price_cols
                processed_df = pd.DataFrame(processed_data, columns=output_cols)
                processed_df.to_excel(writer, sheet_name=sheet_name, index=False)
                format_worksheet(writer.sheets[sheet_name])
    
    except Exception as e:
        pass
    
    output_io.seek(0)
    return output_io

def calculate_difference_streamlit(forecast_file, price_quantity_file):
    """计算差值（仅返回数据字典）"""
    station_coefficient = {
        '风储一期': 0.8*0.725*0.7 ,   
        '风储二期': 0.8*0.725*0.7,
        '栗溪': 0.8*0.725*0.7,
        '峪山一期': 0.8*0.725*0.7 ,
        '圣境山': 0.8*0.725*0.7,
        '襄北农光': 0.8*0.775*0.8,
        '浠水渔光': 0.8*0.775*0.8
    }
    
    # 用于展示的数据字典：{场站名: 数据框}
    result_data = {}
    
    try:
        forecast_xls = pd.ExcelFile(forecast_file, engine='openpyxl')
        price_quantity_xls = pd.ExcelFile(price_quantity_file, engine='openpyxl')
        
        forecast_sheet_names = forecast_xls.sheet_names
        price_quantity_sheet_names = price_quantity_xls.sheet_names

        for sheet_name in forecast_sheet_names:
            if sheet_name == '填写说明':
                continue
            if sheet_name not in price_quantity_sheet_names:
                continue

            try:
                forecast_df = forecast_xls.parse(sheet_name)
                price_quantity_df = price_quantity_xls.parse(sheet_name)
            except Exception as e:
                continue

            if forecast_df.empty or len(forecast_df.columns) < 2:
                continue
            
            current_coeff = station_coefficient.get(sheet_name, 1.0)
            
            time_col = forecast_df.iloc[:, 0]
            forecast_cols = forecast_df.columns[1:]
            quantity_cols = [col for col in price_quantity_df.columns if '电量' in str(col)]
            price_cols = [col for col in price_quantity_df.columns if '电价' in str(col)]
            
            if not quantity_cols:
                continue
            quantity_col = quantity_cols[0]
            price_col = price_cols[0] if price_cols else None

            processed_data = []
            for idx, row in forecast_df.iterrows():
                if idx >= len(price_quantity_df):
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
                        row_data.append(None)
                
                row_data.append(current_price)
                processed_data.append(row_data)

            new_cols = ['时间']
            for col in forecast_cols:
                new_cols.extend([col, f'{col} (修正后差额)'])
            new_cols.append('对应时段电价')
            
            processed_df = pd.DataFrame(processed_data, columns=new_cols)
            if '对应时段电价' in processed_df.columns:
                processed_df['对应时段电价'] = processed_df['对应时段电价'].apply(truncate_to_two_decimal)
            
            # 保存到结果字典
            result_data[sheet_name] = processed_df.copy()
    
    except Exception as e:
        st.error(f"计算差值出错：{str(e)}")
    
    return result_data

# ---------------------- 主页面逻辑（封装为展开式菜单） ----------------------
def main():
    # 侧边栏主菜单：连续竞价调整（展开式）
    with st.sidebar:
        st.title("📋 系统功能菜单")
        
        # 核心：连续竞价调整 展开菜单
        with st.expander("🔧 连续竞价调整", expanded=False):
            st.header("📁 文件管理")
            
            # 1. 批量上传月度合约文件
            st.subheader("1. 批量上传月度合约文件")
            new_contract_files = st.file_uploader(
                "选择合约文件（支持批量上传）",
                type=["xlsx", "xls"],
                accept_multiple_files=True,
                key="new_contract"
            )
            selected_month = st.text_input(
                "文件对应月份（格式：2025-11）",
                value=datetime.now().strftime("%Y-%m"),
                key="contract_month"
            )
            
            # 批量保存逻辑
            if st.button("保存月度文件", key="save_contract"):
                if not new_contract_files:
                    st.warning("⚠️ 请先选择要上传的合约文件！")
                elif not selected_month:
                    st.warning("⚠️ 请输入对应的月份（格式：2025-11）！")
                else:
                    with st.spinner("批量保存文件中..."):
                        saved_files = []
                        failed_files = []
                        for file in new_contract_files:
                            try:
                                save_path = save_monthly_contract_file(file, selected_month)
                                saved_files.append(os.path.basename(save_path))
                            except Exception as e:
                                failed_files.append(f"{file.name} - {str(e)}")
                        
                        if saved_files:
                            st.success(f"✅ 成功保存 {len(saved_files)} 个文件：")
                            for fname in saved_files:
                                st.write(f"  - {fname}")
                        if failed_files:
                            st.error(f"❌ 保存失败 {len(failed_files)} 个文件：")
                            for fname in failed_files:
                                st.write(f"  - {fname}")
            
            # 2. 选择已上传的月份
            st.subheader("2. 选择分析月份")
            uploaded_months = get_uploaded_months()
            if uploaded_months:
                selected_months = st.multiselect(
                    "勾选要分析的月份",
                    options=uploaded_months,
                    default=uploaded_months,
                    key="selected_months"
                )
                # 展示每个月份的文件数量
                st.write("📋 各月份文件统计：")
                for month in uploaded_months:
                    file_count = len(get_files_by_month(month))
                    st.write(f"  • {month}：{file_count} 个文件")
            else:
                selected_months = []
                st.info("暂无已上传的月度合约文件，请先上传")
            
            # 3. 上传功率预测文件
            st.subheader("3. 功率预测文件")
            forecast_file = st.file_uploader(
                "上传功率预测文件（2025功率预测.xlsx）",
                type=["xlsx", "xls"],
                key="forecast"
            )
            
            # 映射配置（折叠展示）
            with st.expander("⚙️ 交易单元映射配置", expanded=False):
                unit_to_station = {
                    "襄阳协合峪山泉水风电": "峪山一期",
                    "荆门协合圣境山风电": "圣境山",
                    "襄阳聚合光伏": "襄北农光",
                    "三王（协合襄北）风电": "风储一期",
                    "荆门协合栗溪风电": "栗溪",
                    "襄州协合三王风光储能电站风电二期": "风储二期",
                    "浠水聚合关口光伏": "浠水渔光"
                }
                for k, v in unit_to_station.items():
                    st.write(f"• {k} → {v}")
        
        # 其他扩展菜单（预留）
        st.divider()
        st.write("📌 其他功能模块（预留）")
        # 可添加更多展开菜单
        # with st.expander("📊 数据报表", expanded=False):
        #     st.write("后续添加报表功能")
        # with st.expander("💾 数据导出", expanded=False):
        #     st.write("后续添加导出功能")
    
    # 主页面内容（仅在选择连续竞价调整后显示）
    st.title("🔧 连续竞价调整")
    st.divider()
    
    # 获取侧边栏的变量（通过key获取）
    selected_months = st.session_state.get("selected_months", [])
    forecast_file = st.session_state.get("forecast")
    
    # 执行按钮（禁用条件：无选中月份/无预测文件）
    run_disabled = not (selected_months and forecast_file)
    if st.button("开始测算", type="primary", disabled=run_disabled):
        with st.spinner("正在测算数据，请稍候..."):
            # 加载选中月份的合约文件
            contract_files = load_contract_files(selected_months)
            
            # 执行核心处理流程
            unit_to_station = {
                "襄阳协合峪山泉水风电": "峪山一期",
                "荆门协合圣境山风电": "圣境山",
                "襄阳聚合光伏": "襄北农光",
                "三王（协合襄北）风电": "风储一期",
                "荆门协合栗溪风电": "栗溪",
                "襄州协合三王风光储能电站风电二期": "风储二期",
                "浠水聚合关口光伏": "浠水渔光"
            }
            integrated_io = generate_integrated_file_streamlit(contract_files, unit_to_station)
            forecast_processed_io = process_power_forecast_streamlit(forecast_file)
            price_quantity_processed_io = process_price_quantity_streamlit(integrated_io)
            result_data = calculate_difference_streamlit(forecast_processed_io, integrated_io)
        
        # 仅展示最终汇总数据
        st.divider()
        st.header("📈 最终汇总数据展示")
        if result_data:
            # 按场站分标签展示
            station_tabs = st.tabs(list(result_data.keys()))
            for tab, (station_name, df) in zip(station_tabs, result_data.items()):
                with tab:
                    st.subheader(f"📍 {station_name}")
                    # 数据展示（支持筛选/排序）
                    st.dataframe(
                        df,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "时间": st.column_config.TextColumn("时段", width="small"),
                            "对应时段电价": st.column_config.NumberColumn("电价(元)", format="%.2f"),
                        }
                    )
                    # 单场站数据下载
                    csv_data = df.to_csv(index=False, encoding="utf-8-sig")
                    st.download_button(
                        label=f"📥 下载{station_name}数据（CSV）",
                        data=csv_data,
                        file_name=f"{station_name}_调整结果.csv",
                        mime="text/csv"
                    )
        else:
            st.warning("暂无可展示的结果数据")
        
        st.success("✅ 测算完成！")
    
    # 提示信息
    if run_disabled:
        if not selected_months:
            st.warning("⚠️ 请先在侧边栏「连续竞价调整」菜单中上传并选择要分析的月度合约文件！")
        elif not forecast_file:
            st.warning("⚠️ 请先在侧边栏「连续竞价调整」菜单中上传功率预测文件！")

if __name__ == "__main__":
    main()
