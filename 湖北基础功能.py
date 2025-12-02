import streamlit as st
import pandas as pd
import os
import zipfile
import re
import json
import uuid
from datetime import datetime, date, time
from openpyxl.styles import Alignment, PatternFill
from io import BytesIO
import shutil
import plotly.express as px

# -------------------------- 全局配置（核心：按省份隔离） --------------------------
st.set_page_config(
    page_title="多省份新能源数据管理系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 省份配置（可扩展）
PROVINCES = ["湖北", "贵州"]
CURRENT_PROVINCE_KEY = "current_province"
CURRENT_FUNCTION_KEY = "current_function"  # 连续竞价调整 / 光伏风电数据管理

# 初始化全局会话状态
if CURRENT_PROVINCE_KEY not in st.session_state:
    st.session_state[CURRENT_PROVINCE_KEY] = "湖北"
if CURRENT_FUNCTION_KEY not in st.session_state:
    st.session_state[CURRENT_FUNCTION_KEY] = "连续竞价调整"
# 按省份隔离的状态存储
if "province_data" not in st.session_state:
    st.session_state.province_data = {
        "湖北": {
            "竞价调整": {},  # 湖北连续竞价调整的所有状态
            "光伏风电": {}   # 湖北光伏风电数据管理的所有状态
        },
        "贵州": {
            "竞价调整": {},  # 贵州连续竞价调整（可自定义）
            "光伏风电": {}   # 贵州光伏风电数据管理（可自定义）
        }
    }

# -------------------------- 工具函数（通用+省份专属） --------------------------
# ========== 通用工具函数 ==========
def standardize_column_name(col):
    """列名标准化"""
    col_str = str(col).strip() if col is not None else f"未知列_{uuid.uuid4().hex[:8]}"
    col_str = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9_]', '_', col_str)
    if col_str == "" or col_str == "_":
        col_str = f"列_{uuid.uuid4().hex[:8]}"
    return col_str

def force_unique_columns(df):
    """强制列名唯一"""
    df.columns = [standardize_column_name(col) for col in df.columns]
    cols = df.columns.tolist()
    unique_cols = []
    col_seen = {}
    
    for col in cols:
        if col not in col_seen:
            col_seen[col] = 0
            unique_cols.append(col)
        else:
            col_seen[col] += 1
            unique_col = f"{col}_{uuid.uuid4().hex[:4]}"
            unique_cols.append(unique_col)
    
    df.columns = unique_cols
    time_col_candidates = [i for i, col in enumerate(df.columns) if "时间" in col or "date" in col.lower()]
    if time_col_candidates:
        df.columns = ["时间" if i == time_col_candidates[0] else col for i, col in enumerate(df.columns)]
    return df

def extract_month_from_file(file, df=None):
    """从文件名/数据中提取月份"""
    file_name = file.name
    month_patterns = [
        r'(\d{4})[-_年](\d{2})',
        r'(\d{6})',
    ]
    for pattern in month_patterns:
        match = re.search(pattern, file_name)
        if match:
            if len(match.groups()) == 2:
                year, month = match.groups()
                return f"{year}-{month}"
            elif len(match.groups()) == 1:
                num_str = match.group(1)
                if len(num_str) == 6:
                    year = num_str[:4]
                    month = num_str[4:]
                    return f"{year}-{month}"
    if df is not None and "时间" in df.columns and not df.empty:
        df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
        if not df["时间"].isna().all():
            first_date = df["时间"].dropna().iloc[0]
            return f"{first_date.year}-{first_date.month:02d}"
    now = datetime.now()
    return f"{now.year}-{now.month:02d}"

def to_excel(df, sheet_name="数据"):
    if df.empty:
        st.warning("⚠️ 数据为空，无法生成Excel文件")
        return BytesIO()
    df_export = force_unique_columns(df.copy())
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

# ========== 湖北专属工具函数（连续竞价调整） ==========
def clean_unit_name(unit_name):
    if pd.isna(unit_name) or unit_name == '':
        return ""
    unit_str = str(unit_name).strip()
    cleaned_str = re.sub(r'(\(.*?\)|（.*?）)', '', unit_str).strip()
    return cleaned_str

def truncate_to_two_decimal(x):
    if pd.isna(x):
        return None
    try:
        return float(int(float(x) * 100)) / 100
    except:
        return None

def format_worksheet(worksheet):
    alignment = Alignment(horizontal='center', vertical='center')
    for row in worksheet.iter_rows():
        for cell in row:
            cell.alignment = alignment
    for col in worksheet.columns:
        worksheet.column_dimensions[col[0].column_letter].width = 30

def extract_key_columns(df):
    key_columns = {
        '日期': None, '时段': None, '时段名称': None, '电量': None, '电价': None
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

# -------------------------- 省份专属配置 --------------------------
# ========== 湖北配置 ==========
def get_hubei_default_params():
    return {
        "风储一期": {"online": 0.8, "prefer": 0.725, "limit": 0.7, "mechanism": 0.0},
        "风储二期": {"online": 0.8, "prefer": 0.725, "limit": 0.7, "mechanism": 0.0},
        "栗溪": {"online": 0.8, "prefer": 0.725, "limit": 0.7, "mechanism": 0.0},
        "峪山一期": {"online": 0.8, "prefer": 0.725, "limit": 0.7, "mechanism": 0.0},
        "圣境山": {"online": 0.8, "prefer": 0.725, "limit": 0.7, "mechanism": 0.0},
        "襄北农光": {"online": 0.8, "prefer": 0.775, "limit": 0.8, "mechanism": 0.0},
        "浠水渔光": {"online": 0.8, "prefer": 0.775, "limit": 0.8, "mechanism": 0.0}
    }

def get_hubei_unit_mapping():
    return {
        "襄阳协合峪山泉水风电": "峪山一期",
        "荆门协合圣境山风电": "圣境山",
        "襄阳聚合光伏": "襄北农光",
        "三王风电": "风储一期",
        "荆门协合栗溪风电": "栗溪",
        "襄州协合三王风光储能电站风电二期": "风储二期",
        "浠水聚合关口光伏": "浠水渔光"
    }

HUBEI_STATION_TYPE_MAP = {
    "风电": ["荆门栗溪", "荆门圣境山", "襄北风储二期", "襄北风储一期", "襄州峪山一期"],
    "光伏": ["襄北农光", "浠水渔光"]
}

# ========== 贵州配置（可自定义） ==========
def get_guizhou_default_params():
    # 贵州连续竞价调整默认参数（按需修改）
    return {
        "贵州风电场1": {"online": 0.85, "prefer": 0.75, "limit": 0.65, "mechanism": 0.0},
        "贵州光伏场1": {"online": 0.82, "prefer": 0.78, "limit": 0.75, "mechanism": 0.0}
    }

def get_guizhou_unit_mapping():
    # 贵州交易单元映射（按需修改）
    return {
        "贵州风电场1": "贵州风电场1",
        "贵州光伏场1": "贵州光伏场1"
    }

GUIZHOU_STATION_TYPE_MAP = {
    "风电": ["贵州风电场1", "贵州风电场2"],
    "光伏": ["贵州光伏场1", "贵州光伏场2"]
}

# -------------------------- 核心功能模块 --------------------------
# ========== 模块1：连续竞价调整（支持多省份） ==========
def bidding_adjustment_module(province):
    st.title(f"🔧 {province} - 连续竞价调整")
    st.divider()
    
    # 加载省份专属配置
    if province == "湖北":
        DEFAULT_PARAMS = get_hubei_default_params()
        UNIT_MAPPING = get_hubei_unit_mapping()
    elif province == "贵州":
        DEFAULT_PARAMS = get_guizhou_default_params()
        UNIT_MAPPING = get_guizhou_unit_mapping()
    
    # 省份专属存储路径
    STORAGE_DIR = os.path.join(os.path.expanduser('~'), f'{province}_power_analysis_storage')
    CONTRACT_DIR = os.path.join(STORAGE_DIR, 'monthly_contracts')
    PARAM_SAVE_PATH = os.path.join(STORAGE_DIR, "station_params.json")
    os.makedirs(CONTRACT_DIR, exist_ok=True)
    
    # 初始化省份专属session_state
    province_data = st.session_state.province_data[province]["竞价调整"]
    if "station_params" not in province_data:
        def load_params():
            if os.path.exists(PARAM_SAVE_PATH):
                try:
                    with open(PARAM_SAVE_PATH, "r", encoding="utf-8") as f:
                        saved = json.load(f)
                    final = {}
                    for name in DEFAULT_PARAMS.keys():
                        final[name] = {**DEFAULT_PARAMS[name], **saved.get(name, {})}
                    return final
                except:
                    return DEFAULT_PARAMS.copy()
            return DEFAULT_PARAMS.copy()
        province_data["station_params"] = load_params()
    
    if "cached_editable_params" not in province_data:
        param_summary = []
        for name, params in province_data["station_params"].items():
            online = float(params.get("online", 0.8))
            prefer = float(params.get("prefer", 0.725))
            limit = float(params.get("limit", 0.7))
            mechanism = float(params.get("mechanism", 0.0))
            final_coeff = round(online - prefer - limit - mechanism, 6)
            param_summary.append({
                "场站名称": str(name),
                "上网电量折算系数": online,
                "优发优购比例": prefer,
                "限电率": limit,
                "机制电量比例": mechanism,
                "最终计算系数": final_coeff
            })
        province_data["cached_editable_params"] = pd.DataFrame(param_summary)
    
    # ========== 侧边栏文件管理 ==========
    with st.sidebar:
        st.subheader(f"📁 {province} - 合约文件管理")
        
        # 1. 批量上传合约文件
        new_contract_files = st.file_uploader(
            "选择合约文件（支持批量）",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key=f"{province}_contract_upload"
        )
        selected_month = st.text_input(
            "文件对应月份（2025-11）",
            value=datetime.now().strftime("%Y-%m"),
            key=f"{province}_contract_month"
        )
        
        def save_contract_file(file, month):
            safe_name = re.sub(r'[^\w\.-]', '_', file.name)
            path = os.path.join(CONTRACT_DIR, f"{month}_{safe_name}")
            with open(path, 'wb') as f:
                f.write(file.getbuffer())
            return path
        
        if st.button("保存月度文件", key=f"{province}_save_contract"):
            if not new_contract_files:
                st.warning("⚠️ 请选择文件！")
            elif not selected_month:
                st.warning("⚠️ 请输入月份！")
            else:
                with st.spinner("保存中..."):
                    saved = []
                    failed = []
                    for f in new_contract_files:
                        try:
                            save_contract_file(f, selected_month)
                            saved.append(f.name)
                        except Exception as e:
                            failed.append(f"{f.name}: {str(e)}")
                    if saved:
                        st.success(f"✅ 保存{len(saved)}个文件")
                    if failed:
                        st.error(f"❌ 失败{len(failed)}个文件")
        
        # 2. 选择分析月份
        def get_uploaded_months():
            months = set()
            for f in os.listdir(CONTRACT_DIR):
                if f.startswith(('2024-', '2025-')) and f.endswith(('.xlsx', '.xls')):
                    month = f.split('_')[0]
                    if len(month) == 7:
                        months.add(month)
            return sorted(list(months))
        
        uploaded_months = get_uploaded_months()
        if uploaded_months:
            province_data["selected_months"] = st.multiselect(
                "勾选分析月份",
                uploaded_months,
                default=uploaded_months,
                key=f"{province}_selected_months"
            )
            for m in uploaded_months:
                st.write(f"• {m}：{len(os.listdir(CONTRACT_DIR))}个文件")
        else:
            province_data["selected_months"] = []
            st.info("暂无上传文件")
        
        # 3. 上传功率预测文件
        province_data["forecast_file"] = st.file_uploader(
            "上传功率预测文件",
            type=["xlsx", "xls"],
            key=f"{province}_forecast_upload"
        )
    
    # ========== 主页面参数编辑 ==========
    editable_df = province_data["cached_editable_params"].copy()
    required_cols = ["场站名称", "上网电量折算系数", "优发优购比例", "限电率", "机制电量比例", "最终计算系数"]
    
    if editable_df.empty or not all(col in editable_df.columns for col in required_cols):
        param_summary = []
        for name, params in DEFAULT_PARAMS.items():
            final = params["online"] - params["prefer"] - params["limit"] - params["mechanism"]
            param_summary.append({
                "场站名称": name,
                "上网电量折算系数": params["online"],
                "优发优购比例": params["prefer"],
                "限电率": params["limit"],
                "机制电量比例": params["mechanism"],
                "最终计算系数": round(final, 6)
            })
        editable_df = pd.DataFrame(param_summary)
        province_data["cached_editable_params"] = editable_df
    
    # 可编辑表格
    st.subheader("📊 场站参数配置（编辑后保存生效）")
    edited_df = st.data_editor(
        editable_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "场站名称": st.column_config.TextColumn(disabled=True),
            "上网电量折算系数": st.column_config.NumberColumn(min_value=0, max_value=1, step=0.001, format="%.3f"),
            "优发优购比例": st.column_config.NumberColumn(min_value=0, max_value=1, step=0.001, format="%.3f"),
            "限电率": st.column_config.NumberColumn(min_value=0, max_value=1, step=0.001, format="%.3f"),
            "机制电量比例": st.column_config.NumberColumn(min_value=0, max_value=1, step=0.001, format="%.3f"),
            "最终计算系数": st.column_config.NumberColumn(disabled=True, format="%.6f")
        },
        key=f"{province}_params_editor"
    )
    
    # 保存参数按钮
    col1, col2 = st.columns([1, 9])
    with col1:
        if st.button("💾 保存参数", type="primary", key=f"{province}_save_params"):
            edited_df["最终计算系数"] = edited_df.apply(
                lambda x: round(float(x["上网电量折算系数"]) - float(x["优发优购比例"]) - float(x["限电率"]) - float(x["机制电量比例"]), 6),
                axis=1
            )
            province_data["cached_editable_params"] = edited_df
            
            updated_params = {}
            for _, row in edited_df.iterrows():
                updated_params[row["场站名称"]] = {
                    "online": float(row["上网电量折算系数"]),
                    "prefer": float(row["优发优购比例"]),
                    "limit": float(row["限电率"]),
                    "mechanism": float(row["机制电量比例"])
                }
            province_data["station_params"] = updated_params
            
            # 保存到本地
            try:
                with open(PARAM_SAVE_PATH, "w", encoding="utf-8") as f:
                    json.dump(updated_params, f, ensure_ascii=False, indent=4)
                st.success("✅ 参数保存成功！")
            except Exception as e:
                st.error(f"❌ 保存失败：{str(e)}")
    
    # ========== 测算功能 ==========
    selected_months = province_data.get("selected_months", [])
    forecast_file = province_data.get("forecast_file")
    run_disabled = not (selected_months and forecast_file)
    
    with col2:
        if st.button("🚀 开始测算", type="secondary", disabled=run_disabled, key=f"{province}_run_calc"):
            with st.spinner("测算中..."):
                # 核心测算逻辑（复用原有代码）
                def load_contract_files(months):
                    files = []
                    for m in months:
                        for f in os.listdir(CONTRACT_DIR):
                            if f.startswith(f"{m}_") and f.endswith(('.xlsx', '.xls')):
                                with open(os.path.join(CONTRACT_DIR, f), 'rb') as fp:
                                    bytes_io = BytesIO(fp.read())
                                    bytes_io.name = f
                                    files.append(bytes_io)
                    return files
                
                def generate_integrated_file(files, mapping):
                    unit_data = {u: [] for u in mapping.keys()}
                    for f in files:
                        try:
                            xls = pd.ExcelFile(f, engine='openpyxl')
                            for sheet in xls.sheet_names:
                                df = xls.parse(sheet)
                                if df.empty:
                                    continue
                                key_df = extract_key_columns(df)
                                if key_df.empty:
                                    continue
                                for idx, row in df.iterrows():
                                    try:
                                        raw_unit = row.iloc[0]
                                        cleaned = clean_unit_name(raw_unit)
                                        if cleaned not in mapping:
                                            continue
                                        key_row = key_df.iloc[idx:idx+1].copy()
                                        key_row['数据来源'] = f"文件：{f.name} | 工作表：{sheet}"
                                        unit_data[cleaned].append(key_row)
                                    except:
                                        continue
                        except:
                            continue
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for unit, name in mapping.items():
                            data = unit_data.get(unit, [])
                            if not data:
                                pd.DataFrame({"提示": [f"无数据：{unit}"]}).to_excel(writer, sheet_name=name, index=False)
                                format_worksheet(writer.sheets[name])
                                continue
                            merged = pd.concat(data, ignore_index=True)
                            if '日期' in merged.columns:
                                merged['日期'] = pd.to_datetime(merged['日期'], errors='coerce')
                                merged = merged.sort_values(by=['日期', '时段']).reset_index(drop=True)
                            for col in merged.columns:
                                if '电量' in col or '电价' in col:
                                    merged[col] = merged[col].apply(truncate_to_two_decimal)
                            merged.to_excel(writer, sheet_name=name, index=False)
                            format_worksheet(writer.sheets[name])
                    output.seek(0)
                    return output
                
                def process_power_forecast(f):
                    output = BytesIO()
                    try:
                        xls = pd.ExcelFile(f, engine='openpyxl')
                        today = date.today()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            for sheet in xls.sheet_names:
                                if sheet == '填写说明':
                                    continue
                                df = xls.parse(sheet)
                                if df.empty:
                                    continue
                                time_col = df.iloc[:, 0]
                                times = []
                                for t in time_col:
                                    try:
                                        times.append(pd.to_datetime(t).time())
                                    except:
                                        times.append(None)
                                valid = [t is not None for t in times]
                                times = [t for t in times if t is not None]
                                df = df[valid].reset_index(drop=True)
                                if not times:
                                    continue
                                processed = []
                                for col in df.columns[1:]:
                                    try:
                                        col_date = pd.to_datetime(col).date()
                                        if col_date >= today:
                                            col_data = df[col]
                                            avg_data = []
                                            for i in range(0, len(col_data), 4):
                                                seg = col_data[i:i+4]
                                                avg = seg.mean() if not seg.isna().all() else None
                                                avg_data.append(truncate_to_two_decimal(avg))
                                            if any(pd.notna(avg_data)):
                                                processed.append([col_date] + avg_data)
                                    except:
                                        continue
                                if not processed:
                                    continue
                                time_points = [time(hour=i) for i in range(24)]
                                cols = ['时间'] + [row[0] for row in processed]
                                proc_df = pd.DataFrame(columns=cols)
                                proc_df['时间'] = [t.strftime('%H:%M:%S') for t in time_points]
                                for i, row in enumerate(processed):
                                    col_name = row[0]
                                    for j in range(min(24, len(row[1:]))):
                                        proc_df.loc[j, col_name] = row[j+1]
                                proc_df = proc_df.dropna(axis=1, how='all')
                                proc_df.to_excel(writer, sheet_name=sheet, index=False)
                                format_worksheet(writer.sheets[sheet])
                    except:
                        pass
                    output.seek(0)
                    return output
                
                def calculate_difference(forecast, integrated, params):
                    coeff = {n: p["online"] - p["prefer"] - p["limit"] - p["mechanism"] for n, p in params.items()}
                    result = {}
                    try:
                        forecast_xls = pd.ExcelFile(forecast, engine='openpyxl')
                        integrated_xls = pd.ExcelFile(integrated, engine='openpyxl')
                        for sheet in forecast_xls.sheet_names:
                            if sheet == '填写说明' or sheet not in integrated_xls.sheet_names or sheet not in coeff:
                                continue
                            try:
                                f_df = forecast_xls.parse(sheet)
                                i_df = integrated_xls.parse(sheet)
                            except:
                                continue
                            if f_df.empty:
                                continue
                            current_coeff = coeff[sheet]
                            time_col = f_df.iloc[:, 0]
                            forecast_cols = f_df.columns[1:]
                            q_cols = [c for c in i_df.columns if '电量' in c]
                            p_cols = [c for c in i_df.columns if '电价' in c]
                            if not q_cols:
                                continue
                            q_col = q_cols[0]
                            p_col = p_cols[0] if p_cols else None
                            processed = []
                            for idx, row in f_df.iterrows():
                                if idx >= len(i_df):
                                    continue
                                current_time = row[0]
                                row_data = [current_time]
                                current_price = truncate_to_two_decimal(i_df.iloc[idx][p_col]) if (p_col and pd.notna(i_df.iloc[idx][p_col])) else None
                                for col in forecast_cols:
                                    f_val = row[col]
                                    row_data.append(f_val)
                                    try:
                                        q_val = i_df.iloc[idx][q_col]
                                        if pd.notna(f_val) and pd.notna(q_val):
                                            corrected = float(f_val) * current_coeff
                                            diff = truncate_to_two_decimal(corrected - float(q_val))
                                            diff = max(diff, -float(q_val)) if diff < 0 else diff
                                            row_data.append(diff)
                                        else:
                                            row_data.append(None)
                                    except:
                                        row_data.append(None)
                                row_data.append(current_price)
                                processed.append(row_data)
                            new_cols = ['时间']
                            for col in forecast_cols:
                                new_cols.extend([col, f'{col} (修正后差额)'])
                            new_cols.append('对应时段电价')
                            proc_df = pd.DataFrame(processed, columns=new_cols)
                            if '对应时段电价' in proc_df.columns:
                                proc_df['对应时段电价'] = proc_df['对应时段电价'].apply(truncate_to_two_decimal)
                            result[sheet] = proc_df
                    except Exception as e:
                        st.error(f"测算出错：{str(e)}")
                    return result, coeff
                
                # 执行测算
                contract_files = load_contract_files(selected_months)
                integrated_io = generate_integrated_file(contract_files, UNIT_MAPPING)
                forecast_processed = process_power_forecast(forecast_file)
                result_data, coeff = calculate_difference(forecast_processed, integrated_io, province_data["station_params"])
                
                # 展示结果
                st.divider()
                st.header("📈 测算结果")
                if result_data:
                    tabs = st.tabs(list(result_data.keys()))
                    for tab, (name, df) in zip(tabs, result_data.items()):
                        with tab:
                            st.subheader(f"📍 {name}（系数：{coeff[name]:.6f}）")
                            st.dataframe(
                                df,
                                use_container_width=True,
                                hide_index=True,
                                column_config={
                                    "时间": st.column_config.TextColumn(width="small"),
                                    "对应时段电价": st.column_config.NumberColumn(format="%.2f")
                                }
                            )
                            csv = df.to_csv(index=False, encoding="utf-8-sig")
                            st.download_button(
                                f"📥 下载{name}数据",
                                data=csv,
                                file_name=f"{name}_测算结果.csv",
                                mime="text/csv"
                            )
                else:
                    st.warning("暂无测算结果（数据不匹配）")
                st.success("✅ 测算完成！")
        
        if run_disabled:
            st.warning("⚠️ 请先上传合约文件+选择月份+上传预测文件")

# ========== 模块2：光伏/风电数据管理（支持多省份） ==========
def pv_wind_module(province):
    st.title(f"📈 {province} - 光伏/风电数据管理工具")
    st.divider()
    
    # 加载省份专属配置
    if province == "湖北":
        STATION_TYPE_MAP = HUBEI_STATION_TYPE_MAP
    elif province == "贵州":
        STATION_TYPE_MAP = GUIZHOU_STATION_TYPE_MAP
    
    # 初始化省份专属状态
    province_data = st.session_state.province_data[province]["光伏风电"]
    if "multi_month_data" not in province_data:
        province_data["multi_month_data"] = {}
    if "current_month" not in province_data:
        province_data["current_month"] = ""
    if "module_config" not in province_data:
        province_data["module_config"] = {
            "generated": {
                "time_col": 4, "wind_power_col": 9, "pv_power_col": 5,
                "pv_list": "浠水渔光,襄北农光" if province == "湖北" else "贵州光伏场1,贵州光伏场2",
                "conv": 1000, "skip_rows": 1, "keyword": "历史趋势"
            },
            "hold": {"hold_col": 3, "skip_rows": 1},
            "price": {"spot_col": 1, "wind_contract_col": 2, "pv_contract_col": 3, "skip_rows": 1}
        }
    
    # 工具函数（光伏风电专属）
    class DataProcessor:
        @staticmethod
        @st.cache_data(show_spinner="清洗数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
        def clean_power_value(value):
            if pd.isna(value):
                return None
            val_str = str(value).strip()
            num_match = re.search(r'(\d+\.?\d*)', val_str)
            if not num_match:
                return None
            try:
                return float(num_match.group(1))
            except:
                return None
        
        @staticmethod
        @st.cache_data(show_spinner="提取实发数据...", hash_funcs={BytesIO: lambda x: x.getvalue()})
        def extract_generated_data(file, config, station_type):
            try:
                power_col = config["wind_power_col"] if station_type == "风电" else config["pv_power_col"]
                suffix = file.name.split(".")[-1].lower()
                engine = "openpyxl" if suffix in ["xlsx", "xlsm"] else "xlrd"
                df = pd.read_excel(
                    BytesIO(file.getvalue()),
                    header=None,
                    usecols=[config["time_col"], power_col],
                    skiprows=config["skip_rows"],
                    engine=engine
                )
                df = force_unique_columns(df)
                df = df.iloc[:, :2]
                df.columns = ["时间", "功率(kW)"]
                df["功率(kW)"] = df["功率(kW)"].apply(DataProcessor.clean_power_value)
                df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
                df = df.dropna(subset=["时间", "功率(kW)"]).sort_values("时间").reset_index(drop=True)
                base_name = file.name.split(".")[0].split("-")[0].strip()
                month = extract_month_from_file(file, df)
                unique_name = f"{standardize_column_name(base_name)}_{month}"
                df[unique_name] = df["功率(kW)"] / config["conv"]
                return df[["时间", unique_name]].copy(), base_name, month
            except Exception as e:
                st.error(f"处理失败：{str(e)}")
                return pd.DataFrame(columns=["时间"]), "", ""
        
        @staticmethod
        @st.cache_data(show_spinner="提取持仓数据...", hash_funcs={BytesIO: lambda x: x.getvalue()})
        def extract_hold_data(file, config):
            try:
                suffix = file.name.split(".")[-1].lower()
                engine = "openpyxl" if suffix in ["xlsx", "xlsm"] else "xlrd"
                df = pd.read_excel(
                    BytesIO(file.getvalue()),
                    header=None,
                    usecols=[config["hold_col"]],
                    skiprows=config["skip_rows"],
                    engine=engine
                )
                df = force_unique_columns(df)
                df.columns = ["净持有电量"]
                df["净持有电量"] = pd.to_numeric(df["净持有电量"], errors="coerce").fillna(0)
                return round(df["净持有电量"].sum(), 2)
            except Exception as e:
                st.error(f"处理失败：{str(e)}")
                return 0.0
        
        @staticmethod
        @st.cache_data(show_spinner="提取电价数据...", hash_funcs={BytesIO: lambda x: x.getvalue()})
        def extract_price_data(file, config):
            try:
                suffix = file.name.split(".")[0].split("-")[-1].lower()
                engine = "openpyxl" if suffix in ["xlsx", "xlsm"] else "xlrd"
                df = pd.read_excel(
                    BytesIO(file.getvalue()),
                    header=None,
                    usecols=[0, config["spot_col"], config["wind_contract_col"], config["pv_contract_col"]],
                    skiprows=config["skip_rows"],
                    engine=engine,
                    nrows=24
                )
                df = force_unique_columns(df)
                df = df.iloc[:, :4]
                df.columns = ["时段", "现货均价(元/MWh)", "风电合约均价(元/MWh)", "光伏合约均价(元/MWh)"]
                df["时段"] = [f"{i:02d}:00" for i in range(24)]
                for col in ["现货均价(元/MWh)", "风电合约均价(元/MWh)", "光伏合约均价(元/MWh)"]:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
                return df
            except Exception as e:
                st.error(f"处理失败：{str(e)}")
                return pd.DataFrame()
        
        @staticmethod
        def calculate_24h_generated(merged_df, config):
            if merged_df.empty:
                st.warning("数据为空")
                return pd.DataFrame(), {}
            merged_df = force_unique_columns(merged_df)
            time_diff = merged_df["时间"].diff().dropna()
            avg_interval = time_diff.dt.total_seconds().mean() / 3600
            avg_interval = avg_interval if avg_interval > 0 else 1/4
            merged_df["时段"] = merged_df["时间"].dt.hour.apply(lambda x: f"{x:02d}:00")
            station_cols = [c for c in merged_df.columns if c not in ["时间", "时段"]]
            try:
                df_24h = merged_df.groupby("时段")[station_cols].apply(
                    lambda x: (x * avg_interval).sum()
                ).round(2).reset_index()
                df_24h = force_unique_columns(df_24h)
                total = {s: round(df_24h[s].sum(), 2) for s in station_cols if s in df_24h.columns}
                return df_24h, total
            except Exception as e:
                st.error(f"汇总失败：{str(e)}")
                return pd.DataFrame(), {}
        
        @staticmethod
        def calculate_excess_profit(gen_24h, hold_total, price_24h, month):
            if gen_24h.empty or not hold_total or price_24h.empty:
                st.warning("数据不完整")
                return pd.DataFrame()
            gen_24h = force_unique_columns(gen_24h)
            price_24h = force_unique_columns(price_24h)
            merged = pd.merge(gen_24h, price_24h, on="时段", how="inner")
            if merged.empty:
                st.warning("时段不匹配")
                return pd.DataFrame()
            result = []
            station_cols = [c for c in gen_24h.columns if c != "时段"]
            for station in station_cols:
                base_station = re.sub(r'_\d{4}-\d{2}$', '', station)
                base_station = re.sub(r'_[a-f0-9]{4,6}$', '', base_station)
                station_type = None
                contract_col = None
                for wind in STATION_TYPE_MAP["风电"]:
                    if wind in base_station or base_station in wind:
                        station_type = "风电"
                        contract_col = "风电合约均价(元/MWh)"
                        break
                if not station_type:
                    for pv in STATION_TYPE_MAP["光伏"]:
                        if pv in base_station or base_station in pv:
                            station_type = "光伏"
                            contract_col = "光伏合约均价(元/MWh)"
                            break
                if not station_type:
                    continue
                total_hold = 0
                for h_station, h_val in hold_total.items():
                    if h_station in base_station or base_station in h_station:
                        total_hold = h_val
                        break
                if total_hold == 0:
                    continue
                hourly_hold = total_hold / 24
                for _, row in merged.iterrows():
                    gen = row.get(station, 0)
                    spot = row.get("现货均价(元/MWh)", 0)
                    contract = row.get(contract_col, 0)
                    excess_qty = max(0, gen - hourly_hold)
                    excess_profit = excess_qty * (spot - contract)
                    if excess_profit > 0:
                        result.append({
                            "场站名称": base_station,
                            "场站类型": station_type,
                            "月份": month,
                            "时段": row["时段"],
                            "时段实发量(MWh)": round(gen, 2),
                            "时段持仓量(MWh)": round(hourly_hold, 2),
                            "超额电量(MWh)": round(excess_qty, 2),
                            "现货均价(元/MWh)": round(spot, 2),
                            "合约均价(元/MWh)": round(contract, 2),
                            "超额获利(元)": round(excess_profit, 2)
                        })
            return pd.DataFrame(result)
    
    # 获取当前月份数据
    def get_current_core_data():
        month = province_data["current_month"]
        if month not in province_data["multi_month_data"]:
            province_data["multi_month_data"][month] = {
                "generated": {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "total": {}},
                "hold": {"total": {}, "config": {}},
                "price": {"24h": pd.DataFrame(), "excess_profit": pd.DataFrame()}
            }
        return province_data["multi_month_data"][month]
    
    # ========== 月份选择 ==========
    col_month, _ = st.columns([2, 8])
    with col_month:
        all_months = list(province_data["multi_month_data"].keys())
        if all_months:
            province_data["current_month"] = st.selectbox(
                "📅 选择月份",
                all_months,
                key=f"{province}_pv_wind_month"
            )
        else:
            st.info("ℹ️ 暂无数据，请先上传文件")
    
    st.divider()
    
    # ========== 模块1：实发配置 ==========
    with st.expander("📊 模块1：场站实发配置", expanded=False):
        st.subheader("1.1 数据上传")
        col1, col2 = st.columns(2)
        with col1:
            station_type = st.radio("选择场站类型", ["风电", "光伏"], key=f"{province}_pv_wind_type")
            gen_files = st.file_uploader(
                f"上传{station_type}实发文件",
                accept_multiple_files=True,
                type=["xlsx", "xls", "xlsm"],
                key=f"{province}_pv_wind_gen_upload"
            )
        with col2:
            if gen_files:
                st.success(f"✅ 已上传{len(gen_files)}个文件")
                if st.button("处理实发数据", key=f"{province}_pv_wind_process_gen"):
                    file_month_map = {}
                    all_raw = {}
                    for f in gen_files:
                        df, station, month = DataProcessor.extract_generated_data(
                            f, province_data["module_config"]["generated"], station_type
                        )
                        if not df.empty and month:
                            if month not in file_month_map:
                                file_month_map[month] = []
                                all_raw[month] = []
                            file_month_map[month].append((df, station))
                            all_raw[month].append(df)
                    for month, dfs in all_raw.items():
                        if dfs:
                            merged = dfs[0].copy()
                            for df in dfs[1:]:
                                merged = pd.merge(merged, df, on="时间", how="outer")
                            merged = merged.sort_values("时间").dropna(subset=["时间"]).reset_index(drop=True)
                            core_data = get_current_core_data() if month == province_data["current_month"] else {
                                "generated": {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "total": {}},
                                "hold": {"total": {}, "config": {}},
                                "price": {"24h": pd.DataFrame(), "excess_profit": pd.DataFrame()}
                            }
                            core_data["generated"]["raw"] = merged
                            gen_24h, gen_total = DataProcessor.calculate_24h_generated(merged, province_data["module_config"]["generated"])
                            core_data["generated"]["24h"] = gen_24h
                            core_data["generated"]["total"] = gen_total
                            province_data["multi_month_data"][month] = core_data
                    st.success(f"✅ 处理完成！识别{len(file_month_map)}个月份")
                    if file_month_map and not province_data["current_month"]:
                        province_data["current_month"] = list(file_month_map.keys())[0]
        
        # 配置项
        st.subheader("1.2 列索引配置")
        col3, col4, col5 = st.columns(3)
        with col3:
            province_data["module_config"]["generated"]["time_col"] = st.number_input(
                "时间列索引", min_value=0, value=province_data["module_config"]["generated"]["time_col"], key=f"{province}_pv_wind_time_col"
            )
        with col4:
            province_data["module_config"]["generated"]["wind_power_col"] = st.number_input(
                "风电功率列索引", min_value=0, value=province_data["module_config"]["generated"]["wind_power_col"], key=f"{province}_pv_wind_wind_col"
            )
        with col5:
            province_data["module_config"]["generated"]["pv_power_col"] = st.number_input(
                "光伏功率列索引", min_value=0, value=province_data["module_config"]["generated"]["pv_power_col"], key=f"{province}_pv_wind_pv_col"
            )
        
        st.subheader("1.3 基础参数")
        col6, col7, col8 = st.columns(3)
        with col6:
            province_data["module_config"]["generated"]["conv"] = st.number_input(
                "功率转换系数", min_value=1, value=province_data["module_config"]["generated"]["conv"], key=f"{province}_pv_wind_conv"
            )
        with col7:
            province_data["module_config"]["generated"]["skip_rows"] = st.number_input(
                "跳过表头行数", min_value=0, value=province_data["module_config"]["generated"]["skip_rows"], key=f"{province}_pv_wind_skip_rows"
            )
        with col8:
            province_data["module_config"]["generated"]["pv_list"] = st.text_input(
                "光伏场站名单", value=province_data["module_config"]["generated"]["pv_list"], key=f"{province}_pv_wind_pv_list"
            )
        
        # 数据预览
        if province_data["current_month"]:
            core_data = get_current_core_data()
            if not core_data["generated"]["raw"].empty:
                st.subheader(f"📋 {province_data['current_month']} 实发数据预览")
                raw = force_unique_columns(core_data["generated"]["raw"].copy())
                gen_24h = force_unique_columns(core_data["generated"]["24h"].copy())
                tab1, tab2 = st.tabs(["原始数据", "24时段汇总"])
                with tab1:
                    st.dataframe(raw, use_container_width=True)
                    st.download_button(
                        f"下载{province_data['current_month']}原始数据",
                        data=to_excel(raw, f"{province_data['current_month']}原始数据"),
                        file_name=f"{province_data['current_month']}_实发原始数据.xlsx",
                        key=f"{province}_pv_wind_download_raw"
                    )
                with tab2:
                    st.dataframe(gen_24h, use_container_width=True)
                    st.download_button(
                        f"下载{province_data['current_month']}汇总数据",
                        data=to_excel(gen_24h, f"{province_data['current_month']}汇总数据"),
                        file_name=f"{province_data['current_month']}_实发汇总数据.xlsx",
                        key=f"{province}_pv_wind_download_24h"
                    )
    
    # ========== 模块2：持仓配置 ==========
    with st.expander("📦 模块2：中长期持仓配置", expanded=False):
        st.subheader("2.1 数据上传")
        col1, col2 = st.columns(2)
        with col1:
            hold_files = st.file_uploader(
                "上传持仓文件",
                accept_multiple_files=True,
                type=["xlsx", "xls", "xlsm"],
                key=f"{province}_pv_wind_hold_upload"
            )
        with col2:
            if hold_files and province_data["current_month"]:
                st.success(f"✅ 已上传{len(hold_files)}个文件")
                if st.button("处理持仓数据", key=f"{province}_pv_wind_process_hold"):
                    core_data = get_current_core_data()
                    hold_total = {}
                    for f in hold_files:
                        month = extract_month_from_file(f)
                        if month != province_data["current_month"]:
                            st.warning(f"文件{f.name}属于{month}，跳过")
                            continue
                        base_name = f.name.split(".")[0].split("-")[0].strip()
                        total = DataProcessor.extract_hold_data(f, province_data["module_config"]["hold"])
                        hold_total[standardize_column_name(base_name)] = total
                    core_data["hold"]["total"] = hold_total
                    province_data["multi_month_data"][province_data["current_month"]] = core_data
                    st.success("✅ 持仓数据处理完成！")
                    st.write(f"📊 {province_data['current_month']} 总持仓：")
                    st.write(hold_total)
        
        st.subheader("2.2 配置参数")
        province_data["module_config"]["hold"]["hold_col"] = st.number_input(
            "净持有电量列索引", min_value=0, value=province_data["module_config"]["hold"]["hold_col"], key=f"{province}_pv_wind_hold_col"
        )
        province_data["module_config"]["hold"]["skip_rows"] = st.number_input(
            "跳过表头行数", min_value=0, value=province_data["module_config"]["hold"]["skip_rows"], key=f"{province}_pv_wind_hold_skip"
        )
    
    # ========== 模块3：电价配置 ==========
    with st.expander("💰 模块3：月度电价配置", expanded=False):
        st.subheader("3.1 数据上传")
        col1, col2 = st.columns(2)
        with col1:
            price_file = st.file_uploader(
                "上传电价文件",
                accept_multiple_files=False,
                type=["xlsx", "xls", "xlsm"],
                key=f"{province}_pv_wind_price_upload"
            )
        with col2:
            if price_file and province_data["current_month"]:
                st.success("✅ 已上传电价文件")
                if st.button("处理电价数据", key=f"{province}_pv_wind_process_price"):
                    core_data = get_current_core_data()
                    price_df = DataProcessor.extract_price_data(price_file, province_data["module_config"]["price"])
                    core_data["price"]["24h"] = price_df
                    province_data["multi_month_data"][province_data["current_month"]] = core_data
                    st.success("✅ 电价数据处理完成！")
        
        st.subheader("3.2 列索引配置")
        col3, col4, col5 = st.columns(3)
        with col3:
            province_data["module_config"]["price"]["spot_col"] = st.number_input(
                "现货均价列索引", min_value=0, value=province_data["module_config"]["price"]["spot_col"], key=f"{province}_pv_wind_spot_col"
            )
        with col4:
            province_data["module_config"]["price"]["wind_contract_col"] = st.number_input(
                "风电合约列索引", min_value=0, value=province_data["module_config"]["price"]["wind_contract_col"], key=f"{province}_pv_wind_wind_contract_col"
            )
        with col5:
            province_data["module_config"]["price"]["pv_contract_col"] = st.number_input(
                "光伏合约列索引", min_value=0, value=province_data["module_config"]["price"]["pv_contract_col"], key=f"{province}_pv_wind_pv_contract_col"
            )
        
        # 电价预览
        if province_data["current_month"]:
            core_data = get_current_core_data()
            if not core_data["price"]["24h"].empty:
                st.subheader(f"📋 {province_data['current_month']} 电价数据预览")
                price_df = force_unique_columns(core_data["price"]["24h"].copy())
                st.dataframe(price_df, use_container_width=True)
                st.download_button(
                    f"下载{province_data['current_month']}电价数据",
                    data=to_excel(price_df, f"{province_data['current_month']}电价数据"),
                    file_name=f"{province_data['current_month']}_电价数据.xlsx",
                    key=f"{province}_pv_wind_download_price"
                )
    
    # ========== 模块4：超额获利计算 ==========
    if province_data["current_month"]:
        st.subheader(f"🎯 {province_data['current_month']} 超额获利计算")
        core_data = get_current_core_data()
        if st.button("计算超额获利", key=f"{province}_pv_wind_calc_profit"):
            profit_df = DataProcessor.calculate_excess_profit(
                core_data["generated"]["24h"],
                core_data["hold"]["total"],
                core_data["price"]["24h"],
                province_data["current_month"]
            )
            core_data["price"]["excess_profit"] = profit_df
            province_data["multi_month_data"][province_data["current_month"]] = core_data
            
            if not profit_df.empty:
                st.success("✅ 计算完成！")
                profit_df = force_unique_columns(profit_df)
                st.dataframe(profit_df, use_container_width=True)
                total_profit = profit_df["超额获利(元)"].sum()
                st.metric(f"💰 总超额获利", value=round(total_profit, 2))
                st.download_button(
                    f"下载{province_data['current_month']}获利数据",
                    data=to_excel(profit_df, f"{province_data['current_month']}获利数据"),
                    file_name=f"{province_data['current_month']}_超额获利数据.xlsx",
                    key=f"{province}_pv_wind_download_profit"
                )
                # 可视化
                fig = px.bar(
                    profit_df,
                    x="时段",
                    y="超额获利(元)",
                    color="场站名称",
                    title=f"{province_data['current_month']} 分时段超额获利",
                    barmode="group"
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("ℹ️ 暂无超额获利数据")

# -------------------------- 主程序入口 --------------------------
def main():
    # 侧边栏：省份选择 + 功能菜单
    with st.sidebar:
        st.title("🌍 多省份新能源管理系统")
        st.divider()
        
        # 1. 省份选择
        st.session_state[CURRENT_PROVINCE_KEY] = st.selectbox(
            "选择省份",
            PROVINCES,
            index=PROVINCES.index(st.session_state[CURRENT_PROVINCE_KEY]),
            key="province_selector"
        )
        
        st.divider()
        
        # 2. 功能菜单
        st.session_state[CURRENT_FUNCTION_KEY] = st.radio(
            "选择功能模块",
            ["连续竞价调整", "光伏/风电数据管理"],
            index=0 if st.session_state[CURRENT_FUNCTION_KEY] == "连续竞价调整" else 1,
            key="function_selector"
        )
        
        st.divider()
        st.info("💡 切换省份/功能后，数据将自动隔离存储")
    
    # 根据选择加载对应模块
    current_province = st.session_state[CURRENT_PROVINCE_KEY]
    current_function = st.session_state[CURRENT_FUNCTION_KEY]
    
    if current_function == "连续竞价调整":
        bidding_adjustment_module(current_province)
    elif current_function == "光伏/风电数据管理":
        pv_wind_module(current_province)

if __name__ == "__main__":
    main()
