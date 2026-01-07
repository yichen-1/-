import streamlit as st
import pandas as pd
import re
import uuid
from io import BytesIO
import datetime
import plotly.express as px

# -------------------------- 1. 页面基础配置 --------------------------
st.set_page_config(
    page_title="光伏/风电数据管理工具（多月份版）",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------- 2. 全局常量与映射 --------------------------
STATION_TYPE_MAP = {
    "风电": ["荆门栗溪", "荆门圣境山", "襄北风储二期", "襄北风储一期", "襄州峪山一期"],
    "光伏": ["襄北农光", "浠水渔光"]
}

# 新增：电价模板列名常量
PRICE_TEMPLATE_COLS = [
    "时段", 
    "风电现货均价(元/MWh)", 
    "风电合约均价(元/MWh)", 
    "光伏现货均价(元/MWh)", 
    "光伏合约均价(元/MWh)"
]

# -------------------------- 3. 核心工具函数 --------------------------
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
    # 固定时间列名
    time_col_candidates = [i for i, col in enumerate(df.columns) if "时间" in col or "date" in col.lower()]
    if time_col_candidates:
        df.columns = ["时间" if i == time_col_candidates[0] else col for i, col in enumerate(df.columns)]
    return df

def extract_month_from_file(file, df=None):
    """从文件名/数据中提取月份（优先级：文件名 > 时间列）"""
    # 1. 从文件名提取（支持202501、2025-01、2025年01月等格式）
    file_name = file.name
    month_patterns = [
        r'(\d{4})[-_年](\d{2})',  # 2025-01 / 2025_01 / 2025年01
        r'(\d{6})',              # 202501（6位数字，前4年+后2月）
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
    
    # 2. 从时间列提取
    if df is not None and "时间" in df.columns and not df.empty:
        df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
        if not df["时间"].isna().all():
            first_date = df["时间"].dropna().iloc[0]
            return f"{first_date.year}-{first_date.month:02d}"
    
    # 3. 默认当前月份
    now = datetime.datetime.now()
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

# 新增：生成电价模板函数
def generate_price_template():
    """生成标准电价模板（24时段，5列）"""
    template_data = []
    for hour in range(24):
        template_data.append({
            "时段": f"{hour:02d}:00",
            "风电现货均价(元/MWh)": 0.0,
            "风电合约均价(元/MWh)": 0.0,
            "光伏现货均价(元/MWh)": 0.0,
            "光伏合约均价(元/MWh)": 0.0
        })
    df_template = pd.DataFrame(template_data)
    return df_template

# -------------------------- 4. 会话状态初始化（按月份存储） --------------------------
if "multi_month_data" not in st.session_state:
    st.session_state.multi_month_data = {}  # 结构：{"2025-01": core_data, "2025-02": core_data}
if "current_month" not in st.session_state:
    st.session_state.current_month = ""  # 当前选中的月份
if "module_config" not in st.session_state:
    st.session_state.module_config = {
        "generated": {
            "time_col": 4, "wind_power_col": 9, "pv_power_col": 5,
            "pv_list": "浠水渔光,襄北农光", "conv": 1000, "skip_rows": 1, "keyword": "历史趋势"
        },
        "hold": {"hold_col": 3, "skip_rows": 1},
        # 修改：更新电价配置列索引（对应5列模板）
        "price": {
            "wind_spot_col": 1,      # 风电现货均价列
            "wind_contract_col": 2,  # 风电合约均价列
            "pv_spot_col": 3,        # 光伏现货均价列
            "pv_contract_col": 4,    # 光伏合约均价列
            "skip_rows": 1
        }
    }

# 获取当前选中月份的核心数据
def get_current_core_data():
    if st.session_state.current_month not in st.session_state.multi_month_data:
        # 初始化空的core_data结构
        st.session_state.multi_month_data[st.session_state.current_month] = {
            "generated": {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "total": {}},
            "hold": {"total": {}, "config": {}},
            "price": {"24h": pd.DataFrame(), "excess_profit": pd.DataFrame()}
        }
    return st.session_state.multi_month_data[st.session_state.current_month]

# -------------------------- 5. 核心数据处理类 --------------------------
class DataProcessor:
    @staticmethod
    @st.cache_data(show_spinner="清洗功率数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
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
    @st.cache_data(show_spinner="提取实发数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
    def extract_generated_data(file, config, station_type):
        try:
            power_col = config["wind_power_col"] if station_type == "风电" else config["pv_power_col"]
            file_suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                usecols=[config["time_col"], power_col],
                skiprows=config["skip_rows"],
                engine=engine,
                nrows=None
            )
            
            df = force_unique_columns(df)
            df = df.iloc[:, :2]
            df.columns = ["时间", "功率(kW)"]

            df["功率(kW)"] = df["功率(kW)"].apply(DataProcessor.clean_power_value)
            df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
            df = df.dropna(subset=["时间", "功率(kW)"]).sort_values("时间").reset_index(drop=True)

            # 生成唯一场站名（包含月份标识）
            base_name = file.name.split(".")[0].split("-")[0].strip()
            month = extract_month_from_file(file, df)
            unique_station_name = f"{standardize_column_name(base_name)}_{month}"
            df[unique_station_name] = df["功率(kW)"] / config["conv"]

            df_result = df[["时间", unique_station_name]].copy()
            df_result = force_unique_columns(df_result)
            
            return df_result, base_name, month
        except Exception as e:
            st.error(f"❌ 实发文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame(columns=["时间"]), "", ""

    @staticmethod
    @st.cache_data(show_spinner="提取持仓数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
    def extract_hold_data(file, config):
        try:
            file_suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                usecols=[config["hold_col"]],
                skiprows=config["skip_rows"],
                engine=engine,
                nrows=None
            )
            df = force_unique_columns(df)
            df.columns = ["净持有电量"]
            df["净持有电量"] = pd.to_numeric(df["净持有电量"], errors="coerce").fillna(0)
            total_hold = round(df["净持有电量"].sum(), 2)
            return total_hold
        except Exception as e:
            st.error(f"❌ 持仓文件[{file.name}]处理失败：{str(e)}")
            return 0.0

    @staticmethod
    @st.cache_data(show_spinner="提取电价数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
    def extract_price_data(file, config):
        try:
            file_suffix = file.name.split(".")[-1].lower()  # 修复：原代码错误取文件名前缀作为后缀
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                # 修改：读取5列（时段+4个价格列）
                usecols=[0, config["wind_spot_col"], config["wind_contract_col"], 
                         config["pv_spot_col"], config["pv_contract_col"]],
                skiprows=config["skip_rows"],
                engine=engine,
                nrows=24
            )
            df = force_unique_columns(df)
            df = df.iloc[:, :5]
            # 修改：使用标准模板列名
            df.columns = PRICE_TEMPLATE_COLS
            
            # 标准化时段列
            df["时段"] = [f"{i:02d}:00" for i in range(24)]
            # 价格列转数值型
            price_cols = df.columns[1:]  # 除时段外的所有列
            for col in price_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            
            return df
        except Exception as e:
            st.error(f"❌ 电价文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame()

    @staticmethod
    def calculate_24h_generated(merged_raw_df, config):
        if merged_raw_df.empty:
            st.warning("⚠️ 实发原始数据为空，无法计算24时段汇总")
            return pd.DataFrame(), {}

        merged_raw_df = force_unique_columns(merged_raw_df)
        
        time_diff = merged_raw_df["时间"].diff().dropna()
        avg_interval_h = time_diff.dt.total_seconds().mean() / 3600
        avg_interval_h = avg_interval_h if avg_interval_h > 0 else 1/4

        merged_raw_df["时段"] = merged_raw_df["时间"].dt.hour.apply(lambda x: f"{x:02d}:00")
        station_cols = [col for col in merged_raw_df.columns if col not in ["时间", "时段"]]
        
        try:
            generated_24h_df = merged_raw_df.groupby("时段")[station_cols].apply(
                lambda x: (x * avg_interval_h).sum()
            ).round(2).reset_index()
            generated_24h_df = force_unique_columns(generated_24h_df)
        except Exception as e:
            st.error(f"❌ 24时段汇总失败：{str(e)}")
            return pd.DataFrame(), {}

        monthly_total = {
            station: round(generated_24h_df[station].sum(), 2)
            for station in station_cols if station in generated_24h_df.columns
        }

        return generated_24h_df, monthly_total

    @staticmethod
    def calculate_excess_profit(generated_24h_df, hold_total_dict, price_24h_df, current_month):
        """
        修改后的超额获利计算逻辑：
        1. 实发电量：24时段分时总电量 × 系数（风电0.7，光伏0.8）
        2. 合约电量校核：
           - 若实发电量 > 合约电量×0.9 → 差额 = 实发电量 - 合约电量×1.1
           - 若实发电量 < 合约电量×0.9 → 差额 = 实发电量 - 合约电量×0.9
        3. 价格差值：现货均价 - 合约均价（风电/光伏分别计算）
        4. 超额获利 = 电量差额 × 价格差值
        """
        if generated_24h_df.empty or not hold_total_dict or price_24h_df.empty:
            st.warning("⚠️ 实发/持仓/电价数据不完整，无法计算超额获利")
            return pd.DataFrame()

        generated_24h_df = force_unique_columns(generated_24h_df)
        price_24h_df = force_unique_columns(price_24h_df)
        
        merged_df = pd.merge(generated_24h_df, price_24h_df, on="时段", how="inner")
        merged_df = force_unique_columns(merged_df)
        if merged_df.empty:
            st.warning("⚠️ 实发与电价数据时段不匹配，无法计算")
            return pd.DataFrame()

        result_rows = []
        station_cols = [col for col in generated_24h_df.columns if col != "时段"]

        for station in station_cols:
            # 提取原始场站名（去掉月份后缀）
            base_station = re.sub(r'_\d{4}-\d{2}$', '', station)
            base_station = re.sub(r'_[a-f0-9]{4,6}$', '', base_station)
            station_type = None
            
            # 匹配场站类型
            for wind_station in STATION_TYPE_MAP["风电"]:
                if wind_station in base_station or base_station in wind_station:
                    station_type = "风电"
                    # 修改：匹配风电价格列
                    spot_col = "风电现货均价(元/MWh)"
                    contract_col = "风电合约均价(元/MWh)"
                    gen_coeff = 0.7  # 风电发电量系数
                    break
            if not station_type:
                for pv_station in STATION_TYPE_MAP["光伏"]:
                    if pv_station in base_station or base_station in pv_station:
                        station_type = "光伏"
                        # 修改：匹配光伏价格列
                        spot_col = "光伏现货均价(元/MWh)"
                        contract_col = "光伏合约均价(元/MWh)"
                        gen_coeff = 0.8  # 光伏发电量系数
                        break
            
            if not station_type:
                continue

            # 匹配当前月份的持仓数据
            total_hold = 0
            for hold_station, hold_value in hold_total_dict.items():
                if hold_station in base_station or base_station in hold_station:
                    total_hold = hold_value
                    break
            if total_hold == 0:
                continue
                
            # 计算分时合约电量（总持仓/24）
            hourly_hold = total_hold / 24

            for _, row in merged_df.iterrows():
                # 1. 计算修正后实发电量（分时电量 × 场站类型系数）
                hourly_generated_raw = row.get(station, 0)
                hourly_generated = hourly_generated_raw * gen_coeff  # 风电×0.7，光伏×0.8
                
                # 2. 合约电量校核
                hold_09 = hourly_hold * 0.9  # 合约电量0.9倍
                hold_11 = hourly_hold * 1.1  # 合约电量1.1倍
                
                if hourly_generated > hold_09:
                    quantity_diff = hourly_generated - hold_11  # 大于0.9倍，与1.1倍做差
                else:
                    quantity_diff = hourly_generated - hold_09  # 小于0.9倍，与0.9倍做差
                
                # 3. 价格差值（现货-合约）
                spot_price = row.get(spot_col, 0)
                contract_price = row.get(contract_col, 0)
                price_diff = spot_price - contract_price
                
                # 4. 计算超额获利
                excess_profit = quantity_diff * price_diff

                result_rows.append({
                    "场站名称": base_station,
                    "场站类型": station_type,
                    "月份": current_month,
                    "时段": row["时段"],
                    "原始分时实发量(MWh)": round(hourly_generated_raw, 2),
                    "修正后实发量(MWh)": round(hourly_generated, 2),  # 新增：修正后电量
                    "分时合约电量(MWh)": round(hourly_hold, 2),
                    "合约电量0.9倍(MWh)": round(hold_09, 2),          # 新增：校核阈值
                    "合约电量1.1倍(MWh)": round(hold_11, 2),          # 新增：校核阈值
                    "电量差额(MWh)": round(quantity_diff, 2),         # 新增：电量差额
                    f"{station_type}现货均价(元/MWh)": round(spot_price, 2),
                    f"{station_type}合约均价(元/MWh)": round(contract_price, 2),
                    f"{station_type}价格差值(元/MWh)": round(price_diff, 2),  # 新增：价格差值
                    "超额获利(元)": round(excess_profit, 2)
                })

        result_df = pd.DataFrame(result_rows)
        result_df = force_unique_columns(result_df)
        return result_df

# -------------------------- 6. 页面布局 --------------------------
st.title("📈 光伏/风电数据管理工具（多月份版）")

# 月份选择器（核心新增）
col_month, col_refresh = st.columns([2, 8])
with col_month:
    all_months = list(st.session_state.multi_month_data.keys())
    if all_months:
        st.session_state.current_month = st.selectbox(
            "📅 选择月份",
            all_months,
            key="month_selector"
        )
    else:
        st.info("ℹ️ 暂无数据，请先上传文件")

st.divider()

# ====================== 模块1：场站实发配置 ======================
with st.expander("📊 模块1：场站实发配置", expanded=False):
    st.subheader("1.1 数据上传")
    col1_1, col1_2 = st.columns(2)
    with col1_1:
        station_type = st.radio("选择场站类型", ["风电", "光伏"], key="gen_station_type")
        gen_files = st.file_uploader(
            f"上传{station_type}实发数据文件（支持多月份）",
            accept_multiple_files=True,
            type=["xlsx", "xls", "xlsm"],
            key="gen_file_upload"
        )
    with col1_2:
        if gen_files:
            st.success(f"✅ 已上传{len(gen_files)}个{station_type}实发文件")
            if st.button("📝 处理实发数据", key="process_gen_data"):
                file_month_map = {}  # 按月份分组存储数据
                all_raw_dfs = {}
                
                # 逐个处理文件并按月份分组
                for file in gen_files:
                    df, station, month = DataProcessor.extract_generated_data(
                        file, st.session_state.module_config["generated"], station_type
                    )
                    if not df.empty and "时间" in df.columns and month:
                        if month not in file_month_map:
                            file_month_map[month] = []
                            all_raw_dfs[month] = []
                        file_month_map[month].append((df, station))
                        all_raw_dfs[month].append(df)
                
                # 按月份合并数据
                for month, dfs in all_raw_dfs.items():
                    if dfs:
                        merged_raw = dfs[0].copy()
                        merged_raw = force_unique_columns(merged_raw)
                        
                        for df in dfs[1:]:
                            df = force_unique_columns(df)
                            merged_raw = pd.merge(merged_raw, df, on="时间", how="outer")
                            merged_raw = force_unique_columns(merged_raw)
                        
                        merged_raw = merged_raw.sort_values("时间").reset_index(drop=True)
                        merged_raw = merged_raw.dropna(subset=["时间"])
                        merged_raw = force_unique_columns(merged_raw)
                        
                        # 存储到对应月份
                        core_data = get_current_core_data() if month == st.session_state.current_month else {
                            "generated": {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "total": {}},
                            "hold": {"total": {}, "config": {}},
                            "price": {"24h": pd.DataFrame(), "excess_profit": pd.DataFrame()}
                        }
                        core_data["generated"]["raw"] = merged_raw
                        
                        # 计算24时段汇总
                        gen_24h, gen_total = DataProcessor.calculate_24h_generated(
                            merged_raw, st.session_state.module_config["generated"]
                        )
                        core_data["generated"]["24h"] = gen_24h
                        core_data["generated"]["total"] = gen_total
                        
                        # 更新会话状态
                        st.session_state.multi_month_data[month] = core_data
                
                st.success(f"✅ 处理完成！共识别{len(file_month_map)}个月份数据：{list(file_month_map.keys())}")
                # 自动选中第一个月份
                if file_month_map and not st.session_state.current_month:
                    st.session_state.current_month = list(file_month_map.keys())[0]

    st.subheader("1.2 列索引配置（索引从0开始）")
    col1_3, col1_4, col1_5 = st.columns(3)
    with col1_3:
        st.session_state.module_config["generated"]["time_col"] = st.number_input(
            "时间列索引", min_value=0, value=st.session_state.module_config["generated"]["time_col"], key="gen_time_col"
        )
    with col1_4:
        st.session_state.module_config["generated"]["wind_power_col"] = st.number_input(
            "风电功率列索引", min_value=0, value=st.session_state.module_config["generated"]["wind_power_col"], key="gen_wind_col"
        )
    with col1_5:
        st.session_state.module_config["generated"]["pv_power_col"] = st.number_input(
            "光伏功率列索引", min_value=0, value=st.session_state.module_config["generated"]["pv_power_col"], key="gen_pv_col"
        )

    st.subheader("1.3 基础参数配置")
    col1_6, col1_7, col1_8 = st.columns(3)
    with col1_6:
        st.session_state.module_config["generated"]["conv"] = st.number_input(
            "功率转换系数（kW→MW）", min_value=1, value=st.session_state.module_config["generated"]["conv"], key="gen_conv"
        )
    with col1_7:
        st.session_state.module_config["generated"]["skip_rows"] = st.number_input(
            "跳过表头行数", min_value=0, value=st.session_state.module_config["generated"]["skip_rows"], key="gen_skip_rows"
        )
    with col1_8:
        st.session_state.module_config["generated"]["pv_list"] = st.text_input(
            "光伏场站名单（逗号分隔）", value=st.session_state.module_config["generated"]["pv_list"], key="gen_pv_list"
        )

    # 数据预览（当前月份）
    if st.session_state.current_month:
        core_data = get_current_core_data()
        if not core_data["generated"]["raw"].empty:
            st.subheader(f"📋 {st.session_state.current_month} 实发数据预览")
            display_raw = force_unique_columns(core_data["generated"]["raw"].copy())
            display_raw.columns = [str(col) for col in display_raw.columns]
            display_raw = display_raw.reset_index(drop=True)
            
            display_24h = force_unique_columns(core_data["generated"]["24h"].copy())
            display_24h.columns = [str(col) for col in display_24h.columns]
            display_24h = display_24h.reset_index(drop=True)
            
            tab1, tab2 = st.tabs(["原始数据", "24时段汇总"])
            with tab1:
                st.dataframe(display_raw, use_container_width=True)
                st.download_button(
                    f"💾 下载{st.session_state.current_month}原始实发数据",
                    data=to_excel(display_raw, f"{st.session_state.current_month}原始实发数据"),
                    file_name=f"实发原始数据_{st.session_state.current_month}.xlsx",
                    key="download_gen_raw"
                )
            with tab2:
                st.dataframe(display_24h, use_container_width=True)
                st.download_button(
                    f"💾 下载{st.session_state.current_month}24时段汇总数据",
                    data=to_excel(display_24h, f"{st.session_state.current_month}24时段实发汇总"),
                    file_name=f"24时段实发汇总_{st.session_state.current_month}.xlsx",
                    key="download_gen_24h"
                )

st.divider()

# ====================== 模块2：中长期持仓配置 ======================
with st.expander("📦 模块2：中长期持仓配置", expanded=False):
    st.subheader("2.1 数据上传")
    col2_1, col2_2 = st.columns(2)
    with col2_1:
        hold_files = st.file_uploader(
            "上传持仓数据文件（支持多月份）",
            accept_multiple_files=True,
            type=["xlsx", "xls", "xlsm"],
            key="hold_file_upload"
        )
    with col2_2:
        if hold_files and st.session_state.current_month:
            st.success(f"✅ 已上传{len(hold_files)}个持仓文件")
            if st.button("📝 处理持仓数据", key="process_hold_data"):
                core_data = get_current_core_data()
                hold_total = {}
                for file in hold_files:
                    # 提取文件对应的月份
                    month = extract_month_from_file(file)
                    if month != st.session_state.current_month:
                        st.warning(f"⚠️ 文件[{file.name}]属于{month}，当前选中{st.session_state.current_month}，跳过")
                        continue
                    base_name = file.name.split(".")[0].split("-")[0].strip()
                    standard_name = standardize_column_name(base_name)
                    total = DataProcessor.extract_hold_data(file, st.session_state.module_config["hold"])
                    hold_total[standard_name] = total
                core_data["hold"]["total"] = hold_total
                st.session_state.multi_month_data[st.session_state.current_month] = core_data
                st.success("✅ 持仓数据处理完成！")
                st.write(f"📊 {st.session_state.current_month} 各场站月度总持仓（MWh）：")
                st.write(hold_total)

    st.subheader("2.2 配置参数")
    col2_3 = st.columns(1)[0]
    with col2_3:
        st.session_state.module_config["hold"]["hold_col"] = st.number_input(
            "净持有电量列索引（0开始）", min_value=0, value=st.session_state.module_config["hold"]["hold_col"], key="hold_col"
        )
        st.session_state.module_config["hold"]["skip_rows"] = st.number_input(
            "跳过表头行数", min_value=0, value=st.session_state.module_config["hold"]["skip_rows"], key="hold_skip_rows"
        )

st.divider()

# ====================== 模块3：月度电价配置（修改后） ======================
with st.expander("💰 模块3：月度电价配置", expanded=False):
    st.subheader("3.1 标准模板下载")
    # 新增：模板导出功能
    col3_0 = st.columns(1)[0]
    with col3_0:
        price_template_df = generate_price_template()
        st.download_button(
            "📥 下载电价标准模板（24时段）",
            data=to_excel(price_template_df, "电价标准模板"),
            file_name="电价标准模板.xlsx",
            key="download_price_template"
        )
        st.info("💡 模板包含5列：时段、风电现货均价、风电合约均价、光伏现货均价、光伏合约均价，请按此格式填写后上传")

    st.subheader("3.2 数据上传")
    col3_1, col3_2 = st.columns(2)
    with col3_1:
        price_file = st.file_uploader(
            "上传电价数据文件（请使用标准模板）",
            accept_multiple_files=False,
            type=["xlsx", "xls", "xlsm"],
            key="price_file_upload"
        )
    with col3_2:
        if price_file and st.session_state.current_month:
            st.success("✅ 已上传电价数据文件")
            if st.button("📝 处理电价数据", key="process_price_data"):
                core_data = get_current_core_data()
                # 提取文件月份
                price_df = DataProcessor.extract_price_data(price_file, st.session_state.module_config["price"])
                price_df = force_unique_columns(price_df)
                core_data["price"]["24h"] = price_df
                st.session_state.multi_month_data[st.session_state.current_month] = core_data
                st.success("✅ 电价数据处理完成！")

    st.subheader("3.3 列索引配置（索引从0开始）")
    # 修改：更新电价列索引配置项
    col3_3, col3_4, col3_5, col3_6 = st.columns(4)
    with col3_3:
        st.session_state.module_config["price"]["wind_spot_col"] = st.number_input(
            "风电现货均价列索引", min_value=0, value=st.session_state.module_config["price"]["wind_spot_col"], key="price_wind_spot_col"
        )
    with col3_4:
        st.session_state.module_config["price"]["wind_contract_col"] = st.number_input(
            "风电合约均价列索引", min_value=0, value=st.session_state.module_config["price"]["wind_contract_col"], key="price_wind_contract_col"
        )
    with col3_5:
        st.session_state.module_config["price"]["pv_spot_col"] = st.number_input(
            "光伏现货均价列索引", min_value=0, value=st.session_state.module_config["price"]["pv_spot_col"], key="price_pv_spot_col"
        )
    with col3_6:
        st.session_state.module_config["price"]["pv_contract_col"] = st.number_input(
            "光伏合约均价列索引", min_value=0, value=st.session_state.module_config["price"]["pv_contract_col"], key="price_pv_contract_col"
        )
    
    # 补充：跳过行数配置
    col3_7 = st.columns(1)[0]
    with col3_7:
        st.session_state.module_config["price"]["skip_rows"] = st.number_input(
            "跳过表头行数", min_value=0, value=st.session_state.module_config["price"]["skip_rows"], key="price_skip_rows"
        )

    # 电价数据预览（当前月份）
    if st.session_state.current_month:
        core_data = get_current_core_data()
        if not core_data["price"]["24h"].empty:
            st.subheader(f"📋 {st.session_state.current_month} 24时段电价数据预览")
            display_price = force_unique_columns(core_data["price"]["24h"].copy())
            display_price.columns = [str(col) for col in display_price.columns]
            display_price = display_price.reset_index(drop=True)
            st.dataframe(display_price, use_container_width=True)
            st.download_button(
                f"💾 下载{st.session_state.current_month}电价数据",
                data=to_excel(display_price, f"{st.session_state.current_month}24时段电价数据"),
                file_name=f"24时段电价数据_{st.session_state.current_month}.xlsx",
                key="download_price_24h"
            )

st.divider()

# ====================== 模块4：超额获利计算 ======================
if st.session_state.current_month:
    st.subheader(f"🎯 {st.session_state.current_month} 超额获利回收计算")
    core_data = get_current_core_data()
    if st.button("🔍 计算超额获利", key="calc_excess_profit"):
        excess_profit_df = DataProcessor.calculate_excess_profit(
            core_data["generated"]["24h"],
            core_data["hold"]["total"],
            core_data["price"]["24h"],
            st.session_state.current_month
        )
        core_data["price"]["excess_profit"] = excess_profit_df
        st.session_state.multi_month_data[st.session_state.current_month] = core_data

        if not excess_profit_df.empty:
            st.success("✅ 超额获利计算完成！")
            display_profit = force_unique_columns(excess_profit_df.copy())
            display_profit.columns = [str(col) for col in display_profit.columns]
            display_profit = display_profit.reset_index(drop=True)
            st.dataframe(display_profit, use_container_width=True)
            
            total_profit = display_profit["超额获利(元)"].sum()
            st.metric(f"💰 {st.session_state.current_month} 总超额获利（元）", value=round(total_profit, 2))
            
            st.download_button(
                f"💾 下载{st.session_state.current_month}超额获利数据",
                data=to_excel(display_profit, f"{st.session_state.current_month}超额获利回收明细"),
                file_name=f"超额获利回收明细_{st.session_state.current_month}.xlsx",
                key="download_excess_profit"
            )

            # 可视化
            st.subheader(f"📊 {st.session_state.current_month} 超额获利可视化")
            fig = px.bar(
                display_profit,
                x="时段",
                y="超额获利(元)",
                color="场站名称",
                title=f"{st.session_state.current_month} 各场站分时段超额获利",
                barmode="group"
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("ℹ️ 暂无超额获利（或数据不完整）")
