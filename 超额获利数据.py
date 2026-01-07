import streamlit as st
import pandas as pd
import re
import uuid
from io import BytesIO
import datetime
import plotly.express as px

# -------------------------- 1. 页面基础配置 --------------------------
st.set_page_config(
    page_title="光伏/风电数据管理工具（2025-11专用版）",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------- 2. 全局常量与映射 --------------------------
STATION_TYPE_MAP = {
    "风电": ["荆门栗溪", "荆门圣境山", "襄北风储二期", "襄北风储一期", "襄州峪山一期"],
    "光伏": ["襄北农光", "浠水渔光"]
}
PRICE_TEMPLATE_COLS = [
    "时段", 
    "风电现货均价(元/MWh)", 
    "风电合约均价(元/MWh)", 
    "光伏现货均价(元/MWh)", 
    "光伏合约均价(元/MWh)"
]

# -------------------------- 3. 核心工具函数 --------------------------
def standardize_column_name(col):
    col_str = str(col).strip() if col is not None else f"未知列_{uuid.uuid4().hex[:8]}"
    col_str = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9_]', '_', col_str)
    if col_str == "" or col_str == "_":
        col_str = f"列_{uuid.uuid4().hex[:8]}"
    return col_str

def force_unique_columns(df):
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

def generate_price_template():
    template_data = []
    for hour in range(24):
        template_data.append({
            "时段": f"{hour:02d}:00",
            "风电现货均价(元/MWh)": 0.0,
            "风电合约均价(元/MWh)": 0.0,
            "光伏现货均价(元/MWh)": 0.0,
            "光伏合约均价(元/MWh)": 0.0
        })
    return pd.DataFrame(template_data)

# -------------------------- 4. 会话状态初始化（简化版） --------------------------
if "target_month" not in st.session_state:
    st.session_state.target_month = "2025-11"  # 默认选中2025-11，不用再选
if "gen_data" not in st.session_state:
    st.session_state.gen_data = {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "total": {}}
if "hold_data" not in st.session_state:
    st.session_state.hold_data = {}
if "price_data" not in st.session_state:
    st.session_state.price_data = {"24h": pd.DataFrame(), "excess_profit": pd.DataFrame()}
if "module_config" not in st.session_state:
    st.session_state.module_config = {
        "generated": {"time_col":4, "wind_power_col":9, "pv_power_col":5, "conv":1000, "skip_rows":1},
        "hold": {"hold_col":3, "skip_rows":1},
        "price": {"wind_spot_col":1, "wind_contract_col":2, "pv_spot_col":3, "pv_contract_col":4, "skip_rows":1}
    }

# -------------------------- 5. 核心数据处理类（简化版） --------------------------
class DataProcessor:
    @staticmethod
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
                engine=engine
            )
            
            df = df.iloc[:, :2]
            df.columns = ["时间", "功率(kW)"]
            df["功率(kW)"] = df["功率(kW)"].apply(DataProcessor.clean_power_value)
            df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
            df = df.dropna(subset=["时间", "功率(kW)"]).sort_values("时间").reset_index(drop=True)

            base_name = file.name.split(".")[0].strip()
            unique_station_name = f"{standardize_column_name(base_name)}"
            df[unique_station_name] = df["功率(kW)"] / config["conv"]
            return df[["时间", unique_station_name]].copy(), base_name
        except Exception as e:
            st.error(f"❌ 实发文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame(columns=["时间"]), ""

    @staticmethod
    def calculate_24h_generated(raw_df, config):
        if raw_df.empty:
            st.warning("⚠️ 实发原始数据为空")
            return pd.DataFrame(), {}

        raw_df["时段"] = raw_df["时间"].dt.hour.apply(lambda x: f"{x:02d}:00")
        station_cols = [col for col in raw_df.columns if col not in ["时间", "时段"]]
        
        time_diff = raw_df["时间"].diff().dropna()
        avg_interval_h = time_diff.dt.total_seconds().mean() / 3600
        avg_interval_h = avg_interval_h if avg_interval_h > 0 else 1/4

        generated_24h_df = raw_df.groupby("时段")[station_cols].apply(
            lambda x: (x * avg_interval_h).sum()
        ).round(2).reset_index()
        
        monthly_total = {station: round(generated_24h_df[station].sum(), 2) for station in station_cols}
        return generated_24h_df, monthly_total

    @staticmethod
    def extract_hold_data(file, config):
        try:
            file_suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                usecols=[config["hold_col"]],
                skiprows=config["skip_rows"],
                engine=engine
            )
            df.columns = ["净持有电量"]
            df["净持有电量"] = pd.to_numeric(df["净持有电量"], errors="coerce").fillna(0)
            return round(df["净持有电量"].sum(), 2)
        except Exception as e:
            st.error(f"❌ 持仓文件[{file.name}]处理失败：{str(e)}")
            return 0.0

    @staticmethod
    def extract_price_data(file, config):
        try:
            file_suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                usecols=[0, config["wind_spot_col"], config["wind_contract_col"], 
                         config["pv_spot_col"], config["pv_contract_col"]],
                skiprows=config["skip_rows"],
                engine=engine,
                nrows=24
            )
            df = df.iloc[:, :5]
            df.columns = PRICE_TEMPLATE_COLS
            df["时段"] = [f"{i:02d}:00" for i in range(24)]
            price_cols = df.columns[1:]
            for col in price_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            return df
        except Exception as e:
            st.error(f"❌ 电价文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame()

    @staticmethod
    def calculate_excess_profit(gen_24h_df, hold_dict, price_df, target_month):
        if gen_24h_df.empty or not hold_dict or price_df.empty:
            st.warning("⚠️ 实发/持仓/电价数据不完整")
            return pd.DataFrame()

        merged_df = pd.merge(gen_24h_df, price_df, on="时段", how="inner")
        if merged_df.empty:
            st.warning("⚠️ 实发与电价时段不匹配")
            return pd.DataFrame()

        result_rows = []
        station_cols = [col for col in gen_24h_df.columns if col != "时段"]

        for station in station_cols:
            base_station = station
            station_type = None
            gen_coeff = 1.0
            spot_col = ""
            contract_col = ""
            
            # 匹配场站类型
            for wind_station in STATION_TYPE_MAP["风电"]:
                if wind_station in base_station or base_station in wind_station:
                    station_type = "风电"
                    spot_col = "风电现货均价(元/MWh)"
                    contract_col = "风电合约均价(元/MWh)"
                    gen_coeff = 0.7
                    break
            if not station_type:
                for pv_station in STATION_TYPE_MAP["光伏"]:
                    if pv_station in base_station or base_station in pv_station:
                        station_type = "光伏"
                        spot_col = "光伏现货均价(元/MWh)"
                        contract_col = "光伏合约均价(元/MWh)"
                        gen_coeff = 0.8
                        break
            if not station_type:
                continue

            # 匹配持仓数据
            total_hold = 0
            for hold_station, hold_value in hold_dict.items():
                if hold_station in base_station or base_station in hold_station:
                    total_hold = hold_value
                    break
            if total_hold == 0:
                continue
                
            hourly_hold = total_hold / 24

            for _, row in merged_df.iterrows():
                hourly_generated_raw = row.get(station, 0)
                hourly_generated = hourly_generated_raw * gen_coeff
                
                hold_09 = hourly_hold * 0.9
                hold_11 = hourly_hold * 1.1
                
                if hourly_generated > hold_09:
                    quantity_diff = hourly_generated - hold_11
                else:
                    quantity_diff = hourly_generated - hold_09
                
                spot_price = row.get(spot_col, 0)
                contract_price = row.get(contract_col, 0)
                price_diff = spot_price - contract_price
                excess_profit = quantity_diff * price_diff

                result_rows.append({
                    "场站名称": base_station,
                    "场站类型": station_type,
                    "月份": target_month,
                    "时段": row["时段"],
                    "原始分时实发量(MWh)": round(hourly_generated_raw, 2),
                    "修正后实发量(MWh)": round(hourly_generated, 2),
                    "分时合约电量(MWh)": round(hourly_hold, 2),
                    "合约电量0.9倍(MWh)": round(hold_09, 2),
                    "合约电量1.1倍(MWh)": round(hold_11, 2),
                    "电量差额(MWh)": round(quantity_diff, 2),
                    f"{station_type}现货均价(元/MWh)": round(spot_price, 2),
                    f"{station_type}合约均价(元/MWh)": round(contract_price, 2),
                    f"{station_type}价格差值(元/MWh)": round(price_diff, 2),
                    "超额获利(元)": round(excess_profit, 2)
                })

        return pd.DataFrame(result_rows)

# -------------------------- 6. 页面布局（极简版，按钮全显示） --------------------------
st.title("📈 光伏/风电超额获利计算工具（2025-11专用）")

# 固定月份选择（不用再选，直接锁定2025-11）
st.sidebar.markdown("### 📅 数据月份")
st.session_state.target_month = st.sidebar.text_input("目标月份", value="2025-11")
st.sidebar.markdown("---")

# ====================== 模块1：场站实发配置 ======================
with st.expander("📊 模块1：场站实发配置", expanded=True):
    col1_1, col1_2 = st.columns([3, 2])
    with col1_1:
        station_type = st.radio("选择场站类型", ["风电", "光伏"], key="gen_type")
        gen_files = st.file_uploader(
            f"上传{station_type}实发数据文件（支持多文件）",
            accept_multiple_files=True,
            type=["xlsx", "xls", "xlsm"],
            key="gen_upload"
        )
        if st.button("📝 处理实发数据", key="btn_gen"):
            if not gen_files:
                st.error("❌ 请先上传实发数据文件")
            else:
                all_dfs = []
                for file in gen_files:
                    df, _ = DataProcessor.extract_generated_data(file, st.session_state.module_config["generated"], station_type)
                    if not df.empty:
                        all_dfs.append(df)
                if all_dfs:
                    merged_raw = all_dfs[0].copy()
                    for df in all_dfs[1:]:
                        merged_raw = pd.merge(merged_raw, df, on="时间", how="outer")
                    merged_raw = merged_raw.sort_values("时间").dropna(subset=["时间"]).reset_index(drop=True)
                    st.session_state.gen_data["raw"] = merged_raw
                    
                    # 计算24h汇总
                    gen_24h, gen_total = DataProcessor.calculate_24h_generated(merged_raw, st.session_state.module_config["generated"])
                    st.session_state.gen_data["24h"] = gen_24h
                    st.session_state.gen_data["total"] = gen_total
                    st.success("✅ 实发数据处理完成！")
    
    with col1_2:
        st.markdown("### ⚙️ 列索引配置（0开始）")
        st.session_state.module_config["generated"]["time_col"] = st.number_input("时间列", 0, value=4)
        if station_type == "风电":
            st.session_state.module_config["generated"]["wind_power_col"] = st.number_input("功率列", 0, value=9)
        else:
            st.session_state.module_config["generated"]["pv_power_col"] = st.number_input("功率列", 0, value=5)
        st.session_state.module_config["generated"]["skip_rows"] = st.number_input("跳过行数", 0, value=1)
        st.session_state.module_config["generated"]["conv"] = st.number_input("转换系数(kW→MW)", 1, value=1000)

    # 数据预览
    if not st.session_state.gen_data["raw"].empty:
        st.markdown("### 📋 实发数据预览")
        tab1, tab2 = st.tabs(["原始数据", "24时段汇总"])
        with tab1:
            st.dataframe(st.session_state.gen_data["raw"], use_container_width=True)
            st.download_button("💾 下载原始数据", to_excel(st.session_state.gen_data["raw"]), f"实发原始数据_{st.session_state.target_month}.xlsx")
        with tab2:
            st.dataframe(st.session_state.gen_data["24h"], use_container_width=True)
            st.download_button("💾 下载24h汇总", to_excel(st.session_state.gen_data["24h"]), f"实发24h汇总_{st.session_state.target_month}.xlsx")

# ====================== 模块2：中长期持仓配置 ======================
with st.expander("📦 模块2：中长期持仓配置", expanded=True):
    col2_1, col2_2 = st.columns([3, 2])
    with col2_1:
        hold_files = st.file_uploader(
            "上传持仓数据文件（支持多文件）",
            accept_multiple_files=True,
            type=["xlsx", "xls", "xlsm"],
            key="hold_upload"
        )
        if st.button("📝 处理持仓数据", key="btn_hold"):
            if not hold_files:
                st.error("❌ 请先上传持仓数据文件")
            else:
                hold_total = {}
                for file in hold_files:
                    base_name = file.name.split(".")[0].strip()
                    total = DataProcessor.extract_hold_data(file, st.session_state.module_config["hold"])
                    hold_total[standardize_column_name(base_name)] = total
                st.session_state.hold_data = hold_total
                st.success("✅ 持仓数据处理完成！")
                st.write(f"📊 总持仓数据：{hold_total}")
    
    with col2_2:
        st.markdown("### ⚙️ 列索引配置（0开始）")
        st.session_state.module_config["hold"]["hold_col"] = st.number_input("净持仓列", 0, value=3)
        st.session_state.module_config["hold"]["skip_rows"] = st.number_input("跳过行数", 0, value=1)

# ====================== 模块3：月度电价配置 ======================
with st.expander("💰 模块3：月度电价配置", expanded=True):
    col3_1, col3_2 = st.columns([3, 2])
    with col3_1:
        st.markdown("### 📥 下载电价标准模板")
        st.download_button("📥 下载模板", to_excel(generate_price_template()), "电价标准模板.xlsx")
        
        price_file = st.file_uploader(
            "上传电价数据文件（用标准模板填写）",
            accept_multiple_files=False,
            type=["xlsx", "xls", "xlsm"],
            key="price_upload"
        )
        if st.button("📝 处理电价数据", key="btn_price"):
            if not price_file:
                st.error("❌ 请先上传电价数据文件")
            else:
                price_df = DataProcessor.extract_price_data(price_file, st.session_state.module_config["price"])
                st.session_state.price_data["24h"] = price_df
                st.success("✅ 电价数据处理完成！")
        
        # 电价预览
        if not st.session_state.price_data["24h"].empty:
            st.markdown("### 📋 电价数据预览")
            st.dataframe(st.session_state.price_data["24h"], use_container_width=True)
            st.download_button("💾 下载电价数据", to_excel(st.session_state.price_data["24h"]), f"电价数据_{st.session_state.target_month}.xlsx")
    
    with col2_2:
        st.markdown("### ⚙️ 列索引配置（0开始）")
        st.session_state.module_config["price"]["wind_spot_col"] = st.number_input("风电现货列", 0, value=1)
        st.session_state.module_config["price"]["wind_contract_col"] = st.number_input("风电合约列", 0, value=2)
        st.session_state.module_config["price"]["pv_spot_col"] = st.number_input("光伏现货列", 0, value=3)
        st.session_state.module_config["price"]["pv_contract_col"] = st.number_input("光伏合约列", 0, value=4)
        st.session_state.module_config["price"]["skip_rows"] = st.number_input("跳过行数", 0, value=1)

# ====================== 模块4：超额获利计算 ======================
st.markdown("### 🎯 超额获利计算")
if st.button("🔍 计算超额获利", key="btn_calc", type="primary"):
    excess_df = DataProcessor.calculate_excess_profit(
        st.session_state.gen_data["24h"],
        st.session_state.hold_data,
        st.session_state.price_data["24h"],
        st.session_state.target_month
    )
    st.session_state.price_data["excess_profit"] = excess_df
    
    if not excess_df.empty:
        st.success("✅ 超额获利计算完成！")
        st.dataframe(excess_df, use_container_width=True)
        total_profit = excess_df["超额获利(元)"].sum()
        st.metric(f"💰 {st.session_state.target_month} 总超额获利", value=f"{round(total_profit, 2)} 元")
        
        # 下载+可视化
        col_down, col_plot = st.columns(2)
        with col_down:
            st.download_button("💾 下载获利明细", to_excel(excess_df), f"超额获利明细_{st.session_state.target_month}.xlsx")
        with col_plot:
            fig = px.bar(excess_df, x="时段", y="超额获利(元)", color="场站名称", title="分时段超额获利")
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("ℹ️ 暂无超额获利数据（检查实发/持仓/电价数据是否完整）")
