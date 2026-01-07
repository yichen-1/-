import streamlit as st
import pandas as pd
import re
import uuid
from io import BytesIO
import datetime
import plotly.express as px
import numpy as np

# -------------------------- 1. 页面基础配置 --------------------------
st.set_page_config(
    page_title="光伏/风电超额获利计算工具（2025-11专用版）",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------- 2. 全局常量与映射 --------------------------
STATION_TYPE_MAP = {
    "风电": ["荆门栗溪", "荆门圣境山", "襄北风储二期", "襄北风储一期", "襄州峪山一期", "风电"],
    "光伏": ["襄北农光", "浠水渔光", "光伏"]
}
PRICE_TEMPLATE_COLS = [
    "时段", 
    "风电现货均价(元/MWh)", 
    "风电合约均价(元/MWh)", 
    "光伏现货均价(元/MWh)", 
    "光伏合约均价(元/MWh)"
]
# 新增：标准时段列表（用于匹配分时段持仓）
STANDARD_HOURS = [f"{i:02d}:00" for i in range(24)]

# -------------------------- 3. 核心工具函数 --------------------------
def standardize_column_name(col):
    col_str = str(col).strip() if col is not None else f"未知列_{uuid.uuid4().hex[:8]}"
    col_str = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9_]', '_', col_str)
    col_str = col_str.lower()
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
    time_col_candidates = [i for i, col in enumerate(df.columns) if "时间" in col or "date" in col.lower() or "时段" in col]
    if time_col_candidates:
        df.columns = ["时段" if i == time_col_candidates[0] else col for i, col in enumerate(df.columns)]
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

# 新增：标准化时段格式（统一为"00:00"格式）
def standardize_hour(hour_str):
    try:
        # 处理"0时"、"1点"、"00:00"等多种格式
        hour_str = str(hour_str).strip().replace("时", "").replace("点", "").replace("：", ":")
        if ":" in hour_str:
            h, _ = hour_str.split(":")
            return f"{int(h):02d}:00"
        else:
            return f"{int(hour_str):02d}:00"
    except:
        return None

# -------------------------- 4. 会话状态初始化 --------------------------
if "target_month" not in st.session_state:
    st.session_state.target_month = "2025-11"
if "gen_data" not in st.session_state:
    st.session_state.gen_data = {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "total": {}}
if "hold_data" not in st.session_state:
    st.session_state.hold_data = {}  # 改为：{场站名称: {时段: 持仓值, ...}}
if "hold_data_df" not in st.session_state:
    st.session_state.hold_data_df = pd.DataFrame()  # 存储分时段持仓的原始DataFrame
if "binded_hold_data" not in st.session_state:
    st.session_state.binded_hold_data = {}  # 改为：{实发场站: 持仓场站}
if "price_data" not in st.session_state:
    st.session_state.price_data = {"24h": pd.DataFrame(), "excess_profit": pd.DataFrame()}
if "module_config" not in st.session_state:
    st.session_state.module_config = {
        "generated": {"time_col":4, "wind_power_col":9, "pv_power_col":5, "conv":1000, "skip_rows":1},
        "hold": {"hour_col":0, "hold_col":1, "skip_rows":1},  # 修改：hour_col=时段列，hold_col=持仓列
        "price": {"wind_spot_col":1, "wind_contract_col":2, "pv_spot_col":3, "pv_contract_col":4, "skip_rows":1}
    }

# -------------------------- 5. 核心数据处理类（适配分时段持仓） --------------------------
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
            st.info(f"✅ 实发文件[{file.name}]提取成功，场站名称：{unique_station_name}")
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
        st.info(f"🔍 识别到实发场站：{station_cols}")
        
        time_diff = raw_df["时间"].diff().dropna()
        avg_interval_h = time_diff.dt.total_seconds().mean() / 3600
        avg_interval_h = avg_interval_h if avg_interval_h > 0 else 1/4

        generated_24h_df = raw_df.groupby("时段")[station_cols].apply(
            lambda x: (x * avg_interval_h).sum()
        ).round(2).reset_index()
        
        monthly_total = {station: round(generated_24h_df[station].sum(), 2) for station in station_cols}
        st.success(f"✅ 24时段汇总完成，各场站总发电量：{monthly_total}")
        return generated_24h_df, monthly_total

    @staticmethod
    def extract_hold_data(file, config):
        """修改：读取分时段持仓数据，返回{场站名称: {时段: 持仓值}}"""
        try:
            file_suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            
            # 读取时段列和持仓列
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                usecols=[config["hour_col"], config["hold_col"]],
                skiprows=config["skip_rows"],
                engine=engine,
                nrows=24  # 仅读取24行（对应24时段）
            )
            
            df = df.iloc[:, :2]
            df.columns = ["时段", "持仓量(MWh)"]
            
            # 标准化时段格式
            df["时段"] = df["时段"].apply(standardize_hour)
            # 清洗持仓值
            df["持仓量(MWh)"] = pd.to_numeric(df["持仓量(MWh)"], errors="coerce").fillna(0)
            # 过滤有效时段（仅保留00:00~23:00）
            df = df[df["时段"].isin(STANDARD_HOURS)].reset_index(drop=True)
            
            # 补充缺失的时段（确保24个时段完整）
            full_hours = pd.DataFrame({"时段": STANDARD_HOURS})
            df = pd.merge(full_hours, df, on="时段", how="left").fillna(0)
            
            # 生成场站名称
            base_name = standardize_column_name(file.name.split(".")[0].strip())
            st.info(f"✅ 持仓文件[{file.name}]提取成功，场站名称：{base_name}，有效时段数：{len(df)}")
            
            # 转换为字典：{时段: 持仓值}
            hold_hourly_dict = dict(zip(df["时段"], df["持仓量(MWh)"]))
            total_hold = round(sum(hold_hourly_dict.values()), 2)
            
            return base_name, hold_hourly_dict, df, total_hold
        except Exception as e:
            st.error(f"❌ 持仓文件[{file.name}]处理失败：{str(e)}")
            return "", {}, pd.DataFrame(), 0.0

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
            st.success(f"✅ 电价文件[{file.name}]提取成功，时段数：{len(df)}")
            return df
        except Exception as e:
            st.error(f"❌ 电价文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame()

    @staticmethod
    def calculate_excess_profit(gen_24h_df, hold_dict, binded_hold, price_df, target_month):
        st.markdown("### 🕵️ 数据检查")
        if gen_24h_df.empty:
            st.error("❌ 实发24h汇总数据为空")
            return pd.DataFrame()
        else:
            st.success(f"✅ 实发24h数据：{len(gen_24h_df)} 行，场站：{[col for col in gen_24h_df.columns if col != '时段']}")
        
        if not hold_dict:
            st.error("❌ 持仓数据为空")
            return pd.DataFrame()
        else:
            st.success(f"✅ 持仓数据：{list(hold_dict.keys())} （均为分时段持仓）")
        
        if price_df.empty:
            st.error("❌ 电价数据为空")
            return pd.DataFrame()
        else:
            st.success(f"✅ 电价数据：{len(price_df)} 行")

        # 过滤有效时段
        gen_24h_df = gen_24h_df[gen_24h_df["时段"].isin(STANDARD_HOURS)]
        price_df = price_df[price_df["时段"].isin(STANDARD_HOURS)]
        
        merged_df = pd.merge(gen_24h_df, price_df, on="时段", how="inner")
        if merged_df.empty:
            st.error("❌ 实发与电价数据时段无法匹配")
            return pd.DataFrame()
        st.success(f"✅ 数据合并成功，有效时段数：{len(merged_df)}")

        result_rows = []
        gen_stations = [col for col in gen_24h_df.columns if col != "时段"]

        for gen_station in gen_stations:
            # 获取绑定的持仓场站
            hold_station = binded_hold.get(gen_station)
            if not hold_station or hold_station not in hold_dict:
                st.warning(f"⚠️ 场站[{gen_station}]无绑定的分时段持仓数据，跳过计算")
                continue
            
            # 获取该场站的分时段持仓字典
            hold_hourly_dict = hold_dict[hold_station]
            base_station = gen_station.lower()
            
            # 匹配场站类型和修正系数
            station_type = None
            gen_coeff = 1.0
            spot_col = ""
            contract_col = ""
            for wind_key in STATION_TYPE_MAP["风电"]:
                if wind_key.lower() in base_station or base_station in wind_key.lower():
                    station_type = "风电"
                    spot_col = "风电现货均价(元/MWh)"
                    contract_col = "风电合约均价(元/MWh)"
                    gen_coeff = 0.7
                    st.info(f"🔍 匹配到场站[{gen_station}]类型：风电，修正系数：{gen_coeff}")
                    break
            if not station_type:
                for pv_key in STATION_TYPE_MAP["光伏"]:
                    if pv_key.lower() in base_station or base_station in pv_key.lower():
                        station_type = "光伏"
                        spot_col = "光伏现货均价(元/MWh)"
                        contract_col = "光伏合约均价(元/MWh)"
                        gen_coeff = 0.8
                        st.info(f"🔍 匹配到场站[{gen_station}]类型：光伏，修正系数：{gen_coeff}")
                        break
            if not station_type:
                st.warning(f"⚠️ 场站[{gen_station}]无法匹配类型，跳过计算")
                continue

            # 逐时段计算
            for _, row in merged_df.iterrows():
                hour = row["时段"]
                # 1. 获取当前时段的实发量
                hourly_generated_raw = row.get(gen_station, 0)
                hourly_generated = hourly_generated_raw * gen_coeff
                
                # 2. 获取当前时段的持仓量（直接读取分时段数据，不再均分）
                hourly_hold = hold_hourly_dict.get(hour, 0)
                if hourly_hold <= 0:
                    continue  # 持仓为0的时段跳过
                
                # 3. 计算电量差额（0.9~1.1倍区间规则不变）
                if hourly_generated > hourly_hold * 1.1:
                    quantity_diff = hourly_generated - hourly_hold * 1.1
                elif hourly_generated < hourly_hold * 0.9:
                    quantity_diff = hourly_generated - hourly_hold * 0.9
                else:
                    quantity_diff = 0
                
                # 4. 计算价格差值
                spot_price = row.get(spot_col, 0)
                contract_price = row.get(contract_col, 0)
                price_diff = spot_price - contract_price
                
                # 5. 计算超额获利（负数归零，只统计正数）
                excess_profit = quantity_diff * price_diff
                if excess_profit < 0:
                    excess_profit = 0

                # 6. 保存结果
                result_rows.append({
                    "场站名称": gen_station,
                    "场站类型": station_type,
                    "月份": target_month,
                    "时段": hour,
                    "原始分时实发量(MWh)": round(hourly_generated_raw, 2),
                    "修正后实发量(MWh)": round(hourly_generated, 2),
                    "分时合约电量(MWh)": round(hourly_hold, 2),
                    "合约电量0.9倍(MWh)": round(hourly_hold * 0.9, 2),
                    "合约电量1.1倍(MWh)": round(hourly_hold * 1.1, 2),
                    "电量差额(MWh)": round(quantity_diff, 2),
                    f"{station_type}现货均价(元/MWh)": round(spot_price, 2),
                    f"{station_type}合约均价(元/MWh)": round(contract_price, 2),
                    f"{station_type}价格差值(元/MWh)": round(price_diff, 2),
                    "超额获利(元)": round(excess_profit, 2)
                })

        # 生成结果表
        result_df = pd.DataFrame(result_rows)
        if not result_df.empty:
            # 总计行（仅统计正数获利）
            total_row = {
                "场站名称": "总计",
                "场站类型": "",
                "月份": target_month,
                "时段": "",
                "原始分时实发量(MWh)": round(result_df["原始分时实发量(MWh)"].sum(), 2),
                "修正后实发量(MWh)": round(result_df["修正后实发量(MWh)"].sum(), 2),
                "分时合约电量(MWh)": round(result_df["分时合约电量(MWh)"].sum(), 2),
                "合约电量0.9倍(MWh)": round(result_df["合约电量0.9倍(MWh)"].sum(), 2),
                "合约电量1.1倍(MWh)": round(result_df["合约电量1.1倍(MWh)"].sum(), 2),
                "电量差额(MWh)": round(result_df["电量差额(MWh)"].sum(), 2),
                "风电现货均价(元/MWh)": "",
                "风电合约均价(元/MWh)": "",
                "光伏现货均价(元/MWh)": "",
                "光伏合约均价(元/MWh)": "",
                "风电价格差值(元/MWh)": "",
                "光伏价格差值(元/MWh)": "",
                "超额获利(元)": round(result_df["超额获利(元)"].sum(), 2)
            }
            result_df = pd.concat([result_df, pd.DataFrame([total_row])], ignore_index=True)
            st.success(f"✅ 超额获利计算完成（仅统计正数），共{len(result_df)-1}行数据 + 1行总计")
        
        return result_df

# -------------------------- 6. 页面布局（适配分时段持仓） --------------------------
st.title("📈 光伏/风电超额获利计算工具（分时段持仓版）")

# 固定月份选择
st.sidebar.markdown("### 📅 数据月份")
st.session_state.target_month = st.sidebar.text_input(
    "目标月份", 
    value="2025-11",
    key="sidebar_target_month"
)
st.sidebar.markdown("---")

# ====================== 模块1：场站实发配置 ======================
with st.expander("📊 模块1：场站实发配置", expanded=True):
    col1_1, col1_2 = st.columns([3, 2])
    with col1_1:
        station_type = st.radio(
            "选择场站类型", 
            ["风电", "光伏"], 
            key="gen_type_radio"
        )
        gen_files = st.file_uploader(
            f"上传{station_type}实发数据文件（支持多文件）",
            accept_multiple_files=True,
            type=["xlsx", "xls", "xlsm"],
            key="gen_upload_file"
        )
        if st.button(
            "📝 处理实发数据", 
            key="btn_process_gen_data"
        ):
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
                    
                    gen_24h, gen_total = DataProcessor.calculate_24h_generated(merged_raw, st.session_state.module_config["generated"])
                    st.session_state.gen_data["24h"] = gen_24h
                    st.session_state.gen_data["total"] = gen_total
                    st.success("✅ 实发数据处理完成！")
    
    with col1_2:
        st.markdown("### ⚙️ 列索引配置（0开始）")
        st.session_state.module_config["generated"]["time_col"] = st.number_input(
            "时间列", 
            0, 
            value=4,
            key="gen_time_col_input"
        )
        if station_type == "风电":
            st.session_state.module_config["generated"]["wind_power_col"] = st.number_input(
                "功率列", 
                0, 
                value=9,
                key="gen_wind_power_col_input"
            )
        else:
            st.session_state.module_config["generated"]["pv_power_col"] = st.number_input(
                "功率列", 
                0, 
                value=5,
                key="gen_pv_power_col_input"
            )
        st.session_state.module_config["generated"]["skip_rows"] = st.number_input(
            "跳过行数", 
            0, 
            value=1,
            key="gen_skip_rows_input"
        )
        st.session_state.module_config["generated"]["conv"] = st.number_input(
            "转换系数(kW→MW)", 
            1, 
            value=1000,
            key="gen_conv_input"
        )

    if not st.session_state.gen_data["raw"].empty:
        st.markdown("### 📋 实发数据预览")
        tab1, tab2 = st.tabs(["原始数据", "24时段汇总"])
        with tab1:
            st.dataframe(st.session_state.gen_data["raw"], use_container_width=True)
            st.download_button(
                "💾 下载原始数据", 
                to_excel(st.session_state.gen_data["raw"]), 
                f"实发原始数据_{st.session_state.target_month}.xlsx",
                key="download_gen_raw"
            )
        with tab2:
            st.dataframe(st.session_state.gen_data["24h"], use_container_width=True)
            st.download_button(
                "💾 下载24h汇总", 
                to_excel(st.session_state.gen_data["24h"]), 
                f"实发24h汇总_{st.session_state.target_month}.xlsx",
                key="download_gen_24h"
            )

# ====================== 模块2：分时段持仓配置（核心修改） ======================
with st.expander("📦 模块2：分时段持仓配置", expanded=True):
    col2_1, col2_2 = st.columns([3, 2])
    with col2_1:
        hold_files = st.file_uploader(
            "上传分时段持仓数据文件（支持多文件）",
            accept_multiple_files=True,
            type=["xlsx", "xls", "xlsm"],
            key="hold_upload_file"
        )
        if st.button(
            "📝 处理分时段持仓数据", 
            key="btn_process_hold_data"
        ):
            if not hold_files:
                st.error("❌ 请先上传分时段持仓数据文件")
            else:
                hold_total_dict = {}  # {持仓场站: {时段: 持仓值}}
                hold_dfs = []
                for file in hold_files:
                    hold_station, hold_hourly, hold_df, total = DataProcessor.extract_hold_data(file, st.session_state.module_config["hold"])
                    if hold_station and total > 0:
                        hold_total_dict[hold_station] = hold_hourly
                        hold_df["场站名称"] = hold_station
                        hold_dfs.append(hold_df)
                
                if hold_dfs:
                    st.session_state.hold_data_df = pd.concat(hold_dfs, ignore_index=True)
                st.session_state.hold_data = hold_total_dict
                st.success("✅ 分时段持仓数据处理完成！")
                # 展示各持仓场站的总持仓
                hold_summary = {k: round(sum(v.values()), 2) for k, v in hold_total_dict.items()}
                st.write(f"📊 持仓汇总（各场站总持仓）：{hold_summary}")
        
        # 手动绑定：实发场站 ↔ 分时段持仓场站
        if st.session_state.hold_data and not st.session_state.gen_data["24h"].empty:
            st.markdown("### 🔗 绑定实发场站到分时段持仓场站")
            gen_stations = [col for col in st.session_state.gen_data["24h"].columns if col != "时段"]
            hold_stations = list(st.session_state.hold_data.keys())
            
            if gen_stations and hold_stations:
                col_bind1, col_bind2 = st.columns(2)
                with col_bind1:
                    selected_gen_station = st.selectbox("选择实发场站", gen_stations, key="bind_gen_station")
                with col_bind2:
                    selected_hold_station = st.selectbox("选择分时段持仓场站", hold_stations, key="bind_hold_station")
                
                if st.button("✅ 确认绑定", key="btn_bind_hold"):
                    st.session_state.binded_hold_data[selected_gen_station] = selected_hold_station
                    st.success(f"✅ 已将实发场站[{selected_gen_station}]绑定到分时段持仓场站[{selected_hold_station}]")
                    st.write(f"当前绑定关系：{st.session_state.binded_hold_data}")
        
        # 展示分时段持仓数据预览
        if not st.session_state.hold_data_df.empty:
            st.markdown("### 📋 分时段持仓数据预览")
            st.dataframe(st.session_state.hold_data_df, use_container_width=True)
            st.download_button(
                "💾 下载分时段持仓数据", 
                to_excel(st.session_state.hold_data_df), 
                f"分时段持仓数据_{st.session_state.target_month}.xlsx",
                key="download_hold_data"
            )
    
    with col2_2:
        st.markdown("### ⚙️ 列索引配置（0开始）")
        st.session_state.module_config["hold"]["hour_col"] = st.number_input(
            "时段列（分时段持仓）", 
            0, 
            value=0,
            key="hold_hour_col_input"
        )
        st.session_state.module_config["hold"]["hold_col"] = st.number_input(
            "持仓量列", 
            0, 
            value=1,
            key="hold_col_input"
        )
        st.session_state.module_config["hold"]["skip_rows"] = st.number_input(
            "跳过行数", 
            0, 
            value=1,
            key="hold_skip_rows_input"
        )

# ====================== 模块3：月度电价配置 ======================
with st.expander("💰 模块3：月度电价配置", expanded=True):
    col3_1, col3_2 = st.columns([3, 2])
    with col3_1:
        st.markdown("### 📥 下载电价标准模板")
        st.download_button(
            "📥 下载模板", 
            to_excel(generate_price_template()), 
            "电价标准模板.xlsx",
            key="download_price_template"
        )
        
        price_file = st.file_uploader(
            "上传电价数据文件（用标准模板填写）",
            accept_multiple_files=False,
            type=["xlsx", "xls", "xlsm"],
            key="price_upload_file"
        )
        if st.button(
            "📝 处理电价数据", 
            key="btn_process_price_data"
        ):
            if not price_file:
                st.error("❌ 请先上传电价数据文件")
            else:
                price_df = DataProcessor.extract_price_data(price_file, st.session_state.module_config["price"])
                st.session_state.price_data["24h"] = price_df
                st.success("✅ 电价数据处理完成！")
        
        if not st.session_state.price_data["24h"].empty:
            st.markdown("### 📋 电价数据预览")
            st.dataframe(st.session_state.price_data["24h"], use_container_width=True)
            st.download_button(
                "💾 下载电价数据", 
                to_excel(st.session_state.price_data["24h"]), 
                f"电价数据_{st.session_state.target_month}.xlsx",
                key="download_price_data"
            )
    
    with col3_2:
        st.markdown("### ⚙️ 列索引配置（0开始）")
        st.session_state.module_config["price"]["wind_spot_col"] = st.number_input(
            "风电现货列", 
            0, 
            value=1,
            key="price_wind_spot_col_input"
        )
        st.session_state.module_config["price"]["wind_contract_col"] = st.number_input(
            "风电合约列", 
            0, 
            value=2,
            key="price_wind_contract_col_input"
        )
        st.session_state.module_config["price"]["pv_spot_col"] = st.number_input(
            "光伏现货列", 
            0, 
            value=3,
            key="price_pv_spot_col_input"
        )
        st.session_state.module_config["price"]["pv_contract_col"] = st.number_input(
            "光伏合约列", 
            0, 
            value=4,
            key="price_pv_contract_col_input"
        )
        st.session_state.module_config["price"]["skip_rows"] = st.number_input(
            "跳过行数", 
            0, 
            value=1,
            key="price_skip_rows_input"
        )

# ====================== 模块4：超额获利计算（适配分时段持仓） ======================
st.markdown("### 🎯 超额获利计算（仅统计正数部分+分时段持仓）")
if st.button(
    "🔍 计算超额获利", 
    key="btn_calc_excess_profit",
    type="primary"
):
    if not st.session_state.binded_hold_data:
        st.error("❌ 请先完成「实发场站 ↔ 分时段持仓场站」的绑定！")
    else:
        excess_df = DataProcessor.calculate_excess_profit(
            st.session_state.gen_data["24h"],
            st.session_state.hold_data,
            st.session_state.binded_hold_data,
            st.session_state.price_data["24h"],
            st.session_state.target_month
        )
        st.session_state.price_data["excess_profit"] = excess_df
        
        if not excess_df.empty:
            st.success("✅ 超额获利计算完成（仅统计正数部分+分时段持仓）！")
            st.dataframe(excess_df, use_container_width=True)
            total_profit = excess_df[excess_df["场站名称"] == "总计"]["超额获利(元)"].iloc[0]
            st.metric(f"💰 {st.session_state.target_month} 总超额获利（仅正数）", value=f"{round(total_profit, 2)} 元")
            
            col_down, col_plot = st.columns(2)
            with col_down:
                st.download_button(
                    "💾 下载获利明细", 
                    to_excel(excess_df), 
                    f"超额获利明细_{st.session_state.target_month}.xlsx",
                    key="download_excess_profit"
                )
            with col_plot:
                plot_df = excess_df[excess_df["场站名称"] != "总计"]
                fig = px.bar(
                    plot_df, 
                    x="时段", 
                    y="超额获利(元)", 
                    color="场站名称", 
                    title=f"{st.session_state.target_month} 各场站分时段超额获利（仅正数）",
                    barmode="group"
                )
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.error("❌ 超额获利计算失败，请检查：")
            st.markdown("""
            1. 是否已完成「实发场站 ↔ 分时段持仓场站」绑定；
            2. 分时段持仓数据是否每个时段都有非0值；
            3. 电价数据是否填写了非0值；
            4. 实发数据是否有非0的发电量；
            5. 是否有至少一个时段的获利为正数。
            """)
