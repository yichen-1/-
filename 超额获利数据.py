import streamlit as st
import pandas as pd
import re
from io import BytesIO
import datetime
import plotly.express as px

# -------------------------- 1. 页面基础配置 --------------------------
st.set_page_config(
    page_title="光伏/风电数据管理工具（最终版）",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------- 2. 全局常量与映射（固定配置） --------------------------
STATION_TYPE_MAP = {
    "风电": ["荆门栗溪", "荆门圣境山", "襄北风储二期", "襄北风储一期", "襄州峪山一期"],
    "光伏": ["襄北农光", "浠水渔光"]
}

# -------------------------- 3. 会话状态初始化 --------------------------
if "core_data" not in st.session_state:
    st.session_state.core_data = {
        "generated": {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "total": {}},
        "hold": {"total": {}, "config": {}},
        "price": {"24h": pd.DataFrame(), "excess_profit": pd.DataFrame()}
    }

if "module_config" not in st.session_state:
    st.session_state.module_config = {
        "generated": {
            "time_col": 4, "wind_power_col": 9, "pv_power_col": 5,
            "pv_list": "浠水渔光,襄北农光", "conv": 1000, "skip_rows": 1, "keyword": "历史趋势"
        },
        "hold": {"hold_col": 3, "skip_rows": 1},
        "price": {"spot_col": 1, "wind_contract_col": 2, "pv_contract_col": 3, "skip_rows": 1}
    }

# -------------------------- 4. 辅助函数 --------------------------
def to_excel(df, sheet_name="数据"):
    if df.empty:
        st.warning("⚠️ 数据为空，无法生成Excel文件")
        return BytesIO()
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

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
            
            # 读取基础数据
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                usecols=[config["time_col"], power_col],
                skiprows=config["skip_rows"],
                engine=engine,
                nrows=None
            )
            df.columns = ["时间", "功率(kW)"]

            # 严格清洗
            df["功率(kW)"] = df["功率(kW)"].apply(DataProcessor.clean_power_value)
            df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
            df = df.dropna(subset=["时间", "功率(kW)"]).sort_values("时间").reset_index(drop=True)

            # 提取场站名（确保唯一）
            station_name = file.name.split(".")[0].split("-")[0].strip()
            # 避免重复列名
            if station_name in df.columns:
                station_name = f"{station_name}_{file.name[:8]}"  # 加文件前缀区分
            df[station_name] = df["功率(kW)"] / config["conv"]

            return df[["时间", station_name]], station_name
        except Exception as e:
            st.error(f"❌ 实发文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame(columns=["时间"]), ""

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
            file_suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                usecols=[0, config["spot_col"], config["wind_contract_col"], config["pv_contract_col"]],
                skiprows=config["skip_rows"],
                engine=engine,
                nrows=24
            )
            df.columns = ["时段", "现货均价(元/MWh)", "风电合约均价(元/MWh)", "光伏合约均价(元/MWh)"]
            df["时段"] = [f"{i:02d}:00" for i in range(24)]
            price_cols = ["现货均价(元/MWh)", "风电合约均价(元/MWh)", "光伏合约均价(元/MWh)"]
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

        time_diff = merged_raw_df["时间"].diff().dropna()
        avg_interval_h = time_diff.dt.total_seconds().mean() / 3600
        avg_interval_h = avg_interval_h if avg_interval_h > 0 else 1/4

        merged_raw_df["时段"] = merged_raw_df["时间"].dt.hour.apply(lambda x: f"{x:02d}:00")
        station_cols = [col for col in merged_raw_df.columns if col not in ["时间", "时段"]]
        
        # 安全分组汇总
        try:
            generated_24h_df = merged_raw_df.groupby("时段")[station_cols].apply(
                lambda x: (x * avg_interval_h).sum()
            ).round(2).reset_index()
        except Exception as e:
            st.error(f"❌ 24时段汇总失败：{str(e)}")
            return pd.DataFrame(), {}

        monthly_total = {
            station: round(generated_24h_df[station].sum(), 2)
            for station in station_cols
        }

        return generated_24h_df, monthly_total

    @staticmethod
    def calculate_excess_profit(generated_24h_df, hold_total_dict, price_24h_df):
        if generated_24h_df.empty or not hold_total_dict or price_24h_df.empty:
            st.warning("⚠️ 实发/持仓/电价数据不完整，无法计算超额获利")
            return pd.DataFrame()

        merged_df = pd.merge(generated_24h_df, price_24h_df, on="时段", how="inner")
        if merged_df.empty:
            st.warning("⚠️ 实发与电价数据时段不匹配，无法计算")
            return pd.DataFrame()

        result_rows = []
        station_cols = [col for col in generated_24h_df.columns if col != "时段"]

        for station in station_cols:
            # 匹配场站类型（兼容重命名的场站名）
            base_station = station.split("_")[0]  # 去掉重复后缀
            station_type = None
            if base_station in STATION_TYPE_MAP["风电"]:
                station_type = "风电"
                contract_col = "风电合约均价(元/MWh)"
            elif base_station in STATION_TYPE_MAP["光伏"]:
                station_type = "光伏"
                contract_col = "光伏合约均价(元/MWh)"
            
            if not station_type:
                st.warning(f"⚠️ 场站[{station}]未配置类型，跳过计算")
                continue

            # 匹配持仓数据（兼容重命名）
            total_hold = hold_total_dict.get(base_station, hold_total_dict.get(station, 0))
            if total_hold == 0:
                st.warning(f"⚠️ 场站[{station}]无持仓数据，跳过计算")
                continue
                
            hourly_hold = total_hold / 24

            for _, row in merged_df.iterrows():
                hourly_generated = row[station]
                spot_price = row["现货均价(元/MWh)"]
                contract_price = row[contract_col]

                excess_quantity = max(0, hourly_generated - hourly_hold)
                excess_profit = excess_quantity * (spot_price - contract_price)

                if excess_profit > 0:
                    result_rows.append({
                        "场站名称": station,
                        "场站类型": station_type,
                        "时段": row["时段"],
                        "时段实发量(MWh)": round(hourly_generated, 2),
                        "时段持仓量(MWh)": round(hourly_hold, 2),
                        "超额电量(MWh)": round(excess_quantity, 2),
                        "现货均价(元/MWh)": round(spot_price, 2),
                        "合约均价(元/MWh)": round(contract_price, 2),
                        "超额获利(元)": round(excess_profit, 2)
                    })

        return pd.DataFrame(result_rows)

# -------------------------- 6. 页面布局 --------------------------
st.title("📈 光伏/风电数据管理工具（最终版）")
st.divider()

# ====================== 模块1：场站实发配置 ======================
with st.expander("📊 模块1：场站实发配置", expanded=False):
    st.subheader("1.1 数据上传")
    col1_1, col1_2 = st.columns(2)
    with col1_1:
        station_type = st.radio("选择场站类型", ["风电", "光伏"], key="gen_station_type")
        gen_files = st.file_uploader(
            f"上传{station_type}实发数据文件（Excel）",
            accept_multiple_files=True,
            type=["xlsx", "xls", "xlsm"],
            key="gen_file_upload"
        )
    with col1_2:
        if gen_files:
            st.success(f"✅ 已上传{len(gen_files)}个{station_type}实发文件")
            if st.button("📝 处理实发数据", key="process_gen_data"):
                all_raw_dfs = []
                all_stations = []
                
                # 逐个处理文件
                for file in gen_files:
                    df, station = DataProcessor.extract_generated_data(
                        file, st.session_state.module_config["generated"], station_type
                    )
                    if not df.empty and "时间" in df.columns and not df["时间"].isna().all():
                        all_raw_dfs.append(df)
                        all_stations.append(station)

                # 安全合并（核心修复点）
                if all_raw_dfs:
                    # 初始化合并数据
                    merged_raw = all_raw_dfs[0].copy()
                    # 循环合并，避免一次性merge多个导致报错
                    for df in all_raw_dfs[1:]:
                        try:
                            merged_raw = pd.merge(
                                merged_raw, df, on="时间", how="outer", suffixes=("", f"_{df.columns[1][:4]}")
                            )
                        except pd.errors.MergeError as e:
                            st.warning(f"⚠️ 合并文件{df.columns[1]}失败：{str(e)}，跳过该文件")
                            continue
                    
                    # 最终清洗
                    merged_raw = merged_raw.sort_values("时间").reset_index(drop=True)
                    merged_raw = merged_raw.dropna(subset=["时间"])  # 过滤无时间的行
                    
                    st.session_state.core_data["generated"]["raw"] = merged_raw
                    
                    # 计算汇总数据
                    gen_24h, gen_total = DataProcessor.calculate_24h_generated(
                        merged_raw, st.session_state.module_config["generated"]
                    )
                    st.session_state.core_data["generated"]["24h"] = gen_24h
                    st.session_state.core_data["generated"]["total"] = gen_total
                    st.success("✅ 实发数据处理完成！")
                else:
                    st.error("❌ 无有效实发数据，请检查文件格式或内容")

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

    # 数据预览
    if not st.session_state.core_data["generated"]["raw"].empty:
        st.subheader("📋 实发数据预览")
        tab1, tab2 = st.tabs(["原始数据", "24时段汇总"])
        with tab1:
            st.dataframe(st.session_state.core_data["generated"]["raw"], use_container_width=True)
            st.download_button(
                "💾 下载原始实发数据",
                data=to_excel(st.session_state.core_data["generated"]["raw"], "原始实发数据"),
                file_name=f"实发原始数据_{datetime.date.today()}.xlsx",
                key="download_gen_raw"
            )
        with tab2:
            st.dataframe(st.session_state.core_data["generated"]["24h"], use_container_width=True)
            st.download_button(
                "💾 下载24时段汇总数据",
                data=to_excel(st.session_state.core_data["generated"]["24h"], "24时段实发汇总"),
                file_name=f"24时段实发汇总_{datetime.date.today()}.xlsx",
                key="download_gen_24h"
            )

st.divider()

# ====================== 模块2：中长期持仓配置 ======================
with st.expander("📦 模块2：中长期持仓配置", expanded=False):
    st.subheader("2.1 数据上传")
    col2_1, col2_2 = st.columns(2)
    with col2_1:
        hold_files = st.file_uploader(
            "上传持仓数据文件（Excel）",
            accept_multiple_files=True,
            type=["xlsx", "xls", "xlsm"],
            key="hold_file_upload"
        )
    with col2_2:
        if hold_files:
            st.success(f"✅ 已上传{len(hold_files)}个持仓文件")
            if st.button("📝 处理持仓数据", key="process_hold_data"):
                hold_total = {}
                for file in hold_files:
                    station_name = file.name.split(".")[0].split("-")[0].strip()
                    total = DataProcessor.extract_hold_data(file, st.session_state.module_config["hold"])
                    hold_total[station_name] = total
                st.session_state.core_data["hold"]["total"] = hold_total
                st.success("✅ 持仓数据处理完成！")
                st.write("📊 各场站月度总持仓（MWh）：")
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

# ====================== 模块3：月度电价配置 ======================
with st.expander("💰 模块3：月度电价配置", expanded=False):
    st.subheader("3.1 数据上传")
    col3_1, col3_2 = st.columns(2)
    with col3_1:
        price_file = st.file_uploader(
            "上传电价数据文件（Excel）",
            accept_multiple_files=False,
            type=["xlsx", "xls", "xlsm"],
            key="price_file_upload"
        )
    with col3_2:
        if price_file:
            st.success("✅ 已上传电价数据文件")
            if st.button("📝 处理电价数据", key="process_price_data"):
                price_24h = DataProcessor.extract_price_data(price_file, st.session_state.module_config["price"])
                st.session_state.core_data["price"]["24h"] = price_24h
                st.success("✅ 电价数据处理完成！")

    st.subheader("3.2 列索引配置（索引从0开始）")
    col3_3, col3_4, col3_5 = st.columns(3)
    with col3_3:
        st.session_state.module_config["price"]["spot_col"] = st.number_input(
            "现货均价列索引", min_value=0, value=st.session_state.module_config["price"]["spot_col"], key="price_spot_col"
        )
    with col3_4:
        st.session_state.module_config["price"]["wind_contract_col"] = st.number_input(
            "风电合约均价列索引", min_value=0, value=st.session_state.module_config["price"]["wind_contract_col"], key="price_wind_col"
        )
    with col3_5:
        st.session_state.module_config["price"]["pv_contract_col"] = st.number_input(
            "光伏合约均价列索引", min_value=0, value=st.session_state.module_config["price"]["pv_contract_col"], key="price_pv_col"
        )

    if not st.session_state.core_data["price"]["24h"].empty:
        st.subheader("📋 24时段电价数据预览")
        st.dataframe(st.session_state.core_data["price"]["24h"], use_container_width=True)
        st.download_button(
            "💾 下载电价数据",
            data=to_excel(st.session_state.core_data["price"]["24h"], "24时段电价数据"),
            file_name=f"24时段电价数据_{datetime.date.today()}.xlsx",
            key="download_price_24h"
        )

st.divider()

# ====================== 模块4：超额获利计算 ======================
st.subheader("🎯 超额获利回收计算")
if st.button("🔍 计算超额获利", key="calc_excess_profit"):
    excess_profit_df = DataProcessor.calculate_excess_profit(
        st.session_state.core_data["generated"]["24h"],
        st.session_state.core_data["hold"]["total"],
        st.session_state.core_data["price"]["24h"]
    )
    st.session_state.core_data["price"]["excess_profit"] = excess_profit_df

    if not excess_profit_df.empty:
        st.success("✅ 超额获利计算完成！")
        st.dataframe(excess_profit_df, use_container_width=True)
        total_profit = excess_profit_df["超额获利(元)"].sum()
        st.metric("💰 总超额获利（元）", value=round(total_profit, 2))
        st.download_button(
            "💾 下载超额获利数据",
            data=to_excel(excess_profit_df, "超额获利回收明细"),
            file_name=f"超额获利回收明细_{datetime.date.today()}.xlsx",
            key="download_excess_profit"
        )

        # 可视化
        st.subheader("📊 超额获利可视化")
        fig = px.bar(
            excess_profit_df,
            x="时段",
            y="超额获利(元)",
            color="场站名称",
            title="各场站分时段超额获利",
            barmode="group"
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("ℹ️ 暂无超额获利（或数据不完整）")
