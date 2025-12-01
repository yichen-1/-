import streamlit as st
import pandas as pd
import re
from io import BytesIO
import datetime
import plotly.express as px

# -------------------------- 页面基础配置 --------------------------
st.set_page_config(
    page_title="光伏/风电数据管理工具（完整版）",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------- 会话状态初始化（数据持久化） --------------------------
# 核心数据存储（刷新页面不丢失，关闭页面重置）
if "core_data" not in st.session_state:
    st.session_state.core_data = {
        "generated": {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "total": {}},  # 实发数据
        "hold": {"total": {}, "config": {}},  # 持仓数据
        "price": {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "config": {}}  # 电价数据
    }
# 配置参数存储（各模块独立配置）
if "module_config" not in st.session_state:
    st.session_state.module_config = {
        "generated": {"time_col": 4, "wind_power_col": 9, "pv_power_col": 5, "pv_list": "浠水渔光,襄北农光", "conv": 1000, "skip_rows": 1, "keyword": "历史趋势"},
        "hold": {"hold_col": 3, "skip_rows": 1},
        "price": {"spot_col": 1, "wind_contract_col": 2, "pv_contract_col": 3, "skip_rows": 1}  # 现货/风电合约/光伏合约列索引
    }
# 场站类型映射（固定配置，可修改）
STATION_TYPE_MAP = {
    "风电": ["荆门栗溪", "荆门圣境山", "襄北风储二期", "襄北风储一期", "襄州峪山一期"],
    "光伏": ["襄北农光", "浠水渔光"]
}

# -------------------------- 侧边栏：功能模块收纳（核心优化） --------------------------
st.sidebar.title("⚙️ 功能模块配置")

# 1. 场站实发配置模块（完整收纳上传+参数）
with st.sidebar.expander("📊 模块1：场站实发配置", expanded=True):
    st.sidebar.subheader("1.1 数据上传")
    uploaded_generated = st.sidebar.file_uploader(
        "上传实发Excel文件（支持多选，含「历史趋势」关键词）",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="gen_upload"
    )

    st.sidebar.subheader("1.2 列索引配置（索引从0开始）")
    st.session_state.module_config["generated"]["time_col"] = st.sidebar.number_input("时间列（E列=4）", value=4, min_value=0, key="gen_time_col")
    st.session_state.module_config["generated"]["wind_power_col"] = st.sidebar.number_input("风电功率列（J列=9）", value=9, min_value=0, key="gen_wind_col")
    st.session_state.module_config["generated"]["pv_power_col"] = st.sidebar.number_input("光伏功率列（F列=5）", value=5, min_value=0, key="gen_pv_col")

    st.sidebar.subheader("1.3 基础参数配置")
    st.session_state.module_config["generated"]["pv_list"] = st.sidebar.text_input("光伏场站名单（逗号分隔）", value="浠水渔光,襄北农光", key="gen_pv_list")
    st.session_state.module_config["generated"]["conv"] = st.sidebar.number_input("功率转换系数（kW→MW）", value=1000, key="gen_conv")
    st.session_state.module_config["generated"]["skip_rows"] = st.sidebar.number_input("跳过表头行数", value=1, min_value=0, key="gen_skip")
    st.session_state.module_config["generated"]["keyword"] = st.sidebar.text_input("文件筛选关键词", value="历史趋势", key="gen_keyword")

# 2. 中长期持仓配置模块
with st.sidebar.expander("📦 模块2：中长期持仓配置", expanded=False):
    st.sidebar.subheader("2.1 数据上传")
    uploaded_hold = st.sidebar.file_uploader(
        "上传持仓Excel文件（D列为净持有电量）",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="hold_upload"
    )

    st.sidebar.subheader("2.2 配置参数")
    st.session_state.module_config["hold"]["hold_col"] = st.sidebar.number_input("净持有电量列（D列=3）", value=3, min_value=0, key="hold_col")
    st.session_state.module_config["hold"]["skip_rows"] = st.sidebar.number_input("跳过表头行数", value=1, min_value=0, key="hold_skip")

    st.sidebar.subheader("2.3 场站关联")
    # 下拉选择已提取的实发场站
    generated_stations = list(st.session_state.core_data["generated"]["total"].keys())
    selected_hold_stations = st.sidebar.multiselect(
        "选择持仓关联的场站（从实发场站中选）",
        options=generated_stations,
        key="hold_stations"
    )

# 3. 月度电价配置模块（新增）
with st.sidebar.expander("💰 模块3：月度电价配置", expanded=False):
    st.sidebar.subheader("3.1 数据上传")
    uploaded_price = st.sidebar.file_uploader(
        "上传月度电价Excel文件（含24时段现货+合约价）",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=False,  # 电价文件仅需1个
        key="price_upload"
    )

    st.sidebar.subheader("3.2 列索引配置（索引从0开始）")
    st.session_state.module_config["price"]["spot_col"] = st.sidebar.number_input("现货均价列", value=1, min_value=0, key="price_spot_col")
    st.session_state.module_config["price"]["wind_contract_col"] = st.sidebar.number_input("风电中长期合约均价列", value=2, min_value=0, key="price_wind_col")
    st.session_state.module_config["price"]["pv_contract_col"] = st.sidebar.number_input("光伏中长期合约均价列", value=3, min_value=0, key="price_pv_col")
    st.session_state.module_config["price"]["skip_rows"] = st.sidebar.number_input("跳过表头行数", value=1, min_value=0, key="price_skip")

# -------------------------- 核心工具函数库 --------------------------
class DataProcessor:
    """数据处理工具类（按模块封装）"""
    @staticmethod
    @st.cache_data(show_spinner="清洗功率数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
    def clean_power(value):
        """清洗功率数据"""
        if pd.isna(value):
            return None
        val_str = str(value).strip()
        if not re.search(r'\d', val_str):
            return None
        match = re.search(r'(\d+\.?\d*)', val_str)
        return float(match.group(1)) if match else None

    @staticmethod
    @st.cache_data(show_spinner="提取实发数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
    def extract_generated(file, config):
        """提取单个实发文件数据"""
        try:
            # 读取文件
            suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                usecols=[config["time_col"], config["power_col"]],
                skiprows=config["skip_rows"],
                engine=engine,
                nrows=None
            )
            df.columns = ["时间", "功率(kW)"]

            # 数据清洗
            df["功率(kW)"] = df["功率(kW)"].apply(DataProcessor.clean_power)
            df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
            df = df.dropna(subset=["时间", "功率(kW)"]).sort_values("时间").reset_index(drop=True)

            # 转换单位（kW→MW）
            station_name = file.name.split(".")[0].split("-")[0].strip()
            df[station_name] = df["功率(kW)"] / config["conv"]
            return df[["时间", station_name]], station_name
        except Exception as e:
            st.error(f"实发文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame(), ""

    @staticmethod
    @st.cache_data(show_spinner="提取持仓数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
    def extract_hold(file, config):
        """提取持仓文件数据（D列净持有电量）"""
        try:
            suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if suffix in ["xlsx", "xlsm"] else "xlrd"
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
            return round(df["净持有电量"].sum(), 2)  # 返回总持仓量
        except Exception as e:
            st.error(f"持仓文件[{file.name}]处理失败：{str(e)}")
            return 0.0

    @staticmethod
    @st.cache_data(show_spinner="提取电价数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
    def extract_price(file, config):
        """提取电价文件数据（24时段现货+合约价）"""
        try:
            suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                usecols=[0, config["spot_col"], config["wind_contract_col"], config["pv_contract_col"]],  # 0列为时段列
                skiprows=config["skip_rows"],
                engine=engine,
                nrows=24  # 仅读取24行（对应0-23时）
            )
            df.columns = ["时段", "现货均价(元/MWh)", "风电合约均价(元/MWh)", "光伏合约均价(元/MWh)"]
            
            # 数据清洗：时段格式化为00:00-23:00，价格转数值
            df["时段"] = df["时段"].apply(lambda x: f"{int(x):02d}:00" if pd.notna(x) and str(x).isdigit() else f"{i:02d}:00" for i in range(24))
            for col in ["现货均价(元/MWh)", "风电合约均价(元/MWh)", "光伏合约均价(元/MWh)"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            
            return df
        except Exception as e:
            st.error(f"电价文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame()

    @staticmethod
    def calculate_24h_generated(merged_raw, config):
        """计算24时段实发汇总"""
        if merged_raw.empty:
            return pd.DataFrame(), {}
        
        # 计算时间间隔
        time_diff = merged_raw["时间"].diff().dropna()
        avg_interval = time_diff.dt.total_seconds().mean() / 3600  # 小时
        merged_raw["时段"] = merged_raw["时间"].dt.hour.apply(lambda x: f"{x:02d}:00")

        # 按时段汇总
        stations = [col for col in merged_raw.columns if col not in ["时间", "时段"]]
        generated_24h = merged_raw.groupby("时段")[stations].apply(
            lambda x: (x * avg_interval).sum()  # 电量=功率*时间
        ).round(2).reset_index()

        # 计算月度总实发
        monthly_total = {station: round(generated_24h[station].sum(), 2) for station in stations}
        return generated_24h, monthly_total

    @staticmethod
    def calculate_excess_profit(generated_24h, hold_total, price_24h):
        """计算超额获利回收（按风电/光伏区分）"""
        if generated_24h.empty or not hold_total or price_24h.empty:
            return pd.DataFrame()
        
        # 合并数据（按时段匹配）
        merged_data = pd.merge(generated_24h, price_24h, on="时段", how="inner")
        result_rows = []

        # 遍历每个场站、每个时段计算
        for station in [col for col in generated_24h.columns if col != "时段"]:
            # 判断场站类型（风电/光伏）
            station_type = "风电" if station in STATION_TYPE_MAP["风电"] else "光伏" if station in STATION_TYPE_MAP["光伏"] else None
            if not station_type:
                st.warning(f"场站[{station}]未配置类型（风电/光伏），跳过计算")
                continue
            
            # 获取该场站的持仓量（均分至24时段）
            station_hold = hold_total.get(station, 0) / 24  # 时段持仓量
            contract_col = "风电合约均价(元/MWh)" if station_type == "风电" else "光伏合约均价(元/MWh)"

            for _, row in merged_data.iterrows():
                spot_price = row["现货均价(元/MWh)"]
                contract_price = row[contract_col]
                price_diff = spot_price - contract_price  # 价差
                generated = row[station]  # 时段实发量

                # 按公式计算超额获利回收
                if price_diff > 0:
                    excess = (generated * 0.8 - station_hold * 0.7) * price_diff
                else:
                    excess = (generated * 0.8 - station_hold * 1.3) * price_diff

                result_rows.append({
                    "时段": row["时段"],
                    "场站名": station,
                    "场站类型": station_type,
                    "时段实发量(MWh)": generated,
                    "时段持仓量(MWh)": round(station_hold, 2),
                    "现货均价(元/MWh)": spot_price,
                    "合约均价(元/MWh)": contract_price,
                    "价差(元/MWh)": round(price_diff, 2),
                    "超额获利回收(元)": round(excess, 2)
                })

        return pd.DataFrame(result_rows)

# -------------------------- 主界面：功能执行与结果展示 --------------------------
st.title("📊 光伏/风电数据管理工具（实发+持仓+电价计算）")
st.markdown("---")
processor = DataProcessor()

# -------------------------- 1. 实发数据处理（模块1执行） --------------------------
st.subheader("📊 模块1：场站实发数据处理")
if uploaded_generated:
    if st.button("🚀 执行实发数据提取与汇总", type="primary", key="exec_gen"):
        with st.spinner("正在处理实发文件..."):
            # 1. 筛选含关键词的文件
            config = st.session_state.module_config["generated"]
            target_files = [f for f in uploaded_generated if config["keyword"] in f.name]
            if not target_files:
                st.error(f"未找到含关键词「{config['keyword']}」的文件")
                st.stop()

            # 2. 批量提取实发数据
            all_generated = []
            for file in target_files:
                # 判断场站类型（风电/光伏），选择对应功率列
                station_name = file.name.split(".")[0].split("-")[0].strip()
                pv_list = [s.strip() for s in config["pv_list"].split(",") if s.strip()]
                config["power_col"] = config["pv_power_col"] if station_name in pv_list else config["wind_power_col"]
                
                file_data, station = processor.extract_generated(file, config)
                if not file_data.empty:
                    all_generated.append(file_data)

            # 3. 合并实发数据
            if not all_generated:
                st.error("未提取到有效实发数据")
                st.stop()
            merged_raw = all_generated[0]
            for df in all_generated[1:]:
                merged_raw = pd.merge(merged_raw, df, on="时间", how="outer")
            merged_raw = merged_raw.sort_values("时间").reset_index(drop=True)

            # 4. 计算24时段汇总与月度总实发
            generated_24h, monthly_total = processor.calculate_24h_generated(merged_raw, config)

            # 5. 保存到会话状态
            st.session_state.core_data["generated"]["raw"] = merged_raw
            st.session_state.core_data["generated"]["24h"] = generated_24h
            st.session_state.core_data["generated"]["total"] = monthly_total

            st.success("✅ 实发数据处理完成！")

# 实发结果展示
if not st.session_state.core_data["generated"]["raw"].empty:
    # 原始数据预览
    with st.expander("查看实发原始数据（前20/后20条）", expanded=False):
        raw = st.session_state.core_data["generated"]["raw"]
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("早期数据（前20条）")
            st.dataframe(raw.head(20), use_container_width=True)
        with col2:
            st.subheader("后期数据（后20条）")
            st.dataframe(raw.tail(20), use_container_width=True)

    # 24时段汇总展示
    with st.expander("查看24时段实发汇总", expanded=True):
        generated_24h = st.session_state.core_data["generated"]["24h"]
        st.dataframe(generated_24h, use_container_width=True)

        # 月度总实发统计
        monthly_total = st.session_state.core_data["generated"]["total"]
        st.subheader("月度实发总量统计")
        total_df = pd.DataFrame([{"场站名": k, "月度实发总量(MWh)": v} for k, v in monthly_total.items()])
        st.dataframe(total_df, use_container_width=True)

    # 下载
    st.subheader("实发数据下载")
    current_month = datetime.datetime.now().strftime("%Y%m")
    # 原始数据
    raw_excel = to_excel(st.session_state.core_data["generated"]["raw"], "实发原始数据")
    st.download_button(
        "下载实发原始数据",
        data=raw_excel,
        file_name=f"实发原始数据_{current_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    # 24时段汇总
    gen24h_excel = to_excel(st.session_state.core_data["generated"]["24h"], "24时段实发汇总")
    st.download_button(
        "下载24时段实发汇总",
        data=gen24h_excel,
        file_name=f"24时段实发汇总_{current_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")

# -------------------------- 2. 持仓数据处理（模块2执行） --------------------------
st.subheader("📦 模块2：中长期持仓数据处理")
if uploaded_hold and selected_hold_stations:
    if st.button("🚀 执行持仓数据提取与关联", type="primary", key="exec_hold"):
        with st.spinner("正在处理持仓文件..."):
            config = st.session_state.module_config["hold"]
            # 1. 提取所有持仓文件的总电量
            total_hold = 0.0
            for file in uploaded_hold:
                file_hold = processor.extract_hold(file, config)
                total_hold += file_hold

            # 2. 均分至所选场站（可修改为自定义分配逻辑）
            hold_per_station = round(total_hold / len(selected_hold_stations), 2) if selected_hold_stations else 0.0
            hold_total = {station: hold_per_station for station in selected_hold_stations}

            # 3. 保存到会话状态
            st.session_state.core_data["hold"]["total"] = hold_total
            st.session_state.core_data["hold"]["config"] = config

            st.success(f"✅ 持仓数据处理完成！总持仓量：{total_hold} MWh，均分至{len(selected_hold_stations)}个场站")

# 持仓结果展示
if st.session_state.core_data["hold"]["total"]:
    st.subheader("持仓数据关联结果")
    hold_df = pd.DataFrame([{"场站名": k, "月度总持仓量(MWh)": v} for k, v in st.session_state.core_data["hold"]["total"].items()])
    st.dataframe(hold_df, use_container_width=True)

st.markdown("---")

# -------------------------- 3. 电价数据处理与超额获利计算（模块3执行，新增） --------------------------
st.subheader("💰 模块3：月度电价处理与超额获利回收计算")
if uploaded_price:
    if st.button("🚀 执行电价提取与超额获利计算", type="primary", key="exec_price"):
        with st.spinner("正在处理电价文件并计算超额获利..."):
            # 1. 检查前置数据（实发+持仓）
            if not st.session_state.core_data["generated"]["24h"].empty and not st.session_state.core_data["hold"]["total"]:
                st.error("请先处理「模块2：中长期持仓数据」，再计算超额获利")
                st.stop()
            if st.session_state.core_data["generated"]["24h"].empty:
                st.error("请先处理「模块1：场站实发数据」，再计算超额获利")
                st.stop()

            # 2. 提取电价数据
            config = st.session_state.module_config["price"]
            price_24h = processor.extract_price(uploaded_price, config)
            if price_24h.empty:
                st.error("未提取到有效电价数据")
                st.stop()

            # 3. 计算超额获利回收
            excess_profit = processor.calculate_excess_profit(
                generated_24h=st.session_state.core_data["generated"]["24h"],
                hold_total=st.session_state.core_data["hold"]["total"],
                price_24h=price_24h
            )

            # 4. 保存到会话状态
            st.session_state.core_data["price"]["raw"] = price_24h
            st.session_state.core_data["price"]["excess_profit"] = excess_profit

            st.success("✅ 电价处理与超额获利计算完成！")

# 电价与超额获利结果展示
if not st.session_state.core_data["price"]["raw"].empty:
    # 电价数据展示
    with st.expander("查看24时段电价数据", expanded=False):
        st.dataframe(st.session_state.core_data["price"]["raw"], use_container_width=True)

# 超额获利结果展示
if "excess_profit" in st.session_state.core_data["price"] and not st.session_state.core_data["price"]["excess_profit"].empty:
    excess_df = st.session_state.core_data["price"]["excess_profit"]
    with st.expander("查看超额获利回收明细（24时段×场站）", expanded=True):
        st.dataframe(excess_df, use_container_width=True)

    # 超额获利汇总（按场站/类型）
    st.subheader("超额获利回收汇总统计")
    # 按场站汇总
    station_excess = excess_df.groupby("场站名")["超额获利回收(元)"].sum().round(2).reset_index()
    station_excess.columns = ["场站名", "月度超额获利回收(元)"]
    # 按类型汇总
    type_excess = excess_df.groupby("场站类型")["超额获利回收(元)"].sum().round(2).reset_index()
    type_excess.columns = ["场站类型", "月度超额获利回收(元)"]

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("按场站汇总")
        st.dataframe(station_excess, use_container_width=True)
    with col2:
        st.subheader("按类型汇总")
        st.dataframe(type_excess, use_container_width=True)

    # 下载
    st.subheader("电价与超额获利数据下载")
    current_month = datetime.datetime.now().strftime("%Y%m")
    # 电价数据
    price_excel = to_excel(st.session_state.core_data["price"]["raw"], "24时段电价数据")
    st.download_button(
        "下载24时段电价数据",
        data=price_excel,
        file_name=f"24时段电价数据_{current_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    # 超额获利明细
    excess_excel = to_excel(excess_df, "超额获利回收明细")
    st.download_button(
        "下载超额获利回收明细",
        data=excess_excel,
        file_name=f"超额获利回收明细_{current_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")

# -------------------------- 全局功能：数据重置 --------------------------
if st.button("🗑️ 重置所有模块数据（实发+持仓+电价）", type="secondary"):
    st.session_state.core_data = {
        "generated": {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "total": {}},
        "hold": {"total": {}, "config": {}},
        "price": {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "config": {}}
    }
    st.success("✅ 所有模块数据已重置！")

# -------------------------- 辅助函数：Excel下载 --------------------------
def to_excel(df, sheet_name="数据"):
    """DataFrame转Excel字节流"""
    if df.empty:
        st.warning("数据为空，无法生成Excel")
        return BytesIO()
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output
