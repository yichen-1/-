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
# 场站类型映射（可直接修改）
STATION_TYPE_MAP = {
    "风电": ["荆门栗溪", "荆门圣境山", "襄北风储二期", "襄北风储一期", "襄州峪山一期"],
    "光伏": ["襄北农光", "浠水渔光"]
}

# -------------------------- 3. 会话状态初始化（数据持久化，刷新不丢失） --------------------------
# 核心业务数据存储
if "core_data" not in st.session_state:
    st.session_state.core_data = {
        "generated": {  # 实发数据
            "raw": pd.DataFrame(),       # 原始实发数据（带时间戳）
            "24h": pd.DataFrame(),       # 24时段实发汇总
            "total": {}                  # 月度总实发（场站名: 总量MWh）
        },
        "hold": {       # 持仓数据
            "total": {},                 # 月度总持仓（场站名: 总量MWh）
            "config": {}                 # 持仓配置参数
        },
        "price": {      # 电价数据
            "24h": pd.DataFrame(),       # 24时段电价数据
            "excess_profit": pd.DataFrame()  # 超额获利回收计算结果
        }
    }

# 模块配置参数存储（独立配置，互不影响）
if "module_config" not in st.session_state:
    st.session_state.module_config = {
        "generated": {  # 实发模块配置
            "time_col": 4,               # 时间列索引（E列）
            "wind_power_col": 9,         # 风电功率列索引（J列）
            "pv_power_col": 5,           # 光伏功率列索引（F列）
            "pv_list": "浠水渔光,襄北农光",  # 光伏场站名单
            "conv": 1000,                # 功率转换系数（kW→MW）
            "skip_rows": 1,              # 跳过表头行数
            "keyword": "历史趋势"        # 实发文件筛选关键词
        },
        "hold": {       # 持仓模块配置
            "hold_col": 3,               # 净持有电量列索引（D列）
            "skip_rows": 1               # 跳过表头行数
        },
        "price": {      # 电价模块配置
            "spot_col": 1,               # 现货均价列索引
            "wind_contract_col": 2,      # 风电合约均价列索引
            "pv_contract_col": 3,        # 光伏合约均价列索引
            "skip_rows": 1               # 跳过表头行数
        }
    }

# -------------------------- 4. 辅助函数（提前定义，避免调用顺序错误） --------------------------
def to_excel(df, sheet_name="数据"):
    """
    DataFrame转Excel字节流（供下载使用）
    :param df: 待转换的DataFrame
    :param sheet_name: Excel工作表名称
    :return: BytesIO字节流
    """
    if df.empty:
        st.warning("⚠️ 数据为空，无法生成Excel文件")
        return BytesIO()
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

# -------------------------- 5. 核心数据处理类（封装所有计算逻辑） --------------------------
class DataProcessor:
    """数据处理核心类（实发/持仓/电价/获利计算）"""

    @staticmethod
    @st.cache_data(show_spinner="清洗功率数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
    def clean_power_value(value):
        """清洗功率数据（过滤非数值、空值）"""
        if pd.isna(value):
            return None
        val_str = str(value).strip()
        # 提取数值部分（支持小数）
        num_match = re.search(r'(\d+\.?\d*)', val_str)
        if not num_match:
            return None
        try:
            return float(num_match.group(1))
        except:
            return None

    @staticmethod
    @st.cache_data(show_spinner="提取实发数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
    def extract_generated_data(file, config):
        """
        提取单个实发文件数据
        :param file: 上传的Excel文件
        :param config: 实发模块配置参数
        :return: (处理后的DataFrame, 场站名)
        """
        try:
            # 1. 读取文件（仅读取时间列+功率列）
            file_suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                usecols=[config["time_col"], config["power_col"]],
                skiprows=config["skip_rows"],
                engine=engine,
                nrows=None
            )
            df.columns = ["时间", "功率(kW)"]

            # 2. 数据清洗
            df["功率(kW)"] = df["功率(kW)"].apply(DataProcessor.clean_power_value)
            df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
            # 过滤空值并排序
            df = df.dropna(subset=["时间", "功率(kW)"]).sort_values("时间").reset_index(drop=True)

            # 3. 提取场站名+转换功率单位（kW→MW）
            station_name = file.name.split(".")[0].split("-")[0].strip()
            df[station_name] = df["功率(kW)"] / config["conv"]

            return df[["时间", station_name]], station_name
        except Exception as e:
            st.error(f"❌ 实发文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame(), ""

    @staticmethod
    @st.cache_data(show_spinner="提取持仓数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
    def extract_hold_data(file, config):
        """
        提取持仓文件的净持有电量（D列求和）
        :param file: 上传的Excel文件
        :param config: 持仓模块配置参数
        :return: 总净持有电量（MWh，保留2位小数）
        """
        try:
            # 1. 读取文件（仅读取净持有电量列）
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

            # 2. 数据清洗（转数值+填充空值）
            df["净持有电量"] = pd.to_numeric(df["净持有电量"], errors="coerce").fillna(0)

            # 3. 计算总净持有电量
            total_hold = round(df["净持有电量"].sum(), 2)
            return total_hold
        except Exception as e:
            st.error(f"❌ 持仓文件[{file.name}]处理失败：{str(e)}")
            return 0.0

    @staticmethod
    @st.cache_data(show_spinner="提取电价数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
    def extract_price_data(file, config):
        """
        提取24时段电价数据（现货+风电合约+光伏合约）
        :param file: 上传的Excel文件
        :param config: 电价模块配置参数
        :return: 24时段电价DataFrame
        """
        try:
            # 1. 读取文件（仅读取时段列+3个价格列，最多24行）
            file_suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(
                BytesIO(file.getvalue()),
                header=None,
                usecols=[0, config["spot_col"], config["wind_contract_col"], config["pv_contract_col"]],
                skiprows=config["skip_rows"],
                engine=engine,
                nrows=24  # 仅读取24行（对应0-23时）
            )
            df.columns = ["时段", "现货均价(元/MWh)", "风电合约均价(元/MWh)", "光伏合约均价(元/MWh)"]

            # 2. 数据清洗
            # 格式化时段（00:00 ~ 23:00）
            df["时段"] = [f"{i:02d}:00" for i in range(24)]
            # 价格转数值+填充空值
            price_cols = ["现货均价(元/MWh)", "风电合约均价(元/MWh)", "光伏合约均价(元/MWh)"]
            for col in price_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

            return df
        except Exception as e:
            st.error(f"❌ 电价文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame()

    @staticmethod
    def calculate_24h_generated(merged_raw_df, config):
        """
        计算24时段实发汇总+月度总实发
        :param merged_raw_df: 合并后的原始实发数据
        :param config: 实发模块配置参数
        :return: (24时段汇总DF, 月度总实发字典)
        """
        if merged_raw_df.empty:
            st.warning("⚠️ 实发原始数据为空，无法计算24时段汇总")
            return pd.DataFrame(), {}

        # 1. 计算时间间隔（小时）：用于功率→电量转换（电量=功率×时间）
        time_diff = merged_raw_df["时间"].diff().dropna()
        avg_interval_h = time_diff.dt.total_seconds().mean() / 3600
        if avg_interval_h == 0:  # 容错：避免时间间隔为0
            avg_interval_h = 1/4  # 默认15分钟

        # 2. 提取时段（00:00 ~ 23:00）
        merged_raw_df["时段"] = merged_raw_df["时间"].dt.hour.apply(lambda x: f"{x:02d}:00")

        # 3. 按时段汇总实发量（电量=功率×时间间隔）
        station_cols = [col for col in merged_raw_df.columns if col not in ["时间", "时段"]]
        generated_24h_df = merged_raw_df.groupby("时段")[station_cols].apply(
            lambda x: (x * avg_interval_h).sum()
        ).round(2).reset_index()

        # 4. 计算月度总实发（各场站24时段求和）
        monthly_total = {
            station: round(generated_24h_df[station].sum(), 2)
            for station in station_cols
        }

        return generated_24h_df, monthly_total

    @staticmethod
    def calculate_excess_profit(generated_24h_df, hold_total_dict, price_24h_df):
        """
        计算超额获利回收（按风电/光伏差异化计算）
        :param generated_24h_df: 24时段实发汇总
        :param hold_total_dict: 月度总持仓（场站名: 总量）
        :param price_24h_df: 24时段电价数据
        :return: 超额获利回收明细DF
        """
        # 前置检查
        if generated_24h_df.empty:
            st.warning("⚠️ 24时段实发数据为空，无法计算超额获利")
            return pd.DataFrame()
        if not hold_total_dict:
            st.warning("⚠️ 持仓数据为空，无法计算超额获利")
            return pd.DataFrame()
        if price_24h_df.empty:
            st.warning("⚠️ 24时段电价数据为空，无法计算超额获利")
            return pd.DataFrame()

        # 1. 合并实发+电价数据（按时段匹配）
        merged_df = pd.merge(generated_24h_df, price_24h_df, on="时段", how="inner")
        if merged_df.empty:
            st.warning("⚠️ 实发与电价数据时段不匹配，无法计算")
            return pd.DataFrame()

        # 2. 逐时段+逐场站计算超额获利
        result_rows = []
        station_cols = [col for col in generated_24h_df.columns if col != "时段"]

        for station in station_cols:
            # 识别场站类型（风电/光伏）
            if station in STATION_TYPE_MAP["风电"]:
                station_type = "风电"
                contract_col = "风电合约均价(元/MWh)"
            elif station in STATION_TYPE_MAP["光伏"]:
                station_type = "光伏"
                contract_col = "光伏合约均价(元/MWh)"
            else:
                st.warning(f"⚠️ 场站[{station}]未配置类型（风电/光伏），跳过计算")
                continue

            # 计算时段持仓量（月度总持仓均分至24时段）
            monthly_hold = hold_total_dict.get(station, 0)
            hourly_hold = round(monthly_hold / 24, 2)  # 时段持仓量

            # 逐时段计算
            for _, row in merged_df.iterrows():
                # 基础数据
                hour_generated = row[station]          # 时段实发量(MWh)
                spot_price = row["现货均价(元/MWh)"]    # 现货均价
                contract_price = row[contract_col]     # 合约均价
                price_diff = spot_price - contract_price  # 价差

                # 核心公式（严格按需求）
                if price_diff > 0:
                    excess_profit = (hour_generated * 0.8 - hourly_hold * 0.7) * price_diff
                else:
                    excess_profit = (hour_generated * 0.8 - hourly_hold * 1.3) * price_diff

                # 存储结果
                result_rows.append({
                    "时段": row["时段"],
                    "场站名": station,
                    "场站类型": station_type,
                    "时段实发量(MWh)": round(hour_generated, 2),
                    "时段持仓量(MWh)": hourly_hold,
                    "现货均价(元/MWh)": round(spot_price, 2),
                    "合约均价(元/MWh)": round(contract_price, 2),
                    "价差(元/MWh)": round(price_diff, 2),
                    "超额获利回收(元)": round(excess_profit, 2)
                })

        # 转换为DataFrame
        result_df = pd.DataFrame(result_rows)
        return result_df

# -------------------------- 6. 侧边栏：模块化配置（核心收纳优化） --------------------------
st.sidebar.title("⚙️ 功能模块配置")

# 6.1 模块1：场站实发配置（完整收纳）
with st.sidebar.expander("📊 模块1：场站实发配置", expanded=True):
    st.sidebar.subheader("1.1 数据上传")
    uploaded_generated_files = st.sidebar.file_uploader(
        "上传实发Excel文件（支持多选，含「历史趋势」关键词）",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="gen_upload"
    )

    st.sidebar.subheader("1.2 列索引配置（索引从0开始）")
    st.session_state.module_config["generated"]["time_col"] = st.sidebar.number_input(
        "时间列（E列=4）", value=4, min_value=0, key="gen_time_col"
    )
    st.session_state.module_config["generated"]["wind_power_col"] = st.sidebar.number_input(
        "风电功率列（J列=9）", value=9, min_value=0, key="gen_wind_col"
    )
    st.session_state.module_config["generated"]["pv_power_col"] = st.sidebar.number_input(
        "光伏功率列（F列=5）", value=5, min_value=0, key="gen_pv_col"
    )

    st.sidebar.subheader("1.3 基础参数配置")
    st.session_state.module_config["generated"]["pv_list"] = st.sidebar.text_input(
        "光伏场站名单（逗号分隔）", value="浠水渔光,襄北农光", key="gen_pv_list"
    )
    st.session_state.module_config["generated"]["conv"] = st.sidebar.number_input(
        "功率转换系数（kW→MW）", value=1000, key="gen_conv"
    )
    st.session_state.module_config["generated"]["skip_rows"] = st.sidebar.number_input(
        "跳过表头行数", value=1, min_value=0, key="gen_skip"
    )
    st.session_state.module_config["generated"]["keyword"] = st.sidebar.text_input(
        "文件筛选关键词", value="历史趋势", key="gen_keyword"
    )

# 6.2 模块2：中长期持仓配置（完整收纳）
with st.sidebar.expander("📦 模块2：中长期持仓配置", expanded=False):
    st.sidebar.subheader("2.1 数据上传")
    uploaded_hold_files = st.sidebar.file_uploader(
        "上传持仓Excel文件（D列为净持有电量）",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="hold_upload"
    )

    st.sidebar.subheader("2.2 配置参数")
    st.session_state.module_config["hold"]["hold_col"] = st.sidebar.number_input(
        "净持有电量列（D列=3）", value=3, min_value=0, key="hold_col"
    )
    st.session_state.module_config["hold"]["skip_rows"] = st.sidebar.number_input(
        "跳过表头行数", value=1, min_value=0, key="hold_skip"
    )

    st.sidebar.subheader("2.3 场站关联")
    # 仅显示已提取的实发场站（避免手动输入错误）
    generated_stations = list(st.session_state.core_data["generated"]["total"].keys())
    selected_hold_stations = st.sidebar.multiselect(
        "选择持仓关联的场站（从实发场站中选）",
        options=generated_stations,
        key="hold_stations"
    )

# 6.3 模块3：月度电价配置（完整收纳）
with st.sidebar.expander("💰 模块3：月度电价配置", expanded=False):
    st.sidebar.subheader("3.1 数据上传")
    uploaded_price_file = st.sidebar.file_uploader(
        "上传月度电价Excel文件（含24时段现货+合约价）",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=False,  # 电价文件仅需1个
        key="price_upload"
    )

    st.sidebar.subheader("3.2 列索引配置（索引从0开始）")
    st.session_state.module_config["price"]["spot_col"] = st.sidebar.number_input(
        "现货均价列", value=1, min_value=0, key="price_spot_col"
    )
    st.session_state.module_config["price"]["wind_contract_col"] = st.sidebar.number_input(
        "风电合约均价列", value=2, min_value=0, key="price_wind_col"
    )
    st.session_state.module_config["price"]["pv_contract_col"] = st.sidebar.number_input(
        "光伏合约均价列", value=3, min_value=0, key="price_pv_col"
    )
    st.session_state.module_config["price"]["skip_rows"] = st.sidebar.number_input(
        "跳过表头行数", value=1, min_value=0, key="price_skip"
    )

# -------------------------- 7. 主界面：功能执行与结果展示 --------------------------
st.title("📊 光伏/风电数据管理工具（实发+持仓+电价计算）")
st.markdown("---")
processor = DataProcessor()
current_month = datetime.datetime.now().strftime("%Y%m")

# ======================== 7.1 模块1：实发数据处理 ========================
st.subheader("📊 模块1：场站实发数据处理")
if uploaded_generated_files:
    if st.button("🚀 执行实发数据提取与汇总", type="primary", key="exec_gen"):
        with st.spinner("正在处理实发文件..."):
            # 1. 筛选含关键词的文件
            gen_config = st.session_state.module_config["generated"]
            target_files = [f for f in uploaded_generated_files if gen_config["keyword"] in f.name]
            if not target_files:
                st.error(f"❌ 未找到含关键词「{gen_config['keyword']}」的实发文件")
                st.stop()

            # 2. 批量提取实发数据
            all_gen_dfs = []
            pv_list = [s.strip() for s in gen_config["pv_list"].split(",") if s.strip()]
            
            for file in target_files:
                # 识别场站类型，选择对应功率列
                station_name = file.name.split(".")[0].split("-")[0].strip()
                gen_config["power_col"] = gen_config["pv_power_col"] if station_name in pv_list else gen_config["wind_power_col"]
                
                # 提取单个文件数据
                file_df, station = processor.extract_generated_data(file, gen_config)
                if not file_df.empty:
                    all_gen_dfs.append(file_df)

            # 3. 合并实发数据
            if not all_gen_dfs:
                st.error("❌ 未提取到任何有效实发数据")
                st.stop()
            merged_raw_df = all_gen_dfs[0]
            for df in all_gen_dfs[1:]:
                merged_raw_df = pd.merge(merged_raw_df, df, on="时间", how="outer")
            merged_raw_df = merged_raw_df.sort_values("时间").reset_index(drop=True)

            # 4. 计算24时段汇总+月度总实发
            generated_24h_df, monthly_total_dict = processor.calculate_24h_generated(merged_raw_df, gen_config)

            # 5. 保存到会话状态
            st.session_state.core_data["generated"]["raw"] = merged_raw_df
            st.session_state.core_data["generated"]["24h"] = generated_24h_df
            st.session_state.core_data["generated"]["total"] = monthly_total_dict

            st.success("✅ 实发数据处理完成！")

# 实发结果展示
gen_core = st.session_state.core_data["generated"]
if not gen_core["raw"].empty:
    # 原始数据预览
    with st.expander("📜 查看实发原始数据（前20/后20条）", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("早期数据（前20条）")
            st.dataframe(gen_core["raw"].head(20), use_container_width=True)
        with col2:
            st.subheader("后期数据（后20条）")
            st.dataframe(gen_core["raw"].tail(20), use_container_width=True)

    # 24时段汇总展示
    with st.expander("📈 查看24时段实发汇总", expanded=True):
        st.dataframe(gen_core["24h"], use_container_width=True)

        # 月度总实发统计
        st.subheader("📊 月度实发总量统计")
        total_df = pd.DataFrame([
            {"场站名": k, "月度实发总量(MWh)": v} for k, v in gen_core["total"].items()
        ])
        st.dataframe(total_df, use_container_width=True)

    # 实发数据下载
    st.subheader("💾 实发数据下载")
    col1, col2 = st.columns(2)
    with col1:
        # 原始数据下载
        raw_excel = to_excel(gen_core["raw"], "实发原始数据")
        st.download_button(
            "下载实发原始数据",
            data=raw_excel,
            file_name=f"实发原始数据_{current_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_gen_raw"
        )
    with col2:
        # 24时段汇总下载
        gen24h_excel = to_excel(gen_core["24h"], "24时段实发汇总")
        st.download_button(
            "下载24时段实发汇总",
            data=gen24h_excel,
            file_name=f"24时段实发汇总_{current_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_gen_24h"
        )

st.markdown("---")

# ======================== 7.2 模块2：持仓数据处理 ========================
st.subheader("📦 模块2：中长期持仓数据处理")
if uploaded_hold_files and selected_hold_stations:
    if st.button("🚀 执行持仓数据提取与关联", type="primary", key="exec_hold"):
        with st.spinner("正在处理持仓文件..."):
            # 1. 提取所有持仓文件的总电量
            hold_config = st.session_state.module_config["hold"]
            total_hold_amount = 0.0
            for file in uploaded_hold_files:
                file_hold = processor.extract_hold_data(file, hold_config)
                total_hold_amount += file_hold

            # 2. 均分至所选场站（可自定义分配逻辑）
            if len(selected_hold_stations) == 0:
                st.error("❌ 未选择持仓关联的场站")
                st.stop()
            hold_per_station = round(total_hold_amount / len(selected_hold_stations), 2)
            hold_total_dict = {station: hold_per_station for station in selected_hold_stations}

            # 3. 保存到会话状态
            st.session_state.core_data["hold"]["total"] = hold_total_dict
            st.session_state.core_data["hold"]["config"] = hold_config

            st.success(f"✅ 持仓数据处理完成！")
            st.info(f"📊 总持仓量：{total_hold_amount} MWh，均分至{len(selected_hold_stations)}个场站（每个场站{hold_per_station} MWh）")

# 持仓结果展示
hold_core = st.session_state.core_data["hold"]
if hold_core["total"]:
    st.subheader("📊 持仓数据关联结果")
    hold_df = pd.DataFrame([
        {"场站名": k, "月度总持仓量(MWh)": v} for k, v in hold_core["total"].items()
    ])
    st.dataframe(hold_df, use_container_width=True)

st.markdown("---")

# ======================== 7.3 模块3：电价+超额获利计算 ========================
st.subheader("💰 模块3：月度电价处理与超额获利回收计算")
if uploaded_price_file:
    if st.button("🚀 执行电价提取与超额获利计算", type="primary", key="exec_price"):
        with st.spinner("正在处理电价文件并计算超额获利..."):
            # 前置检查：实发+持仓数据是否存在
            if gen_core["24h"].empty:
                st.error("❌ 请先完成「模块1：实发数据处理」，再计算超额获利")
                st.stop()
            if not hold_core["total"]:
                st.error("❌ 请先完成「模块2：持仓数据处理」，再计算超额获利")
                st.stop()

            # 1. 提取电价数据
            price_config = st.session_state.module_config["price"]
            price_24h_df = processor.extract_price_data(uploaded_price_file, price_config)
            if price_24h_df.empty:
                st.error("❌ 未提取到有效电价数据")
                st.stop()

            # 2. 计算超额获利回收
            excess_profit_df = processor.calculate_excess_profit(
                generated_24h_df=gen_core["24h"],
                hold_total_dict=hold_core["total"],
                price_24h_df=price_24h_df
            )

            # 3. 保存到会话状态
            st.session_state.core_data["price"]["24h"] = price_24h_df
            st.session_state.core_data["price"]["excess_profit"] = excess_profit_df

            st.success("✅ 电价处理与超额获利计算完成！")

# 电价结果展示
price_core = st.session_state.core_data["price"]
if not price_core["24h"].empty:
    with st.expander("📈 查看24时段电价数据", expanded=False):
        st.dataframe(price_core["24h"], use_container_width=True)

# 超额获利结果展示
if not price_core["excess_profit"].empty:
    excess_df = price_core["excess_profit"]
    with st.expander("📊 查看超额获利回收明细（24时段×场站）", expanded=True):
        st.dataframe(excess_df, use_container_width=True)

    # 超额获利汇总统计
    st.subheader("📊 超额获利回收汇总")
    col1, col2 = st.columns(2)
    with col1:
        # 按场站汇总
        station_excess = excess_df.groupby("场站名")["超额获利回收(元)"].sum().round(2).reset_index()
        station_excess.columns = ["场站名", "月度超额获利回收(元)"]
        st.subheader("按场站汇总")
        st.dataframe(station_excess, use_container_width=True)
    with col2:
        # 按类型汇总
        type_excess = excess_df.groupby("场站类型")["超额获利回收(元)"].sum().round(2).reset_index()
        type_excess.columns = ["场站类型", "月度超额获利回收(元)"]
        st.subheader("按类型汇总")
        st.dataframe(type_excess, use_container_width=True)

    # 电价+超额获利下载
    st.subheader("💾 电价与超额获利数据下载")
    col1, col2 = st.columns(2)
    with col1:
        # 电价数据下载
        price_excel = to_excel(price_core["24h"], "24时段电价数据")
        st.download_button(
            "下载24时段电价数据",
            data=price_excel,
            file_name=f"24时段电价数据_{current_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_price"
        )
    with col2:
        # 超额获利明细下载
        excess_excel = to_excel(excess_df, "超额获利回收明细")
        st.download_button(
            "下载超额获利回收明细",
            data=excess_excel,
            file_name=f"超额获利回收明细_{current_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excess"
        )

# ======================== 7.4 全局功能：数据重置 ========================
st.markdown("---")
if st.button("🗑️ 重置所有模块数据（实发+持仓+电价）", type="secondary", key="reset_all"):
    # 清空所有会话状态数据
    st.session_state.core_data = {
        "generated": {"raw": pd.DataFrame(), "24h": pd.DataFrame(), "total": {}},
        "hold": {"total": {}, "config": {}},
        "price": {"24h": pd.DataFrame(), "excess_profit": pd.DataFrame()}
    }
    st.success("✅ 所有模块数据已重置！")

# -------------------------- 8. 侧边栏使用说明 --------------------------
st.sidebar.markdown("---")
st.sidebar.markdown("### 📝 使用流程指引")
st.sidebar.markdown("""
1. **模块1（实发）**：上传文件→确认配置→执行提取→生成24时段汇总
2. **模块2（持仓）**：上传文件→选择关联场站→执行提取→持仓均分至场站
3. **模块3（电价）**：上传文件→执行计算→生成超额获利明细+汇总

⚠️ 执行顺序：模块1 → 模块2 → 模块3（前置模块未完成无法执行后续）
""")

st.sidebar.markdown("### ℹ️ 关键说明")
st.sidebar.markdown("""
- 场站类型：风电（荆门栗溪等5个）、光伏（襄北农光/浠水渔光）
- 超额获利公式：
  - 价差>0：(实发×0.8 - 持仓×0.7) × 价差
  - 价差<0：(实发×0.8 - 持仓×1.3) × 价差
- 所有数据存储在会话中，刷新页面不丢失（关闭页面重置）
""")
