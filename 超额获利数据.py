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
    page_title="光伏/风电超额获利计算工具（纯时段匹配版）",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------- 2. 全局常量 --------------------------
STATION_TYPE_MAP = {
    "风电": ["荆门栗溪", "荆门圣境山", "襄北风储二期", "襄北风储一期", "襄州峪山一期", "风电"],
    "光伏": ["襄北农光", "浠水渔光", "光伏"]
}
PRICE_TEMPLATE_COLS = ["时段", "风电现货均价(元/MWh)", "风电合约均价(元/MWh)", "光伏现货均价(元/MWh)", "光伏合约均价(元/MWh)"]
STANDARD_HOURS = [f"{i:02d}:00" for i in range(24)]  # 标准24时段

# -------------------------- 3. 核心工具函数 --------------------------
def standardize_column_name(col):
    col_str = str(col).strip() if col is not None else f"未知列_{uuid.uuid4().hex[:8]}"
    col_str = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9_]', '_', col_str).lower()
    if col_str == "" or col_str == "_":
        col_str = f"列_{uuid.uuid4().hex[:8]}"
    return col_str

def force_unique_columns(df):
    df.columns = [standardize_column_name(col) for col in df.columns]
    time_col_candidates = [i for i, col in enumerate(df.columns) if "时间" in col or "时段" in col]
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
    return pd.DataFrame([{"时段": h, "风电现货均价(元/MWh)":0.0, "风电合约均价(元/MWh)":0.0, "光伏现货均价(元/MWh)":0.0, "光伏合约均价(元/MWh)":0.0} for h in STANDARD_HOURS])

def standardize_hour(hour_str):
    try:
        hour_str = str(hour_str).strip().replace("时", "").replace("点", "").replace("：", ":")
        return f"{int(hour_str.split(':')[0] if ':' in hour_str else hour_str):02d}:00"
    except:
        return None

# -------------------------- 4. 会话状态初始化 --------------------------
if "target_month" not in st.session_state:
    st.session_state.target_month = "2025-12"
if "gen_data" not in st.session_state:
    st.session_state.gen_data = {"raw": pd.DataFrame(), "24h": pd.DataFrame()}
if "hold_data" not in st.session_state:
    st.session_state.hold_data = {}  # {时段: 持仓值}
if "hold_data_df" not in st.session_state:
    st.session_state.hold_data_df = pd.DataFrame()
if "price_data" not in st.session_state:
    st.session_state.price_data = {"24h": pd.DataFrame(), "excess_profit": pd.DataFrame()}
if "module_config" not in st.session_state:
    st.session_state.module_config = {
        "generated": {"time_col":4, "power_col":9, "conv":1000, "skip_rows":1},
        "hold": {"hour_col":0, "hold_col":1, "skip_rows":1},
        "price": {"wind_spot_col":1, "wind_contract_col":2, "pv_spot_col":3, "pv_contract_col":4, "skip_rows":1}
    }

# -------------------------- 5. 核心数据处理类（纯时段计算） --------------------------
class DataProcessor:
    @staticmethod
    def extract_generated_data(file, config, station_type):
        try:
            file_suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(BytesIO(file.getvalue()), header=None, usecols=[config["time_col"], config["power_col"]], skiprows=config["skip_rows"], engine=engine)
            df.columns = ["时间", "功率(kW)"]
            df["功率(kW)"] = pd.to_numeric(df["功率(kW)"], errors="coerce").fillna(0)
            df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
            df = df.dropna(subset=["时间"]).sort_values("时间")
            
            # 计算24时段实发量
            df["时段"] = df["时间"].dt.hour.apply(lambda x: f"{x:02d}:00")
            time_diff = df["时间"].diff().dropna()
            avg_interval_h = time_diff.dt.total_seconds().mean() / 3600 if not time_diff.empty else 1/4
            gen_24h = df.groupby("时段")["功率(kW)"].sum() * avg_interval_h / config["conv"]  # 转换为MWh
            
            # 补全24时段
            gen_24h_df = pd.DataFrame({"时段": STANDARD_HOURS})
            gen_24h_df["实发量(MWh)"] = gen_24h_df["时段"].map(gen_24h).fillna(0)
            st.success(f"✅ 实发文件[{file.name}]处理成功，已生成24时段实发数据")
            return gen_24h_df
        except Exception as e:
            st.error(f"❌ 实发文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame()

    @staticmethod
    def extract_hold_data(file, config):
        try:
            file_suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(BytesIO(file.getvalue()), header=None, usecols=[config["hour_col"], config["hold_col"]], skiprows=config["skip_rows"], engine=engine, nrows=24)
            df.columns = ["时段", "持仓量(MWh)"]
            
            # 标准化时段+补全24时段
            df["时段"] = df["时段"].apply(standardize_hour)
            df["持仓量(MWh)"] = pd.to_numeric(df["持仓量(MWh)"], errors="coerce").fillna(0)
            hold_24h_df = pd.DataFrame({"时段": STANDARD_HOURS})
            hold_24h_df["持仓量(MWh)"] = hold_24h_df["时段"].map(dict(zip(df["时段"], df["持仓量(MWh)"]))).fillna(0)
            
            st.success(f"✅ 持仓文件[{file.name}]处理成功，已生成24时段持仓数据")
            return hold_24h_df
        except Exception as e:
            st.error(f"❌ 持仓文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame()

    @staticmethod
    def extract_price_data(file, config):
        try:
            file_suffix = file.name.split(".")[-1].lower()
            engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
            df = pd.read_excel(BytesIO(file.getvalue()), header=None, usecols=[0, config["wind_spot_col"], config["wind_contract_col"], config["pv_spot_col"], config["pv_contract_col"]], skiprows=config["skip_rows"], engine=engine, nrows=24)
            df.columns = ["时段", "风电现货", "风电合约", "光伏现货", "光伏合约"]
            
            # 标准化时段+补全24时段
            df["时段"] = df["时段"].apply(standardize_hour)
            price_24h_df = pd.DataFrame({"时段": STANDARD_HOURS})
            for col in ["风电现货", "风电合约", "光伏现货", "光伏合约"]:
                price_24h_df[col] = price_24h_df["时段"].map(dict(zip(df["时段"], df[col]))).fillna(0)
            
            st.success(f"✅ 电价文件[{file.name}]处理成功，已生成24时段电价数据")
            return price_24h_df
        except Exception as e:
            st.error(f"❌ 电价文件[{file.name}]处理失败：{str(e)}")
            return pd.DataFrame()

    @staticmethod
    def calculate_profit(gen_df, hold_df, price_df, station_type):
        if gen_df.empty or hold_df.empty or price_df.empty:
            st.error("❌ 实发/持仓/电价数据不完整")
            return pd.DataFrame()
        
        # 合并数据（纯时段匹配，不看名称）
        merged_df = pd.merge(gen_df, hold_df, on="时段", how="inner")
        merged_df = pd.merge(merged_df, price_df, on="时段", how="inner")
        
        # 选择对应类型的电价
        if station_type == "风电":
            merged_df["现货价"] = merged_df["风电现货"]
            merged_df["合约价"] = merged_df["风电合约"]
        else:
            merged_df["现货价"] = merged_df["光伏现货"]
            merged_df["合约价"] = merged_df["光伏合约"]
        
        # 计算核心逻辑
        merged_df["修正后实发量"] = merged_df["实发量(MWh)"] * (0.7 if station_type == "风电" else 0.8)  # 修正系数
        merged_df["合约0.9倍"] = merged_df["持仓量(MWh)"] * 0.9
        merged_df["合约1.1倍"] = merged_df["持仓量(MWh)"] * 1.1
        
        # 电量差额（超出1.1倍或低于0.9倍的部分）
        merged_df["电量差额"] = np.where(merged_df["修正后实发量"] > merged_df["合约1.1倍"], 
                                       merged_df["修正后实发量"] - merged_df["合约1.1倍"],
                                       np.where(merged_df["修正后实发量"] < merged_df["合约0.9倍"], 
                                                merged_df["修正后实发量"] - merged_df["合约0.9倍"], 0))
        
        # 超额获利（负数归零）
        merged_df["价格差"] = merged_df["现货价"] - merged_df["合约价"]
        merged_df["超额获利(元)"] = merged_df["电量差额"] * merged_df["价格差"]
        merged_df["超额获利(元)"] = merged_df["超额获利(元)"].apply(lambda x: max(x, 0))  # 负数归零
        
        # 整理结果
        result_df = merged_df[["时段", "实发量(MWh)", "修正后实发量", "持仓量(MWh)", "合约0.9倍", "合约1.1倍", "电量差额", "现货价", "合约价", "价格差", "超额获利(元)"]].round(2)
        
        # 总计行
        total_row = {
            "时段": "总计",
            "实发量(MWh)": result_df["实发量(MWh)"].sum(),
            "修正后实发量": result_df["修正后实发量"].sum(),
            "持仓量(MWh)": result_df["持仓量(MWh)"].sum(),
            "合约0.9倍": result_df["合约0.9倍"].sum(),
            "合约1.1倍": result_df["合约1.1倍"].sum(),
            "电量差额": result_df["电量差额"].sum(),
            "现货价": "",
            "合约价": "",
            "价格差": "",
            "超额获利(元)": result_df["超额获利(元)"].sum()
        }
        result_df = pd.concat([result_df, pd.DataFrame([total_row])], ignore_index=True)
        return result_df

# -------------------------- 6. 页面布局（极简版） --------------------------
st.title("📈 超额获利计算工具（纯时段匹配）")

# 侧边栏配置
st.sidebar.markdown("### ⚙️ 基础配置")
station_type = st.sidebar.radio("场站类型", ["风电", "光伏"], key="station_type")
st.session_state.target_month = st.sidebar.text_input("数据月份", value="2025-12", key="month")

# ====================== 1. 上传实发数据 ======================
st.markdown("### 1️⃣ 上传实发数据")
gen_file = st.file_uploader("上传实发数据Excel", type=["xlsx", "xls", "xlsm"], key="gen_file")
if st.button("处理实发数据", key="btn_gen"):
    if gen_file:
        gen_df = DataProcessor.extract_generated_data(gen_file, st.session_state.module_config["generated"], station_type)
        st.session_state.gen_data["24h"] = gen_df
        st.dataframe(gen_df)

# ====================== 2. 上传分时段持仓数据 ======================
st.markdown("### 2️⃣ 上传分时段持仓数据")
hold_file = st.file_uploader("上传分时段持仓Excel（24行，列1=时段，列2=持仓量）", type=["xlsx", "xls", "xlsm"], key="hold_file")
if st.button("处理持仓数据", key="btn_hold"):
    if hold_file:
        hold_df = DataProcessor.extract_hold_data(hold_file, st.session_state.module_config["hold"])
        st.session_state.hold_data_df = hold_df
        st.session_state.hold_data = dict(zip(hold_df["时段"], hold_df["持仓量(MWh)"]))
        st.dataframe(hold_df)

# ====================== 3. 上传电价数据 ======================
st.markdown("### 3️⃣ 上传电价数据")
price_file = st.file_uploader("上传电价Excel（24行，列1=时段，列2-5=风电/光伏现货/合约价）", type=["xlsx", "xls", "xlsm"], key="price_file")
if st.button("处理电价数据", key="btn_price"):
    if price_file:
        price_df = DataProcessor.extract_price_data(price_file, st.session_state.module_config["price"])
        st.session_state.price_data["24h"] = price_df
        st.dataframe(price_df)

# ====================== 4. 计算超额获利（纯时段匹配） ======================
st.markdown("### 4️⃣ 计算超额获利")
if st.button("🔍 立即计算", type="primary", key="btn_calc"):
    gen_df = st.session_state.gen_data["24h"]
    hold_df = st.session_state.hold_data_df
    price_df = st.session_state.price_data["24h"]
    
    if not gen_df.empty and not hold_df.empty and not price_df.empty:
        result_df = DataProcessor.calculate_profit(gen_df, hold_df, price_df, station_type)
        st.session_state.price_data["excess_profit"] = result_df
        
        st.success("✅ 计算完成！")
        st.dataframe(result_df, use_container_width=True)
        
        # 显示总获利
        total_profit = result_df.iloc[-1]["超额获利(元)"]
        st.metric(f"💰 {st.session_state.target_month} 总超额获利（仅正数）", value=f"{round(total_profit, 2)} 元")
        
        # 下载+绘图
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("💾 下载计算结果", to_excel(result_df), f"超额获利计算结果_{st.session_state.target_month}.xlsx", key="download_result")
        with col2:
            plot_df = result_df[result_df["时段"] != "总计"]
            fig = px.bar(plot_df, x="时段", y="超额获利(元)", title="各时段超额获利", width=500, key="profit_chart")
            st.plotly_chart(fig)
    else:
        st.error("❌ 请先上传并处理实发、持仓、电价数据！")

# -------------------------- 7. 配置项（可选） --------------------------
with st.expander("🔧 高级配置（默认值适配你的场景）"):
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("#### 实发数据配置")
        st.session_state.module_config["generated"]["time_col"] = st.number_input("时间列索引", 0, value=4, key="gen_time_col")
        st.session_state.module_config["generated"]["power_col"] = st.number_input("功率列索引", 0, value=9, key="gen_power_col")
        st.session_state.module_config["generated"]["skip_rows"] = st.number_input("实发数据跳过行数", 0, value=1, key="gen_skip_rows")  # 唯一label+key
    with col2:
        st.markdown("#### 持仓数据配置")
        st.session_state.module_config["hold"]["hour_col"] = st.number_input("时段列索引", 0, value=0, key="hold_hour_col")
        st.session_state.module_config["hold"]["hold_col"] = st.number_input("持仓列索引", 0, value=1, key="hold_hold_col")
        st.session_state.module_config["hold"]["skip_rows"] = st.number_input("持仓数据跳过行数", 0, value=1, key="hold_skip_rows")  # 唯一label+key
    with col3:
        st.markdown("#### 电价数据配置")
        st.session_state.module_config["price"]["wind_spot_col"] = st.number_input("风电现货列", 0, value=1, key="price_wind_spot")
        st.session_state.module_config["price"]["wind_contract_col"] = st.number_input("风电合约列", 0, value=2, key="price_wind_contract")
        st.session_state.module_config["price"]["skip_rows"] = st.number_input("电价数据跳过行数", 0, value=1, key="price_skip_rows")  # 唯一label+key
