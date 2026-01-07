import streamlit as st
import pandas as pd
import re
import uuid
from io import BytesIO
import plotly.express as px
import numpy as np

# -------------------------- 1. 页面基础配置 --------------------------
st.set_page_config(
    page_title="超额获利计算工具（正确提取+时段匹配）",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------- 2. 全局常量 --------------------------
STANDARD_HOURS = [f"{i:02d}:00" for i in range(24)]  # 标准24时段
# 修正系数：风电0.7，光伏0.8
CORRECTION_FACTOR = {"风电": 0.7, "光伏": 0.8}

# -------------------------- 3. 核心工具函数 --------------------------
def standardize_column_name(col):
    """标准化列名，避免重复"""
    col_str = str(col).strip() if col is not None else f"列_{uuid.uuid4().hex[:6]}"
    col_str = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9_]', '_', col_str).lower()
    return col_str if col_str else f"列_{uuid.uuid4().hex[:6]}"

def force_unique_columns(df):
    """强制列名唯一"""
    df.columns = [standardize_column_name(col) for col in df.columns]
    # 识别时段列并改名
    time_cols = [col for col in df.columns if "时段" in col or "时间" in col or "hour" in col]
    if time_cols:
        df.rename(columns={time_cols[0]: "时段"}, inplace=True)
    return df

def to_excel(df, sheet_name="计算结果"):
    """导出Excel"""
    if df.empty:
        st.warning("⚠️ 数据为空，无法导出")
        return BytesIO()
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

def standardize_hour(hour_str):
    """标准化时段格式为00:00"""
    try:
        hour_str = str(hour_str).strip().replace("时", "").replace("点", "").replace("：", ":")
        if ":" in hour_str:
            h = int(hour_str.split(":")[0])
        else:
            h = int(hour_str)
        return f"{h:02d}:00"
    except:
        return None

# -------------------------- 4. 会话状态初始化 --------------------------
if "gen_24h_df" not in st.session_state:
    st.session_state.gen_24h_df = pd.DataFrame()  # 实发24时段数据
if "hold_24h_df" not in st.session_state:
    st.session_state.hold_24h_df = pd.DataFrame()  # 持仓24时段数据
if "price_24h_df" not in st.session_state:
    st.session_state.price_24h_df = pd.DataFrame()  # 电价24时段数据
if "result_df" not in st.session_state:
    st.session_state.result_df = pd.DataFrame()     # 计算结果
if "config" not in st.session_state:
    # 配置项（默认值适配常见Excel结构，你可根据实际调整）
    st.session_state.config = {
        # 实发数据：time_col=时间列索引，power_col=功率列索引，skip_rows=跳过行数
        "gen": {"time_col": 0, "power_col": 1, "skip_rows": 0},
        # 持仓数据：hour_col=时段列索引，hold_col=持仓列索引，skip_rows=跳过行数
        "hold": {"hour_col": 0, "hold_col": 1, "skip_rows": 0},
        # 电价数据：hour_col=时段列索引，spot_col=现货列索引，contract_col=合约列索引，skip_rows=跳过行数
        "price": {"hour_col": 0, "spot_col": 1, "contract_col": 2, "skip_rows": 0},
        # 场站类型
        "station_type": "风电"
    }

# -------------------------- 5. 数据提取函数（恢复完整逻辑） --------------------------
def extract_generated_data(file, config):
    """提取实发数据并生成24时段数据（恢复完整逻辑）"""
    try:
        # 读取Excel（兼容xlsx/xlsm/xls）
        file_suffix = file.name.split(".")[-1].lower()
        engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
        df = pd.read_excel(
            BytesIO(file.getvalue()),
            header=None,
            usecols=[config["time_col"], config["power_col"]],
            skiprows=config["skip_rows"],
            engine=engine
        )
        df.columns = ["时间", "功率(kW)"]
        
        # 数据清洗
        df["功率(kW)"] = pd.to_numeric(df["功率(kW)"], errors="coerce").fillna(0)
        df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
        df = df.dropna(subset=["时间"]).sort_values("时间")
        
        # 生成24时段实发量（核心：按小时分组计算）
        df["时段"] = df["时间"].dt.hour.apply(lambda x: f"{x:02d}:00")
        # 计算时间间隔（小时）
        df["时间差(h)"] = df["时间"].diff().dt.total_seconds() / 3600
        avg_interval = df["时间差(h)"].mean() if not df["时间差(h)"].empty else 1/4
        # 按时段求和并转换为MWh（1MWh=1000kWh）
        gen_hourly = df.groupby("时段")["功率(kW)"].sum() * avg_interval / 1000
        
        # 补全24时段（确保每个时段都有数据）
        gen_24h_df = pd.DataFrame({"时段": STANDARD_HOURS})
        gen_24h_df["实发量(MWh)"] = gen_24h_df["时段"].map(gen_hourly).fillna(0)
        
        st.success(f"✅ 实发数据提取成功！共{len(gen_24h_df)}个时段")
        st.dataframe(gen_24h_df, use_container_width=True)
        return gen_24h_df
    except Exception as e:
        st.error(f"❌ 实发数据提取失败：{str(e)}")
        return pd.DataFrame()

def extract_hold_data(file, config):
    """提取分时段持仓数据（恢复完整逻辑）"""
    try:
        file_suffix = file.name.split(".")[-1].lower()
        engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
        df = pd.read_excel(
            BytesIO(file.getvalue()),
            header=None,
            usecols=[config["hour_col"], config["hold_col"]],
            skiprows=config["skip_rows"],
            engine=engine,
            nrows=24  # 只读取前24行（对应24时段）
        )
        df.columns = ["时段", "持仓量(MWh)"]
        
        # 数据清洗+标准化时段
        df["时段"] = df["时段"].apply(standardize_hour)
        df["持仓量(MWh)"] = pd.to_numeric(df["持仓量(MWh)"], errors="coerce").fillna(0)
        # 补全24时段
        hold_24h_df = pd.DataFrame({"时段": STANDARD_HOURS})
        hold_24h_df["持仓量(MWh)"] = hold_24h_df["时段"].map(dict(zip(df["时段"], df["持仓量(MWh)"]))).fillna(0)
        
        st.success(f"✅ 持仓数据提取成功！共{len(hold_24h_df)}个时段")
        st.dataframe(hold_24h_df, use_container_width=True)
        return hold_24h_df
    except Exception as e:
        st.error(f"❌ 持仓数据提取失败：{str(e)}")
        return pd.DataFrame()

def extract_price_data(file, config, station_type):
    """提取分时段电价数据（恢复完整逻辑）"""
    try:
        file_suffix = file.name.split(".")[-1].lower()
        engine = "openpyxl" if file_suffix in ["xlsx", "xlsm"] else "xlrd"
        df = pd.read_excel(
            BytesIO(file.getvalue()),
            header=None,
            usecols=[config["hour_col"], config["spot_col"], config["contract_col"]],
            skiprows=config["skip_rows"],
            engine=engine,
            nrows=24
        )
        df.columns = ["时段", "现货价(元/MWh)", "合约价(元/MWh)"]
        
        # 数据清洗+标准化时段
        df["时段"] = df["时段"].apply(standardize_hour)
        df["现货价(元/MWh)"] = pd.to_numeric(df["现货价(元/MWh)"], errors="coerce").fillna(0)
        df["合约价(元/MWh)"] = pd.to_numeric(df["合约价(元/MWh)"], errors="coerce").fillna(0)
        # 补全24时段
        price_24h_df = pd.DataFrame({"时段": STANDARD_HOURS})
        price_24h_df["现货价(元/MWh)"] = price_24h_df["时段"].map(dict(zip(df["时段"], df["现货价(元/MWh)"]))).fillna(0)
        price_24h_df["合约价(元/MWh)"] = price_24h_df["时段"].map(dict(zip(df["时段"], df["合约价(元/MWh)"]))).fillna(0)
        
        st.success(f"✅ {station_type}电价数据提取成功！共{len(price_24h_df)}个时段")
        st.dataframe(price_24h_df, use_container_width=True)
        return price_24h_df
    except Exception as e:
        st.error(f"❌ 电价数据提取失败：{str(e)}")
        return pd.DataFrame()

# -------------------------- 6. 计算函数（仅时段匹配） --------------------------
def calculate_profit(gen_df, hold_df, price_df, station_type):
    """核心计算：仅按时段匹配，不搞场站名称匹配"""
    if gen_df.empty or hold_df.empty or price_df.empty:
        st.error("❌ 实发/持仓/电价数据不完整！")
        return pd.DataFrame()
    
    # 1. 纯时段匹配合并数据（核心：只按时段列合并）
    merged_df = pd.merge(gen_df, hold_df, on="时段", how="inner")
    merged_df = pd.merge(merged_df, price_df, on="时段", how="inner")
    
    # 2. 计算核心逻辑（保留所有原有规则）
    # 修正后实发量
    merged_df["修正后实发量(MWh)"] = merged_df["实发量(MWh)"] * CORRECTION_FACTOR[station_type]
    # 合约量0.9倍/1.1倍
    merged_df["合约0.9倍(MWh)"] = merged_df["持仓量(MWh)"] * 0.9
    merged_df["合约1.1倍(MWh)"] = merged_df["持仓量(MWh)"] * 1.1
    # 电量差额（超出1.1倍或低于0.9倍的部分）
    merged_df["电量差额(MWh)"] = np.where(
        merged_df["修正后实发量(MWh)"] > merged_df["合约1.1倍(MWh)"],
        merged_df["修正后实发量(MWh)"] - merged_df["合约1.1倍(MWh)"],
        np.where(
            merged_df["修正后实发量(MWh)"] < merged_df["合约0.9倍(MWh)"],
            merged_df["修正后实发量(MWh)"] - merged_df["合约0.9倍(MWh)"],
            0
        )
    )
    # 价格差
    merged_df["价格差(元/MWh)"] = merged_df["现货价(元/MWh)"] - merged_df["合约价(元/MWh)"]
    # 超额获利（负数归零）
    merged_df["超额获利(元)"] = merged_df["电量差额(MWh)"] * merged_df["价格差(元/MWh)"]
    merged_df["超额获利(元)"] = merged_df["超额获利(元)"].apply(lambda x: max(x, 0))
    
    # 3. 整理结果
    result_df = merged_df[[
        "时段", "实发量(MWh)", "修正后实发量(MWh)", "持仓量(MWh)",
        "合约0.9倍(MWh)", "合约1.1倍(MWh)", "电量差额(MWh)",
        "现货价(元/MWh)", "合约价(元/MWh)", "价格差(元/MWh)", "超额获利(元)"
    ]].round(2)
    
    # 4. 添加总计行
    total_row = {
        "时段": "总计",
        "实发量(MWh)": result_df["实发量(MWh)"].sum(),
        "修正后实发量(MWh)": result_df["修正后实发量(MWh)"].sum(),
        "持仓量(MWh)": result_df["持仓量(MWh)"].sum(),
        "合约0.9倍(MWh)": result_df["合约0.9倍(MWh)"].sum(),
        "合约1.1倍(MWh)": result_df["合约1.1倍(MWh)"].sum(),
        "电量差额(MWh)": result_df["电量差额(MWh)"].sum(),
        "现货价(元/MWh)": "",
        "合约价(元/MWh)": "",
        "价格差(元/MWh)": "",
        "超额获利(元)": result_df["超额获利(元)"].sum()
    }
    result_df = pd.concat([result_df, pd.DataFrame([total_row])], ignore_index=True)
    
    st.success("✅ 超额获利计算完成！")
    return result_df

# -------------------------- 7. 页面布局 --------------------------
st.title("📈 超额获利计算工具（正确提取+纯时段匹配）")

# 侧边栏：基础配置（场站类型+索引调整）
with st.sidebar:
    st.markdown("### ⚙️ 基础配置")
    # 选择场站类型
    st.session_state.config["station_type"] = st.radio(
        "场站类型", ["风电", "光伏"], 
        index=0 if st.session_state.config["station_type"] == "风电" else 1,
        key="station_type_radio"
    )
    
    st.markdown("### 📌 列索引配置（关键！按你的Excel调整）")
    # 实发数据配置
    st.markdown("#### 实发数据")
    st.session_state.config["gen"]["time_col"] = st.number_input(
        "时间列索引", 0, value=st.session_state.config["gen"]["time_col"],
        key="gen_time_col"
    )
    st.session_state.config["gen"]["power_col"] = st.number_input(
        "功率列索引", 0, value=st.session_state.config["gen"]["power_col"],
        key="gen_power_col"
    )
    st.session_state.config["gen"]["skip_rows"] = st.number_input(
        "跳过行数", 0, value=st.session_state.config["gen"]["skip_rows"],
        key="gen_skip_rows"
    )
    
    # 持仓数据配置
    st.markdown("#### 持仓数据")
    st.session_state.config["hold"]["hour_col"] = st.number_input(
        "时段列索引", 0, value=st.session_state.config["hold"]["hour_col"],
        key="hold_hour_col"
    )
    st.session_state.config["hold"]["hold_col"] = st.number_input(
        "持仓列索引", 0, value=st.session_state.config["hold"]["hold_col"],
        key="hold_hold_col"
    )
    st.session_state.config["hold"]["skip_rows"] = st.number_input(
        "跳过行数", 0, value=st.session_state.config["hold"]["skip_rows"],
        key="hold_skip_rows"
    )
    
    # 电价数据配置
    st.markdown("#### 电价数据")
    st.session_state.config["price"]["hour_col"] = st.number_input(
        "时段列索引", 0, value=st.session_state.config["price"]["hour_col"],
        key="price_hour_col"
    )
    st.session_state.config["price"]["spot_col"] = st.number_input(
        "现货价列索引", 0, value=st.session_state.config["price"]["spot_col"],
        key="price_spot_col"
    )
    st.session_state.config["price"]["contract_col"] = st.number_input(
        "合约价列索引", 0, value=st.session_state.config["price"]["contract_col"],
        key="price_contract_col"
    )
    st.session_state.config["price"]["skip_rows"] = st.number_input(
        "跳过行数", 0, value=st.session_state.config["price"]["skip_rows"],
        key="price_skip_rows"
    )

# 主页面：分步操作
# 1. 上传实发数据
st.markdown("### 1️⃣ 上传实发数据")
gen_file = st.file_uploader("选择实发数据Excel文件", type=["xlsx", "xls", "xlsm"], key="gen_file")
if st.button("提取实发数据", key="btn_extract_gen"):
    if gen_file:
        st.session_state.gen_24h_df = extract_generated_data(gen_file, st.session_state.config["gen"])
    else:
        st.warning("⚠️ 请先上传实发数据文件！")

st.divider()

# 2. 上传持仓数据
st.markdown("### 2️⃣ 上传分时段持仓数据")
hold_file = st.file_uploader("选择分时段持仓Excel文件", type=["xlsx", "xls", "xlsm"], key="hold_file")
if st.button("提取持仓数据", key="btn_extract_hold"):
    if hold_file:
        st.session_state.hold_24h_df = extract_hold_data(hold_file, st.session_state.config["hold"])
    else:
        st.warning("⚠️ 请先上传持仓数据文件！")

st.divider()

# 3. 上传电价数据
st.markdown("### 3️⃣ 上传分时段电价数据")
price_file = st.file_uploader(
    f"选择{st.session_state.config['station_type']}电价Excel文件", 
    type=["xlsx", "xls", "xlsm"], 
    key="price_file"
)
if st.button("提取电价数据", key="btn_extract_price"):
    if price_file:
        st.session_state.price_24h_df = extract_price_data(
            price_file, 
            st.session_state.config["price"], 
            st.session_state.config["station_type"]
        )
    else:
        st.warning("⚠️ 请先上传电价数据文件！")

st.divider()

# 4. 计算超额获利（仅时段匹配）
st.markdown("### 4️⃣ 计算超额获利（纯时段匹配）")
if st.button("🔍 立即计算", type="primary", key="btn_calc"):
    # 调用计算函数（仅时段匹配）
    st.session_state.result_df = calculate_profit(
        st.session_state.gen_24h_df,
        st.session_state.hold_24h_df,
        st.session_state.price_24h_df,
        st.session_state.config["station_type"]
    )
    
    # 显示结果
    if not st.session_state.result_df.empty:
        st.dataframe(st.session_state.result_df, use_container_width=True)
        
        # 显示总获利
        total_profit = st.session_state.result_df.iloc[-1]["超额获利(元)"]
        st.metric("💰 总超额获利（仅正数）", value=f"{round(total_profit, 2)} 元")
        
        # 下载+绘图
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "💾 下载计算结果",
                to_excel(st.session_state.result_df),
                f"超额获利计算结果_{st.session_state.config['station_type']}.xlsx",
                key="download_btn"
            )
        with col2:
            # 绘图（排除总计行）
            plot_df = st.session_state.result_df[st.session_state.result_df["时段"] != "总计"]
            fig = px.bar(
                plot_df,
                x="时段",
                y="超额获利(元)",
                title=f"{st.session_state.config['station_type']}各时段超额获利",
                width=500
            )
            st.plotly_chart(fig, key="profit_chart")
