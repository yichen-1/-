import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime

# -------------------------- 初始化配置（无缩进错误） --------------------------
st.set_page_config(
    page_title="新能源场站年度方案设计系统",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 初始化Session State（仅一次，顶格）
if "initialized" not in st.session_state:
    st.session_state.initialized = True
    st.session_state.site_data = {}
    st.session_state.current_region = "总部"
    st.session_state.current_province = ""
    st.session_state.current_month = 1
    st.session_state.current_site = ""
    st.session_state.trade_power_data = None
    st.session_state.adjusted_trade_power = None
    st.session_state.total_trade_power = 0.0
    st.session_state.mechanism_mode = "小时数"
    st.session_state.guaranteed_mode = "小时数"
    st.session_state.manual_market_hours = 0.0
    st.session_state.auto_calculate = True
    st.session_state.calculated = False
    st.session_state.market_hours = 0.0
    st.session_state.gen_hours = 0.0
    
    # 初始化24小时数据（顶格）
    hours = list(range(1, 25))
    st.session_state.current_24h_data = pd.DataFrame({
        "时段": hours,
        "平均发电量(MWh)": [0.0]*24,
        "当月各时段累计发电量(MWh)": [0.0]*24,
        "现货价格(元/MWh)": [0.0]*24,
        "中长期价格(元/MWh)": [0.0]*24
    })

# -------------------------- 核心工具函数（顶格） --------------------------
def calculate_core_params(installed_capacity, power_limit_rate, mechanism_value, mechanism_mode, 
                         guaranteed_value, guaranteed_mode):
    # 函数内缩进4个空格
    total_generation = st.session_state.current_24h_data["当月各时段累计发电量(MWh)"].sum()
    gen_hours = round(total_generation / installed_capacity, 2) if installed_capacity > 0 else 0.0
    
    if gen_hours <= 0:
        market_hours = 0.0
    else:
        available_hours = gen_hours * (1 - power_limit_rate / 100)
        
        if mechanism_mode == "小时数":
            available_hours -= mechanism_value
        else:
            available_hours -= gen_hours * (mechanism_value / 100)
        
        if guaranteed_mode == "小时数":
            available_hours -= guaranteed_value
        else:
            available_hours -= gen_hours * (guaranteed_value / 100)
        
        market_hours = max(round(available_hours, 2), 0.0)
    
    return gen_hours, market_hours

def calculate_trade_power_distribution(avg_generation_24h, market_hours, installed_capacity):
    # 函数内缩进4个空格
    total_trade_power = market_hours * installed_capacity
    total_avg_generation = sum(avg_generation_24h)
    
    if installed_capacity <= 0 or market_hours <= 0 or total_avg_generation <= 0:
        raise ValueError("计算条件不满足：装机容量/市场化小时数/平均发电量总和必须大于0")
    
    trade_power_data = []
    for hour, avg_gen in enumerate(avg_generation_24h, 1):
        proportion = avg_gen / total_avg_generation
        trade_power = total_trade_power * proportion
        trade_power_data.append({
            "时段": hour,
            "平均发电量(MWh)": avg_gen,
            "时段比重(%)": round(proportion * 100, 4),
            "市场化交易电量(MWh)": round(trade_power, 2)
        })
    
    return pd.DataFrame(trade_power_data), round(total_trade_power, 2)

def adjust_trade_power_by_price(trade_power_df, spot_price_24h, total_trade_power):
    # 函数内缩进4个空格
    spot_price_24h = [max(p, 0.01) for p in spot_price_24h]
    total_price = sum(spot_price_24h)
    price_weights = [p / total_price for p in spot_price_24h]
    
    adjusted_data = trade_power_df.copy()
    for idx, weight in enumerate(price_weights):
        adjusted_data.loc[idx, "市场化交易电量(MWh)"] = round(total_trade_power * weight, 2)
        adjusted_data.loc[idx, "时段比重(%)"] = round(weight * 100, 4)
    
    sum_adjusted = adjusted_data["市场化交易电量(MWh)"].sum()
    diff = total_trade_power - sum_adjusted
    if abs(diff) > 0.01:
        max_price_idx = spot_price_24h.index(max(spot_price_24h))
        adjusted_data.loc[max_price_idx, "市场化交易电量(MWh)"] += round(diff, 2)
    
    return adjusted_data

def calibrate_trade_power(adjusted_df, total_trade_power):
    # 函数内缩进4个空格
    calibrated_df = adjusted_df.copy()
    current_sum = calibrated_df["市场化交易电量(MWh)"].sum()
    diff = total_trade_power - current_sum
    
    if abs(diff) <= 0.01:
        return calibrated_df
    
    positive_mask = calibrated_df["市场化交易电量(MWh)"] > 0
    positive_qty = calibrated_df.loc[positive_mask, "市场化交易电量(MWh)"]
    total_positive = positive_qty.sum()
    
    if total_positive <= 0:
        avg_qty = total_trade_power / 24
        calibrated_df["市场化交易电量(MWh)"] = round(avg_qty, 2)
        calibrated_df["时段比重(%)"] = round((avg_qty / total_trade_power) * 100, 4)
    else:
        for idx in calibrated_df.index:
            if positive_mask[idx]:
                ratio = calibrated_df.loc[idx, "市场化交易电量(MWh)"] / total_positive
                calibrated_df.loc[idx, "市场化交易电量(MWh)"] += round(diff * ratio, 2)
                calibrated_df.loc[idx, "市场化交易电量(MWh)"] = max(0.0, calibrated_df.loc[idx, "市场化交易电量(MWh)"])
                calibrated_df.loc[idx, "时段比重(%)"] = round((calibrated_df.loc[idx, "市场化交易电量(MWh)"] / total_trade_power) * 100, 4)
    
    final_diff = total_trade_power - calibrated_df["市场化交易电量(MWh)"].sum()
    if abs(final_diff) > 0.01:
        non_zero_idx = calibrated_df[calibrated_df["市场化交易电量(MWh)"] > 0].index[0]
        calibrated_df.loc[non_zero_idx, "市场化交易电量(MWh)"] += round(final_diff, 2)
    
    return calibrated_df

def save_data_to_file(province, month, site_name, data, trade_power_data=None, total_trade_power=0.0):
    # 函数内缩进4个空格
    save_dir = f"./新能源场站数据/{province}/{site_name}"
    os.makedirs(save_dir, exist_ok=True)
    
    if trade_power_data is not None:
        data = pd.merge(
            data, 
            trade_power_data[["时段", "时段比重(%)", "市场化交易电量(MWh)"]],
            on="时段", 
            how="left"
        )
    
    filepath = os.path.join(save_dir, f"{month}月数据.csv")
    data.to_csv(filepath, index=False, encoding="utf-8-sig")
    return filepath

def load_data_from_file(province, month, site_name):
    # 函数内缩进4个空格
    filepath = f"./新能源场站数据/{province}/{site_name}/{month}月数据.csv"
    return pd.read_csv(filepath, encoding="utf-8-sig") if os.path.exists(filepath) else None

# -------------------------- 区域-省份字典（顶格） --------------------------
REGIONS = {
    "总部": ["北京"],
    "华北": ["首都", "河北", "冀北", "山东", "山西", "天津"],
    "华东": ["安徽", "福建", "江苏", "上海", "浙江"],
    "华中": ["湖北", "河南", "湖南", "江西"],
    "东北": ["吉林", "黑龙江", "辽宁", "蒙东"],
    "西北": ["甘肃", "宁夏", "青海", "陕西", "新疆"],
    "西南": ["重庆", "四川", "西藏"],
    "南方": ["广东", "广西", "云南", "海南", "贵州"],
    "内蒙古电网": ["蒙西"]
}

MONTHS = list(range(1, 13))

# -------------------------- 侧边栏配置（顶格） --------------------------
st.sidebar.header("⚙️ 基础信息配置")

# 1. 基础信息（顶格）
selected_region = st.sidebar.selectbox(
    "选择区域", list(REGIONS.keys()),
    index=list(REGIONS.keys()).index(st.session_state.current_region),
    key="sidebar_region_select"
)
st.session_state.current_region = selected_region

current_province_list = REGIONS[st.session_state.current_region]
if not st.session_state.current_province or st.session_state.current_province not in current_province_list:
    st.session_state.current_province = current_province_list[0]

selected_province = st.sidebar.selectbox(
    "选择省份/地区", current_province_list,
    index=current_province_list.index(st.session_state.current_province),
    key="sidebar_province_select"
)
st.session_state.current_province = selected_province

selected_month = st.sidebar.selectbox(
    "选择月份", MONTHS,
    index=st.session_state.current_month - 1,
    key="sidebar_month_select"
)
st.session_state.current_month = selected_month

site_name = st.sidebar.text_input(
    "场站名称", value=st.session_state.current_site,
    key="sidebar_site_name", placeholder="如：张家口风电场"
)
st.session_state.current_site = site_name

installed_capacity = st.sidebar.number_input(
    "装机容量(MW)", min_value=0.0, value=0.0, step=0.1,
    key="sidebar_installed_capacity", help="场站总装机容量，单位：兆瓦"
)

# 2. 电量参数配置（顶格）
st.sidebar.subheader("⚡ 电量参数配置")

# 机制电量（顶格）
st.sidebar.write("#### 机制电量")
col_mech1, col_mech2 = st.sidebar.columns([2, 1])
# 无with嵌套，直接操作列（规避缩进）
col_mech1_select = col_mech1.selectbox(
    "输入模式", ["小时数", "比例(%)"],
    index=0 if st.session_state.mechanism_mode == "小时数" else 1,
    key="sidebar_mechanism_mode"
)
st.session_state.mechanism_mode = col_mech1_select

mech_min = 0.0
mech_max = 100.0 if st.session_state.mechanism_mode == "比例(%)" else 1000000.0
col_mech2_input = col_mech2.number_input(
    "数值", min_value=mech_min, max_value=mech_max, value=0.0, step=0.1,
    key="sidebar_mechanism_value", help=f"机制电量{st.session_state.mechanism_mode}"
)
mechanism_value = col_mech2_input

# 保障性电量（顶格）
st.sidebar.write("#### 保障性电量")
col_gua1, col_gua2 = st.sidebar.columns([2, 1])
col_gua1_select = col_gua1.selectbox(
    "输入模式", ["小时数", "比例(%)"],
    index=0 if st.session_state.guaranteed_mode == "小时数" else 1,
    key="sidebar_guaranteed_mode"
)
st.session_state.guaranteed_mode = col_gua1_select

gua_min = 0.0
gua_max = 100.0 if st.session_state.guaranteed_mode == "比例(%)" else 1000000.0
col_gua2_input = col_gua2.number_input(
    "数值", min_value=gua_min, max_value=gua_max, value=0.0, step=0.1,
    key="sidebar_guaranteed_value", help=f"保障性电量{st.session_state.guaranteed_mode}"
)
guaranteed_value = col_gua2_input

# 限电率（顶格）
power_limit_rate = st.sidebar.number_input(
    "限电率(%)", min_value=0.0, max_value=100.0, value=0.0, step=0.1,
    key="sidebar_power_limit_rate", help="场站当月限电比例，0-100%"
)

# 市场化交易小时数（顶格）
st.sidebar.write("#### 市场化交易小时数")
auto_calculate = st.sidebar.toggle(
    "自动计算", value=st.session_state.auto_calculate,
    key="sidebar_auto_calculate", help="勾选：按公式自动计算；取消：手动输入"
)
st.session_state.auto_calculate = auto_calculate

if st.session_state.auto_calculate:
    st.sidebar.number_input(
        "计算结果（点击生成方案后更新）",
        value=st.session_state.market_hours,
        step=0.1,
        disabled=True,
        key="sidebar_market_hours_auto",
        min_value=0.0,
        max_value=1000000.0
    )
else:
    manual_market_hours = st.sidebar.number_input(
        "手动输入", min_value=0.0, max_value=1000000.0,
        value=st.session_state.manual_market_hours, step=0.1,
        key="sidebar_market_hours_manual"
    )
    st.session_state.manual_market_hours = manual_market_hours

# -------------------------- 主页面内容（顶格） --------------------------
st.title("⚡ 新能源场站年度方案设计系统")
st.subheader(
    f"当前配置：{st.session_state.current_region} | {st.session_state.current_province} | "
    f"{st.session_state.current_month}月 | {st.session_state.current_site}"
)

# 数据操作按钮（顶格，无with嵌套）
btn_col1, btn_col2, btn_col3, btn_col4, btn_col5 = st.columns(5)

# 按钮1：初始化
if btn_col1.button("📋 初始化24时段数据模板", use_container_width=True, key="main_init_btn"):
    hours = list(range(1, 25))
    st.session_state.current_24h_data = pd.DataFrame({
        "时段": hours,
        "平均发电量(MWh)": [0.0]*24,
        "当月各时段累计发电量(MWh)": [0.0]*24,
        "现货价格(元/MWh)": [0.0]*24,
        "中长期价格(元/MWh)": [0.0]*24
    })
    st.session_state.trade_power_data = None
    st.session_state.adjusted_trade_power = None
    st.session_state.total_trade_power = 0.0
    st.session_state.calculated = False
    st.success("✅ 已初始化24时段数据模板！")

# 按钮2：导入数据
import_btn = btn_col2.file_uploader(
    "📤 导入数据(CSV/Excel)", type=["csv", "xlsx"], key="main_import_btn"
)
if import_btn is not None:
    try:
        df = pd.read_csv(import_btn) if import_btn.name.endswith(".csv") else pd.read_excel(import_btn)
        required_cols = ["时段", "平均发电量(MWh)", "当月各时段累计发电量(MWh)", "现货价格(元/MWh)", "中长期价格(元/MWh)"]
        if all(col in df.columns for col in required_cols) and len(df) == 24:
            st.session_state.current_24h_data = df
            st.session_state.trade_power_data = None
            st.session_state.adjusted_trade_power = None
            st.session_state.total_trade_power = 0.0
            st.session_state.calculated = False
            st.success("✅ 数据导入成功！")
        else:
            st.error("❌ 导入文件格式错误，请检查列名和24时段数据")
    except Exception as e:
        st.error(f"❌ 导入失败：{str(e)}")

# 按钮3：保存数据
if btn_col3.button("💾 保存当前数据", use_container_width=True, key="main_save_btn"):
    if not st.session_state.current_province or not st.session_state.current_site or installed_capacity <= 0:
        st.warning("⚠️ 请完善省份、场站名称、装机容量信息")
    else:
        final_data = st.session_state.current_24h_data.copy()
        trade_power_data = st.session_state.adjusted_trade_power if st.session_state.adjusted_trade_power is not None else st.session_state.trade_power_data
        
        final_data["区域"] = st.session_state.current_region
        final_data["省份/地区"] = st.session_state.current_province
        final_data["月份"] = st.session_state.current_month
        final_data["场站名称"] = st.session_state.current_site
        final_data["装机容量(MW)"] = installed_capacity
        final_data["预估发电小时数"] = st.session_state.gen_hours
        final_data["机制电量模式"] = st.session_state.mechanism_mode
        final_data["机制电量值"] = mechanism_value
        final_data["保障性电量模式"] = st.session_state.guaranteed_mode
        final_data["保障性电量值"] = guaranteed_value
        final_data["限电率(%)"] = power_limit_rate
        final_data["市场化交易小时数"] = st.session_state.market_hours if st.session_state.auto_calculate else st.session_state.manual_market_hours
        final_data["总交易电量(MWh)"] = st.session_state.total_trade_power
        final_data["保存时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            filepath = save_data_to_file(
                st.session_state.current_province,
                st.session_state.current_month,
                st.session_state.current_site,
                final_data,
                trade_power_data,
                st.session_state.total_trade_power
            )
            st.success(f"✅ 数据保存成功！文件路径：{filepath}")
        except Exception as e:
            st.error(f"❌ 保存失败：{str(e)}")

# 按钮4：加载历史数据
if btn_col4.button("📥 加载历史数据", use_container_width=True, key="main_load_btn"):
    if not st.session_state.current_province or not st.session_state.current_month or not st.session_state.current_site:
        st.warning("⚠️ 请先填写省份、月份和场站名称")
    else:
        loaded_data = load_data_from_file(
            st.session_state.current_province,
            st.session_state.current_month,
            st.session_state.current_site
        )
        if loaded_data is not None:
            st.session_state.current_24h_data = loaded_data
            if "市场化交易电量(MWh)" in loaded_data.columns:
                trade_power_cols = ["时段", "平均发电量(MWh)", "时段比重(%)", "市场化交易电量(MWh)"]
                st.session_state.trade_power_data = loaded_data[trade_power_cols].copy()
                st.session_state.adjusted_trade_power = st.session_state.trade_power_data.copy()
                if "总交易电量(MWh)" in loaded_data.columns:
                    st.session_state.total_trade_power = loaded_data["总交易电量(MWh)"].iloc[0]
                else:
                    st.session_state.total_trade_power = loaded_data["市场化交易电量(MWh)"].sum()
            st.success("✅ 历史数据加载成功！")
        else:
            st.warning("⚠️ 未找到该场站的历史数据")

# 按钮5：生成初始方案
if btn_col5.button("📝 生成初始交易方案", use_container_width=True, type="primary", key="main_generate_btn"):
    with st.spinner("🔢 正在计算交易方案，请稍候..."):
        try:
            if st.session_state.auto_calculate:
                gen_hours, market_hours = calculate_core_params(
                    installed_capacity, power_limit_rate,
                    mechanism_value, st.session_state.mechanism_mode,
                    guaranteed_value, st.session_state.guaranteed_mode
                )
                st.session_state.gen_hours = gen_hours
                st.session_state.market_hours = market_hours
            else:
                market_hours = st.session_state.manual_market_hours
                total_generation = st.session_state.current_24h_data["当月各时段累计发电量(MWh)"].sum()
                st.session_state.gen_hours = round(total_generation / installed_capacity, 2) if installed_capacity > 0 else 0.0
            
            avg_generation_list = st.session_state.current_24h_data["平均发电量(MWh)"].tolist()
            trade_power_df, total_trade_power = calculate_trade_power_distribution(
                avg_generation_list, market_hours, installed_capacity
            )
            
            st.session_state.trade_power_data = trade_power_df
            st.session_state.total_trade_power = total_trade_power
            st.session_state.adjusted_trade_power = None
            st.session_state.calculated = True
            
            st.success(f"✅ 初始交易方案生成成功！总交易电量：{total_trade_power:.2f} MWh")
            
        except ValueError as e:
            st.error(f"❌ 生成方案失败：{str(e)}")
        except Exception as e:
            st.error(f"❌ 生成方案失败：未知错误 - {str(e)}")

# 24时段数据编辑（顶格）
st.divider()
st.header("📊 24时段基础数据编辑")
edited_df = st.data_editor(
    st.session_state.current_24h_data,
    column_config={
        "时段": st.column_config.NumberColumn("时段", disabled=True),
        "平均发电量(MWh)": st.column_config.NumberColumn("平均发电量(MWh)", min_value=0.0, step=0.1),
        "当月各时段累计发电量(MWh)": st.column_config.NumberColumn("当月各时段累计发电量(MWh)", min_value=0.0, step=0.1),
        "现货价格(元/MWh)": st.column_config.NumberColumn("现货价格(元/MWh)", min_value=0.0, step=0.1),
        "中长期价格(元/MWh)": st.column_config.NumberColumn("中长期价格(元/MWh)", min_value=0.0, step=0.1)
    },
    use_container_width=True,
    num_rows="fixed",
    key="main_data_editor"
)
st.session_state.current_24h_data = edited_df

# 关键指标展示（顶格）
st.divider()
st.header("📈 关键指标（生成方案后更新）")

ind_col1, ind_col2, ind_col3, ind_col4, ind_col5, ind_col6 = st.columns(6)
ind_col1.metric("当月总发电量(MWh)", f"{edited_df['当月各时段累计发电量(MWh)'].sum():.2f}")
ind_col2.metric("预估发电小时数", f"{st.session_state.gen_hours:.2f}")
ind_col3.metric("装机容量(MW)", f"{installed_capacity:.1f}")
ind_col4.metric("限电率(%)", f"{power_limit_rate:.1f}")
ind_col5.metric(f"机制电量({st.session_state.mechanism_mode})", f"{mechanism_value:.2f}")
display_market_hours = st.session_state.market_hours if st.session_state.auto_calculate else st.session_state.manual_market_hours
ind_col6.metric("市场化交易小时数", f"{display_market_hours:.2f}")

# 交易电量调整模块（顶格）
if st.session_state.calculated and st.session_state.trade_power_data is not None:
    st.divider()
    st.header("💰 交易电量智能调整（总电量锁定）")
    st.info(f"🔒 锁定总交易电量：{st.session_state.total_trade_power:.2f} MWh")
    
    if st.session_state.adjusted_trade_power is None:
        st.session_state.adjusted_trade_power = st.session_state.trade_power_data.copy()
    
    # 调整功能按钮（顶格，无with嵌套）
    adj_btn_col1, adj_btn_col2, adj_btn_col3 = st.columns(3)
    
    # 按钮1：按电价优化
    if adj_btn_col1.button("📈 按现货电价自动优化", use_container_width=True, key="adjust_by_price_btn"):
        spot_price_list = st.session_state.current_24h_data["现货价格(元/MWh)"].tolist()
        if sum(spot_price_list) <= 0:
            st.warning("⚠️ 现货电价数据全为0，无法按电价优化！")
        else:
            adjusted_df = adjust_trade_power_by_price(
                st.session_state.trade_power_data,
                spot_price_list,
                st.session_state.total_trade_power
            )
            st.session_state.adjusted_trade_power = adjusted_df
            st.success("✅ 已按现货电价优化！")
    
    # 按钮2：重置
    if adj_btn_col2.button("🔄 重置为初始分配", use_container_width=True, key="reset_adjust_btn"):
        st.session_state.adjusted_trade_power = st.session_state.trade_power_data.copy()
        st.success("✅ 已重置为初始分配方案！")
    
    # 按钮3：校准
    if adj_btn_col3.button("🎯 自动校准总和", use_container_width=True, key="calibrate_btn"):
        calibrated_df = calibrate_trade_power(
            st.session_state.adjusted_trade_power,
            st.session_state.total_trade_power
        )
        st.session_state.adjusted_trade_power = calibrated_df
        st.success("✅ 已校准！总和已匹配锁定值！")
    
    # 手动调整表格（顶格）
    st.subheader("✏️ 手动调整各时段交易电量")
    adjust_df = st.data_editor(
        st.session_state.adjusted_trade_power,
        column_config={
            "时段": st.column_config.NumberColumn("时段", disabled=True),
            "平均发电量(MWh)": st.column_config.NumberColumn("平均发电量(MWh)", disabled=True),
            "时段比重(%)": st.column_config.NumberColumn("时段比重(%)", disabled=True, format="%.4f"),
            "市场化交易电量(MWh)": st.column_config.NumberColumn(
                "市场化交易电量(MWh)", 
                min_value=0.0, 
                step=0.1,
                format="%.2f"
            )
        },
        use_container_width=True,
        hide_index=True,
        key="adjust_data_editor"
    )
    st.session_state.adjusted_trade_power = adjust_df

    # -------------------------- 实时状态显示（核心修复，无缩进错误） --------------------------
    # 完全扁平化，无with嵌套，无复杂缩进
    current_sum = adjust_df["市场化交易电量(MWh)"].sum()
    diff = st.session_state.total_trade_power - current_sum
    
    # 直接创建列对象（无嵌套）
    status_col1 = st.columns(3)[0]
    status_col2 = st.columns(3)[1]
    status_col3 = st.columns(3)[2]
    
    # 列1：当前总和（顶格）
    status_col1.metric("当前总和(MWh)", f"{current_sum:.2f}", delta=f"{diff:.2f}")
    
    # 列2：校准状态（单层if/else，仅内部缩进4空格）
    if abs(diff) <= 0.01:
        status_col2.metric("校准状态", "✅ 已匹配", delta="0.00")
    else:
        status_col2.metric("校准状态", "⚠️ 未匹配", delta=f"{diff:.2f}")
    
    # 列3：最高电价时段（单层if/else，仅内部缩进4空格）
    spot_price_list = st.session_state.current_24h_data["现货价格(元/MWh)"].tolist()
    if sum(spot_price_list) <= 0:
        status_col3.metric("最高电价时段", "无有效电价", delta="0.00元/MWh")
    else:
        max_price_hour = spot_price_list.index(max(spot_price_list)) + 1
        max_price = max(spot_price_list)
        status_col3.metric("最高电价时段", f"{max_price_hour}时", delta=f"{max_price:.2f}元/MWh")

    # 对比展示（顶格）
    st.subheader("📊 调整前后对比")
    compare_df = pd.DataFrame({
        "时段": st.session_state.trade_power_data["时段"],
        "初始电量(MWh)": st.session_state.trade_power_data["市场化交易电量(MWh)"],
        "调整后电量(MWh)": st.session_state.adjusted_trade_power["市场化交易电量(MWh)"],
        "差值(MWh)": st.session_state.adjusted_trade_power["市场化交易电量(MWh)"] - st.session_state.trade_power_data["市场化交易电量(MWh)"]
    })
    st.dataframe(compare_df, use_container_width=True, hide_index=True)

    # 可视化对比（顶格，无with嵌套）
    chart_col1 = st.columns(2)[0]
    chart_col2 = st.columns(2)[1]
    
    chart_col1.write("初始分配电量分布")
    chart_col1.bar_chart(
        st.session_state.trade_power_data.set_index("时段")["市场化交易电量(MWh)"],
        use_container_width=True,
        y_label="电量(MWh)"
    )
    
    chart_col2.write("调整后分配电量分布")
    chart_col2.bar_chart(
        st.session_state.adjusted_trade_power.set_index("时段")["市场化交易电量(MWh)"],
        use_container_width=True,
        y_label="电量(MWh)"
    )

# 历史数据查询（顶格）
st.divider()
st.header("🗂️ 历史数据查询")
query_col1, query_col2, query_col3, query_col4 = st.columns(4)

query_region = query_col1.selectbox("查询区域", list(REGIONS.keys()), key="query_region")
query_province_list = REGIONS[query_region]
query_province = query_col2.selectbox("查询省份/地区", query_province_list, key="query_province")
query_month = query_col3.selectbox("查询月份", MONTHS, key="query_month")
query_site = query_col4.text_input("查询场站名称", key="query_site", placeholder="输入场站名称")

if st.button("🔍 查询数据", use_container_width=True, key="query_btn"):
    if not query_province or not query_site:
        st.warning("⚠️ 请填写查询省份/地区和场站名称")
    else:
        query_data = load_data_from_file(query_province, query_month, query_site)
        if query_data is not None:
            st.subheader(f"查询结果：{query_region} | {query_province} | {query_month}月 | {query_site}")
            st.dataframe(query_data, use_container_width=True)
            
            query_total_gen = query_data["当月各时段累计发电量(MWh)"].sum()
            query_installed_cap = query_data["装机容量(MW)"].iloc[0] if "装机容量(MW)" in query_data.columns else 0
            query_gen_hours = round(query_total_gen / query_installed_cap, 2) if query_installed_cap > 0 else 0.0
            
            q_ind_col1, q_ind_col2, q_ind_col3 = st.columns(3)
            q_ind_col1.metric("总发电量(MWh)", f"{query_total_gen:.2f}")
            q_ind_col2.metric("装机容量(MW)", f"{query_installed_cap:.1f}")
            q_ind_col3.metric("预估发电小时数", f"{query_gen_hours:.2f}")
            
            if "市场化交易电量(MWh)" in query_data.columns:
                st.subheader("市场化交易电量分配")
                trade_cols = ["时段", "平均发电量(MWh)", "时段比重(%)", "市场化交易电量(MWh)"]
                st.dataframe(query_data[trade_cols], use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ 未查询到该条件下的历史数据")

# 页脚（顶格）
st.divider()
st.caption("© 2025 新能源场站年度方案设计系统 | 数据自动保存至本地CSV文件")
