import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime

# -------------------------- 初始化配置 --------------------------
st.set_page_config(
    page_title="新能源场站年度方案设计系统",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 初始化会话状态（修复初始值不匹配问题）
if "site_data" not in st.session_state:
    st.session_state.site_data = {}
if "current_region" not in st.session_state:
    st.session_state.current_region = "总部"  # 默认选中总部
if "current_province" not in st.session_state:
    st.session_state.current_province = ""  # 先置空，后续自动匹配
if "current_month" not in st.session_state:
    st.session_state.current_month = 1
if "current_site" not in st.session_state:
    st.session_state.current_site = ""

# 重新定义区域-省份字典（核心调整：内蒙古电网为顶级区域，蒙西在其下）
REGIONS = {
    "总部": ["北京"],
    "华北": ["首都", "河北", "冀北", "山东", "山西", "天津"],
    "华东": ["安徽", "福建", "江苏", "上海", "浙江"],
    "华中": ["湖北", "河南", "湖南", "江西"],
    "东北": ["吉林", "黑龙江", "辽宁", "蒙东"],
    "西北": ["甘肃", "宁夏", "青海", "陕西", "新疆"],
    "西南": ["重庆", "四川", "西藏"],
    "南方": ["广东", "广西", "云南", "海南", "贵州"],
    "内蒙古电网": ["蒙西"]  # 提升为顶级区域，包含蒙西子选项
}

# 月份列表
MONTHS = list(range(1, 13))

# -------------------------- 工具函数 --------------------------
def init_24h_data():
    """初始化24时段数据模板"""
    hours = list(range(1, 25))
    data = {
        "时段": hours,
        "平均发电量(MWh)": [0.0]*24,
        "当月各时段累计发电量(MWh)": [0.0]*24,
        "现货价格(元/MWh)": [0.0]*24,
        "中长期价格(元/MWh)": [0.0]*24
    }
    return pd.DataFrame(data)

def calculate_generation_hours(total_generation, installed_capacity):
    """计算发电小时数"""
    if installed_capacity <= 0:
        return 0.0
    return round(total_generation / installed_capacity, 2)

def save_data_to_file(province, month, site_name, data):
    """保存数据到CSV文件"""
    # 创建保存目录（按省份+场站分层）
    save_dir = f"./新能源场站数据/{province}/{site_name}"
    os.makedirs(save_dir, exist_ok=True)
    
    # 生成文件名
    filename = f"{month}月数据.csv"
    filepath = os.path.join(save_dir, filename)
    
    # 保存数据
    data.to_csv(filepath, index=False, encoding="utf-8-sig")
    return filepath

def load_data_from_file(province, month, site_name):
    """从文件加载数据"""
    filepath = f"./新能源场站数据/{province}/{site_name}/{month}月数据.csv"
    if os.path.exists(filepath):
        return pd.read_csv(filepath, encoding="utf-8-sig")
    return None

# -------------------------- 侧边栏配置 --------------------------
st.sidebar.header("⚙️ 基础信息配置")

# 区域选择（包含内蒙古电网顶级选项）
st.session_state.current_region = st.sidebar.selectbox(
    "选择区域",
    list(REGIONS.keys()),
    index=list(REGIONS.keys()).index(st.session_state.current_region),
    key="region_select"
)

# 获取当前区域的省份/地区列表
current_province_list = REGIONS[st.session_state.current_region]

# 自动匹配初始省份（修复索引错误核心逻辑）
if not st.session_state.current_province or st.session_state.current_province not in current_province_list:
    st.session_state.current_province = current_province_list[0]  # 默认选中第一个

# 省份/地区选择（安全的索引处理）
st.session_state.current_province = st.sidebar.selectbox(
    "选择省份/地区",
    current_province_list,
    index=current_province_list.index(st.session_state.current_province),  # 此时值一定在列表中
    key="province_select"
)

# 月份选择
st.session_state.current_month = st.sidebar.selectbox(
    "选择月份",
    MONTHS,
    index=st.session_state.current_month-1,
    key="month_select"
)

# 场站名称
st.session_state.current_site = st.sidebar.text_input(
    "场站名称",
    value=st.session_state.current_site,
    key="site_name_input",
    placeholder="请输入场站名称（如：张家口风电场）"
)

# 装机容量
installed_capacity = st.sidebar.number_input(
    "装机容量(MW)",
    min_value=0.0,
    value=0.0,
    step=0.1,
    key="installed_capacity",
    help="场站总装机容量，单位：兆瓦"
)

# 其他关键参数
st.sidebar.subheader("⚡ 电量相关参数")
mechanism_hours = st.sidebar.number_input(
    "机制电量小时数",
    min_value=0.0,
    value=0.0,
    step=0.1,
    key="mechanism_hours"
)

guaranteed_hours = st.sidebar.number_input(
    "保障性小时数",
    min_value=0.0,
    value=0.0,
    step=0.1,
    key="guaranteed_hours"
)

power_limit_rate = st.sidebar.number_input(
    "限电率(%)",
    min_value=0.0,
    max_value=100.0,
    value=0.0,
    step=0.1,
    key="power_limit_rate"
)

market_hours = st.sidebar.number_input(
    "市场化交易小时数",
    min_value=0.0,
    value=0.0,
    step=0.1,
    key="market_hours"
)

# -------------------------- 主页面内容 --------------------------
st.title("⚡ 新能源场站年度方案设计系统")
st.subheader(f"当前配置：{st.session_state.current_region} | {st.session_state.current_province} | {st.session_state.current_month}月 | {st.session_state.current_site}")

# 数据操作区域
col1, col2, col3, col4 = st.columns(4)

with col1:
    init_btn = st.button("📋 初始化24时段数据模板", use_container_width=True)
with col2:
    import_btn = st.file_uploader(
        "📤 导入数据(CSV/Excel)",
        type=["csv", "xlsx"],
        key="data_import"
    )
with col3:
    save_btn = st.button("💾 保存当前数据", use_container_width=True)
with col4:
    load_btn = st.button("📥 加载历史数据", use_container_width=True)

# 初始化数据
if init_btn:
    st.session_state.current_24h_data = init_24h_data()
elif "current_24h_data" not in st.session_state:
    st.session_state.current_24h_data = init_24h_data()

# 导入数据处理
if import_btn is not None:
    try:
        if import_btn.name.endswith(".csv"):
            df = pd.read_csv(import_btn, encoding="utf-8-sig")
        else:
            df = pd.read_excel(import_btn)
        
        # 验证数据格式
        required_cols = ["时段", "平均发电量(MWh)", "当月各时段累计发电量(MWh)", "现货价格(元/MWh)", "中长期价格(元/MWh)"]
        if all(col in df.columns for col in required_cols) and len(df) == 24:
            st.session_state.current_24h_data = df
            st.success("✅ 数据导入成功！")
        else:
            st.error("❌ 导入文件格式错误，请检查列名和数据行数（必须包含24时段）")
    except Exception as e:
        st.error(f"❌ 导入失败：{str(e)}")

# 加载历史数据
if load_btn:
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
            st.success("✅ 历史数据加载成功！")
        else:
            st.warning("⚠️ 未找到该场站的历史数据")

# 24时段数据编辑区域
st.divider()
st.header("📊 24时段数据编辑")

# 数据编辑表格
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
    num_rows="fixed"
)

# 更新会话状态中的数据
st.session_state.current_24h_data = edited_df

# -------------------------- 数据计算与展示 --------------------------
st.divider()
st.header("📈 关键指标计算")

# 计算总发电量
total_generation = edited_df["当月各时段累计发电量(MWh)"].sum()
# 计算发电小时数
generation_hours = calculate_generation_hours(total_generation, installed_capacity)

# 展示计算结果
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("当月总发电量(MWh)", f"{total_generation:.2f}")
with col2:
    st.metric("装机容量(MW)", f"{installed_capacity:.1f}")
with col3:
    st.metric("当月发电小时数", f"{generation_hours:.2f}")
with col4:
    st.metric("限电率(%)", f"{power_limit_rate:.1f}")
with col5:
    st.metric("市场化交易小时数", f"{market_hours:.2f}")

# 展示其他参数
st.write("### 补充参数信息")
param_df = pd.DataFrame({
    "参数名称": ["机制电量小时数", "保障性小时数", "限电率", "市场化交易小时数"],
    "数值": [mechanism_hours, guaranteed_hours, f"{power_limit_rate}%", market_hours],
    "说明": [
        "机制电量对应的发电小时数",
        "保障性收购电量对应的小时数",
        "场站当月限电比例",
        "参与市场化交易的电量小时数"
    ]
})
st.dataframe(param_df, use_container_width=True, hide_index=True)

# -------------------------- 数据保存 --------------------------
if save_btn:
    # 验证必填信息
    if not st.session_state.current_province:
        st.warning("⚠️ 请选择省份/地区")
    elif not st.session_state.current_site:
        st.warning("⚠️ 请输入场站名称")
    elif installed_capacity <= 0:
        st.warning("⚠️ 装机容量必须大于0")
    else:
        # 整合所有数据
        final_data = edited_df.copy()
        # 添加元数据
        final_data["区域"] = st.session_state.current_region
        final_data["省份/地区"] = st.session_state.current_province
        final_data["月份"] = st.session_state.current_month
        final_data["场站名称"] = st.session_state.current_site
        final_data["装机容量(MW)"] = installed_capacity
        final_data["当月发电小时数"] = generation_hours
        final_data["机制电量小时数"] = mechanism_hours
        final_data["保障性小时数"] = guaranteed_hours
        final_data["限电率(%)"] = power_limit_rate
        final_data["市场化交易小时数"] = market_hours
        final_data["保存时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 保存到文件
        try:
            filepath = save_data_to_file(
                st.session_state.current_province,
                st.session_state.current_month,
                st.session_state.current_site,
                final_data
            )
            # 保存到会话状态
            key = f"{st.session_state.current_region}_{st.session_state.current_province}_{st.session_state.current_month}_{st.session_state.current_site}"
            st.session_state.site_data[key] = final_data
            
            st.success(f"✅ 数据保存成功！\n文件路径：{filepath}")
        except Exception as e:
            st.error(f"❌ 保存失败：{str(e)}")

# -------------------------- 数据查询与管理 --------------------------
st.divider()
st.header("🗂️ 历史数据查询")

# 数据查询区域（匹配新的区域-省份层级）
query_col1, query_col2, query_col3, query_col4 = st.columns(4)
with query_col1:
    query_region = st.selectbox("查询区域", list(REGIONS.keys()), key="query_region")
with query_col2:
    # 查询省份也做安全处理
    query_province_list = REGIONS[query_region]
    query_province = st.selectbox("查询省份/地区", query_province_list, key="query_province")
with query_col3:
    query_month = st.selectbox("查询月份", MONTHS, key="query_month")
with query_col4:
    query_site = st.text_input("查询场站名称", key="query_site", placeholder="输入要查询的场站名称")

query_btn = st.button("🔍 查询数据", use_container_width=True)

if query_btn:
    if not query_province or not query_site:
        st.warning("⚠️ 请填写查询省份/地区和场站名称")
    else:
        query_data = load_data_from_file(query_province, query_month, query_site)
        if query_data is not None:
            st.subheader(f"查询结果：{query_region} | {query_province} | {query_month}月 | {query_site}")
            st.dataframe(query_data, use_container_width=True)
            
            # 重新计算关键指标用于展示
            query_total_gen = query_data["当月各时段累计发电量(MWh)"].sum()
            query_installed_cap = query_data["装机容量(MW)"].iloc[0] if "装机容量(MW)" in query_data.columns else 0
            query_gen_hours = calculate_generation_hours(query_total_gen, query_installed_cap)
            
            # 展示查询数据的关键指标
            st.subheader("关键指标")
            q_col1, q_col2, q_col3 = st.columns(3)
            with q_col1:
                st.metric("总发电量(MWh)", f"{query_total_gen:.2f}")
            with q_col2:
                st.metric("装机容量(MW)", f"{query_installed_cap:.1f}")
            with q_col3:
                st.metric("发电小时数", f"{query_gen_hours:.2f}")
        else:
            st.info("ℹ️ 未查询到该条件下的历史数据")

# -------------------------- 页脚信息 --------------------------
st.divider()
st.caption("© 2025 新能源场站年度方案设计系统 | 数据自动保存至本地CSV文件")
