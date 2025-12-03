import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# -------------------------- 初始化配置 --------------------------
st.set_page_config(
    page_title="新能源电厂年度方案设计系统",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 初始化Session State（适配新需求）
if "initialized" not in st.session_state:
    st.session_state.initialized = True
    st.session_state.site_data = {}
    st.session_state.current_region = "总部"
    st.session_state.current_province = ""
    st.session_state.current_year = 2025  # 新增：年份
    st.session_state.current_power_plant = ""  # 修改：站点→电厂
    st.session_state.current_plant_type = "风电"  # 新增：电厂类型（风电/光伏）
    st.session_state.monthly_data = {}  # 新增：存储各月份数据（key:月份，value:DataFrame）
    st.session_state.selected_months = []  # 新增：多选月份
    st.session_state.trade_power_typical = {}  # 新增：典型出力曲线方案（分月份）
    st.session_state.trade_power_linear = {}   # 新增：直线方案（平均分配，分月份）
    st.session_state.total_annual_trade = 0.0  # 新增：年度总交易电量
    st.session_state.mechanism_mode = "小时数"
    st.session_state.guaranteed_mode = "小时数"
    st.session_state.manual_market_hours = 0.0
    st.session_state.auto_calculate = True
    st.session_state.calculated = False
    st.session_state.market_hours = {}  # 分月份市场化小时数
    st.session_state.gen_hours = {}     # 分月份预估发电小时数

# -------------------------- 核心工具函数 --------------------------
def get_days_in_month(year, month):
    """根据年份和月份获取天数（处理闰年）"""
    if month == 2:
        return 29 if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0) else 28
    elif month in [4, 6, 9, 11]:
        return 30
    else:
        return 31

def init_month_template(month):
    """初始化单个月份的模板数据"""
    hours = list(range(1, 25))
    return pd.DataFrame({
        "时段": hours,
        "平均发电量(MWh)": [0.0]*24,
        "当月各时段累计发电量(MWh)": [0.0]*24,
        "现货价格(元/MWh)": [0.0]*24,
        "中长期价格(元/MWh)": [0.0]*24,
        "年份": st.session_state.current_year,
        "月份": month,
        "电厂名称": st.session_state.current_power_plant,
        "电厂类型": st.session_state.current_plant_type,
        "区域": st.session_state.current_region,
        "省份": st.session_state.current_province
    })

def export_template():
    """导出Excel模板（包含12个月份子表）"""
    wb = Workbook()
    # 删除默认工作表
    wb.remove(wb.active)
    # 为每个月份创建子表
    for month in range(1, 13):
        ws = wb.create_sheet(title=f"{month}月")
        template_df = init_month_template(month)
        # 写入数据
        for r in dataframe_to_rows(template_df, index=False, header=True):
            ws.append(r)
    # 保存到 BytesIO
    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def batch_import_excel(file):
    """批量导入Excel（按子表名称匹配月份）"""
    monthly_data = {}
    try:
        # 读取所有子表
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            # 从子表名称提取月份（如“1月”→1）
            if not sheet_name.endswith("月"):
                st.warning(f"跳过无效子表：{sheet_name}（需命名为“1月”-“12月”）")
                continue
            try:
                month = int(sheet_name.replace("月", ""))
                if month < 1 or month > 12:
                    st.warning(f"跳过无效月份子表：{sheet_name}（需1-12月）")
                    continue
                # 读取子表数据
                df = pd.read_excel(file, sheet_name=sheet_name)
                # 验证必要列
                required_cols = ["时段", "平均发电量(MWh)", "当月各时段累计发电量(MWh)", "现货价格(元/MWh)", "中长期价格(元/MWh)"]
                if not all(col in df.columns for col in required_cols):
                    st.warning(f"子表{sheet_name}缺少必要列，跳过")
                    continue
                # 补充元数据
                df["年份"] = st.session_state.current_year
                df["电厂名称"] = st.session_state.current_power_plant
                df["电厂类型"] = st.session_state.current_plant_type
                df["区域"] = st.session_state.current_region
                df["省份"] = st.session_state.current_province
                monthly_data[month] = df
            except Exception as e:
                st.warning(f"处理子表{sheet_name}失败：{str(e)}")
        return monthly_data
    except Exception as e:
        st.error(f"批量导入失败：{str(e)}")
        return None

def calculate_core_params_monthly(month, installed_capacity, power_limit_rate, mechanism_value, mechanism_mode, guaranteed_value, guaranteed_mode):
    """按月份计算核心参数（市场化小时数、发电小时数）"""
    if month not in st.session_state.monthly_data:
        return 0.0, 0.0
    df = st.session_state.monthly_data[month]
    total_generation = df["当月各时段累计发电量(MWh)"].sum()
    # 预估发电小时数
    gen_hours = round(total_generation / installed_capacity, 2) if installed_capacity > 0 else 0.0
    # 市场化交易小时数
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

def calculate_trade_power_typical(month, market_hours, installed_capacity):
    """计算典型出力曲线方案（按发电权重分配）"""
    if month not in st.session_state.monthly_data:
        return None, 0.0
    df = st.session_state.monthly_data[month]
    avg_generation_list = df["平均发电量(MWh)"].tolist()
    total_trade_power = market_hours * installed_capacity
    total_avg_generation = sum(avg_generation_list)
    
    if installed_capacity <= 0 or market_hours <= 0 or total_avg_generation <= 0:
        return None, 0.0
    
    trade_data = []
    for hour, avg_gen in enumerate(avg_generation_list, 1):
        proportion = avg_gen / total_avg_generation
        trade_power = total_trade_power * proportion
        trade_data.append({
            "时段": hour,
            "平均发电量(MWh)": avg_gen,
            "时段比重(%)": round(proportion * 100, 4),
            "市场化交易电量(MWh)": round(trade_power, 2)
        })
    trade_df = pd.DataFrame(trade_data)
    # 补充月份信息
    trade_df["年份"] = st.session_state.current_year
    trade_df["月份"] = month
    trade_df["电厂名称"] = st.session_state.current_power_plant
    # 数据清洗：填充NaN，确保数值类型
    trade_df = trade_df.fillna(0.0)
    trade_df["市场化交易电量(MWh)"] = trade_df["市场化交易电量(MWh)"].astype(float)
    return trade_df, round(total_trade_power, 2)

def calculate_trade_power_linear(month, total_trade_power):
    """计算直线方案（平均分配，各时段电量一致）"""
    if month not in st.session_state.monthly_data:
        return None
    df = st.session_state.monthly_data[month]
    avg_generation_list = df["平均发电量(MWh)"].tolist()
    # 平均分配：总电量/24
    hourly_trade = total_trade_power / 24
    proportion = 1 / 24  # 每个时段占比1/24
    
    trade_data = []
    for hour, avg_gen in enumerate(avg_generation_list, 1):
        trade_data.append({
            "时段": hour,
            "平均发电量(MWh)": avg_gen,
            "时段比重(%)": round(proportion * 100, 4),
            "市场化交易电量(MWh)": round(hourly_trade, 2)
        })
    trade_df = pd.DataFrame(trade_data)
    # 补充月份信息
    trade_df["年份"] = st.session_state.current_year
    trade_df["月份"] = month
    trade_df["电厂名称"] = st.session_state.current_power_plant
    # 数据清洗：填充NaN，确保数值类型
    trade_df = trade_df.fillna(0.0)
    trade_df["市场化交易电量(MWh)"] = trade_df["市场化交易电量(MWh)"].astype(float)
    return trade_df

def decompose_to_daily(trade_df, year, month):
    """将月度24时段电量分解到每天（按月份天数平均）"""
    days = get_days_in_month(year, month)
    df = trade_df.copy()
    # 计算每日该时段电量：月度时段电量 / 天数
    df["每日时段电量(MWh)"] = round(df["市场化交易电量(MWh)"] / days, 4)
    df["月份天数"] = days
    # 数据清洗
    df = df.fillna(0.0)
    return df

def export_annual_plan():
    """导出年度方案Excel（包含两种方案+日分解+模板内容）"""
    wb = Workbook()
    wb.remove(wb.active)
    total_annual = 0.0
    
    # 1. 年度汇总表
    summary_data = []
    for month in st.session_state.selected_months:
        if month not in st.session_state.trade_power_typical:
            continue
        typical_df = st.session_state.trade_power_typical[month]
        linear_df = st.session_state.trade_power_linear[month]
        total_typical = typical_df["市场化交易电量(MWh)"].sum()
        total_linear = linear_df["市场化交易电量(MWh)"].sum()
        total_annual += total_typical
        summary_data.append({
            "年份": st.session_state.current_year,
            "月份": month,
            "电厂名称": st.session_state.current_power_plant,
            "电厂类型": st.session_state.current_plant_type,
            "典型方案总电量(MWh)": total_typical,
            "直线方案总电量(MWh)": total_linear,
            "月份天数": get_days_in_month(st.session_state.current_year, month),
            "市场化小时数": st.session_state.market_hours.get(month, 0.0)
        })
    summary_df = pd.DataFrame(summary_data)
    ws_summary = wb.create_sheet(title="年度汇总")
    for r in dataframe_to_rows(summary_df, index=False, header=True):
        ws_summary.append(r)
    
    # 2. 各月份详细表（模板内容+两种方案+日分解）
    for month in st.session_state.selected_months:
        if month not in st.session_state.monthly_data:
            continue
        # 模板基础数据
        base_df = st.session_state.monthly_data[month].copy()
        # 典型方案数据（含日分解）
        typical_df = st.session_state.trade_power_typical[month].copy()
        typical_daily = decompose_to_daily(typical_df, st.session_state.current_year, month)
        # 直线方案数据（含日分解）
        linear_df = st.session_state.trade_power_linear[month].copy()
        linear_daily = decompose_to_daily(linear_df, st.session_state.current_year, month)
        
        # 合并数据（按时段）
        merged_df = base_df.merge(
            typical_daily[["时段", "时段比重(%)", "市场化交易电量(MWh)", "每日时段电量(MWh)"]],
            on="时段", suffixes=("", "_典型")
        ).merge(
            linear_daily[["时段", "时段比重(%)", "市场化交易电量(MWh)", "每日时段电量(MWh)"]],
            on="时段", suffixes=("", "_直线")
        )
        
        # 创建子表
        ws_month = wb.create_sheet(title=f"{month}月详情")
        for r in dataframe_to_rows(merged_df, index=False, header=True):
            ws_month.append(r)
    
    # 3. 方案说明表
    ws_desc = wb.create_sheet(title="方案说明")
    desc_content = [
        ["新能源电厂年度交易方案说明"],
        [""],
        ["基础信息："],
        [f"电厂名称：{st.session_state.current_power_plant}"],
        [f"电厂类型：{st.session_state.current_plant_type}"],
        [f"年份：{st.session_state.current_year}"],
        [f"区域：{st.session_state.current_region}"],
        [f"省份：{st.session_state.current_province}"],
        [""],
        ["方案说明："],
        ["1. 典型出力曲线方案：按各时段平均发电量权重分配交易电量"],
        ["2. 直线方案：各时段交易电量平均分配（总电量与典型方案一致）"],
        ["3. 日分解电量：月度时段电量 ÷ 当月天数，用于日常执行"],
        [""],
        [f"年度总交易电量（典型方案）：{round(total_annual, 2)} MWh"]
    ]
    for row in desc_content:
        ws_desc.append(row)
    
    # 保存到BytesIO
    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -------------------------- 区域-省份字典 --------------------------
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

# -------------------------- 侧边栏配置（适配新需求） --------------------------
st.sidebar.header("⚙️ 基础信息配置")

# 1. 年份选择（新增）
years = list(range(2020, 2031))
st.session_state.current_year = st.sidebar.selectbox(
    "选择年份", years,
    index=years.index(st.session_state.current_year),
    key="sidebar_year"
)

# 2. 区域/省份
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

# 3. 电厂信息（修改+新增）
plant_name = st.sidebar.text_input(
    "电厂名称", value=st.session_state.current_power_plant,
    key="sidebar_plant_name", placeholder="如：张家口风电场"
)
st.session_state.current_power_plant = plant_name

st.session_state.current_plant_type = st.sidebar.selectbox(
    "电厂类型", ["风电", "光伏"],
    index=["风电", "光伏"].index(st.session_state.current_plant_type),
    key="sidebar_plant_type"
)

# 4. 装机容量
installed_capacity = st.sidebar.number_input(
    "装机容量(MW)", min_value=0.0, value=0.0, step=0.1,
    key="sidebar_installed_capacity", help="电厂总装机容量，单位：兆瓦"
)

# 5. 电量参数配置
st.sidebar.subheader("⚡ 电量参数配置")

# 机制电量
st.sidebar.write("#### 机制电量")
col_mech1, col_mech2 = st.sidebar.columns([2, 1])
mech_mode = col_mech1.selectbox(
    "输入模式", ["小时数", "比例(%)"],
    index=0 if st.session_state.mechanism_mode == "小时数" else 1,
    key="sidebar_mechanism_mode"
)
st.session_state.mechanism_mode = mech_mode

mech_min = 0.0
mech_max = 100.0 if st.session_state.mechanism_mode == "比例(%)" else 1000000.0
mechanism_value = col_mech2.number_input(
    "数值", min_value=mech_min, max_value=mech_max, value=0.0, step=0.1,
    key="sidebar_mechanism_value", help=f"机制电量{st.session_state.mechanism_mode}"
)

# 保障性电量
st.sidebar.write("#### 保障性电量")
col_gua1, col_gua2 = st.sidebar.columns([2, 1])
gua_mode = col_gua1.selectbox(
    "输入模式", ["小时数", "比例(%)"],
    index=0 if st.session_state.guaranteed_mode == "小时数" else 1,
    key="sidebar_guaranteed_mode"
)
st.session_state.guaranteed_mode = gua_mode

gua_min = 0.0
gua_max = 100.0 if st.session_state.guaranteed_mode == "比例(%)" else 1000000.0
guaranteed_value = col_gua2.number_input(
    "数值", min_value=gua_min, max_value=gua_max, value=0.0, step=0.1,
    key="sidebar_guaranteed_value", help=f"保障性电量{st.session_state.guaranteed_mode}"
)

# 限电率
power_limit_rate = st.sidebar.number_input(
    "限电率(%)", min_value=0.0, max_value=100.0, value=0.0, step=0.1,
    key="sidebar_power_limit_rate", help="电厂当月限电比例，0-100%"
)

# 市场化交易小时数（自动/手动）
st.sidebar.write("#### 市场化交易小时数")
auto_calculate = st.sidebar.toggle(
    "自动计算", value=st.session_state.auto_calculate,
    key="sidebar_auto_calculate"
)
st.session_state.auto_calculate = auto_calculate

if not st.session_state.auto_calculate:
    manual_market_hours = st.sidebar.number_input(
        "手动输入（适用于所有选中月份）", min_value=0.0, max_value=1000000.0,
        value=st.session_state.manual_market_hours, step=0.1,
        key="sidebar_market_hours_manual"
    )
    st.session_state.manual_market_hours = manual_market_hours

# -------------------------- 主页面内容 --------------------------
st.title("⚡ 新能源电厂年度方案设计系统")
st.subheader(
    f"当前配置：{st.session_state.current_year}年 | {st.session_state.current_region} | {st.session_state.current_province} | "
    f"{st.session_state.current_plant_type} | {st.session_state.current_power_plant}"
)

# 一、模板导出与批量导入区域
st.divider()
st.header("📤 模板导出与批量导入")
col_import1, col_import2, col_import3 = st.columns(3)

# 1. 导出模板按钮
with col_import1:
    template_output = export_template()
    st.download_button(
        "📥 导出Excel模板（含12个月）",
        data=template_output,
        file_name=f"{st.session_state.current_power_plant}_{st.session_state.current_year}年方案模板.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# 2. 批量导入按钮
with col_import2:
    batch_file = st.file_uploader(
        "📥 批量导入Excel（含多月份子表）",
        type=["xlsx"],
        key="batch_import_file"
    )
    if batch_file is not None:
        monthly_data = batch_import_excel(batch_file)
        if monthly_data:
            st.session_state.monthly_data = monthly_data
            # 自动选中导入的月份
            st.session_state.selected_months = sorted(list(monthly_data.keys()))
            st.success(f"✅ 批量导入成功！共导入{len(monthly_data)}个月份数据")

# 3. 月份多选（从侧边栏移到此处）
with col_import3:
    st.session_state.selected_months = st.multiselect(
        "选择需要处理的月份",
        list(range(1, 13)),
        default=st.session_state.selected_months,
        key="month_multiselect"
    )
    if st.session_state.selected_months:
        st.info(f"当前选中月份：{', '.join([f'{m}月' for m in st.session_state.selected_months])}")
    else:
        st.warning("⚠️ 请选择需要处理的月份")

# 二、数据操作按钮
st.divider()
st.header("🔧 数据操作")
col_data1, col_data2, col_data3 = st.columns(3)

# 1. 初始化选中月份模板
with col_data1:
    if st.button("📋 初始化选中月份模板", use_container_width=True, key="init_selected_months"):
        if not st.session_state.selected_months:
            st.warning("⚠️ 请先选择月份")
        else:
            for month in st.session_state.selected_months:
                st.session_state.monthly_data[month] = init_month_template(month)
            st.success(f"✅ 已初始化{len(st.session_state.selected_months)}个月份模板")

# 2. 生成年度方案（含两种方案）
with col_data2:
    if st.button("📝 生成年度双方案", use_container_width=True, type="primary", key="generate_annual_plan"):
        if not st.session_state.selected_months or not st.session_state.monthly_data:
            st.warning("⚠️ 请先导入/初始化月份数据并选择月份")
        elif installed_capacity <= 0:
            st.warning("⚠️ 请填写装机容量")
        else:
            with st.spinner("🔄 正在计算年度方案（含典型/直线双方案）..."):
                try:
                    trade_typical = {}  # 典型方案
                    trade_linear = {}    # 直线方案
                    market_hours = {}   # 分月份市场化小时数
                    gen_hours = {}       # 分月份发电小时数
                    total_annual = 0.0   # 年度总电量
                    
                    for month in st.session_state.selected_months:
                        # 计算核心参数（市场化小时数）
                        if st.session_state.auto_calculate:
                            gh, mh = calculate_core_params_monthly(
                                month, installed_capacity, power_limit_rate,
                                mechanism_value, st.session_state.mechanism_mode,
                                guaranteed_value, st.session_state.guaranteed_mode
                            )
                        else:
                            gh = calculate_core_params_monthly(month, installed_capacity, 0, 0, "小时数", 0, "小时数")[0]
                            mh = st.session_state.manual_market_hours
                        market_hours[month] = mh
                        gen_hours[month] = gh
                        
                        # 计算典型方案
                        typical_df, total_typical = calculate_trade_power_typical(month, mh, installed_capacity)
                        if typical_df is None:
                            st.error(f"❌ 月份{month}典型方案计算失败")
                            continue
                        trade_typical[month] = typical_df
                        total_annual += total_typical
                        
                        # 计算直线方案（总电量与典型方案一致）
                        linear_df = calculate_trade_power_linear(month, total_typical)
                        if linear_df is None:
                            st.error(f"❌ 月份{month}直线方案计算失败")
                            continue
                        trade_linear[month] = linear_df
                    
                    # 保存到session_state
                    st.session_state.trade_power_typical = trade_typical
                    st.session_state.trade_power_linear = trade_linear
                    st.session_state.market_hours = market_hours
                    st.session_state.gen_hours = gen_hours
                    st.session_state.total_annual_trade = total_annual
                    st.session_state.calculated = True
                    
                    st.success(
                        f"✅ 年度双方案生成成功！\n"
                        f"年度总交易电量：{round(total_annual, 2)} MWh\n"
                        f"涉及月份：{', '.join([f'{m}月' for m in st.session_state.selected_months])}"
                    )
                except Exception as e:
                    st.error(f"❌ 生成方案失败：{str(e)}")

# 3. 导出年度方案
with col_data3:
    if st.session_state.calculated and st.session_state.selected_months:
        annual_output = export_annual_plan()
        st.download_button(
            "💾 导出年度方案（含双方案+日分解）",
            data=annual_output,
            file_name=f"{st.session_state.current_power_plant}_{st.session_state.current_year}年交易方案.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.button(
            "💾 导出年度方案（含双方案+日分解）",
            use_container_width=True,
            disabled=True,
            help="请先生成年度方案"
        )

# 三、选中月份数据编辑
if st.session_state.selected_months and st.session_state.monthly_data:
    st.divider()
    st.header("📊 选中月份数据编辑")
    # 选择要编辑的月份
    edit_month = st.selectbox(
        "选择要编辑的月份",
        st.session_state.selected_months,
        key="edit_month_select"
    )
    if edit_month in st.session_state.monthly_data:
        edit_df = st.data_editor(
            st.session_state.monthly_data[edit_month],
            column_config={
                "时段": st.column_config.NumberColumn("时段", disabled=True),
                "平均发电量(MWh)": st.column_config.NumberColumn("平均发电量(MWh)", min_value=0.0, step=0.1),
                "当月各时段累计发电量(MWh)": st.column_config.NumberColumn("当月各时段累计发电量(MWh)", min_value=0.0, step=0.1),
                "现货价格(元/MWh)": st.column_config.NumberColumn("现货价格(元/MWh)", min_value=0.0, step=0.1),
                "中长期价格(元/MWh)": st.column_config.NumberColumn("中长期价格(元/MWh)", min_value=0.0, step=0.1),
                "年份": st.column_config.NumberColumn("年份", disabled=True),
                "月份": st.column_config.NumberColumn("月份", disabled=True),
                "电厂名称": st.column_config.TextColumn("电厂名称", disabled=True),
                "电厂类型": st.column_config.TextColumn("电厂类型", disabled=True),
                "区域": st.column_config.TextColumn("区域", disabled=True),
                "省份": st.column_config.TextColumn("省份", disabled=True)
            },
            use_container_width=True,
            num_rows="fixed",
            key=f"edit_df_{edit_month}"
        )
        st.session_state.monthly_data[edit_month] = edit_df

# 四、年度方案展示（双方案对比）
if st.session_state.calculated and st.session_state.selected_months:
    st.divider()
    st.header(f"📈 {st.session_state.current_year}年度方案展示（双方案对比）")
    
    # 1. 年度汇总
    st.subheader("1. 年度汇总")
    summary_data = []
    for month in st.session_state.selected_months:
        typical_total = st.session_state.trade_power_typical[month]["市场化交易电量(MWh)"].sum()
        linear_total = st.session_state.trade_power_linear[month]["市场化交易电量(MWh)"].sum()
        days = get_days_in_month(st.session_state.current_year, month)
        summary_data.append({
            "月份": f"{month}月",
            "月份天数": days,
            "市场化小时数": st.session_state.market_hours[month],
            "预估发电小时数": st.session_state.gen_hours[month],
            "典型方案电量(MWh)": typical_total,
            "直线方案电量(MWh)": linear_total,
            "占年度比重(%)": round(typical_total / st.session_state.total_annual_trade * 100, 2)
        })
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)
    st.metric("年度总交易电量（典型方案）", f"{st.session_state.total_annual_trade:.2f} MWh")
    
    # 2. 月份方案详情（双方案对比）
    st.subheader("2. 月份方案详情（双方案对比）")
    view_month = st.selectbox(
        "选择查看的月份",
        st.session_state.selected_months,
        key="view_month_select"
    )
    
    # 典型方案展示
    st.write(f"### 典型出力曲线方案（{view_month}月）")
    typical_df = st.session_state.trade_power_typical[view_month][["时段", "平均发电量(MWh)", "时段比重(%)", "市场化交易电量(MWh)"]].copy()
    # 数据清洗：确保无NaN，类型正确
    typical_df = typical_df.fillna(0.0)
    typical_df["市场化交易电量(MWh)"] = typical_df["市场化交易电量(MWh)"].astype(float)
    st.dataframe(typical_df, use_container_width=True, hide_index=True)
    # 典型方案图表（修复参数，适配Streamlit API）
    chart_data_typical = typical_df.set_index("时段")["市场化交易电量(MWh)"]
    st.bar_chart(
        chart_data_typical,
        use_container_width=True,
        ylabel="交易电量(MWh)"  # 修复：y_label → ylabel
    )
    st.caption(f"{view_month}月典型方案电量分布")
    
    # 直线方案展示
    st.write(f"### 直线方案（平均分配，{view_month}月）")
    linear_df = st.session_state.trade_power_linear[view_month][["时段", "平均发电量(MWh)", "时段比重(%)", "市场化交易电量(MWh)"]].copy()
    # 数据清洗：确保无NaN，类型正确
    linear_df = linear_df.fillna(0.0)
    linear_df["市场化交易电量(MWh)"] = linear_df["市场化交易电量(MWh)"].astype(float)
    st.dataframe(linear_df, use_container_width=True, hide_index=True)
    # 直线方案图表（修复参数，适配Streamlit API）
    chart_data_linear = linear_df.set_index("时段")["市场化交易电量(MWh)"]
    st.bar_chart(
        chart_data_linear,
        use_container_width=True,
        ylabel="交易电量(MWh)"  # 修复：y_label → ylabel
    )
    st.caption(f"{view_month}月直线方案电量分布")
    
    # 3. 日分解展示（当前查看月份）
    st.subheader(f"3. {view_month}月日分解电量（按天数平均）")
    # 典型方案日分解
    typical_daily = decompose_to_daily(st.session_state.trade_power_typical[view_month], st.session_state.current_year, view_month)
    linear_daily = decompose_to_daily(st.session_state.trade_power_linear[view_month], st.session_state.current_year, view_month)
    
    daily_compare = pd.DataFrame({
        "时段": typical_daily["时段"],
        "典型方案日电量(MWh)": typical_daily["每日时段电量(MWh)"],
        "直线方案日电量(MWh)": linear_daily["每日时段电量(MWh)"],
        "月份天数": typical_daily["月份天数"]
    })
    # 数据清洗
    daily_compare = daily_compare.fillna(0.0)
    st.dataframe(daily_compare, use_container_width=True, hide_index=True)
    st.info(f"注：日电量 = 月度时段电量 ÷ {view_month}月天数（{get_days_in_month(st.session_state.current_year, view_month)}天）")

# 页脚
st.divider()
st.caption(f"© {st.session_state.current_year} 新能源电厂年度方案设计系统 | 支持风电/光伏双类型 | 双方案对比导出")
