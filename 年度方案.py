# 第一步：调整导入顺序（Streamlit必须放在最顶部）+ 规范格式
import streamlit as st  # 核心库优先导入
import uuid  # 生成唯一标识
import pandas as pd
import numpy as np
import os
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import matplotlib.pyplot as plt  # 绘图库
from io import BytesIO

# -------------------------- 工具函数：生成调整说明文本（移到顶部） --------------------------
def generate_adjustment_description(month, scheme):
    """生成指定月份和方案的调整说明文本"""
    logs = st.session_state.adjustment_logs[month][scheme]
    if not logs:
        return "无调整记录，使用原始生成数据"
    
    desc = []
    desc.append(f"{month}月{scheme}调整记录：")
    for idx, log in enumerate(logs, 1):
        if log["调整类型"] == "月度同步比例调整":
            desc.append(
                f"{idx}. 【比例调整】{log['调整时间']}：按×{log['调整比例']}缩放，"
                f"原始总量{log[f'{scheme}原始总量']:.2f} MWh → 调整后总量{log[f'{scheme}调整后总量']:.2f} MWh"
            )
        elif log["调整类型"] == "分时段电量微调":
            desc.append(
                f"{idx}. 【时段微调】{log['调整时间']}：锁定基准总量{log['锁定基准总量']:.2f} MWh，"
                f"修改后初始总量{log['修改后初始总量']:.2f} MWh，差额{log['差额']:.2f} MWh，"
                f"已按比例分摊至各时段，最终总量{log['分摊后最终总量']:.2f} MWh"
            )
    return "\n".join(desc)

# -------------------------- 全局Session State初始化 --------------------------
# 唯一前缀（避免多组件key冲突）
unique_prefix_ratio_tune = "power_tune"

# 初始化方案数据
if "scheme_power_data" not in st.session_state:
    st.session_state.scheme_power_data = {
        month: {
            "方案一": {"periods": {}, "base_total": 0.0, "original_periods": {}, "original_base_total": 0.0},
            "方案二": {"periods": {}, "base_total": 0.0, "original_periods": {}, "original_base_total": 0.0}
        } for month in range(1, 13)
    }

# 初始化调整比例记录
if "adjust_ratio_records" not in st.session_state:
    st.session_state.adjust_ratio_records = {month: {"方案一": 1.0, "方案二": 1.0} for month in range(1, 13)}

# 初始化调整日志
if "adjustment_logs" not in st.session_state:
    st.session_state.adjustment_logs = {
        month: {"方案一": [], "方案二": []} for month in range(1, 13)
    }

# 月份选择状态
if "selected_months" not in st.session_state:
    st.session_state.selected_months = []

# 市场化小时数相关
if "auto_calculate" not in st.session_state:
    st.session_state.auto_calculate = True
if "manual_market_hours_global" not in st.session_state:
    st.session_state.manual_market_hours_global = 0.0
if "manual_market_hours_monthly" not in st.session_state:
    st.session_state.manual_market_hours_monthly = {month: 0.0 for month in range(1, 13)}

# 分月参数初始化
if "monthly_params" not in st.session_state:
    st.session_state.monthly_params = {
        month: {
            "mechanism_mode": "小时数",
            "mechanism_value": 0.0,
            "guaranteed_mode": "小时数",
            "guaranteed_value": 0.0,
            "power_limit_rate": 0.0,
            "mechanism_price": 0.0,
            "guaranteed_price": 0.0
        } for month in range(1, 13)
    }

# 核心基础状态
if "installed_capacity" not in st.session_state:
    st.session_state.installed_capacity = 0.0
if "monthly_data" not in st.session_state:
    st.session_state.monthly_data = {}

# -------------------------- 必备：区域-省份映射字典 --------------------------
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

# -------------------------- 全局配置（页面样式） --------------------------
st.set_page_config(
    page_title="新能源电厂年度方案设计系统",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

if "initialized" not in st.session_state:
    st.session_state.current_year = 2025
    st.session_state.current_region = "总部"
    st.session_state.current_province = "北京"
    st.session_state.current_power_plant = "示例电厂"
    st.session_state.current_plant_type = "风电"
    st.session_state.installed_capacity = 0.0
    st.session_state.batch_mech_price = 0.0
    st.session_state.batch_gua_price = 0.0
    
    # 光伏套利时段默认配置
    st.session_state["pv_core_start_key"] = 11
    st.session_state["pv_core_end_key"] = 14
    st.session_state["pv_edge_start_key"] = 6
    st.session_state["pv_edge_end_key"] = 18
    
    # 数据存储容器
    st.session_state.monthly_data = {}
    st.session_state.selected_months = []
    st.session_state.trade_power_typical = {}
    st.session_state.trade_power_arbitrage = {}
    st.session_state.market_hours = {}
    st.session_state.gen_hours = {}
    st.session_state.total_annual_trade = 0.0
    st.session_state.calculated = False

    st.session_state.initialized = True

# -------------------------- 核心工具函数 --------------------------
def get_days_in_month(year, month):
    """根据年份和月份获取天数（处理闰年）"""
    if month == 2:
        return 29 if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0) else 28
    elif month in [4, 6, 9, 11]:
        return 30
    else:
        return 31

def get_pv_arbitrage_hours():
    """获取光伏套利曲线的时段划分"""
    core_start = int(st.session_state.get("pv_core_start_key", 11))
    core_end = int(st.session_state.get("pv_core_end_key", 14))
    edge_start = int(st.session_state.get("pv_edge_start_key", 6))
    edge_end = int(st.session_state.get("pv_edge_end_key", 18))
    
    core_start = max(1, min(24, core_start))
    core_end = max(1, min(24, core_end))
    edge_start = max(1, min(24, edge_start))
    edge_end = max(1, min(24, edge_end))
    
    if core_start > core_end:
        core_start, core_end = core_end, core_start
    if edge_start > edge_end:
        edge_start, edge_end = edge_end, edge_start
    
    core_hours = list(range(core_start, core_end + 1))
    edge_hours = [h for h in range(edge_start, edge_end + 1) if h not in core_hours]
    invalid_hours = [h for h in range(1, 25) if h not in range(edge_start, edge_end + 1)]
    
    return {
        "core": core_hours,
        "edge": edge_hours,
        "invalid": invalid_hours,
        "config": {"core_start": core_start, "core_end": core_end, "edge_start": edge_start, "edge_end": edge_end}
    }

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
    """导出Excel模板"""
    wb = Workbook()
    wb.remove(wb.active)
    for month in range(1, 13):
        ws = wb.create_sheet(title=f"{month}月")
        template_df = init_month_template(month)
        for r in dataframe_to_rows(template_df, index=False, header=True):
            ws.append(r)
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def batch_import_excel(file):
    """批量导入Excel"""
    monthly_data = {}
    try:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            if not sheet_name.endswith("月"):
                st.warning(f"跳过无效子表：{sheet_name}")
                continue
            try:
                month = int(sheet_name.replace("月", ""))
                if 1 > month > 12:
                    st.warning(f"跳过无效月份：{sheet_name}")
                    continue
                df = pd.read_excel(file, sheet_name=sheet_name)
                required_cols = ["时段", "平均发电量(MWh)", "当月各时段累计发电量(MWh)", "现货价格(元/MWh)", "中长期价格(元/MWh)"]
                if not all(col in df.columns for col in required_cols):
                    st.warning(f"子表{sheet_name}缺少必要列")
                    continue
                monthly_data[month] = df
            except Exception as e:
                st.warning(f"处理子表{sheet_name}失败：{str(e)}")
        return monthly_data
    except Exception as e:
        st.error(f"批量导入失败：{str(e)}")
        return None

def calculate_core_params_monthly(month, installed_capacity):
    """按月份计算核心参数"""
    month_params = st.session_state.monthly_params.get(month, {})
    power_limit_rate = month_params.get("power_limit_rate", 0.0)
    mechanism_mode = month_params.get("mechanism_mode", "小时数")
    mechanism_value = month_params.get("mechanism_value", 0.0)
    guaranteed_mode = month_params.get("guaranteed_mode", "小时数")
    guaranteed_value = month_params.get("guaranteed_value", 0.0)
    
    if month not in st.session_state.monthly_data:
        return 0.0, 0.0
    
    df = st.session_state.monthly_data[month]
    total_generation = df["当月各时段累计发电量(MWh)"].sum()
    gen_hours = round(total_generation / installed_capacity, 2) if installed_capacity > 0 else 0.0
    
    if gen_hours <= 0:
        return 0.0, 0.0
    
    available_hours = gen_hours * (1 - power_limit_rate / 100)
    
    if mechanism_mode == "小时数":
        available_hours -= mechanism_value
    else:
        available_hours -= gen_hours * (mechanism_value / 100)
    
    if guaranteed_mode == "小时数":
        available_hours -= guaranteed_value
    else:
        available_hours -= gen_hours * (guaranteed_value / 100)
    
    available_hours = max(available_hours, 0.0)
    
    if st.session_state.auto_calculate:
        market_hours = max(round(available_hours, 2), 0.0)
    else:
        manual_hours = st.session_state.manual_market_hours_monthly.get(month, 0.0)
        market_hours = max(round(manual_hours, 2), 0.0) if manual_hours <= available_hours else max(round(available_hours, 2), 0.0)
    
    return gen_hours, market_hours

def calculate_trade_power_typical(month, market_hours, installed_capacity):
    """方案一：典型出力曲线"""
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
            "方案一月度电量(MWh)": round(trade_power, 2)
        })
    trade_df = pd.DataFrame(trade_data)
    trade_df["年份"] = st.session_state.current_year
    trade_df["月份"] = month
    trade_df["电厂名称"] = st.session_state.current_power_plant
    return trade_df, round(total_trade_power, 2)

def calculate_trade_power_arbitrage(month, total_trade_power, typical_df):
    """方案二：光伏套利/风电直线曲线"""
    if month not in st.session_state.monthly_data or typical_df is None or typical_df.empty:
        return None
    
    if st.session_state.current_plant_type == "光伏":
        pv_hours = get_pv_arbitrage_hours()
        core_hours = pv_hours["core"]
        edge_hours = pv_hours["edge"]
        invalid_hours = pv_hours["invalid"]
        
        edge_total = typical_df[typical_df["时段"].isin(edge_hours)]["方案一月度电量(MWh)"].sum()
        core_count = len(core_hours) if len(core_hours) > 0 else 1
        core_add = edge_total / core_count
        
        trade_data = []
        for idx, row in typical_df.iterrows():
            hour = row["时段"]
            avg_gen = row["平均发电量(MWh)"]
            if hour in invalid_hours:
                trade_power = 0.0
            elif hour in edge_hours:
                trade_power = 0.0
            elif hour in core_hours:
                trade_power = row["方案一月度电量(MWh)"] + core_add
            else:
                trade_power = row["方案一月度电量(MWh)"]
            trade_data.append({"时段": hour, "平均发电量(MWh)": avg_gen, "方案二月度电量(MWh)": round(trade_power, 2)})
        trade_df = pd.DataFrame(trade_data)
    else:
        hourly_trade = total_trade_power / 24 if total_trade_power > 0 else 0.0
        trade_data = [{"时段": h, "平均发电量(MWh)": 0.0, "方案二月度电量(MWh)": round(hourly_trade, 2)} for h in range(1,25)]
        trade_df = pd.DataFrame(trade_data)
    
    trade_df["年份"] = st.session_state.current_year
    trade_df["月份"] = month
    trade_df["电厂名称"] = st.session_state.current_power_plant
    return trade_df

def decompose_double_scheme(typical_df, arbitrage_df, year, month):
    """双方案日分解"""
    days = get_days_in_month(year, month)
    df = pd.DataFrame({
        "时段": typical_df["时段"],
        "方案一月度电量(MWh)": typical_df["方案一月度电量(MWh)"],
        "方案一日分解电量(MWh)": round(typical_df["方案一月度电量(MWh)"] / days, 4),
        "方案二月度电量(MWh)": arbitrage_df["方案二月度电量(MWh)"],
        "方案二日分解电量(MWh)": round(arbitrage_df["方案二月度电量(MWh)"] / days, 4)
    })
    return df

def export_annual_plan():
    """导出年度方案Excel"""
    valid_months = [m for m in st.session_state.selected_months if m in st.session_state.trade_power_typical and m in st.session_state.trade_power_arbitrage]
    if not valid_months:
        st.error("无有效数据可导出")
        return None
    
    wb = Workbook()
    wb.remove(wb.active)
    total_annual = 0.0
    summary_data = []
    
    for month in valid_months:
        typ_df = st.session_state.trade_power_typical[month]
        arb_df = st.session_state.trade_power_arbitrage[month]
        total_typ = typ_df["方案一月度电量(MWh)"].sum()
        total_arb = arb_df["方案二月度电量(MWh)"].sum()
        total_annual += total_typ
        summary_data.append({"月份": month, "方案一总电量": total_typ, "方案二总电量": total_arb})
    
    # 汇总表
    ws_sum = wb.create_sheet("年度汇总")
    for r in dataframe_to_rows(pd.DataFrame(summary_data), index=False, header=True):
        ws_sum.append(r)
    
    # 月份表
    for month in valid_months:
        merged = st.session_state.monthly_data[month][["时段", "现货价格(元/MWh)", "中长期价格(元/MWh)"]].merge(
            st.session_state.trade_power_typical[month][["时段", "方案一月度电量(MWh)"]], on="时段"
        ).merge(
            st.session_state.trade_power_arbitrage[month][["时段", "方案二月度电量(MWh)"]], on="时段"
        )
        ws = wb.create_sheet(f"{month}月")
        for r in dataframe_to_rows(merged, index=False, header=True):
            ws.append(r)
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -------------------------- 侧边栏配置 --------------------------
with st.sidebar:
    st.header("⚙️ 基础信息配置")
    years = list(range(2020, 2031))
    st.session_state.current_year = st.selectbox("选择年份", years, index=years.index(st.session_state.current_year))
    
    # 区域省份
    selected_region = st.selectbox("选择区域", list(REGIONS.keys()), index=list(REGIONS.keys()).index(st.session_state.current_region))
    st.session_state.current_region = selected_region
    provinces = REGIONS[selected_region]
    selected_province = st.selectbox("选择省份", provinces, index=provinces.index(st.session_state.current_province))
    st.session_state.current_province = selected_province
    
    # 电厂信息
    st.session_state.current_power_plant = st.text_input("电厂名称", value=st.session_state.current_power_plant)
    st.session_state.current_plant_type = st.selectbox("电厂类型", ["风电", "光伏", "水光互补", "风光互补"], index=["风电", "光伏", "水光互补", "风光互补"].index(st.session_state.current_plant_type))
    
    # 光伏配置
    if st.session_state.current_plant_type == "光伏":
        st.subheader("☀️ 光伏套利配置")
        st.session_state["pv_core_start_key"] = st.number_input("核心起始", 1,24, value=st.session_state["pv_core_start_key"])
        st.session_state["pv_core_end_key"] = st.number_input("核心结束", 1,24, value=st.session_state["pv_core_end_key"])
        st.session_state["pv_edge_start_key"] = st.number_input("边缘起始", 1,24, value=st.session_state["pv_edge_start_key"])
        st.session_state["pv_edge_end_key"] = st.number_input("边缘结束", 1,24, value=st.session_state["pv_edge_end_key"])
    
    # 装机容量
    st.session_state.installed_capacity = st.number_input("装机容量(MW)", 0.0, value=st.session_state.installed_capacity)
    
    # 市场化小时数
    st.session_state.auto_calculate = st.toggle("自动计算市场化小时数", value=st.session_state.auto_calculate)

# -------------------------- 主页面 --------------------------
st.title("⚡ 新能源电厂年度方案设计系统")
st.subheader(f"{st.session_state.current_year}年 | {st.session_state.current_region} | {st.session_state.current_province} | {st.session_state.current_plant_type}")

# 电量参数配置
st.subheader("⚡ 电量参数配置")
col1, col2 = st.columns(2)
with col1:
    st.session_state.batch_mech_mode = st.selectbox("机制电量模式", ["小时数", "比例(%)"])
    st.session_state.batch_mech_value = st.number_input("机制电量数值", 0.0)
    st.session_state.batch_mech_price = st.number_input("机制电价(元/MWh)", 0.0)
with col2:
    st.session_state.batch_gua_mode = st.selectbox("保障性电量模式", ["小时数", "比例(%)"])
    st.session_state.batch_gua_value = st.number_input("保障性电量数值", 0.0)
    st.session_state.batch_gua_price = st.number_input("保障性电价(元/MWh)", 0.0)

st.session_state.batch_limit_rate = st.number_input("限电率(%)", 0.0, 100.0)

if st.button("一键应用到所有月份"):
    for m in range(1,13):
        st.session_state.monthly_params[m] = {
            "mechanism_mode": st.session_state.batch_mech_mode,
            "mechanism_value": st.session_state.batch_mech_value,
            "guaranteed_mode": st.session_state.batch_gua_mode,
            "guaranteed_value": st.session_state.batch_gua_value,
            "power_limit_rate": st.session_state.batch_limit_rate,
            "mechanism_price": st.session_state.batch_mech_price,
            "guaranteed_price": st.session_state.batch_gua_price
        }
    st.success("同步完成！")

# 模板导入导出
st.divider()
st.header("📤 模板导入导出")
col_a, col_b, col_c = st.columns(3)
with col_a:
    st.download_button("导出模板", data=export_template(), file_name="模板.xlsx")
with col_b:
    file = st.file_uploader("批量导入Excel", type="xlsx")
    if file:
        data = batch_import_excel(file)
        if data:
            st.session_state.monthly_data = data
            st.session_state.selected_months = list(data.keys())
            st.success("导入成功！")
with col_c:
    if st.button("全选1-12月"):
        st.session_state.selected_months = list(range(1,13))
    if st.button("取消全选"):
        st.session_state.selected_months = []

# 生成方案
st.divider()
st.header("🔧 生成方案")
if st.button("生成年度双方案", type="primary"):
    if not st.session_state.selected_months or st.session_state.installed_capacity <=0:
        st.warning("请完善数据")
    else:
        with st.spinner("计算中..."):
            for m in st.session_state.selected_months:
                gh, mh = calculate_core_params_monthly(m, st.session_state.installed_capacity)
                typ_df, total = calculate_trade_power_typical(m, mh, st.session_state.installed_capacity)
                arb_df = calculate_trade_power_arbitrage(m, total, typ_df)
                st.session_state.trade_power_typical[m] = typ_df
                st.session_state.trade_power_arbitrage[m] = arb_df
                
                # 写入原始数据
                st.session_state.scheme_power_data[m]["方案一"]["original_periods"] = typ_df.set_index("时段")["方案一月度电量(MWh)"].to_dict()
                st.session_state.scheme_power_data[m]["方案一"]["original_base_total"] = total
                st.session_state.scheme_power_data[m]["方案一"]["periods"] = typ_df.set_index("时段")["方案一月度电量(MWh)"].to_dict()
                st.session_state.scheme_power_data[m]["方案一"]["base_total"] = total
                
                st.session_state.scheme_power_data[m]["方案二"]["original_periods"] = arb_df.set_index("时段")["方案二月度电量(MWh)"].to_dict()
                st.session_state.scheme_power_data[m]["方案二"]["original_base_total"] = total
                st.session_state.scheme_power_data[m]["方案二"]["periods"] = arb_df.set_index("时段")["方案二月度电量(MWh)"].to_dict()
                st.session_state.scheme_power_data[m]["方案二"]["base_total"] = total
            
            st.session_state.calculated = True
            st.success("方案生成完成！")

# 导出方案
if st.session_state.calculated:
    out = export_annual_plan()
    if out:
        st.download_button("导出年度方案", data=out, file_name="年度方案.xlsx")

# 方案展示（完整恢复电价曲线双轴图表）
st.divider()
if st.session_state.calculated:
    st.header("📈 方案展示（含价格对比）")
    view_month = st.selectbox("选择查看的月份", st.session_state.selected_months)
    import plotly.graph_objects as go
    
    # 方案一：交易量+价格双轴图表
    st.subheader("1. 方案一（典型曲线）")
    typical_df = st.session_state.trade_power_typical[view_month]
    base_df = st.session_state.monthly_data[view_month]
    
    # 准备数据（确保24时段对齐）
    merged_data = typical_df[["时段", "方案一月度电量(MWh)"]].copy()
    if len(base_df) >= 24:
        merged_data["现货价格"] = base_df["现货价格(元/MWh)"].head(24).values
        merged_data["中长期价格"] = base_df["中长期价格(元/MWh)"].head(24).values
    else:
        merged_data["现货价格"] = 0.0
        merged_data["中长期价格"] = 0.0
    
    # 创建双轴图表
    fig1 = go.Figure()
    
    # 交易量柱状图（左轴）
    fig1.add_trace(go.Bar(
        x=merged_data["时段"],
        y=merged_data["方案一月度电量(MWh)"],
        name="方案一交易量",
        yaxis="y1",
        marker_color="#4299e1",  # 蓝色
        opacity=0.8
    ))
    
    # 现货价格折线（右轴）
    fig1.add_trace(go.Scatter(
        x=merged_data["时段"],
        y=merged_data["现货价格"],
        name="现货价格",
        yaxis="y2",
        mode="lines+markers",
        line=dict(color="#9f7aea", width=2),  # 紫色
        marker=dict(size=4)
    ))
    
    # 中长期价格折线（右轴）
    fig1.add_trace(go.Scatter(
        x=merged_data["时段"],
        y=merged_data["中长期价格"],
        name="中长期价格",
        yaxis="y2",
        mode="lines+markers",
        line=dict(color="#38b2ac", width=2),  # 青色
        marker=dict(size=4)
    ))
    
    # 布局优化（不使用weight属性）
    fig1.update_layout(
        title=f"{view_month}月 方案一交易量与价格对比",
        title_font=dict(size=13, family="Arial"),  # 修复：删除weight属性
        title_x=0.5,
        plot_bgcolor="white",
        yaxis1=dict(
            title="交易量（MWh）",
            title_font=dict(color="#4299e1"),
            tickfont=dict(color="#4299e1")
        ),
        yaxis2=dict(
            title="价格（元/MWh）",
            title_font=dict(color="#9f7aea"),
            tickfont=dict(color="#9f7aea"),
            overlaying="y",
            side="right"
        ),
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
        margin=dict(l=20, r=20, t=30, b=60)
    )
    st.plotly_chart(fig1, use_container_width=True)
    
    # 方案二：交易量+价格双轴图表
    st.subheader("2. 方案二（套利/直线曲线）")
    arbitrage_df = st.session_state.trade_power_arbitrage[view_month]
    
    # 准备数据
    merged_data2 = arbitrage_df[["时段", "方案二月度电量(MWh)"]].copy()
    if len(base_df) >= 24:
        merged_data2["现货价格"] = base_df["现货价格(元/MWh)"].head(24).values
        merged_data2["中长期价格"] = base_df["中长期价格(元/MWh)"].head(24).values
    else:
        merged_data2["现货价格"] = 0.0
        merged_data2["中长期价格"] = 0.0
    
    # 创建双轴图表
    fig2 = go.Figure()
    
    # 交易量柱状图（左轴）
    fig2.add_trace(go.Bar(
        x=merged_data2["时段"],
        y=merged_data2["方案二月度电量(MWh)"],
        name="方案二交易量",
        yaxis="y1",
        marker_color="#e53e3e",  # 红色
        opacity=0.8
    ))
    
    # 现货价格折线（右轴）
    fig2.add_trace(go.Scatter(
        x=merged_data2["时段"],
        y=merged_data2["现货价格"],
        name="现货价格",
        yaxis="y2",
        mode="lines+markers",
        line=dict(color="#9f7aea", width=2),
        marker=dict(size=4)
    ))
    
    # 中长期价格折线（右轴）
    fig2.add_trace(go.Scatter(
        x=merged_data2["时段"],
        y=merged_data2["中长期价格"],
        name="中长期价格",
        yaxis="y2",
        mode="lines+markers",
        line=dict(color="#38b2ac", width=2),
        marker=dict(size=4)
    ))
    
    # 布局优化
    fig2.update_layout(
        title=f"{view_month}月 方案二交易量与价格对比",
        title_font=dict(size=13, family="Arial"),
        title_x=0.5,
        plot_bgcolor="white",
        yaxis1=dict(
            title="交易量（MWh）",
            title_font=dict(color="#e53e3e"),
            tickfont=dict(color="#e53e3e")
        ),
        yaxis2=dict(
            title="价格（元/MWh）",
            title_font=dict(color="#9f7aea"),
            tickfont=dict(color="#9f7aea"),
            overlaying="y",
            side="right"
        ),
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
        margin=dict(l=20, r=20, t=30, b=60)
    )
    st.plotly_chart(fig2, use_container_width=True)

# -------------------------- 电量调整功能 --------------------------
st.divider()
st.header("⚙️ 电量调整")
# 比例调整
col_x, col_y = st.columns(2)
with col_x:
    adjust_m = st.selectbox("调整月份", range(1,13))
    ratio = st.number_input("调整比例", 0.1, 2.0, 1.0)
    if st.button("执行比例调整"):
        scheme1 = st.session_state.scheme_power_data[adjust_m]["方案一"]
        scheme2 = st.session_state.scheme_power_data[adjust_m]["方案二"]
        
        scheme1["periods"] = {k: round(v*ratio,2) for k,v in scheme1["original_periods"].items()}
        scheme1["base_total"] = round(scheme1["original_base_total"]*ratio,2)
        scheme2["periods"] = {k: round(v*ratio,2) for k,v in scheme2["original_periods"].items()}
        scheme2["base_total"] = round(scheme2["original_base_total"]*ratio,2)
        
        st.session_state.trade_power_typical[adjust_m]["方案一月度电量(MWh)"] = st.session_state.trade_power_typical[adjust_m]["时段"].map(scheme1["periods"])
        st.session_state.trade_power_arbitrage[adjust_m]["方案二月度电量(MWh)"] = st.session_state.trade_power_arbitrage[adjust_m]["时段"].map(scheme2["periods"])
        st.success("调整完成！")

# -------------------------- 收益计算 --------------------------
st.divider()
st.header("💰 收益计算")
if st.session_state.calculated:
    valid = [m for m in st.session_state.selected_months if m in st.session_state.monthly_data]
    if valid:
        select = st.multiselect("选择月份", valid, default=valid)
        if select:
            total1 = 0.0
            total2 = 0.0
            for m in select:
                price = st.session_state.monthly_data[m]["现货价格(元/MWh)"].head(24).mean()
                p1 = st.session_state.trade_power_typical[m]["方案一月度电量(MWh)"].sum() * price
                p2 = st.session_state.trade_power_arbitrage[m]["方案二月度电量(MWh)"].sum() * price
                total1 += p1
                total2 += p2
            col1, col2, col3 = st.columns(3)
            col1.metric("方案一收益", f"¥{total1:.2f}")
            col2.metric("方案二收益", f"¥{total2:.2f}")
            col3.metric("收益差", f"¥{total2-total1:.2f}")
    else:
        st.info("无有效数据")
else:
    st.warning("请先生成方案")

st.divider()
st.caption("© 2025 新能源电厂年度方案设计系统")
