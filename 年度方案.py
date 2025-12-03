import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# -------------------------- 全局配置 & Session State 初始化 --------------------------
st.set_page_config(
    page_title="新能源电厂年度方案设计系统",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 初始化Session State（放在最顶部，所有widget之前）
if "initialized" not in st.session_state:
    # 基础配置
    st.session_state.initialized = True
    st.session_state.current_region = "总部"
    st.session_state.current_province = ""
    st.session_state.current_year = 2025
    st.session_state.current_power_plant = ""
    st.session_state.current_plant_type = "风电"
    
    # 光伏套利时段配置（独立key，避免直接赋值冲突）
    st.session_state["pv_core_start_key"] = 11   # 中午核心时段起始
    st.session_state["pv_core_end_key"] = 14     # 中午核心时段结束
    st.session_state["pv_edge_start_key"] = 6    # 两端边缘时段起始
    st.session_state["pv_edge_end_key"] = 18     # 两端边缘时段结束
    
    # 业务数据
    st.session_state.monthly_data = {}
    st.session_state.selected_months = []
    st.session_state.trade_power_typical = {}  # 方案一：典型曲线
    st.session_state.trade_power_arbitrage = {} # 方案二：光伏套利/风电直线
    st.session_state.total_annual_trade = 0.0
    st.session_state.mechanism_mode = "小时数"
    st.session_state.guaranteed_mode = "小时数"
    st.session_state.manual_market_hours = 0.0
    st.session_state.auto_calculate = True
    st.session_state.calculated = False
    st.session_state.market_hours = {}
    st.session_state.gen_hours = {}

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
    """获取光伏套利曲线的时段划分（从session state读取配置）"""
    # 安全获取配置值（转为整数）
    core_start = int(st.session_state.get("pv_core_start_key", 11))
    core_end = int(st.session_state.get("pv_core_end_key", 14))
    edge_start = int(st.session_state.get("pv_edge_start_key", 6))
    edge_end = int(st.session_state.get("pv_edge_end_key", 18))
    
    # 校验时段有效性
    core_start = max(1, min(24, core_start))
    core_end = max(1, min(24, core_end))
    edge_start = max(1, min(24, edge_start))
    edge_end = max(1, min(24, edge_end))
    
    if core_start > core_end:
        core_start, core_end = core_end, core_start
    if edge_start > edge_end:
        edge_start, edge_end = edge_end, edge_start
    
    # 核心时段（中午，电量接收端）
    core_hours = list(range(core_start, core_end + 1))
    # 边缘时段（两端，电量转出端）
    edge_hours = [h for h in range(edge_start, edge_end + 1) if h not in core_hours]
    # 无效时段（非发电时段）
    invalid_hours = [h for h in range(1, 25) if h not in range(edge_start, edge_end + 1)]
    
    return {
        "core": core_hours,       # 中午核心时段
        "edge": edge_hours,       # 两端边缘时段
        "invalid": invalid_hours, # 无效时段
        "config": {
            "core_start": core_start,
            "core_end": core_end,
            "edge_start": edge_start,
            "edge_end": edge_end
        }
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
    """导出Excel模板（包含12个月份子表）"""
    wb = Workbook()
    wb.remove(wb.active)
    for month in range(1, 13):
        ws = wb.create_sheet(title=f"{month}月")
        template_df = init_month_template(month)
        for r in dataframe_to_rows(template_df, index=False, header=True):
            ws.append(r)
    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def batch_import_excel(file):
    """批量导入Excel（按子表名称匹配月份）"""
    monthly_data = {}
    try:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            if not sheet_name.endswith("月"):
                st.warning(f"跳过无效子表：{sheet_name}（需命名为“1月”-“12月”）")
                continue
            try:
                month = int(sheet_name.replace("月", ""))
                if month < 1 or month > 12:
                    st.warning(f"跳过无效月份子表：{sheet_name}（需1-12月）")
                    continue
                df = pd.read_excel(file, sheet_name=sheet_name)
                required_cols = ["时段", "平均发电量(MWh)", "当月各时段累计发电量(MWh)", "现货价格(元/MWh)", "中长期价格(元/MWh)"]
                if not all(col in df.columns for col in required_cols):
                    st.warning(f"子表{sheet_name}缺少必要列，跳过")
                    continue
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

def calculate_trade_power_typical(month, market_hours, installed_capacity):
    """方案一：典型出力曲线（按发电权重分配）"""
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
    trade_df = trade_df.fillna(0.0)
    trade_df["方案一月度电量(MWh)"] = trade_df["方案一月度电量(MWh)"].astype(np.float64)
    return trade_df, round(total_trade_power, 2)

def calculate_trade_power_arbitrage(month, total_trade_power, typical_df):
    """方案二：光伏套利曲线/风电直线曲线"""
    if month not in st.session_state.monthly_data:
        return None
    
    if st.session_state.current_plant_type == "光伏":
        # 光伏方案二：套利曲线（两端电量转移到中午核心时段）
        pv_hours = get_pv_arbitrage_hours()
        core_hours = pv_hours["core"]
        edge_hours = pv_hours["edge"]
        invalid_hours = pv_hours["invalid"]
        
        # 1. 计算典型曲线中边缘时段的总电量（要转移的电量）
        edge_total = typical_df[typical_df["时段"].isin(edge_hours)]["方案一月度电量(MWh)"].sum()
        # 2. 核心时段数量
        core_count = len(core_hours)
        core_count = core_count if core_count > 0 else 1
        # 3. 每个核心时段增加的电量
        core_add = edge_total / core_count
        
        trade_data = []
        for idx, row in typical_df.iterrows():
            hour = row["时段"]
            avg_gen = row["平均发电量(MWh)"]
            
            if hour in invalid_hours:
                # 无效时段：电量=0
                trade_power = 0.0
                proportion = 0.0
            elif hour in edge_hours:
                # 边缘时段：电量=0（全部转移）
                trade_power = 0.0
                proportion = 0.0
            elif hour in core_hours:
                # 核心时段：原典型电量 + 转移电量
                trade_power = row["方案一月度电量(MWh)"] + core_add
                proportion = trade_power / total_trade_power
            else:
                # 其他时段：保持典型电量
                trade_power = row["方案一月度电量(MWh)"]
                proportion = trade_power / total_trade_power
            
            trade_data.append({
                "时段": hour,
                "平均发电量(MWh)": avg_gen,
                "时段比重(%)": round(proportion * 100, 4),
                "方案二月度电量(MWh)": round(trade_power, 2)
            })
        
        trade_df = pd.DataFrame(trade_data)
    
    else:
        # 风电方案二：24时段直线平均
        avg_generation_list = st.session_state.monthly_data[month]["平均发电量(MWh)"].tolist()
        hourly_trade = total_trade_power / 24
        proportion = 1 / 24
        
        trade_data = []
        for hour, avg_gen in enumerate(avg_generation_list, 1):
            trade_data.append({
                "时段": hour,
                "平均发电量(MWh)": avg_gen,
                "时段比重(%)": round(proportion * 100, 4),
                "方案二月度电量(MWh)": round(hourly_trade, 2)
            })
        trade_df = pd.DataFrame(trade_data)
    
    # 数据清洗和补充
    trade_df["年份"] = st.session_state.current_year
    trade_df["月份"] = month
    trade_df["电厂名称"] = st.session_state.current_power_plant
    trade_df = trade_df.fillna(0.0)
    trade_df["方案二月度电量(MWh)"] = trade_df["方案二月度电量(MWh)"].astype(np.float64)
    
    # 确保方案二总电量和方案一一致
    trade_df["方案二月度电量(MWh)"] = trade_df["方案二月度电量(MWh)"] * (total_trade_power / trade_df["方案二月度电量(MWh)"].sum())
    return trade_df

def decompose_double_scheme(typical_df, arbitrage_df, year, month):
    """双方案日分解（返回四列数据：方案一/二月度+日分解）"""
    days = get_days_in_month(year, month)
    df = pd.DataFrame({
        "时段": typical_df["时段"],
        "方案一月度电量(MWh)": typical_df["方案一月度电量(MWh)"],
        "方案一日分解电量(MWh)": round(typical_df["方案一月度电量(MWh)"] / days, 4),
        "方案二月度电量(MWh)": arbitrage_df["方案二月度电量(MWh)"],
        "方案二日分解电量(MWh)": round(arbitrage_df["方案二月度电量(MWh)"] / days, 4),
        "月份天数": days
    })
    df = df.fillna(0.0)
    return df

def export_annual_plan():
    """导出年度方案Excel（双方案月度+日分解四列数据）"""
    wb = Workbook()
    wb.remove(wb.active)
    total_annual = 0.0
    
    # 1. 年度汇总表（双方案总量）
    summary_data = []
    scheme2_note = "套利曲线（两端转中午）" if st.session_state.current_plant_type == "光伏" else "直线曲线（24小时平均）"
    pv_config = get_pv_arbitrage_hours()["config"] if st.session_state.current_plant_type == "光伏" else {}
    for month in st.session_state.selected_months:
        if month not in st.session_state.trade_power_typical:
            continue
        typical_df = st.session_state.trade_power_typical[month]
        arbitrage_df = st.session_state.trade_power_arbitrage[month]
        total_typical = typical_df["方案一月度电量(MWh)"].sum()
        total_arbitrage = arbitrage_df["方案二月度电量(MWh)"].sum()
        total_annual += total_typical
        summary_data.append({
            "年份": st.session_state.current_year,
            "月份": month,
            "电厂名称": st.session_state.current_power_plant,
            "电厂类型": st.session_state.current_plant_type,
            "光伏核心时段": f"{pv_config.get('core_start', '')}-{pv_config.get('core_end', '')}点" if st.session_state.current_plant_type == "光伏" else "-",
            "光伏边缘时段": f"{pv_config.get('edge_start', '')}-{pv_config.get('edge_end', '')}点" if st.session_state.current_plant_type == "光伏" else "-",
            "方案一（典型曲线）总电量(MWh)": total_typical,
            "方案二（{}）总电量(MWh)".format(scheme2_note): total_arbitrage,
            "月份天数": get_days_in_month(st.session_state.current_year, month),
            "市场化小时数": st.session_state.market_hours.get(month, 0.0),
            "占年度比重(%)": round(total_typical / st.session_state.total_annual_trade * 100, 2)
        })
    summary_df = pd.DataFrame(summary_data)
    ws_summary = wb.create_sheet(title="年度汇总")
    for r in dataframe_to_rows(summary_df, index=False, header=True):
        ws_summary.append(r)
    
    # 2. 各月份详细表（双方案月度+日分解四列）
    for month in st.session_state.selected_months:
        if month not in st.session_state.monthly_data:
            continue
        # 基础数据
        base_df = st.session_state.monthly_data[month][["时段", "平均发电量(MWh)", "现货价格(元/MWh)", "中长期价格(元/MWh)"]].copy()
        # 典型曲线（方案一）
        typical_df = st.session_state.trade_power_typical[month][["时段", "方案一月度电量(MWh)", "时段比重(%)"]].copy()
        typical_df.rename(columns={"时段比重(%)": "方案一时段比重(%)"}, inplace=True)
        # 套利/直线曲线（方案二）
        arbitrage_df = st.session_state.trade_power_arbitrage[month][["时段", "方案二月度电量(MWh)", "时段比重(%)"]].copy()
        arbitrage_df.rename(columns={"时段比重(%)": "方案二时段比重(%)"}, inplace=True)
        # 双方案日分解
        decompose_df = decompose_double_scheme(
            st.session_state.trade_power_typical[month],
            st.session_state.trade_power_arbitrage[month],
            st.session_state.current_year,
            month
        )[["时段", "方案一日分解电量(MWh)", "方案二日分解电量(MWh)", "月份天数"]].copy()
        
        # 合并所有数据
        merged_df = base_df.merge(typical_df, on="时段")
        merged_df = merged_df.merge(arbitrage_df, on="时段")
        merged_df = merged_df.merge(decompose_df, on="时段")
        
        # 创建子表
        ws_month = wb.create_sheet(title=f"{month}月详情")
        for r in dataframe_to_rows(merged_df, index=False, header=True):
            ws_month.append(r)
    
    # 3. 方案说明表
    ws_desc = wb.create_sheet(title="方案说明")
    pv_hours = get_pv_arbitrage_hours()
    pv_desc = f"""
    光伏方案二（套利曲线）配置：
    - 核心时段（接收电量）：{pv_hours['core']}点
    - 边缘时段（转出电量）：{pv_hours['edge']}点
    - 无效时段：{pv_hours['invalid']}点
    - 逻辑：将边缘时段的市场化交易电量全部转移至核心时段，总电量保持不变
    """ if st.session_state.current_plant_type == "光伏" else """
    风电方案二（直线曲线）：
    - 逻辑：24小时平均分配市场化交易电量，总电量与典型曲线一致
    """
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
        ["方案一（典型曲线）：按各时段平均发电量权重分配市场化交易电量"],
        [pv_desc],
        [""],
        [f"年度总交易电量（典型方案）：{round(total_annual, 2)} MWh"]
    ]
    for row in desc_content:
        ws_desc.append(row)
    
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

# -------------------------- 侧边栏配置 --------------------------
with st.sidebar:
    st.header("⚙️ 基础信息配置")
    
    # 1. 年份选择
    years = list(range(2020, 2031))
    st.session_state.current_year = st.selectbox(
        "选择年份", years,
        index=years.index(st.session_state.current_year),
        key="sidebar_year"
    )
    
    # 2. 区域/省份
    selected_region = st.selectbox(
        "选择区域", list(REGIONS.keys()),
        index=list(REGIONS.keys()).index(st.session_state.current_region),
        key="sidebar_region_select"
    )
    st.session_state.current_region = selected_region
    
    current_province_list = REGIONS[st.session_state.current_region]
    if not st.session_state.current_province or st.session_state.current_province not in current_province_list:
        st.session_state.current_province = current_province_list[0]
    
    selected_province = st.selectbox(
        "选择省份/地区", current_province_list,
        index=current_province_list.index(st.session_state.current_province),
        key="sidebar_province_select"
    )
    st.session_state.current_province = selected_province
    
    # 3. 电厂信息
    plant_name = st.text_input(
        "电厂名称", value=st.session_state.current_power_plant,
        key="sidebar_plant_name", placeholder="如：张家口风电场/青海光伏电站"
    )
    st.session_state.current_power_plant = plant_name
    
    st.session_state.current_plant_type = st.selectbox(
        "电厂类型", ["风电", "光伏"],
        index=["风电", "光伏"].index(st.session_state.current_plant_type),
        key="sidebar_plant_type"
    )
    
    # 光伏套利时段配置（仅光伏显示）
    if st.session_state.current_plant_type == "光伏":
        st.subheader("☀️ 光伏套利曲线配置")
        st.write("核心时段（中午，接收电量）")
        col_pv1, col_pv2 = st.columns(2)
        with col_pv1:
            # 使用独立key，避免直接赋值session state
            st.number_input(
                "核心起始（点）", min_value=1, max_value=24,
                value=st.session_state["pv_core_start_key"],
                key="input_pv_core_start"
            )
        with col_pv2:
            st.number_input(
                "核心结束（点）", min_value=1, max_value=24,
                value=st.session_state["pv_core_end_key"],
                key="input_pv_core_end"
            )
        
        st.write("边缘时段（两端，转出电量）")
        col_pv3, col_pv4 = st.columns(2)
        with col_pv3:
            st.number_input(
                "边缘起始（点）", min_value=1, max_value=24,
                value=st.session_state["pv_edge_start_key"],
                key="input_pv_edge_start"
            )
        with col_pv4:
            st.number_input(
                "边缘结束（点）", min_value=1, max_value=24,
                value=st.session_state["pv_edge_end_key"],
                key="input_pv_edge_end"
            )
        
        # 同步input值到session state（关键修复：避免直接赋值）
        st.session_state["pv_core_start_key"] = st.session_state.get("input_pv_core_start", 11)
        st.session_state["pv_core_end_key"] = st.session_state.get("input_pv_core_end", 14)
        st.session_state["pv_edge_start_key"] = st.session_state.get("input_pv_edge_start", 6)
        st.session_state["pv_edge_end_key"] = st.session_state.get("input_pv_edge_end", 18)
        
        # 显示时段划分
        pv_hours = get_pv_arbitrage_hours()
        st.info(f"""
        时段划分：
        - 核心时段（接收）：{pv_hours['core']}点
        - 边缘时段（转出）：{pv_hours['edge']}点
        - 无效时段：{pv_hours['invalid']}点
        """)
    
    # 4. 装机容量
    installed_capacity = st.number_input(
        "装机容量(MW)", min_value=0.0, value=0.0, step=0.1,
        key="sidebar_installed_capacity", help="电厂总装机容量，单位：兆瓦"
    )
    
    # 5. 电量参数配置
    st.subheader("⚡ 电量参数配置")
    
    # 机制电量
    st.write("#### 机制电量")
    col_mech1, col_mech2 = st.columns([2, 1])
    with col_mech1:
        mech_mode = st.selectbox(
            "输入模式", ["小时数", "比例(%)"],
            index=0 if st.session_state.mechanism_mode == "小时数" else 1,
            key="sidebar_mechanism_mode"
        )
        st.session_state.mechanism_mode = mech_mode
    with col_mech2:
        mech_min = 0.0
        mech_max = 100.0 if st.session_state.mechanism_mode == "比例(%)" else 1000000.0
        mechanism_value = st.number_input(
            "数值", min_value=mech_min, max_value=mech_max, value=0.0, step=0.1,
            key="sidebar_mechanism_value", help=f"机制电量{st.session_state.mechanism_mode}"
        )
    
    # 保障性电量
    st.write("#### 保障性电量")
    col_gua1, col_gua2 = st.columns([2, 1])
    with col_gua1:
        gua_mode = st.selectbox(
            "输入模式", ["小时数", "比例(%)"],
            index=0 if st.session_state.guaranteed_mode == "小时数" else 1,
            key="sidebar_guaranteed_mode"
        )
        st.session_state.guaranteed_mode = gua_mode
    with col_gua2:
        gua_min = 0.0
        gua_max = 100.0 if st.session_state.guaranteed_mode == "比例(%)" else 1000000.0
        guaranteed_value = st.number_input(
            "数值", min_value=gua_min, max_value=gua_max, value=0.0, step=0.1,
            key="sidebar_guaranteed_value", help=f"保障性电量{st.session_state.guaranteed_mode}"
        )
    
    # 限电率
    power_limit_rate = st.number_input(
        "限电率(%)", min_value=0.0, max_value=100.0, value=0.0, step=0.1,
        key="sidebar_power_limit_rate", help="电厂当月限电比例，0-100%"
    )
    
    # 市场化交易小时数
    st.write("#### 市场化交易小时数")
    auto_calculate = st.toggle(
        "自动计算", value=st.session_state.auto_calculate,
        key="sidebar_auto_calculate"
    )
    st.session_state.auto_calculate = auto_calculate
    
    manual_market_hours = 0.0
    if not st.session_state.auto_calculate:
        manual_market_hours = st.number_input(
            "手动输入（适用于所有选中月份）", min_value=0.0, max_value=1000000.0,
            value=st.session_state.manual_market_hours, step=0.1,
            key="sidebar_market_hours_manual"
        )
        st.session_state.manual_market_hours = manual_market_hours

# -------------------------- 主页面内容 --------------------------
st.title("⚡ 新能源电厂年度方案设计系统")
scheme2_title = "套利曲线（光伏）/直线曲线（风电）"
st.subheader(
    f"当前配置：{st.session_state.current_year}年 | {st.session_state.current_region} | {st.session_state.current_province} | "
    f"{st.session_state.current_plant_type} | {st.session_state.current_power_plant}"
)
st.caption(f"方案一：典型出力曲线 | 方案二：{scheme2_title}")

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
            st.session_state.selected_months = sorted(list(monthly_data.keys()))
            st.success(f"✅ 批量导入成功！共导入{len(monthly_data)}个月份数据")

# 3. 月份多选
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
        st.warning("⚠️ 请先选择需要处理的月份")

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

# 2. 生成年度双方案
with col_data2:
    if st.button("📝 生成年度双方案", use_container_width=True, type="primary", key="generate_annual_plan"):
        if not st.session_state.selected_months or not st.session_state.monthly_data:
            st.warning("⚠️ 请先导入/初始化月份数据并选择月份")
        elif installed_capacity <= 0:
            st.warning("⚠️ 请填写装机容量")
        else:
            with st.spinner("🔄 正在计算年度双方案..."):
                try:
                    trade_typical = {}
                    trade_arbitrage = {}
                    market_hours = {}
                    gen_hours = {}
                    total_annual = 0.0
                    
                    for month in st.session_state.selected_months:
                        # 计算核心参数
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
                        
                        # 方案一：典型曲线
                        typical_df, total_typical = calculate_trade_power_typical(month, mh, installed_capacity)
                        if typical_df is None:
                            st.error(f"❌ 月份{month}典型方案计算失败")
                            continue
                        trade_typical[month] = typical_df
                        total_annual += total_typical
                        
                        # 方案二：光伏套利/风电直线
                        arbitrage_df = calculate_trade_power_arbitrage(month, total_typical, typical_df)
                        if arbitrage_df is None:
                            st.error(f"❌ 月份{month}方案二计算失败")
                            continue
                        trade_arbitrage[month] = arbitrage_df
                    
                    # 保存到session_state
                    st.session_state.trade_power_typical = trade_typical
                    st.session_state.trade_power_arbitrage = trade_arbitrage
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
            "💾 导出年度方案（双方案+日分解）",
            data=annual_output,
            file_name=f"{st.session_state.current_power_plant}_{st.session_state.current_year}年双方案交易数据.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.button(
            "💾 导出年度方案（双方案+日分解）",
            use_container_width=True,
            disabled=True,
            help="请先生成年度方案"
        )

# 三、选中月份数据编辑
if st.session_state.selected_months and st.session_state.monthly_data:
    st.divider()
    st.header("📊 选中月份数据编辑")
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

# 四、年度方案展示
if st.session_state.calculated and st.session_state.selected_months:
    st.divider()
    st.header(f"📈 {st.session_state.current_year}年度方案展示（双方案对比）")
    
    # 1. 年度汇总
    st.subheader("1. 年度汇总")
    summary_data = []
    scheme2_note = "套利曲线" if st.session_state.current_plant_type == "光伏" else "直线曲线"
    for month in st.session_state.selected_months:
        typical_total = st.session_state.trade_power_typical[month]["方案一月度电量(MWh)"].sum()
        arbitrage_total = st.session_state.trade_power_arbitrage[month]["方案二月度电量(MWh)"].sum()
        days = get_days_in_month(st.session_state.current_year, month)
        summary_data.append({
            "月份": f"{month}月",
            "月份天数": days,
            "市场化小时数": st.session_state.market_hours[month],
            "预估发电小时数": st.session_state.gen_hours[month],
            "方案一总电量(MWh)": typical_total,
            "方案二总电量(MWh)": arbitrage_total,
            "方案二类型": scheme2_note,
            "占年度比重(%)": round(typical_total / st.session_state.total_annual_trade * 100, 2)
        })
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)
    st.metric("年度总交易电量（方案一）", f"{st.session_state.total_annual_trade:.2f} MWh")
    
    # 2. 月份方案详情
    st.subheader("2. 月份方案详情（双方案对比）")
    view_month = st.selectbox(
        "选择查看的月份",
        st.session_state.selected_months,
        key="view_month_select"
    )
    
    # 方案一展示
    st.write(f"### 方案一：典型出力曲线（{view_month}月）")
    typical_df = st.session_state.trade_power_typical[view_month][["时段", "平均发电量(MWh)", "时段比重(%)", "方案一月度电量(MWh)"]].copy()
    typical_df = typical_df.fillna(0.0)
    typical_df["方案一月度电量(MWh)"] = typical_df["方案一月度电量(MWh)"].astype(np.float64)
    typical_df = typical_df.reset_index(drop=True)
    st.dataframe(typical_df, use_container_width=True, hide_index=True)
    
    try:
        chart_data = typical_df[["时段", "方案一月度电量(MWh)"]].set_index("时段")
        if not chart_data.empty and chart_data["方案一月度电量(MWh)"].sum() > 0:
            st.write(f"#### {view_month}月方案一电量分布")
            st.bar_chart(chart_data, use_container_width=True)
        else:
            st.info("⚠️ 暂无有效数据生成图表")
    except Exception as e:
        st.warning(f"📊 方案一图表生成失败：{str(e)}（不影响数据导出）")
    
    # 方案二展示
    st.write(f"### 方案二：{scheme2_note}（{view_month}月）")
    arbitrage_df = st.session_state.trade_power_arbitrage[view_month][["时段", "平均发电量(MWh)", "时段比重(%)", "方案二月度电量(MWh)"]].copy()
    arbitrage_df = arbitrage_df.fillna(0.0)
    arbitrage_df["方案二月度电量(MWh)"] = arbitrage_df["方案二月度电量(MWh)"].astype(np.float64)
    arbitrage_df = arbitrage_df.reset_index(drop=True)
    st.dataframe(arbitrage_df, use_container_width=True, hide_index=True)
    
    # 方案二说明
    if st.session_state.current_plant_type == "光伏":
        pv_hours = get_pv_arbitrage_hours()
        edge_total = typical_df[typical_df["时段"].isin(pv_hours["edge"])]["方案一月度电量(MWh)"].sum()
        core_avg_add = edge_total / len(pv_hours["core"]) if len(pv_hours["core"]) > 0 else 0
        st.info(f"""
        光伏套利曲线说明：
        - 转出时段：{pv_hours['edge']}点（总转出电量={edge_total:.2f} MWh）
        - 接收时段：{pv_hours['core']}点（每时段增加={core_avg_add:.2f} MWh）
        - 总电量：{arbitrage_df['方案二月度电量(MWh)'].sum():.2f} MWh（与方案一一致）
        """)
    else:
        st.info(f"""
        风电直线曲线说明：
        - 24时段平均分配，每时段电量={arbitrage_df['方案二月度电量(MWh)'].iloc[0]:.2f} MWh
        - 总电量：{arbitrage_df['方案二月度电量(MWh)'].sum():.2f} MWh（与方案一一致）
        """)
    
    try:
        chart_data = arbitrage_df[["时段", "方案二月度电量(MWh)"]].set_index("时段")
        if not chart_data.empty and chart_data["方案二月度电量(MWh)"].sum() > 0:
            st.write(f"#### {view_month}月方案二电量分布")
            st.bar_chart(chart_data, use_container_width=True)
        else:
            st.info("⚠️ 暂无有效数据生成图表")
    except Exception as e:
        st.warning(f"📊 方案二图表生成失败：{str(e)}（不影响数据导出）")
    
    # 3. 双方案日分解展示（四列数据）
    st.subheader(f"3. {view_month}月双方案日分解电量（四列数据）")
    decompose_df = decompose_double_scheme(
        st.session_state.trade_power_typical[view_month],
        st.session_state.trade_power_arbitrage[view_month],
        st.session_state.current_year,
        view_month
    )
    decompose_df = decompose_df.fillna(0.0)
    # 显示四列核心数据
    display_df = decompose_df[["时段", "方案一月度电量(MWh)", "方案一日分解电量(MWh)", "方案二月度电量(MWh)", "方案二日分解电量(MWh)"]].copy()
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    st.info(f"""
    日分解说明：
    - 日分解电量 = 月度电量 ÷ {view_month}月天数（{get_days_in_month(st.session_state.current_year, view_month)}天）
    - 方案一/二月度总电量保持一致，日分解电量同步匹配
    """)

# 页脚
st.divider()
st.caption(f"© {st.session_state.current_year} 新能源电厂年度方案设计系统 | 双方案（典型/套利/直线）+ 四列日分解数据 | 总量一致")
