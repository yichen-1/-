import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import matplotlib.pyplot as plt  # 新增：导入matplotlib

# -------------------------- 必备：区域-省份映射字典（合并去重，保留详细版本） --------------------------
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

# -------------------------- 全局配置 & Session State 初始化（完善缺失默认值） --------------------------
st.set_page_config(
    page_title="新能源电厂年度方案设计系统",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

if "initialized" not in st.session_state:
    # 基础信息默认值
    st.session_state.current_year = 2025
    st.session_state.current_region = "总部"
    st.session_state.current_province = REGIONS["总部"][0]  # 联动区域默认省份
    st.session_state.current_power_plant = "示例电厂"
    st.session_state.current_plant_type = "风电"
    st.session_state.installed_capacity = 0.0
    st.session_state.current_region = "总部"
    st.session_state.current_province = "北京"
    
    # 光伏套利时段默认配置（首次运行不报错）
    st.session_state["pv_core_start_key"] = 11
    st.session_state["pv_core_end_key"] = 14
    st.session_state["pv_edge_start_key"] = 6
    st.session_state["pv_edge_end_key"] = 18
    
    # 市场化小时数相关
    st.session_state.auto_calculate = True  # 默认自动计算
    st.session_state.manual_market_hours = 0.0
    
    # 数据存储容器
    st.session_state.monthly_data = {}  # 分月基础数据
    st.session_state.selected_months = []  # 选中的月份
    st.session_state.trade_power_typical = {}  # 方案一结果
    st.session_state.trade_power_arbitrage = {}  # 方案二结果
    st.session_state.market_hours = {}  # 分月市场化小时数
    st.session_state.gen_hours = {}  # 分月发电小时数
    st.session_state.total_annual_trade = 0.0  # 年度总电量
    st.session_state.calculated = False  # 是否已生成方案
    
    # 分月电量参数（每个月独立存储）
    st.session_state.monthly_params = {
        month: {  # 1-12月，每个月对应独立参数
            "mechanism_mode": "小时数",    # 机制电量输入模式
            "mechanism_value": 0.0,        # 机制电量数值
            "guaranteed_mode": "小时数",   # 保障性电量输入模式
            "guaranteed_value": 0.0,       # 保障性电量数值
            "power_limit_rate": 0.0        # 限电率(%)
        } for month in range(1, 13)
    }
    
    # 批量应用的默认参数（用于批量设置时的初始值）
    st.session_state.batch_mech_mode = "小时数"
    st.session_state.batch_mech_value = 0.0
    st.session_state.batch_gua_mode = "小时数"
    st.session_state.batch_gua_value = 0.0
    st.session_state.batch_limit_rate = 0.0
    
    # 标记初始化完成
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
    """获取光伏套利曲线的时段划分（从session state读取配置）"""
    # 安全获取配置值（转为整数，避免类型错误）
    core_start = int(st.session_state.get("pv_core_start_key", 11))
    core_end = int(st.session_state.get("pv_core_end_key", 14))
    edge_start = int(st.session_state.get("pv_edge_start_key", 6))
    edge_end = int(st.session_state.get("pv_edge_end_key", 18))
    
    # 校验时段有效性（防止超出1-24范围）
    core_start = max(1, min(24, core_start))
    core_end = max(1, min(24, core_end))
    edge_start = max(1, min(24, edge_start))
    edge_end = max(1, min(24, edge_end))
    
    # 确保起始<=结束
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

def calculate_core_params_monthly(month, installed_capacity):
    """按月份计算核心参数（市场化小时数、发电小时数）- 内部读取分月参数"""
    # 安全获取该月份的分月参数（避免KeyError）
    month_params = st.session_state.monthly_params.get(month, {
        "power_limit_rate": 0.0,
        "mechanism_mode": "小时数",
        "mechanism_value": 0.0,
        "guaranteed_mode": "小时数",
        "guaranteed_value": 0.0
    })
    
    # 提取参数（带默认值，防止参数缺失）
    power_limit_rate = month_params.get("power_limit_rate", 0.0)
    mechanism_mode = month_params.get("mechanism_mode", "小时数")
    mechanism_value = month_params.get("mechanism_value", 0.0)
    guaranteed_mode = month_params.get("guaranteed_mode", "小时数")
    guaranteed_value = month_params.get("guaranteed_value", 0.0)
    
    # 校验基础数据是否存在
    if month not in st.session_state.monthly_data:
        st.warning(f"⚠️ 月份{month}无基础数据，发电小时数按0计算")
        return 0.0, 0.0
    
    df = st.session_state.monthly_data[month]
    total_generation = df["当月各时段累计发电量(MWh)"].sum()
    
    # 计算发电小时数（避免装机容量为0）
    gen_hours = round(total_generation / installed_capacity, 2) if installed_capacity > 0 else 0.0
    if gen_hours <= 0:
        st.warning(f"⚠️ 月份{month}发电小时数为0（累计发电量：{total_generation:.2f} MWh，装机容量：{installed_capacity:.2f} MW）")
        market_hours = 0.0
    else:
        # 计算可用小时数（扣除限电）
        available_hours = gen_hours * (1 - power_limit_rate / 100)
        
        # 扣除机制电量
        if mechanism_mode == "小时数":
            available_hours -= mechanism_value
        else:  # 比例(%)
            available_hours -= gen_hours * (mechanism_value / 100)
        
        # 扣除保障性电量
        if guaranteed_mode == "小时数":
            available_hours -= guaranteed_value
        else:  # 比例(%)
            available_hours -= gen_hours * (guaranteed_value / 100)
        
        # 市场化小时数不能为负
        market_hours = max(round(available_hours, 2), 0.0)
    
    return gen_hours, market_hours

def calculate_trade_power_typical(month, market_hours, installed_capacity):
    """方案一：典型出力曲线（按发电权重分配）"""
    # 先校验基础数据
    if month not in st.session_state.monthly_data:
        st.warning(f"⚠️ 月份{month}无基础数据，方案一计算失败")
        return None, 0.0
    
    df = st.session_state.monthly_data[month]
    avg_generation_list = df["平均发电量(MWh)"].tolist()
    total_trade_power = market_hours * installed_capacity
    total_avg_generation = sum(avg_generation_list)
    
    if installed_capacity <= 0 or market_hours <= 0 or total_avg_generation <= 0:
        st.warning(f"⚠️ 月份{month}参数异常（装机容量/市场化小时数/平均发电量不能为0）")
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
    
    # 校验列是否存在（防止生成失败）
    if "方案一月度电量(MWh)" not in trade_df.columns:
        st.error(f"❌ 月份{month}方案一数据列缺失")
        return None, 0.0
    
    return trade_df, round(total_trade_power, 2)

def calculate_trade_power_arbitrage(month, total_trade_power, typical_df):
    """方案二：光伏套利曲线/风电直线曲线"""
    # 先校验基础数据和典型方案数据
    if month not in st.session_state.monthly_data:
        st.warning(f"⚠️ 月份{month}无基础数据，方案二计算失败")
        return None
    if typical_df is None or typical_df.empty:
        st.warning(f"⚠️ 月份{month}典型方案数据无效，方案二计算失败")
        return None
    
    if st.session_state.current_plant_type == "光伏":
        # 光伏方案二：套利曲线（两端电量转移到中午核心时段）
        pv_hours = get_pv_arbitrage_hours()
        core_hours = pv_hours["core"]
        edge_hours = pv_hours["edge"]
        invalid_hours = pv_hours["invalid"]
        
        # 1. 计算典型曲线中边缘时段的总电量（要转移的电量）
        edge_total = typical_df[typical_df["时段"].isin(edge_hours)]["方案一月度电量(MWh)"].sum()
        # 2. 核心时段数量（避免除以0）
        core_count = len(core_hours) if len(core_hours) > 0 else 1
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
                proportion = trade_power / total_trade_power if total_trade_power > 0 else 0.0
            else:
                # 其他时段：保持典型电量
                trade_power = row["方案一月度电量(MWh)"]
                proportion = trade_power / total_trade_power if total_trade_power > 0 else 0.0
            
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
        hourly_trade = total_trade_power / 24 if total_trade_power > 0 else 0.0
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
    
    # 确保方案二总电量和方案一一致（修正浮点数误差）
    if total_trade_power > 0:
        trade_df["方案二月度电量(MWh)"] = trade_df["方案二月度电量(MWh)"] * (total_trade_power / trade_df["方案二月度电量(MWh)"].sum())
    
    # 校验列是否存在
    if "方案二月度电量(MWh)" not in trade_df.columns:
        st.error(f"❌ 月份{month}方案二数据列缺失")
        return None
    
    return trade_df

def decompose_double_scheme(typical_df, arbitrage_df, year, month):
    """双方案日分解（返回四列数据：方案一/二月度+日分解）"""
    days = get_days_in_month(year, month)
    df = pd.DataFrame({
        "时段": typical_df["时段"],
        "方案一月度电量(MWh)": typical_df["方案一月度电量(MWh)"],
        "方案一日分解电量(MWh)": round(typical_df["方案一月度电量(MWh)"] / days, 4) if days > 0 else 0.0,
        "方案二月度电量(MWh)": arbitrage_df["方案二月度电量(MWh)"],
        "方案二日分解电量(MWh)": round(arbitrage_df["方案二月度电量(MWh)"] / days, 4) if days > 0 else 0.0,
        "月份天数": days
    })
    df = df.fillna(0.0)
    return df

def export_annual_plan():
    """导出年度方案Excel（双方案月度+日分解四列数据）"""
    # 第一步：过滤有效月份（仅保留生成方案成功的月份）
    valid_months = []
    for month in st.session_state.selected_months:
        # 校验该月份是否同时存在方案一和方案二的数据，且包含目标列
        if (month in st.session_state.trade_power_typical 
            and month in st.session_state.trade_power_arbitrage
            and not st.session_state.trade_power_typical[month].empty
            and not st.session_state.trade_power_arbitrage[month].empty
            and "方案一月度电量(MWh)" in st.session_state.trade_power_typical[month].columns
            and "方案二月度电量(MWh)" in st.session_state.trade_power_arbitrage[month].columns):
            valid_months.append(month)
        else:
            st.warning(f"⚠️ 跳过无效月份 {month}月（方案数据未生成或列缺失）")
    
    if not valid_months:
        st.error("❌ 无有效方案数据可导出，请先确保所有选中月份的方案生成成功！")
        return None  # 无有效数据时返回None，避免后续报错
    
    wb = Workbook()
    wb.remove(wb.active)
    total_annual = 0.0
    
    # 1. 年度汇总表（双方案总量）
    summary_data = []
    scheme2_note = "套利曲线（两端转中午）" if st.session_state.current_plant_type == "光伏" else "直线曲线（24小时平均）"
    pv_config = get_pv_arbitrage_hours()["config"] if st.session_state.current_plant_type == "光伏" else {}
    
    # 循环有效月份（而非全部选中月份）
    for month in valid_months:
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
            "占年度比重(%)": round(total_typical / total_annual * 100, 2) if total_annual > 0 else 0.0
        })
    
    # 生成年度汇总表（确保有数据才生成）
    if summary_data:
        summary_df = pd.DataFrame(summary_data)
        ws_summary = wb.create_sheet(title="年度汇总")
        for r in dataframe_to_rows(summary_df, index=False, header=True):
            ws_summary.append(r)
    else:
        st.error("❌ 无有效汇总数据，导出失败！")
        return None
    
    # 2. 各月份详细表（双方案月度+日分解四列）
    for month in valid_months:
        typical_df = st.session_state.trade_power_typical[month]
        arbitrage_df = st.session_state.trade_power_arbitrage[month]
        
        # 基础数据（安全访问）
        base_df = st.session_state.monthly_data.get(month, None)
        if base_df is None:
            st.warning(f"⚠️ 月份 {month}月 基础数据缺失，跳过详细表")
            continue
        
        # 基础数据列校验
        base_cols = ["时段", "平均发电量(MWh)", "现货价格(元/MWh)", "中长期价格(元/MWh)"]
        if not all(col in base_df.columns for col in base_cols):
            st.warning(f"⚠️ 月份 {month}月 基础数据缺少必要列，跳过详细表")
            continue
        
        # 典型曲线（方案一）- 只保留需要的列
        typical_df_selected = typical_df[["时段", "方案一月度电量(MWh)", "时段比重(%)"]].copy()
        typical_df_selected.rename(columns={"时段比重(%)": "方案一时段比重(%)"}, inplace=True)
        
        # 套利/直线曲线（方案二）- 只保留需要的列
        arbitrage_df_selected = arbitrage_df[["时段", "方案二月度电量(MWh)", "时段比重(%)"]].copy()
        arbitrage_df_selected.rename(columns={"时段比重(%)": "方案二时段比重(%)"}, inplace=True)
        
        # 双方案日分解
        decompose_df = decompose_double_scheme(typical_df, arbitrage_df, st.session_state.current_year, month)
        decompose_df = decompose_df[["时段", "方案一日分解电量(MWh)", "方案二日分解电量(MWh)", "月份天数"]].copy()
        
        # 合并所有数据（按时段关联）
        merged_df = base_df[base_cols].merge(typical_df_selected, on="时段")
        merged_df = merged_df.merge(arbitrage_df_selected, on="时段")
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
        [f"装机容量：{st.session_state.installed_capacity} MW"],
        [""],
        ["方案说明："],
        ["方案一（典型曲线）：按各时段平均发电量权重分配市场化交易电量"],
        [pv_desc],
        [""],
        [f"年度总交易电量（典型方案）：{round(total_annual, 2)} MWh"]
    ]
    for row in desc_content:
        ws_desc.append(row)
    
    # 导出Excel
    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -------------------------- 侧边栏配置 --------------------------
with st.sidebar:
    st.header("⚙️ 基础信息配置")
    
    # 1. 年份选择
    years = list(range(2020, 2031))
    current_year = st.session_state.get("current_year", 2025)
    if current_year not in years:
        current_year = 2025
    st.session_state.current_year = st.selectbox(
        "选择年份", years,
        index=years.index(current_year),
        key="sidebar_year"
    )
    
    # 2. 区域/省份选择（联动+兜底）
    current_region = st.session_state.get("current_region", "总部")
    if current_region not in REGIONS.keys():
        current_region = "总部"
    selected_region = st.selectbox(
        "选择区域",
        list(REGIONS.keys()),
        index=list(REGIONS.keys()).index(current_region),
        key="sidebar_region_select"
    )
    st.session_state.current_region = selected_region
    
    # 省份选择（联动区域）
    provinces = REGIONS[selected_region]
    current_province = st.session_state.get("current_province", provinces[0])
    if current_province not in provinces:
        current_province = provinces[0]
    selected_province = st.selectbox(
        "选择省份",
        provinces,
        index=provinces.index(current_province),
        key="sidebar_province_select"
    )
    st.session_state.current_province = selected_province
    
    # 3. 电厂信息
    st.session_state.current_power_plant = st.text_input(
        "电厂名称",
        value=st.session_state.current_power_plant,
        key="sidebar_power_plant"
    )
    st.session_state.current_plant_type = st.selectbox(
        "电厂类型",
        ["风电", "光伏", "水光互补", "风光互补"],
        index=["风电", "光伏", "水光互补", "风光互补"].index(st.session_state.current_plant_type),
        key="sidebar_plant_type"
    )
    
    # 光伏套利时段配置（仅光伏显示）
    if st.session_state.current_plant_type == "光伏":
        st.subheader("☀️ 光伏套利曲线配置")
        st.write("核心时段（中午，接收电量）")
        col_pv1, col_pv2 = st.columns(2)
        with col_pv1:
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
        
        # 同步input值到session state
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
        "装机容量(MW)", min_value=0.0, value=st.session_state.installed_capacity, step=0.1,
        key="sidebar_installed_capacity", help="电厂总装机容量，单位：兆瓦"
    )
    st.session_state.installed_capacity = installed_capacity  # 同步到session state
    
    # 5. 市场化交易小时数
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

# -------------------------- 主页面：电量参数配置 --------------------------
st.subheader("⚡ 电量参数配置")

# 1. 批量应用参数（一键同步到所有月份）
st.write("#### 批量应用（同步到所有月份）")
col_mech1, col_mech2 = st.columns([2, 1])
with col_mech1:
    st.session_state.batch_mech_mode = st.selectbox(
        "机制电量输入模式", ["小时数", "比例(%)"],
        index=0 if st.session_state.batch_mech_mode == "小时数" else 1,
        key="batch_mech_mode_sel"
    )
with col_mech2:
    mech_max = 100.0 if st.session_state.batch_mech_mode == "比例(%)" else 1000000.0
    st.session_state.batch_mech_value = st.number_input(
        "机制电量数值", min_value=0.0, max_value=mech_max, 
        value=st.session_state.batch_mech_value, step=0.1,
        key="batch_mech_val_inp"
    )

col_gua1, col_gua2 = st.columns([2, 1])
with col_gua1:
    st.session_state.batch_gua_mode = st.selectbox(
        "保障性电量输入模式", ["小时数", "比例(%)"],
        index=0 if st.session_state.batch_gua_mode == "小时数" else 1,
        key="batch_gua_mode_sel"
    )
with col_gua2:
    gua_max = 100.0 if st.session_state.batch_gua_mode == "比例(%)" else 1000000.0
    st.session_state.batch_gua_value = st.number_input(
        "保障性电量数值", min_value=0.0, max_value=gua_max,
        value=st.session_state.batch_gua_value, step=0.1,
        key="batch_gua_val_inp"
    )

st.session_state.batch_limit_rate = st.number_input(
    "限电率(%)", min_value=0.0, max_value=100.0,
    value=st.session_state.batch_limit_rate, step=0.1,
    key="batch_limit_rate_inp"
)

# 批量应用按钮
if st.button("📌 一键应用到所有月份", type="primary", key="batch_apply_btn"):
    for month in range(1, 13):
        st.session_state.monthly_params[month] = {
            "mechanism_mode": st.session_state.batch_mech_mode,
            "mechanism_value": st.session_state.batch_mech_value,
            "guaranteed_mode": st.session_state.batch_gua_mode,
            "guaranteed_value": st.session_state.batch_gua_value,
            "power_limit_rate": st.session_state.batch_limit_rate
        }
    st.success("✅ 已将当前参数同步到所有月份！")

# 2. 分月参数调整（单独修改某月份）
with st.expander("🔧 分月参数调整（单独修改）", expanded=False):
    # 选择要修改的月份
    selected_month = st.selectbox("选择要修改的月份", range(1, 13), key="month_param_sel")
    current_params = st.session_state.monthly_params[selected_month]  # 获取该月当前参数

    # 分月-机制电量
    st.write(f"##### {selected_month}月 · 机制电量")
    col_m1, col_m2 = st.columns([2, 1])
    with col_m1:
        mech_mode = st.selectbox(
            "输入模式", ["小时数", "比例(%)"],
            index=0 if current_params["mechanism_mode"] == "小时数" else 1,
            key=f"mech_mode_{selected_month}"
        )
    with col_m2:
        m_max = 100.0 if mech_mode == "比例(%)" else 1000000.0
        mech_val = st.number_input(
            "数值", min_value=0.0, max_value=m_max,
            value=current_params["mechanism_value"], step=0.1,
            key=f"mech_val_{selected_month}"
        )

    # 分月-保障性电量
    st.write(f"##### {selected_month}月 · 保障性电量")
    col_g1, col_g2 = st.columns([2, 1])
    with col_g1:
        gua_mode = st.selectbox(
            "输入模式", ["小时数", "比例(%)"],
            index=0 if current_params["guaranteed_mode"] == "小时数" else 1,
            key=f"gua_mode_{selected_month}"
        )
    with col_g2:
        g_max = 100.0 if gua_mode == "比例(%)" else 1000000.0
        gua_val = st.number_input(
            "数值", min_value=0.0, max_value=g_max,
            value=current_params["guaranteed_value"], step=0.1,
            key=f"gua_val_{selected_month}"
        )

    # 分月-限电率
    st.write(f"##### {selected_month}月 · 限电率")
    limit_rate = st.number_input(
        "限电率(%)", min_value=0.0, max_value=100.0,
        value=current_params["power_limit_rate"], step=0.1,
        key=f"limit_rate_{selected_month}"
    )

    # 保存分月参数
    if st.button(f"💾 保存{selected_month}月参数", key=f"save_{selected_month}_param"):
        st.session_state.monthly_params[selected_month] = {
            "mechanism_mode": mech_mode,
            "mechanism_value": mech_val,
            "guaranteed_mode": gua_mode,
            "guaranteed_value": gua_val,
            "power_limit_rate": limit_rate
        }
        st.success(f"✅ 已保存{selected_month}月的参数！")

    # 3. 所有月份参数预览表格
    st.divider()
    st.write("#### 所有月份参数预览")
    param_preview = []
    for month in range(1, 13):
        p = st.session_state.monthly_params[month]
        param_preview.append({
            "月份": f"{month}月",
            "机制电量": f"{p['mechanism_mode']} · {p['mechanism_value']:.2f}",
            "保障性电量": f"{p['guaranteed_mode']} · {p['guaranteed_value']:.2f}",
            "限电率": f"{p['power_limit_rate']:.2f}%"
        })
    preview_df = pd.DataFrame(param_preview)
    st.dataframe(preview_df, use_container_width=True, hide_index=True)

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

# 2. 生成年度双方案（重点修复：严格过滤无效数据）
with col_data2:
    if st.button("📝 生成年度双方案", use_container_width=True, type="primary", key="generate_annual_plan"):
        if not st.session_state.selected_months or not st.session_state.monthly_data:
            st.warning("⚠️ 请先导入/初始化月份数据并选择月份")
        elif st.session_state.installed_capacity <= 0:
            st.warning("⚠️ 请填写有效的装机容量（>0）")
        else:
            with st.spinner("🔄 正在计算年度双方案..."):
                try:
                    trade_typical = {}
                    trade_arbitrage = {}
                    market_hours = {}
                    gen_hours = {}
                    total_annual = 0.0
                    valid_calculated_months = []  # 记录成功计算的月份
                    
                    for month in st.session_state.selected_months:
                        # 计算核心参数（仅传2个参数，内部读取分月参数）
                        if st.session_state.auto_calculate:
                            gh, mh = calculate_core_params_monthly(month, st.session_state.installed_capacity)
                        else:
                            # 手动模式：发电小时数按分月参数计算，市场化小时数用手动输入
                            gh, _ = calculate_core_params_monthly(month, st.session_state.installed_capacity)
                            mh = st.session_state.manual_market_hours
                        
                        # 校验市场化小时数有效性
                        if mh <= 0:
                            st.warning(f"⚠️ 月份{month}市场化小时数为0，跳过该月份")
                            continue
                        
                        market_hours[month] = mh   
                        gen_hours[month] = gh
                        
                        # 方案一：典型曲线（校验返回结果）
                        typical_df, total_typical = calculate_trade_power_typical(month, mh, st.session_state.installed_capacity)
                        if typical_df is None or typical_df.empty or "方案一月度电量(MWh)" not in typical_df.columns:
                            st.error(f"❌ 月份{month}典型方案计算失败，跳过该月份")
                            continue
                        
                        # 方案二：光伏套利/风电直线（校验返回结果）
                        arbitrage_df = calculate_trade_power_arbitrage(month, total_typical, typical_df)
                        if arbitrage_df is None or arbitrage_df.empty or "方案二月度电量(MWh)" not in arbitrage_df.columns:
                            st.error(f"❌ 月份{month}方案二计算失败，跳过该月份")
                            continue
                        
                        # 只有两个方案都成功才存入会话状态
                        trade_typical[month] = typical_df
                        trade_arbitrage[month] = arbitrage_df
                        total_annual += total_typical
                        valid_calculated_months.append(month)
                    
                    # 只有有有效计算结果才更新会话状态
                    if valid_calculated_months:
                        st.session_state.trade_power_typical = trade_typical
                        st.session_state.trade_power_arbitrage = trade_arbitrage
                        st.session_state.market_hours = market_hours
                        st.session_state.gen_hours = gen_hours
                        st.session_state.total_annual_trade = total_annual
                        st.session_state.calculated = True
                        
                        st.success(
                            f"✅ 年度双方案生成成功！\n"
                            f"成功计算月份：{', '.join([f'{m}月' for m in valid_calculated_months])}\n"
                            f"年度总交易电量：{round(total_annual, 2)} MWh"
                        )
                    else:
                        st.error("❌ 所有选中月份的方案计算均失败，请检查基础数据和参数配置！")
                        st.session_state.calculated = False  # 标记为未计算成功
                    
                except Exception as e:
                    st.error(f"❌ 生成方案失败：{str(e)}")
                    st.session_state.calculated = False

# 3. 导出年度方案
with col_data3:
    if st.session_state.calculated and st.session_state.trade_power_typical:
        annual_output = export_annual_plan()
        if annual_output:  # 确保有有效数据才显示下载按钮
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
            help="请先生成有效的年度方案"
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

# 四、年度方案展示（重点修复：只展示有效月份）
if st.session_state.calculated and st.session_state.trade_power_typical:
    st.divider()
    st.header(f"📈 {st.session_state.current_year}年度方案展示（双方案对比）")
    
    # 过滤有效展示月份（存在于两个方案中且列名齐全）
    valid_display_months = [
        month for month in st.session_state.selected_months
        if month in st.session_state.trade_power_typical
        and month in st.session_state.trade_power_arbitrage
        and not st.session_state.trade_power_typical[month].empty
        and not st.session_state.trade_power_arbitrage[month].empty
        and "方案一月度电量(MWh)" in st.session_state.trade_power_typical[month].columns
        and "方案二月度电量(MWh)" in st.session_state.trade_power_arbitrage[month].columns
    ]
    
    if not valid_display_months:
        st.warning("⚠️ 无有效方案数据可展示，请重新生成方案")
    else:
        # 1. 年度汇总
        st.subheader("1. 年度汇总")
        summary_data = []
        scheme2_note = "套利曲线" if st.session_state.current_plant_type == "光伏" else "直线曲线"
        for month in valid_display_months:
            typical_df = st.session_state.trade_power_typical[month]
            arbitrage_df = st.session_state.trade_power_arbitrage[month]
            
            typical_total = typical_df["方案一月度电量(MWh)"].sum()
            arbitrage_total = arbitrage_df["方案二月度电量(MWh)"].sum()
            days = get_days_in_month(st.session_state.current_year, month)
            summary_data.append({
                "月份": f"{month}月",
                "月份天数": days,
                "市场化小时数": st.session_state.market_hours.get(month, 0.0),
                "预估发电小时数": st.session_state.gen_hours.get(month, 0.0),
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
            valid_display_months,
            key="view_month_select"
        )
        
        try:
            # 保持联动逻辑：实时读最新数据
            typical_df = st.session_state.trade_power_typical.get(view_month, pd.DataFrame())
            required_cols = ["时段", "方案一月度电量(MWh)"]
            if typical_df.empty or not all(col in typical_df.columns for col in required_cols):
                st.info("⚠️ 暂无有效方案一数据（数据为空或缺少必要列）")
                pass
            else:
                base_df = st.session_state.monthly_data.get(view_month, None)
                if base_df is None or base_df.empty:
                    st.info("⚠️ 缺少基础价格数据，仅展示交易量图表")
                    # 纯交易量交互式柱状图
                    import plotly.express as px
                    fig = px.bar(
                        typical_df,
                        x="时段",
                        y="方案一月度电量(MWh)",
                        title=f"{view_month}月 方案一交易量",
                        labels={"方案一月度电量(MWh)": "交易量（MWh）", "时段": "时段（点）"},
                        color_discrete_sequence=["#4299e1"],  # 柔和蓝色
                        height=350
                    )
                    # 视觉优化：去除背景网格、调整字体
                    fig.update_layout(
                        plot_bgcolor="white",
                        xaxis_showgrid=False,
                        yaxis_showgrid=True,
                        yaxis_gridcolor="#f0f0f0",
                        font=dict(family="Arial", size=11),
                        title_font=dict(size=13, weight="bold"),
                        margin=dict(l=10, r=10, t=30, b=10)  # 紧凑边距
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    # 准备数据（确保长度一致）
                    merged_data = typical_df[["时段", "方案一月度电量(MWh)"]].copy()
                    if len(base_df) >= 24:
                        merged_data["现货价格"] = base_df["现货价格(元/MWh)"].head(24).values
                        merged_data["中长期价格"] = base_df["中长期价格(元/MWh)"].head(24).values
                    else:
                        merged_data["现货价格"] = 0.0
                        merged_data["中长期价格"] = 0.0

                    # 用 Plotly 创建双轴交互式图表
                    import plotly.graph_objects as go
                    fig = go.Figure()

                    # 1. 交易量柱状图（左轴）
                    fig.add_trace(go.Bar(
                        x=merged_data["时段"],
                        y=merged_data["方案一月度电量(MWh)"],
                        name="方案一交易量",
                        yaxis="y1",
                        marker_color="#4299e1",  # 柔和蓝
                        opacity=0.8,
                        hovertemplate="时段：%{x}点<br>交易量：%{y:.2f} MWh<extra></extra>"
                    ))

                    # 2. 现货价格折线（右轴）
                    fig.add_trace(go.Scatter(
                        x=merged_data["时段"],
                        y=merged_data["现货价格"],
                        name="现货价格",
                        yaxis="y2",
                        mode="lines+markers",
                        line=dict(color="#9f7aea", width=2),  # 柔和紫
                        marker=dict(size=4),
                        hovertemplate="时段：%{x}点<br>现货价格：%{y:.2f} 元/MWh<extra></extra>"
                    ))

                    # 3. 中长期价格折线（右轴）
                    fig.add_trace(go.Scatter(
                        x=merged_data["时段"],
                        y=merged_data["中长期价格"],
                        name="中长期价格",
                        yaxis="y2",
                        mode="lines+markers",
                        line=dict(color="#38b2ac", width=2),  # 柔和青
                        marker=dict(size=4),
                        hovertemplate="时段：%{x}点<br>中长期价格：%{y:.2f} 元/MWh<extra></extra>"
                    ))

                    # 视觉+布局优化（核心！）
                    fig.update_layout(
                        # 标题
                        title=f"{view_month}月 方案一交易量与价格对比",
                        title_font=dict(size=13, weight="bold", family="Arial"),
                        title_x=0.5,  # 居中
                        # 背景
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        # 双轴设置
                        yaxis1=dict(
                            title="交易量（MWh）",
                            title_font=dict(color="#4299e1"),
                            tickfont=dict(color="#4299e1"),
                            gridcolor="#f0f0f0"  # 淡灰网格
                        ),
                        yaxis2=dict(
                            title="价格（元/MWh）",
                            title_font=dict(color="#9f7aea"),
                            tickfont=dict(color="#9f7aea"),
                            overlaying="y",
                            side="right",
                            gridcolor="rgba(0,0,0,0)"  # 隐藏右轴网格，避免重叠
                        ),
                        # 图例
                        legend=dict(
                            orientation="h",  # 水平排列
                            yanchor="bottom",
                            y=-0.2,  # 放在图表下方，不挡数据
                            xanchor="center",
                            x=0.5
                        ),
                        # 边距（紧凑不浪费空间）
                        margin=dict(l=20, r=20, t=30, b=60),
                        # x轴优化
                        xaxis=dict(
                            title="时段（点）",
                            tickmode="array",
                            tickvals=merged_data["时段"],  # 显示所有24时段
                            gridcolor="#f0f0f0"
                        )
                    )

                    # 在 Streamlit 中显示（支持交互）
                    st.plotly_chart(fig, use_container_width=True)

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
            # 保持联动逻辑：实时读最新数据
            arbitrage_df = st.session_state.trade_power_arbitrage.get(view_month, pd.DataFrame())
            required_cols = ["时段", "方案二月度电量(MWh)"]
            if arbitrage_df.empty or not all(col in arbitrage_df.columns for col in required_cols):
                st.info("⚠️ 暂无有效方案二数据（数据为空或缺少必要列）")
                pass
            else:
                base_df = st.session_state.monthly_data.get(view_month, None)
                if base_df is None or base_df.empty:
                    st.info("⚠️ 缺少基础价格数据，仅展示交易量图表")
                    import plotly.express as px
                    fig = px.bar(
                        arbitrage_df,
                        x="时段",
                        y="方案二月度电量(MWh)",
                        title=f"{view_month}月 方案二交易量",
                        labels={"方案二月度电量(MWh)": "交易量（MWh）", "时段": "时段（点）"},
                        color_discrete_sequence=["#e53e3e"],  # 柔和红
                        height=350
                    )
                    fig.update_layout(
                        plot_bgcolor="white",
                        xaxis_showgrid=False,
                        yaxis_showgrid=True,
                        yaxis_gridcolor="#f0f0f0",
                        font=dict(family="Arial", size=11),
                        title_font=dict(size=13, weight="bold"),
                        margin=dict(l=10, r=10, t=30, b=10)
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    # 准备数据
                    merged_data = arbitrage_df[["时段", "方案二月度电量(MWh)"]].copy()
                    if len(base_df) >= 24:
                        merged_data["现货价格"] = base_df["现货价格(元/MWh)"].head(24).values
                        merged_data["中长期价格"] = base_df["中长期价格(元/MWh)"].head(24).values
                    else:
                        merged_data["现货价格"] = 0.0
                        merged_data["中长期价格"] = 0.0

                    import plotly.graph_objects as go
                    fig = go.Figure()

                    # 交易量柱状图（左轴）
                    fig.add_trace(go.Bar(
                        x=merged_data["时段"],
                        y=merged_data["方案二月度电量(MWh)"],
                        name="方案二交易量",
                        yaxis="y1",
                        marker_color="#e53e3e",  # 柔和红
                        opacity=0.8,
                        hovertemplate="时段：%{x}点<br>交易量：%{y:.2f} MWh<extra></extra>"
                    ))

                    # 现货价格折线（右轴）
                    fig.add_trace(go.Scatter(
                        x=merged_data["时段"],
                        y=merged_data["现货价格"],
                        name="现货价格",
                        yaxis="y2",
                        mode="lines+markers",
                        line=dict(color="#9f7aea", width=2),
                        marker=dict(size=4),
                        hovertemplate="时段：%{x}点<br>现货价格：%{y:.2f} 元/MWh<extra></extra>"
                    ))

                    # 中长期价格折线（右轴）
                    fig.add_trace(go.Scatter(
                        x=merged_data["时段"],
                        y=merged_data["中长期价格"],
                        name="中长期价格",
                        yaxis="y2",
                        mode="lines+markers",
                        line=dict(color="#38b2ac", width=2),
                        marker=dict(size=4),
                        hovertemplate="时段：%{x}点<br>中长期价格：%{y:.2f} 元/MWh<extra></extra>"
                    ))

                    # 视觉优化（和方案一保持风格统一）
                    fig.update_layout(
                        title=f"{view_month}月 方案二交易量与价格对比",
                        title_font=dict(size=13, weight="bold", family="Arial"),
                        title_x=0.5,
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        yaxis1=dict(
                            title="交易量（MWh）",
                            title_font=dict(color="#e53e3e"),
                            tickfont=dict(color="#e53e3e"),
                            gridcolor="#f0f0f0"
                        ),
                        yaxis2=dict(
                            title="价格（元/MWh）",
                            title_font=dict(color="#9f7aea"),
                            tickfont=dict(color="#9f7aea"),
                            overlaying="y",
                            side="right",
                            gridcolor="rgba(0,0,0,0)"
                        ),
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=-0.2,
                            xanchor="center",
                            x=0.5
                        ),
                        margin=dict(l=20, r=20, t=30, b=60),
                        xaxis=dict(
                            title="时段（点）",
                            tickmode="array",
                            tickvals=merged_data["时段"],
                            gridcolor="#f0f0f0"
                        )
                    )

                    st.plotly_chart(fig, use_container_width=True)

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
else:
    if st.session_state.calculated and not st.session_state.trade_power_typical:
        st.warning("⚠️ 无有效方案数据，请重新生成方案")

# -------------------------- 方案电量手动调增调减（多时段+非实时同步） --------------------------
st.divider()
st.header("✏️ 方案电量手动调增调减（总量保持不变）")

# 初始化临时调整数据（按“月份+方案”区分，仅存储已应用/原始数据）
if "temp_adjust_data" not in st.session_state:
    st.session_state.temp_adjust_data = {}  # 结构：{("月份", "方案"): 已应用/原始DataFrame}
if "original_adjust_data" not in st.session_state:
    st.session_state.original_adjust_data = {}  # 存储原始数据，用于重置

if st.session_state.calculated and st.session_state.trade_power_typical:
    # 过滤有效调整月份
    valid_adjust_months = [
        month for month in st.session_state.selected_months
        if month in st.session_state.trade_power_typical
        and month in st.session_state.trade_power_arbitrage
        and not st.session_state.trade_power_typical[month].empty
        and not st.session_state.trade_power_arbitrage[month].empty
        and "方案一月度电量(MWh)" in st.session_state.trade_power_typical[month].columns
        and "方案二月度电量(MWh)" in st.session_state.trade_power_arbitrage[month].columns
    ]
    
    if not valid_adjust_months:
        st.warning("⚠️ 无有效方案数据可调整，请重新生成方案")
    else:
        # 1. 选择调整的月份和方案
        col_adj1, col_adj2 = st.columns(2)
        with col_adj1:
            adj_month = st.selectbox(
                "选择要调整的月份",
                valid_adjust_months,
                key="adj_month_select"
            )
        with col_adj2:
            adj_scheme = st.selectbox(
                "选择要调整的方案",
                ["方案一（典型曲线）", "方案二（套利/直线曲线）"],
                key="adj_scheme_select"
            )

        # 2. 获取对应方案的原始数据（绑定到“月份+方案”唯一键）
        data_key = (adj_month, adj_scheme)
        if adj_scheme == "方案一（典型曲线）":
            scheme_final_df = st.session_state.trade_power_typical.get(adj_month, None)
            scheme_col = "方案一月度电量(MWh)"
        else:
            scheme_final_df = st.session_state.trade_power_arbitrage.get(adj_month, None)
            scheme_col = "方案二月度电量(MWh)"
        base_df = st.session_state.monthly_data.get(adj_month, None)

        if scheme_final_df is None or scheme_final_df.empty or base_df is None or base_df.empty:
            st.warning("⚠️ 该月份方案数据缺失，请重新生成方案")
        else:
            avg_gen_list = base_df["平均发电量(MWh)"].tolist()
            avg_gen_total = sum(avg_gen_list)
            
            if avg_gen_total <= 0:
                st.error("❌ 该月份原始平均发电量总和为0，无法按权重分摊调整量")
            else:
                # 3. 初始化临时数据和原始数据（仅切换月份/方案时同步，不实时更新）
                if data_key not in st.session_state.original_adjust_data:
                    # 保存原始数据（用于重置，仅初始化1次）
                    st.session_state.original_adjust_data[data_key] = scheme_final_df.copy()
                    # 初始化临时数据为原始数据（未应用任何修改时）
                    st.session_state.temp_adjust_data[data_key] = scheme_final_df.copy()
                
                # 当前显示的临时数据（仅从session_state读取，不实时写入）
                temp_df = st.session_state.temp_adjust_data[data_key].copy()
                # 原始数据（用于对比修改和重置）
                original_df = st.session_state.original_adjust_data[data_key].copy()
                total_fixed = original_df[scheme_col].sum()  # 总量固定（以原始总量为准）

                # 4. 显示可编辑表格（编辑时仅在内存中修改，不实时同步到session_state）
                st.write(f"### {adj_scheme} - {adj_month}月电量调整（固定总量：{total_fixed:.2f} MWh）")
                st.caption(
                    "📌 支持多时段修改：可同时编辑任意多个时段 → 点击「应用调整」生效（刷新页面仅保留已应用数据）"
                )
                edit_temp_df = st.data_editor(
                    temp_df[["时段", "平均发电量(MWh)", "时段比重(%)", scheme_col]],
                    column_config={
                        "时段": st.column_config.NumberColumn("时段", disabled=True),
                        "平均发电量(MWh)": st.column_config.NumberColumn("原始平均发电量(MWh)", disabled=True),
                        "时段比重(%)": st.column_config.NumberColumn("时段比重(%)", disabled=True),
                        scheme_col: st.column_config.NumberColumn(
                            f"{scheme_col}（可编辑）",
                            min_value=0.0,
                            step=0.1,
                            format="%.2f",
                            help="可同时修改多个时段，点击「应用调整」后未修改时段自动分摊调整量"
                        )
                    },
                    use_container_width=True,
                    num_rows="fixed",
                    key=f"edit_adjust_scheme_{data_key}"
                )

                # 5. 应用+重置按钮（并排布局）
                col_apply, col_reset, col_empty = st.columns([1, 1, 8])
                with col_apply:
                    apply_adjust = st.button("应用调整", key=f"apply_adjust_{data_key}", type="primary")
                with col_reset:
                    reset_adjust = st.button("重置调整", key=f"reset_adjust_{data_key}")

                # 6. 重置按钮逻辑（恢复到原始数据，不保留未应用修改）
                if reset_adjust:
                    st.session_state.temp_adjust_data[data_key] = original_df.copy()
                    st.success(f"✅ 已重置为{adj_month}月{adj_scheme}原始数据！（未应用的修改已丢弃）")
                    st.rerun()

                # 7. 应用按钮逻辑（仅点击时同步数据，刷新页面后保留）
                if apply_adjust:
                    # 检测是否有修改
                    if edit_temp_df[scheme_col].equals(original_df[scheme_col]):
                        st.info("ℹ️ 未检测到任何修改，无需应用！")
                    else:
                        # 步骤1：识别修改时段和未修改时段
                        delta_series = edit_temp_df[scheme_col] - original_df[scheme_col]
                        modified_indices = delta_series[delta_series != 0].index.tolist()
                        unmodified_indices = [idx for idx in range(24) if idx not in modified_indices]

                        # 步骤2：计算总调整量
                        total_delta = delta_series.sum()

                        # 步骤3：边界处理1：所有时段都修改
                        if len(unmodified_indices) == 0:
                            modified_total = edit_temp_df[scheme_col].sum()
                            if np.isclose(modified_total, total_fixed, atol=0.01):
                                adjusted_df = edit_temp_df.copy()
                                adjusted_df["时段比重(%)"] = round(adjusted_df[scheme_col] / total_fixed * 100, 4)
                                st.success(
                                    f"✅ 调整成功！\n"
                                    f"- 修改时段数量：{len(modified_indices)}个（所有时段均修改）\n"
                                    f"- 总电量保持：{total_fixed:.2f} MWh"
                                )
                            else:
                                correction = total_fixed - modified_total
                                last_mod_idx = modified_indices[-1]
                                adjusted_df = edit_temp_df.copy()
                                adjusted_df.loc[last_mod_idx, scheme_col] = max(
                                    round(adjusted_df.loc[last_mod_idx, scheme_col] + correction, 2),
                                    0.0
                                )
                                adjusted_df["时段比重(%)"] = round(adjusted_df[scheme_col] / total_fixed * 100, 4)
                                st.success(
                                    f"✅ 调整成功（已自动修正总量）！\n"
                                    f"- 修改时段数量：{len(modified_indices)}个（所有时段均修改）\n"
                                    f"- 修正量：{correction:.2f} MWh（最后修改时段）\n"
                                    f"- 总电量保持：{total_fixed:.2f} MWh"
                                )

                        # 步骤4：边界处理2：未修改时段无发电量
                        else:
                            unmodified_avg_gen = [avg_gen_list[idx] for idx in unmodified_indices]
                            unmodified_avg_total = sum(unmodified_avg_gen)
                            
                            if unmodified_avg_total <= 0:
                                st.error("❌ 未修改时段的原始平均发电量总和为0，无法分摊调整量！请至少保留1个有发电量的时段不修改")
                            else:
                                # 步骤5：未修改时段分摊总调整量（仅unmodified_avg_total>0时执行）
                                adjusted_df = edit_temp_df.copy()
                                for idx in unmodified_indices:
                                    weight_ratio = avg_gen_list[idx] / unmodified_avg_total
                                    share_amount = -total_delta * weight_ratio
                                    new_val = adjusted_df.loc[idx, scheme_col] + share_amount
                                    adjusted_df.loc[idx, scheme_col] = max(round(new_val, 2), 0.0)

                                # 步骤6：修正浮点数误差
                                current_total = adjusted_df[scheme_col].sum()
                                if not np.isclose(current_total, total_fixed, atol=0.01):
                                    last_unmod_idx = unmodified_indices[-1]
                                    correction = total_fixed - current_total
                                    adjusted_df.loc[last_unmod_idx, scheme_col] = max(
                                        round(adjusted_df.loc[last_unmod_idx, scheme_col] + correction, 2),
                                        0.0
                                    )

                                # 步骤7：更新时段比重
                                adjusted_df["时段比重(%)"] = round(adjusted_df[scheme_col] / total_fixed * 100, 4)

                                # 步骤8：反馈结果
                                modified_hours = [str(adjusted_df.loc[idx, "时段"]) for idx in modified_indices]
                                st.success(
                                    f"✅ 调整成功！（刷新页面后保留此状态）\n"
                                    f"- 修改时段：{len(modified_indices)}个（{', '.join(modified_hours)}点）\n"
                                    f"- 总调整量：{total_delta:.2f} MWh\n"
                                    f"- 分摊方式：未修改的{len(unmodified_indices)}个时段按权重分摊\n"
                                    f"- 总电量保持：{total_fixed:.2f} MWh"
                                )

                                # 关键：仅应用时同步数据到session_state（最终数据+临时显示数据）
                                if adj_scheme == "方案一（典型曲线）":
                                    st.session_state.trade_power_typical[adj_month] = adjusted_df
                                else:
                                    st.session_state.trade_power_arbitrage[adj_month] = adjusted_df
                                # 更新临时显示数据（下次打开表格显示调整后的数据）
                                st.session_state.temp_adjust_data[data_key] = adjusted_df.copy()

else:
    st.warning("⚠️ 请先生成年度方案后再进行电量调整")

# 页脚
st.divider()
st.caption(f"© {st.session_state.current_year} 新能源电厂年度方案设计系统 | 双方案（典型/套利/直线）+ 四列日分解数据 | 总量一致")
