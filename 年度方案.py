import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# -------------------------- 必备：区域-省份映射字典（唯一定义） --------------------------
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

# -------------------------- 全局配置（唯一调用） --------------------------
st.set_page_config(
    page_title="新能源电厂年度方案设计系统",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------- Session State 完整初始化（含分月配置） --------------------------
if "initialized" not in st.session_state:
    # 基础信息配置
    st.session_state.initialized = True
    st.session_state.current_year = 2025
    st.session_state.current_region = "总部"
    st.session_state.current_province = REGIONS["总部"][0]
    st.session_state.current_power_plant = ""
    st.session_state.current_plant_type = "风电"
    
    # 业务数据存储
    st.session_state.monthly_data = {}  # 月度基础数据
    st.session_state.selected_months = []  # 选中的月份
    st.session_state.trade_power_typical = {}  # 方案一：典型曲线
    st.session_state.trade_power_arbitrage = {}  # 方案二：套利/直线曲线
    st.session_state.total_annual_trade = 0.0  # 年度总交易电量
    st.session_state.market_hours = {}  # 各月市场化小时数
    st.session_state.gen_hours = {}  # 各月发电小时数
    st.session_state.calculated = False  # 是否已生成方案
    
    # 市场化小时数配置
    st.session_state.auto_calculate = True  # 自动计算开关
    st.session_state.manual_market_hours = 0.0  # 手动输入值
    
    # 1. 分月电量参数配置（1-12月独立）
    st.session_state.monthly_params = {
        month: {
            "mechanism_mode": "小时数",
            "mechanism_value": 0.0,
            "guaranteed_mode": "小时数",
            "guaranteed_value": 0.0,
            "power_limit_rate": 0.0
        } for month in range(1, 13)
    }
    
    # 2. 分月光伏套利曲线配置（1-12月独立，支持季节变化）
    st.session_state.monthly_pv_params = {
        month: {
            "core_start": 11,   # 核心时段起始（默认）
            "core_end": 14,     # 核心时段结束（默认）
            "edge_start": 6,    # 边缘时段起始（默认）
            "edge_end": 18      # 边缘时段结束（默认）
        } for month in range(1, 13)
    }
    
    # 批量应用参数（电量+光伏）
    st.session_state.batch_mech_mode = "小时数"
    st.session_state.batch_mech_value = 0.0
    st.session_state.batch_gua_mode = "小时数"
    st.session_state.batch_gua_value = 0.0
    st.session_state.batch_limit_rate = 0.0
    st.session_state.batch_pv_core_start = 11
    st.session_state.batch_pv_core_end = 14
    st.session_state.batch_pv_edge_start = 6
    st.session_state.batch_pv_edge_end = 18

# -------------------------- 核心工具函数（完善分月逻辑+兜底） --------------------------
def get_days_in_month(year, month):
    """根据年份和月份获取天数（处理闰年）"""
    if month == 2:
        return 29 if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0) else 28
    elif month in [4, 6, 9, 11]:
        return 30
    else:
        return 31

def get_pv_arbitrage_hours(month):
    """获取指定月份的光伏套利曲线时段划分（支持分月配置）"""
    # 安全读取该月份的光伏配置（兜底默认值）
    pv_params = st.session_state.monthly_pv_params.get(month, {
        "core_start": 11,
        "core_end": 14,
        "edge_start": 6,
        "edge_end": 18
    })
    # 有效性校验（确保在1-24之间）
    core_start = max(1, min(24, int(pv_params["core_start"])))
    core_end = max(1, min(24, int(pv_params["core_end"])))
    edge_start = max(1, min(24, int(pv_params["edge_start"])))
    edge_end = max(1, min(24, int(pv_params["edge_end"])))
    
    # 处理起始>结束的情况
    if core_start > core_end:
        core_start, core_end = core_end, core_start
    if edge_start > edge_end:
        edge_start, edge_end = edge_end, edge_start
    
    # 计算各时段
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
    """初始化单个月份的模板数据（安全读取Session State）"""
    hours = list(range(1, 25))
    return pd.DataFrame({
        "时段": hours,
        "平均发电量(MWh)": [0.0]*24,
        "当月各时段累计发电量(MWh)": [0.0]*24,
        "现货价格(元/MWh)": [0.0]*24,
        "中长期价格(元/MWh)": [0.0]*24,
        "年份": st.session_state.get("current_year", 2025),
        "月份": month,
        "电厂名称": st.session_state.get("current_power_plant", ""),
        "电厂类型": st.session_state.get("current_plant_type", "风电"),
        "区域": st.session_state.get("current_region", "总部"),
        "省份": st.session_state.get("current_province", "北京")
    })

def export_basic_template():
    """导出基础数据模板（含12个月基础数据）"""
    wb = Workbook()
    wb.remove(wb.active)
    for month in range(1, 13):
        ws = wb.create_sheet(title=f"{month}月基础数据")
        template_df = init_month_template(month)
        for r in dataframe_to_rows(template_df, index=False, header=True):
            ws.append(r)
    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def batch_import_basic_data(file):
    """批量导入基础数据（Excel）"""
    monthly_data = {}
    try:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            if not sheet_name.endswith("基础数据"):
                st.warning(f"跳过无效子表：{sheet_name}（需命名为“1月基础数据”-“12月基础数据”）")
                continue
            try:
                month = int(sheet_name.replace("月基础数据", ""))
                if month < 1 or month > 12:
                    st.warning(f"跳过无效月份子表：{sheet_name}（需1-12月）")
                    continue
                df = pd.read_excel(file, sheet_name=sheet_name)
                required_cols = ["时段", "平均发电量(MWh)", "当月各时段累计发电量(MWh)", "现货价格(元/MWh)", "中长期价格(元/MWh)"]
                if not all(col in df.columns for col in required_cols):
                    st.warning(f"子表{sheet_name}缺少必要列（需包含{', '.join(required_cols)}），跳过")
                    continue
                df["年份"] = st.session_state.get("current_year", 2025)
                df["电厂名称"] = st.session_state.get("current_power_plant", "")
                df["电厂类型"] = st.session_state.get("current_plant_type", "风电")
                df["区域"] = st.session_state.get("current_region", "总部")
                df["省份"] = st.session_state.get("current_province", "北京")
                monthly_data[month] = df
            except Exception as e:
                st.warning(f"处理子表{sheet_name}失败：{str(e)}")
        return monthly_data
    except Exception as e:
        st.error(f"批量导入基础数据失败：{str(e)}")
        return None

def export_config_template():
    """导出配置模板（含分月电量参数+分月光伏配置）"""
    wb = Workbook()
    
    # 1. 分月电量参数表
    ws_power = wb.active
    ws_power.title = "分月电量参数"
    power_data = []
    for month in range(1, 13):
        params = st.session_state.monthly_params[month]
        power_data.append({
            "月份": month,
            "机制电量模式": params["mechanism_mode"],
            "机制电量数值": params["mechanism_value"],
            "保障性电量模式": params["guaranteed_mode"],
            "保障性电量数值": params["guaranteed_value"],
            "限电率(%)": params["power_limit_rate"]
        })
    power_df = pd.DataFrame(power_data)
    for r in dataframe_to_rows(power_df, index=False, header=True):
        ws_power.append(r)
    
    # 2. 分月光伏配置表（仅光伏电厂显示相关字段）
    ws_pv = wb.create_sheet(title="分月光伏配置")
    pv_data = []
    for month in range(1, 13):
        params = st.session_state.monthly_pv_params[month]
        pv_data.append({
            "月份": month,
            "核心起始（点）": params["core_start"],
            "核心结束（点）": params["core_end"],
            "边缘起始（点）": params["edge_start"],
            "边缘结束（点）": params["edge_end"]
        })
    pv_df = pd.DataFrame(pv_data)
    for r in dataframe_to_rows(pv_df, index=False, header=True):
        ws_pv.append(r)
    
    # 导出文件
    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def import_config(file):
    """导入配置模板（同步分月电量参数+分月光伏配置）"""
    try:
        xls = pd.ExcelFile(file)
        success_count = 0
        
        # 1. 导入分月电量参数
        if "分月电量参数" in xls.sheet_names:
            power_df = pd.read_excel(file, sheet_name="分月电量参数")
            required_cols = ["月份", "机制电量模式", "机制电量数值", "保障性电量模式", "保障性电量数值", "限电率(%)"]
            if all(col in power_df.columns for col in required_cols):
                for _, row in power_df.iterrows():
                    month = int(row["月份"])
                    if 1 <= month <= 12:
                        st.session_state.monthly_params[month] = {
                            "mechanism_mode": row["机制电量模式"] if row["机制电量模式"] in ["小时数", "比例(%)"] else "小时数",
                            "mechanism_value": float(row["机制电量数值"]) if pd.notna(row["机制电量数值"]) else 0.0,
                            "guaranteed_mode": row["保障性电量模式"] if row["保障性电量模式"] in ["小时数", "比例(%)"] else "小时数",
                            "guaranteed_value": float(row["保障性电量数值"]) if pd.notna(row["保障性电量数值"]) else 0.0,
                            "power_limit_rate": float(row["限电率(%)"]) if pd.notna(row["限电率(%)"]) else 0.0
                        }
                success_count += 1
            else:
                st.warning("分月电量参数表缺少必要列，跳过")
        
        # 2. 导入分月光伏配置
        if "分月光伏配置" in xls.sheet_names:
            pv_df = pd.read_excel(file, sheet_name="分月光伏配置")
            required_cols = ["月份", "核心起始（点）", "核心结束（点）", "边缘起始（点）", "边缘结束（点）"]
            if all(col in pv_df.columns for col in required_cols):
                for _, row in pv_df.iterrows():
                    month = int(row["月份"])
                    if 1 <= month <= 12:
                        st.session_state.monthly_pv_params[month] = {
                            "core_start": int(row["核心起始（点）"]) if pd.notna(row["核心起始（点）"]) else 11,
                            "core_end": int(row["核心结束（点）"]) if pd.notna(row["核心结束（点）"]) else 14,
                            "edge_start": int(row["边缘起始（点）"]) if pd.notna(row["边缘起始（点）"]) else 6,
                            "edge_end": int(row["边缘结束（点）"]) if pd.notna(row["边缘结束（点）"]) else 18
                        }
                success_count += 1
            else:
                st.warning("分月光伏配置表缺少必要列，跳过")
        
        if success_count > 0:
            st.success(f"✅ 配置导入成功！共导入{success_count}类配置")
        else:
            st.warning("⚠️ 未导入任何有效配置，请检查模板格式")
        return True
    except Exception as e:
        st.error(f"配置导入失败：{str(e)}")
        return False

def calculate_core_params_monthly(month, installed_capacity):
    """按月份计算核心参数（市场化小时数、发电小时数）"""
    if month not in st.session_state.monthly_data:
        return 0.0, 0.0
    
    # 安全读取分月电量参数
    params = st.session_state.monthly_params.get(month, {
        "mechanism_mode": "小时数",
        "mechanism_value": 0.0,
        "guaranteed_mode": "小时数",
        "guaranteed_value": 0.0,
        "power_limit_rate": 0.0
    })

    df = st.session_state.monthly_data[month]
    total_generation = df["当月各时段累计发电量(MWh)"].sum()
    gen_hours = round(total_generation / installed_capacity, 2) if installed_capacity > 0 else 0.0
    
    if gen_hours <= 0:
        market_hours = 0.0
    else:
        available_hours = gen_hours * (1 - params["power_limit_rate"] / 100)
        # 扣减机制电量
        if params["mechanism_mode"] == "小时数":
            available_hours -= params["mechanism_value"]
        else:
            available_hours -= gen_hours * (params["mechanism_value"] / 100)
        # 扣减保障性电量
        if params["guaranteed_mode"] == "小时数":
            available_hours -= params["guaranteed_value"]
        else:
            available_hours -= gen_hours * (params["guaranteed_value"] / 100)
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
    trade_df["年份"] = st.session_state.get("current_year", 2025)
    trade_df["月份"] = month
    trade_df["电厂名称"] = st.session_state.get("current_power_plant", "")
    trade_df = trade_df.fillna(0.0)
    trade_df["方案一月度电量(MWh)"] = trade_df["方案一月度电量(MWh)"].astype(np.float64)
    return trade_df, round(total_trade_power, 2)

def calculate_trade_power_arbitrage(month, total_trade_power, typical_df):
    """方案二：光伏套利曲线（分月配置）/风电直线曲线"""
    if month not in st.session_state.monthly_data or typical_df is None:
        return None
    
    plant_type = st.session_state.get("current_plant_type", "风电")
    if plant_type == "光伏":
        # 光伏：使用该月份的分月光伏配置
        pv_hours = get_pv_arbitrage_hours(month)
        core_hours = pv_hours["core"]
        edge_hours = pv_hours["edge"]
        invalid_hours = pv_hours["invalid"]
        
        # 计算边缘时段总电量（要转移的电量）
        edge_total = typical_df[typical_df["时段"].isin(edge_hours)]["方案一月度电量(MWh)"].sum()
        core_count = len(core_hours) if len(core_hours) > 0 else 1
        core_add = edge_total / core_count
        
        trade_data = []
        for idx, row in typical_df.iterrows():
            hour = row["时段"]
            avg_gen = row["平均发电量(MWh)"]
            
            if hour in invalid_hours or hour in edge_hours:
                trade_power = 0.0
                proportion = 0.0
            elif hour in core_hours:
                trade_power = row["方案一月度电量(MWh)"] + core_add
                proportion = trade_power / total_trade_power
            else:
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
        # 风电：24小时平均分配
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
    
    # 数据补充和校验（确保总量一致）
    trade_df["年份"] = st.session_state.get("current_year", 2025)
    trade_df["月份"] = month
    trade_df["电厂名称"] = st.session_state.get("current_power_plant", "")
    trade_df = trade_df.fillna(0.0)
    trade_df["方案二月度电量(MWh)"] = trade_df["方案二月度电量(MWh)"].astype(np.float64)
    
    # 修正浮点数精度误差
    total_current = trade_df["方案二月度电量(MWh)"].sum()
    if total_current > 0:
        trade_df["方案二月度电量(MWh)"] = trade_df["方案二月度电量(MWh)"] * (total_trade_power / total_current)
    
    return trade_df

def decompose_double_scheme(typical_df, arbitrage_df, year, month):
    """双方案日分解（返回四列数据）"""
    days = get_days_in_month(year, month)
    df = pd.DataFrame({
        "时段": typical_df["时段"],
        "方案一月度电量(MWh)": typical_df["方案一月度电量(MWh)"],
        "方案一日分解电量(MWh)": round(typical_df["方案一月度电量(MWh)"] / days, 4),
        "方案二月度电量(MWh)": arbitrage_df["方案二月度电量(MWh)"],
        "方案二日分解电量(MWh)": round(arbitrage_df["方案二月度电量(MWh)"] / days, 4),
        "月份天数": days
    })
    return df.fillna(0.0)

def export_annual_plan():
    """导出年度方案Excel（含基础数据、双方案、配置信息）"""
    wb = Workbook()
    wb.remove(wb.active)
    total_annual = 0.0
    plant_type = st.session_state.get("current_plant_type", "风电")
    
    # 1. 年度汇总表（含分月光伏配置）
    summary_data = []
    scheme2_note = "套利曲线（分月配置）" if plant_type == "光伏" else "直线曲线（24小时平均）"
    
    for month in st.session_state.selected_months:
        if month not in st.session_state.trade_power_typical or month not in st.session_state.trade_power_arbitrage:
            continue
        typical_df = st.session_state.trade_power_typical[month]
        arbitrage_df = st.session_state.trade_power_arbitrage[month]
        total_typical = typical_df["方案一月度电量(MWh)"].sum()
        total_arbitrage = arbitrage_df["方案二月度电量(MWh)"].sum()
        total_annual += total_typical
        
        # 光伏配置信息（非光伏电厂显示"-"）
        if plant_type == "光伏":
            pv_config = get_pv_arbitrage_hours(month)["config"]
            pv_core = f"{pv_config['core_start']}-{pv_config['core_end']}点"
            pv_edge = f"{pv_config['edge_start']}-{pv_config['edge_end']}点"
        else:
            pv_core = "-"
            pv_edge = "-"
        
        summary_data.append({
            "年份": st.session_state.get("current_year", 2025),
            "月份": month,
            "电厂名称": st.session_state.get("current_power_plant", ""),
            "电厂类型": plant_type,
            "光伏核心时段": pv_core,
            "光伏边缘时段": pv_edge,
            "方案一（典型曲线）总电量(MWh)": total_typical,
            "方案二（{}）总电量(MWh)".format(scheme2_note): total_arbitrage,
            "月份天数": get_days_in_month(st.session_state.get("current_year", 2025), month),
            "市场化小时数": st.session_state.market_hours.get(month, 0.0),
            "预估发电小时数": st.session_state.gen_hours.get(month, 0.0),
            "占年度比重(%)": round(total_typical / st.session_state.total_annual_trade * 100, 2) if st.session_state.total_annual_trade > 0 else 0.0
        })
    
    # 写入汇总表
    summary_df = pd.DataFrame(summary_data)
    ws_summary = wb.create_sheet(title="年度汇总")
    for r in dataframe_to_rows(summary_df, index=False, header=True):
        ws_summary.append(r)
    
    # 2. 各月详情表（双方案+日分解）
    for month in st.session_state.selected_months:
        if month not in st.session_state.monthly_data:
            continue
        base_df = st.session_state.monthly_data[month][["时段", "平均发电量(MWh)", "现货价格(元/MWh)", "中长期价格(元/MWh)"]].copy()
        typical_df = st.session_state.trade_power_typical[month][["时段", "方案一月度电量(MWh)", "时段比重(%)"]].copy()
        typical_df.rename(columns={"时段比重(%)": "方案一时段比重(%)"}, inplace=True)
        arbitrage_df = st.session_state.trade_power_arbitrage[month][["时段", "方案二月度电量(MWh)", "时段比重(%)"]].copy()
        arbitrage_df.rename(columns={"时段比重(%)": "方案二时段比重(%)"}, inplace=True)
        decompose_df = decompose_double_scheme(
            st.session_state.trade_power_typical[month],
            st.session_state.trade_power_arbitrage[month],
            st.session_state.get("current_year", 2025),
            month
        )[["时段", "方案一日分解电量(MWh)", "方案二日分解电量(MWh)", "月份天数"]].copy()
        
        # 合并数据
        merged_df = base_df.merge(typical_df, on="时段")
        merged_df = merged_df.merge(arbitrage_df, on="时段")
        merged_df = merged_df.merge(decompose_df, on="时段")
        
        # 创建子表
        ws_month = wb.create_sheet(title=f"{month}月详情")
        for r in dataframe_to_rows(merged_df, index=False, header=True):
            ws_month.append(r)
    
    # 3. 配置信息表（电量参数+光伏配置）
    ws_config = wb.create_sheet(title="配置信息")
    config_data = []
    for month in range(1, 13):
        power_params = st.session_state.monthly_params[month]
        pv_params = st.session_state.monthly_pv_params[month]
        config_data.append({
            "月份": month,
            "机制电量（模式-数值）": f"{power_params['mechanism_mode']}-{power_params['mechanism_value']:.2f}",
            "保障性电量（模式-数值）": f"{power_params['guaranteed_mode']}-{power_params['guaranteed_value']:.2f}",
            "限电率(%)": power_params['power_limit_rate'],
            "光伏核心时段": f"{pv_params['core_start']}-{pv_params['core_end']}点" if plant_type == "光伏" else "-",
            "光伏边缘时段": f"{pv_params['edge_start']}-{pv_params['edge_end']}点" if plant_type == "光伏" else "-"
        })
    config_df = pd.DataFrame(config_data)
    for r in dataframe_to_rows(config_df, index=False, header=True):
        ws_config.append(r)
    
    # 4. 方案说明表
    ws_desc = wb.create_sheet(title="方案说明")
    desc_content = [
        ["新能源电厂年度交易方案说明"],
        [""],
        ["基础信息："],
        [f"电厂名称：{st.session_state.get('current_power_plant', '')}"],
        [f"电厂类型：{plant_type}"],
        [f"年份：{st.session_state.get('current_year', 2025)}"],
        [f"区域：{st.session_state.get('current_region', '总部')}"],
        [f"省份：{st.session_state.get('current_province', '北京')}"],
        [f"装机容量：{st.session_state.get('sidebar_installed_capacity', 0.0)} MW"],
        [""],
        ["方案说明："],
        ["方案一（典型曲线）：按各时段平均发电量权重分配市场化交易电量"],
        [f"方案二（{scheme2_note}）：总电量与方案一一致，按配置规则分配时段电量"],
        [""],
        [f"年度总交易电量（典型方案）：{round(total_annual, 2)} MWh"]
    ]
    for row in desc_content:
        ws_desc.append(row)
    
    # 导出文件
    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -------------------------- 侧边栏配置（仅保留基础信息，精简布局） --------------------------
with st.sidebar:
    st.header("⚙️ 基础信息配置")
    
    # 1. 年份选择
    years = list(range(2020, 2031))
    current_year = st.session_state.get("current_year", 2025)
    current_year = current_year if current_year in years else 2025
    st.session_state.current_year = st.selectbox(
        "选择年份", years, index=years.index(current_year), key="sidebar_year"
    )
    
    # 2. 区域/省份选择
    current_region = st.session_state.get("current_region", "总部")
    current_region = current_region if current_region in REGIONS.keys() else "总部"
    selected_region = st.selectbox(
        "选择区域", list(REGIONS.keys()), index=list(REGIONS.keys()).index(current_region), key="sidebar_region_select"
    )
    st.session_state.current_region = selected_region
    
    provinces = REGIONS[selected_region]
    current_province = st.session_state.get("current_province", provinces[0])
    current_province = current_province if current_province in provinces else provinces[0]
    selected_province = st.selectbox(
        "选择省份", provinces, index=provinces.index(current_province), key="sidebar_province_select"
    )
    st.session_state.current_province = selected_province
    
    # 3. 电厂信息
    st.session_state.current_power_plant = st.text_input(
        "电厂名称", value=st.session_state.get("current_power_plant", ""), key="sidebar_power_plant"
    )
    plant_types = ["风电", "光伏", "水光互补", "风光互补"]
    current_plant_type = st.session_state.get("current_plant_type", "风电")
    current_plant_type = current_plant_type if current_plant_type in plant_types else "风电"
    st.session_state.current_plant_type = st.selectbox(
        "电厂类型", plant_types, index=plant_types.index(current_plant_type), key="sidebar_plant_type"
    )
    
    # 4. 装机容量
    installed_capacity = st.number_input(
        "装机容量(MW)", min_value=0.0, value=0.0, step=0.1,
        key="sidebar_installed_capacity", help="电厂总装机容量，单位：兆瓦"
    )
    st.session_state.sidebar_installed_capacity = installed_capacity
    
    # 5. 市场化交易小时数配置
    st.write("#### 市场化交易小时数")
    auto_calculate = st.toggle(
        "自动计算", value=st.session_state.get("auto_calculate", True), key="sidebar_auto_calculate"
    )
    st.session_state.auto_calculate = auto_calculate
    
    if not auto_calculate:
        manual_market_hours = st.number_input(
            "手动输入（适用于所有选中月份）", min_value=0.0, max_value=1000000.0,
            value=st.session_state.manual_market_hours, step=0.1, key="sidebar_market_hours_manual"
        )
        st.session_state.manual_market_hours = manual_market_hours

# -------------------------- 主页面布局（左右分栏：左侧数据操作，右侧分月配置） --------------------------
st.title("⚡ 新能源电厂年度方案设计系统")
plant_type = st.session_state.get("current_plant_type", "风电")
scheme2_title = "套利曲线（分月配置）" if plant_type == "光伏" else "直线曲线（24小时平均）"
st.subheader(
    f"当前配置：{st.session_state.current_year}年 | {st.session_state.current_region} | {st.session_state.current_province} | "
    f"{plant_type} | {st.session_state.current_power_plant}"
)
st.caption(f"方案一：典型出力曲线 | 方案二：{scheme2_title}")

# 主页面分栏（左侧60%：数据操作+方案展示；右侧40%：分月配置+导入导出）
col_left, col_right = st.columns([3, 2])

# -------------------------- 左侧栏：数据操作+方案展示 --------------------------
with col_left:
    # 一、模板导出与批量导入（基础数据）
    st.divider()
    st.header("📤 基础数据导入导出")
    col_import1, col_import2, col_import3 = st.columns(3)
    
    # 1. 导出基础数据模板
    with col_import1:
        basic_template_output = export_basic_template()
        st.download_button(
            "📥 导出基础数据模板",
            data=basic_template_output,
            file_name=f"{st.session_state.current_power_plant}_{st.session_state.current_year}年基础数据模板.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    # 2. 批量导入基础数据
    with col_import2:
        basic_file = st.file_uploader(
            "📥 批量导入基础数据", type=["xlsx"], key="basic_data_uploader"
        )
        if basic_file is not None:
            monthly_data = batch_import_basic_data(basic_file)
            if monthly_data:
                st.session_state.monthly_data = monthly_data
                st.session_state.selected_months = sorted(list(monthly_data.keys()))
                st.success(f"✅ 基础数据导入成功！共导入{len(monthly_data)}个月份")
    
    # 3. 月份多选
    with col_import3:
        st.session_state.selected_months = st.multiselect(
            "选择处理月份", list(range(1, 13)), default=st.session_state.selected_months, key="month_multiselect"
        )
        if st.session_state.selected_months:
            st.info(f"当前选中：{', '.join([f'{m}月' for m in st.session_state.selected_months])}")
        else:
            st.warning("⚠️ 请先选择需要处理的月份")
    
    # 二、数据操作按钮
    st.divider()
    st.header("🔧 数据操作")
    col_data1, col_data2, col_data3 = st.columns(3)
    
    # 1. 初始化选中月份基础数据
    with col_data1:
        if st.button("📋 初始化基础数据", use_container_width=True, key="init_basic_data"):
            if not st.session_state.selected_months:
                st.warning("⚠️ 请先选择月份")
            else:
                for month in st.session_state.selected_months:
                    st.session_state.monthly_data[month] = init_month_template(month)
                st.success(f"✅ 已初始化{len(st.session_state.selected_months)}个月份基础数据")
    
    # 2. 生成年度双方案
    with col_data2:
        if st.button("📝 生成年度方案", use_container_width=True, type="primary", key="generate_annual_plan"):
            if not st.session_state.selected_months or not st.session_state.monthly_data:
                st.warning("⚠️ 请先导入/初始化基础数据并选择月份")
            elif installed_capacity <= 0:
                st.warning("⚠️ 请填写装机容量")
            else:
                with st.spinner("🔄 正在计算年度方案..."):
                    try:
                        trade_typical = {}
                        trade_arbitrage = {}
                        market_hours = {}
                        gen_hours = {}
                        total_annual = 0.0
                        
                        for month in st.session_state.selected_months:
                            # 计算核心参数
                            if st.session_state.auto_calculate:
                                gh, mh = calculate_core_params_monthly(month, installed_capacity)
                            else:
                                gh = calculate_core_params_monthly(month, installed_capacity)[0]
                                mh = st.session_state.manual_market_hours
                            
                            market_hours[month] = mh
                            gen_hours[month] = gh
                            
                            # 方案一计算
                            typical_df, total_typical = calculate_trade_power_typical(month, mh, installed_capacity)
                            if typical_df is None:
                                st.error(f"❌ 月份{month}典型方案计算失败")
                                continue
                            trade_typical[month] = typical_df
                            total_annual += total_typical
                            
                            # 方案二计算（传入月份，支持分月光伏配置）
                            arbitrage_df = calculate_trade_power_arbitrage(month, total_typical, typical_df)
                            if arbitrage_df is None:
                                st.error(f"❌ 月份{month}方案二计算失败")
                                continue
                            trade_arbitrage[month] = arbitrage_df
                        
                        # 保存结果
                        st.session_state.trade_power_typical = trade_typical
                        st.session_state.trade_power_arbitrage = trade_arbitrage
                        st.session_state.market_hours = market_hours
                        st.session_state.gen_hours = gen_hours
                        st.session_state.total_annual_trade = total_annual
                        st.session_state.calculated = True
                        
                        st.success(
                            f"✅ 年度方案生成成功！\n"
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
                "💾 导出年度方案",
                data=annual_output,
                file_name=f"{st.session_state.current_power_plant}_{st.session_state.current_year}年双方案交易数据.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.button(
                "💾 导出年度方案",
                use_container_width=True, disabled=True, help="请先生成年度方案"
            )
    
    # 三、选中月份基础数据编辑
    if st.session_state.selected_months and st.session_state.monthly_data:
        st.divider()
        st.header("📊 基础数据编辑")
        edit_month = st.selectbox(
            "选择编辑月份", st.session_state.selected_months, key="edit_month_select"
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
                use_container_width=True, num_rows="fixed", key=f"edit_df_{edit_month}"
            )
            st.session_state.monthly_data[edit_month] = edit_df
    
    # 四、年度方案展示
    if st.session_state.calculated and st.session_state.selected_months:
        st.divider()
        st.header(f"📈 {st.session_state.current_year}年度方案展示")
        
        # 1. 年度汇总
        st.subheader("1. 年度汇总")
        summary_data = []
        scheme2_note = "套利曲线（分月配置）" if plant_type == "光伏" else "直线曲线"
        for month in st.session_state.selected_months:
            typical_total = st.session_state.trade_power_typical[month]["方案一月度电量(MWh)"].sum()
            arbitrage_total = st.session_state.trade_power_arbitrage[month]["方案二月度电量(MWh)"].sum()
            days = get_days_in_month(st.session_state.current_year, month)
            
            # 光伏配置信息
            if plant_type == "光伏":
                pv_config = get_pv_arbitrage_hours(month)["config"]
                pv_core = f"{pv_config['core_start']}-{pv_config['core_end']}点"
            else:
                pv_core = "-"
            
            summary_data.append({
                "月份": f"{month}月",
                "天数": days,
                "市场化小时数": st.session_state.market_hours.get(month, 0.0),
                "发电小时数": st.session_state.gen_hours.get(month, 0.0),
                "方案一总电量(MWh)": typical_total,
                "方案二总电量(MWh)": arbitrage_total,
                "光伏核心时段": pv_core,
                "占年度比重(%)": round(typical_total / st.session_state.total_annual_trade * 100, 2) if st.session_state.total_annual_trade > 0 else 0.0
            })
        summary_df = pd.DataFrame(summary_data)
        st.dataframe(summary_df, use_container_width=True, hide_index=True)
        st.metric("年度总交易电量", f"{st.session_state.total_annual_trade:.2f} MWh")
        
        # 2. 月份方案详情
        st.subheader("2. 月份方案详情")
        view_month = st.selectbox(
            "选择查看月份", st.session_state.selected_months, key="view_month_select"
        )
        
        # 方案一展示
        st.write(f"### 方案一：典型出力曲线（{view_month}月）")
        typical_df = st.session_state.trade_power_typical[view_month][["时段", "平均发电量(MWh)", "时段比重(%)", "方案一月度电量(MWh)"]].copy()
        typical_df = typical_df.fillna(0.0).reset_index(drop=True)
        st.dataframe(typical_df, use_container_width=True, hide_index=True)
        
        # 方案一图表
        try:
            chart_data = typical_df[["时段", "方案一月度电量(MWh)"]].set_index("时段")
            if not chart_data.empty and chart_data["方案一月度电量(MWh)"].sum() > 0:
                st.bar_chart(chart_data, use_container_width=True)
            else:
                st.info("⚠️ 暂无有效数据生成图表")
        except Exception as e:
            st.warning(f"📊 方案一图表生成失败：{str(e)}")
        
        # 方案二展示
        st.write(f"### 方案二：{scheme2_note}（{view_month}月）")
        arbitrage_df = st.session_state.trade_power_arbitrage[view_month][["时段", "平均发电量(MWh)", "时段比重(%)", "方案二月度电量(MWh)"]].copy()
        arbitrage_df = arbitrage_df.fillna(0.0).reset_index(drop=True)
        st.dataframe(arbitrage_df, use_container_width=True, hide_index=True)
        
        # 方案二说明（含分月光伏配置）
        if plant_type == "光伏":
            pv_hours = get_pv_arbitrage_hours(view_month)
            edge_total = typical_df[typical_df["时段"].isin(pv_hours["edge"])]["方案一月度电量(MWh)"].sum()
            core_avg_add = edge_total / len(pv_hours["core"]) if len(pv_hours["core"]) > 0 else 0
            st.info(f"""
            光伏套利曲线（{view_month}月配置）：
            - 核心时段（接收）：{pv_hours['core']}点
            - 边缘时段（转出）：{pv_hours['edge']}点
            - 转出总电量：{edge_total:.2f} MWh
            - 核心时段每小时增加：{core_avg_add:.2f} MWh
            - 总电量：{arbitrage_df['方案二月度电量(MWh)'].sum():.2f} MWh（与方案一一致）
            """)
        else:
            hourly_trade = arbitrage_df["方案二月度电量(MWh)"].iloc[0] if not arbitrage_df.empty else 0.0
            st.info(f"""
            风电直线曲线说明：
            - 24时段平均分配，每时段电量：{hourly_trade:.2f} MWh
            - 总电量：{arbitrage_df['方案二月度电量(MWh)'].sum():.2f} MWh（与方案一一致）
            """)
        
        # 方案二图表
        try:
            chart_data = arbitrage_df[["时段", "方案二月度电量(MWh)"]].set_index("时段")
            if not chart_data.empty and chart_data["方案二月度电量(MWh)"].sum() > 0:
                st.bar_chart(chart_data, use_container_width=True)
            else:
                st.info("⚠️ 暂无有效数据生成图表")
        except Exception as e:
            st.warning(f"📊 方案二图表生成失败：{str(e)}")
        
        # 3. 双方案日分解
        st.subheader(f"3. {view_month}月日分解电量")
        decompose_df = decompose_double_scheme(
            st.session_state.trade_power_typical[view_month],
            st.session_state.trade_power_arbitrage[view_month],
            st.session_state.current_year,
            view_month
        )
        display_df = decompose_df[["时段", "方案一月度电量(MWh)", "方案一日分解电量(MWh)", "方案二月度电量(MWh)", "方案二日分解电量(MWh)"]].copy()
        st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    # 五、电量手动调增调减
    if st.session_state.calculated and st.session_state.selected_months:
        st.divider()
        st.header("✏️ 方案电量手动调整（总量不变）")
        
        col_adj1, col_adj2 = st.columns(2)
        with col_adj1:
            adj_month = st.selectbox("选择调整月份", st.session_state.selected_months, key="adj_month_select")
        with col_adj2:
            adj_scheme = st.selectbox("选择调整方案", ["方案一", "方案二"], key="adj_scheme_select")
        
        # 获取方案数据
        if adj_scheme == "方案一":
            scheme_df = st.session_state.trade_power_typical.get(adj_month, None)
            scheme_col = "方案一月度电量(MWh)"
        else:
            scheme_df = st.session_state.trade_power_arbitrage.get(adj_month, None)
            scheme_col = "方案二月度电量(MWh)"
        base_df = st.session_state.monthly_data.get(adj_month, None)
        
        if scheme_df is None or base_df is None:
            st.warning("⚠️ 该月份数据无效，请重新生成方案")
        else:
            avg_gen_list = base_df["平均发电量(MWh)"].tolist()
            avg_gen_total = sum(avg_gen_list)
            if avg_gen_total <= 0:
                st.error("❌ 原始平均发电量总和为0，无法分摊调整量")
            else:
                old_scheme_df = scheme_df.copy()
                total_fixed = old_scheme_df[scheme_col].sum()
                
                # 可编辑表格
                st.write(f"### {adj_scheme} - {adj_month}月调整（固定总量：{total_fixed:.2f} MWh）")
                edit_df = st.data_editor(
                    scheme_df[["时段", "平均发电量(MWh)", "时段比重(%)", scheme_col]],
                    column_config={
                        "时段": st.column_config.NumberColumn("时段", disabled=True),
                        "平均发电量(MWh)": st.column_config.NumberColumn("原始平均发电量", disabled=True),
                        "时段比重(%)": st.column_config.NumberColumn("时段比重", disabled=True),
                        scheme_col: st.column_config.NumberColumn(f"{scheme_col}（可编辑）", min_value=0.0, step=0.1)
                    },
                    use_container_width=True, num_rows="fixed",
                    key=f"edit_adjust_{adj_month}_{adj_scheme}"
                )
                
                # 检测修改并分摊调整量
                if not edit_df.equals(old_scheme_df):
                    delta_series = edit_df[scheme_col] - old_scheme_df[scheme_col]
                    modified_indices = delta_series[delta_series != 0].index.tolist()
                    
                    if len(modified_indices) > 1:
                        st.warning("⚠️ 暂支持单次修改1个时段")
                        # 恢复原始数据
                        if adj_scheme == "方案一":
                            st.session_state.trade_power_typical[adj_month] = old_scheme_df
                        else:
                            st.session_state.trade_power_arbitrage[adj_month] = old_scheme_df
                    elif len(modified_indices) == 1:
                        mod_idx = modified_indices[0]
                        delta = delta_series.iloc[0]
                        
                        # 分摊调整量
                        other_indices = [idx for idx in range(24) if idx != mod_idx]
                        other_avg_gen = [avg_gen_list[idx] for idx in other_indices]
                        other_avg_total = sum(other_avg_gen)
                        
                        adjusted_df = edit_df.copy()
                        for idx in other_indices:
                            weight_ratio = other_avg_gen[idx] / other_avg_total
                            share_amount = -delta * weight_ratio
                            new_val = adjusted_df.loc[idx, scheme_col] + share_amount
                            adjusted_df.loc[idx, scheme_col] = max(round(new_val, 2), 0.0)
                        
                        # 修正精度误差
                        current_total = adjusted_df[scheme_col].sum()
                        if not np.isclose(current_total, total_fixed, atol=0.01):
                            last_idx = other_indices[-1]
                            adjusted_df.loc[last_idx, scheme_col] += total_fixed - current_total
                        
                        # 更新比重
                        adjusted_df["时段比重(%)"] = round(adjusted_df[scheme_col] / total_fixed * 100, 4)
                        
                        # 保存数据
                        if adj_scheme == "方案一":
                            st.session_state.trade_power_typical[adj_month] = adjusted_df
                        else:
                            st.session_state.trade_power_arbitrage[adj_month] = adjusted_df
                        
                        st.success(f"✅ 调整成功！修改时段：{adjusted_df.loc[mod_idx, '时段']}点，变化量：{delta:.2f} MWh")

# -------------------------- 右侧栏：分月配置+配置导入导出 --------------------------
with col_right:
    st.divider()
    st.header("⚙️ 分月配置中心")
    
    # 一、配置模板导入导出（电量参数+光伏配置）
    st.subheader("📤 配置模板导入导出")
    col_config1, col_config2 = st.columns(2)
    
    # 1. 导出配置模板
    with col_config1:
        config_template_output = export_config_template()
        st.download_button(
            "📥 导出配置模板",
            data=config_template_output,
            file_name=f"{st.session_state.current_power_plant}_{st.session_state.current_year}年配置模板.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    # 2. 导入配置模板
    with col_config2:
        config_file = st.file_uploader(
            "📥 导入配置模板", type=["xlsx"], key="config_file_uploader"
        )
        if config_file is not None:
            import_config(config_file)
    
    # 二、批量应用配置（电量+光伏）
    st.divider()
    st.subheader("📌 批量应用配置")
    
    # 1. 批量电量参数
    st.write("#### 电量参数批量设置")
    col_batch1, col_batch2 = st.columns([2, 1])
    with col_batch1:
        batch_mech_mode = st.selectbox(
            "机制电量模式", ["小时数", "比例(%)"],
            index=0 if st.session_state.batch_mech_mode == "小时数" else 1,
            key="batch_mech_mode_sel"
        )
    with col_batch2:
        mech_max = 100.0 if batch_mech_mode == "比例(%)" else 1000000.0
        batch_mech_value = st.number_input(
            "机制电量数值", min_value=0.0, max_value=mech_max,
            value=st.session_state.batch_mech_value, step=0.1, key="batch_mech_val_inp"
        )
    
    col_batch3, col_batch4 = st.columns([2, 1])
    with col_batch3:
        batch_gua_mode = st.selectbox(
            "保障性电量模式", ["小时数", "比例(%)"],
            index=0 if st.session_state.batch_gua_mode == "小时数" else 1,
            key="batch_gua_mode_sel"
        )
    with col_batch4:
        gua_max = 100.0 if batch_gua_mode == "比例(%)" else 1000000.0
        batch_gua_value = st.number_input(
            "保障性电量数值", min_value=0.0, max_value=gua_max,
            value=st.session_state.batch_gua_value, step=0.1, key="batch_gua_val_inp"
        )
    
    batch_limit_rate = st.number_input(
        "限电率(%)", min_value=0.0, max_value=100.0,
        value=st.session_state.batch_limit_rate, step=0.1, key="batch_limit_rate_inp"
    )
    
    # 2. 批量光伏配置（仅光伏电厂显示）
    batch_pv_params = {}
    if plant_type == "光伏":
        st.write("#### 光伏配置批量设置")
        col_batch_pv1, col_batch_pv2 = st.columns(2)
        with col_batch_pv1:
            batch_pv_core_start = st.number_input(
                "核心起始（点）", min_value=1, max_value=24,
                value=st.session_state.batch_pv_core_start, key="batch_pv_core_start_inp"
            )
            batch_pv_edge_start = st.number_input(
                "边缘起始（点）", min_value=1, max_value=24,
                value=st.session_state.batch_pv_edge_start, key="batch_pv_edge_start_inp"
            )
        with col_batch_pv2:
            batch_pv_core_end = st.number_input(
                "核心结束（点）", min_value=1, max_value=24,
                value=st.session_state.batch_pv_core_end, key="batch_pv_core_end_inp"
            )
            batch_pv_edge_end = st.number_input(
                "边缘结束（点）", min_value=1, max_value=24,
                value=st.session_state.batch_pv_edge_end, key="batch_pv_edge_end_inp"
            )
        batch_pv_params = {
            "core_start": batch_pv_core_start,
            "core_end": batch_pv_core_end,
            "edge_start": batch_pv_edge_start,
            "edge_end": batch_pv_edge_end
        }
    
    # 批量应用按钮
    col_batch_btn1, col_batch_btn2 = st.columns(2)
    with col_batch_btn1:
        if st.button("✅ 批量应用电量参数", use_container_width=True, key="batch_apply_power"):
            for month in range(1, 13):
                st.session_state.monthly_params[month] = {
                    "mechanism_mode": batch_mech_mode,
                    "mechanism_value": batch_mech_value,
                    "guaranteed_mode": batch_gua_mode,
                    "guaranteed_value": batch_gua_value,
                    "power_limit_rate": batch_limit_rate
                }
            st.success("✅ 电量参数已同步到所有月份！")
    
    with col_batch_btn2:
        if plant_type == "光伏" and st.button("✅ 批量应用光伏配置", use_container_width=True, key="batch_apply_pv"):
            for month in range(1, 13):
                st.session_state.monthly_pv_params[month] = batch_pv_params
            st.success("✅ 光伏配置已同步到所有月份！")
    
    # 三、分月配置调整（电量+光伏）
    st.divider()
    st.subheader("🔧 分月配置调整")
    
    # 选择要调整的月份
    selected_config_month = st.selectbox(
        "选择配置月份", range(1, 13), key="selected_config_month"
    )
    
    # 1. 分月电量参数调整
    st.write(f"##### {selected_config_month}月 · 电量参数")
    current_power_params = st.session_state.monthly_params[selected_config_month]
    
    col_month1, col_month2 = st.columns([2, 1])
    with col_month1:
        month_mech_mode = st.selectbox(
            "机制电量模式", ["小时数", "比例(%)"],
            index=0 if current_power_params["mechanism_mode"] == "小时数" else 1,
            key=f"month_mech_mode_{selected_config_month}"
        )
    with col_month2:
        month_mech_max = 100.0 if month_mech_mode == "比例(%)" else 1000000.0
        month_mech_value = st.number_input(
            "机制电量数值", min_value=0.0, max_value=month_mech_max,
            value=current_power_params["mechanism_value"], step=0.1,
            key=f"month_mech_val_{selected_config_month}"
        )
    
    col_month3, col_month4 = st.columns([2, 1])
    with col_month3:
        month_gua_mode = st.selectbox(
            "保障性电量模式", ["小时数", "比例(%)"],
            index=0 if current_power_params["guaranteed_mode"] == "小时数" else 1,
            key=f"month_gua_mode_{selected_config_month}"
        )
    with col_month4:
        month_gua_max = 100.0 if month_gua_mode == "比例(%)" else 1000000.0
        month_gua_value = st.number_input(
            "保障性电量数值", min_value=0.0, max_value=month_gua_max,
            value=current_power_params["guaranteed_value"], step=0.1,
            key=f"month_gua_val_{selected_config_month}"
        )
    
    month_limit_rate = st.number_input(
        "限电率(%)", min_value=0.0, max_value=100.0,
        value=current_power_params["power_limit_rate"], step=0.1,
        key=f"month_limit_rate_{selected_config_month}"
    )
    
    # 2. 分月光伏配置调整（仅光伏电厂显示）
    current_pv_params = {}
    if plant_type == "光伏":
        st.write(f"##### {selected_config_month}月 · 光伏配置")
        current_pv_params = st.session_state.monthly_pv_params[selected_config_month]
        
        col_month_pv1, col_month_pv2 = st.columns(2)
        with col_month_pv1:
            month_pv_core_start = st.number_input(
                "核心起始（点）", min_value=1, max_value=24,
                value=current_pv_params["core_start"], key=f"month_pv_core_start_{selected_config_month}"
            )
            month_pv_edge_start = st.number_input(
                "边缘起始（点）", min_value=1, max_value=24,
                value=current_pv_params["edge_start"], key=f"month_pv_edge_start_{selected_config_month}"
            )
        with col_month_pv2:
            month_pv_core_end = st.number_input(
                "核心结束（点）", min_value=1, max_value=24,
                value=current_pv_params["core_end"], key=f"month_pv_core_end_{selected_config_month}"
            )
            month_pv_edge_end = st.number_input(
                "边缘结束（点）", min_value=1, max_value=24,
                value=current_pv_params["edge_end"], key=f"month_pv_edge_end_{selected_config_month}"
            )
        
        # 实时预览光伏时段划分
        preview_pv_hours = get_pv_arbitrage_hours(selected_config_month)
        st.info(f"""
        时段预览：
        - 核心时段：{preview_pv_hours['core']}点
        - 边缘时段：{preview_pv_hours['edge']}点
        """)
    
    # 保存分月配置按钮
    if st.button(f"💾 保存{selected_config_month}月配置", use_container_width=True, key=f"save_month_config_{selected_config_month}"):
        # 保存电量参数
        st.session_state.monthly_params[selected_config_month] = {
            "mechanism_mode": month_mech_mode,
            "mechanism_value": month_mech_value,
            "guaranteed_mode": month_gua_mode,
            "guaranteed_value": month_gua_value,
            "power_limit_rate": month_limit_rate
        }
        # 保存光伏配置（仅光伏）
        if plant_type == "光伏":
            st.session_state.monthly_pv_params[selected_config_month] = {
                "core_start": month_pv_core_start,
                "core_end": month_pv_core_end,
                "edge_start": month_pv_edge_start,
                "edge_end": month_pv_edge_end
            }
        st.success(f"✅ {selected_config_month}月配置已保存！")
    
    # 四、配置预览表格
    st.divider()
    st.subheader("📋 配置预览")
    
    # 生成配置预览数据
    preview_data = []
    for month in range(1, 13):
        power_params = st.session_state.monthly_params[month]
        pv_params = st.session_state.monthly_pv_params[month]
        
        if plant_type == "光伏":
            pv_info = f"核心：{pv_params['core_start']}-{pv_params['core_end']}点 | 边缘：{pv_params['edge_start']}-{pv_params['edge_end']}点"
        else:
            pv_info = "-"
        
        preview_data.append({
            "月份": f"{month}月",
            "机制电量": f"{power_params['mechanism_mode']}-{power_params['mechanism_value']:.2f}",
            "保障性电量": f"{power_params['guaranteed_mode']}-{power_params['guaranteed_value']:.2f}",
            "限电率(%)": power_params['power_limit_rate'],
            "光伏配置": pv_info
        })
    
    preview_df = pd.DataFrame(preview_data)
    st.dataframe(preview_df, use_container_width=True, hide_index=True)

# 页脚
st.divider()
st.caption(f"© {st.session_state.current_year} 新能源电厂年度方案设计系统 | 分月配置支持 | 双方案交易数据")
