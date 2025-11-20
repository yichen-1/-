import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, date, timedelta
import os
from io import BytesIO

# 设置页面配置
st.set_page_config(
    page_title="贵州新能源按日调整策略系统",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------- 核心配置：24时段与15分钟点映射关系 ----------------------
HOUR_TO_TIMEPTS = {
    0: ["00:15", "00:30", "00:45", "01:00"],
    1: ["01:15", "01:30", "01:45", "02:00"],
    2: ["02:15", "02:30", "02:45", "03:00"],
    3: ["03:15", "03:30", "03:45", "04:00"],
    4: ["04:15", "04:30", "04:45", "05:00"],
    5: ["05:15", "05:30", "05:45", "06:00"],
    6: ["06:15", "06:30", "06:45", "07:00"],
    7: ["07:15", "07:30", "07:45", "08:00"],
    8: ["08:15", "08:30", "08:45", "09:00"],
    9: ["09:15", "09:30", "09:45", "10:00"],
    10: ["10:15", "10:30", "10:45", "11:00"],
    11: ["11:15", "11:30", "11:45", "12:00"],
    12: ["12:15", "12:30", "12:45", "13:00"],
    13: ["13:15", "13:30", "13:45", "14:00"],
    14: ["14:15", "14:30", "14:45", "15:00"],
    15: ["15:15", "15:30", "15:45", "16:00"],
    16: ["16:15", "16:30", "16:45", "17:00"],
    17: ["17:15", "17:30", "17:45", "18:00"],
    18: ["18:15", "18:30", "18:45", "19:00"],
    19: ["19:15", "19:30", "19:45", "20:00"],
    20: ["20:15", "20:30", "20:45", "21:00"],
    21: ["21:15", "21:30", "21:45", "22:00"],
    22: ["22:15", "22:30", "22:45", "23:00"],
    23: ["23:15", "23:30", "23:45", "00:00"]
}

FULL_96_TIMEPTS = []
for hour in range(24):
    FULL_96_TIMEPTS.extend(HOUR_TO_TIMEPTS[hour])
FULL_96_TIMEPTS = list(dict.fromkeys(FULL_96_TIMEPTS))[:96]

# ---------------------- 全局工具函数 ----------------------
def standardize_time(time_str):
    time_str = str(time_str).strip()
    try:
        if ":" in time_str:
            parts = time_str.split(":")[:2]
            hour = parts[0].zfill(2)
            minute = parts[1].zfill(2)
            return "00:00" if (hour == "24" and minute == "00") else f"{hour}:{minute}"
        return time_str
    except:
        return time_str

def standardize_date(date_input):
    try:
        if isinstance(date_input, str) and len(date_input.split("-")) == 3:
            return pd.to_datetime(date_input).date()
        if isinstance(date_input, (int, float)):
            return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(date_input) - 2).date()
        return pd.to_datetime(date_input).date()
    except:
        st.warning(f"⚠️ 日期格式错误：{date_input}，已跳过该数据")
        return None

def get_y_axis_range(values):
    if len(values) == 0:
        return [0, 100]
    min_val = values.min()
    max_val = values.max()
    range_val = max_val - min_val
    if range_val == 0:
        return [min_val * 0.95, max_val * 1.05] if min_val != 0 else [0, 1]
    y_min = max(min_val - range_val * 0.05, 0)
    y_max = max_val + range_val * 0.05
    return [y_min, y_max]

# ---------------------- 初始化数据存储 ----------------------
def init_session_state():
    if "energy_data" not in st.session_state:
        st.session_state.energy_data = pd.DataFrame({
            "日期": [],
            "时刻": [],
            "日前节点电价(元/MWh)": [],
            "实时节点电价(元/MWh)": [],
            "日前预测出力(MW)": [],
            "实时出力(MW)": [],
            "新能源全省预测(MW)": [],
            "新能源全省实测(MW)": [],
            "非市场化机组预测(MW)": [],
            "非市场化机组实测(MW)": [],
            "日前调整后出力(MW)": pd.Series([], dtype=str)
        })
    
    # 按日存储调整系数（核心：date_str -> {hour: ratio}）
    if "daily_hourly_ratios" not in st.session_state:
        st.session_state.daily_hourly_ratios = {}
    
    # 多日选择器默认值（默认选中当天）
    if "selected_dates" not in st.session_state:
        st.session_state.selected_dates = [date.today()]
    
    # 加载本地备份
    if os.path.exists("energy_data_backup.csv"):
        try:
            backup_data = pd.read_csv("energy_data_backup.csv")
            field_mapping = {
                "日前出清电量(MWh)": "日前预测出力(MW)",
                "实时出清电量(MWh)": "实时出力(MW)",
                "日前调整后电量(MWh)": "日前调整后出力(MW)"
            }
            backup_data.rename(columns=field_mapping, inplace=True)
            backup_data["日期"] = backup_data["日期"].apply(standardize_date)
            backup_data = backup_data.dropna(subset=["日期"])
            backup_data["时刻"] = backup_data["时刻"].apply(standardize_time)
            backup_data["日前调整后出力(MW)"] = backup_data["日前调整后出力(MW)"].astype(str).fillna("未计算")
            st.session_state.energy_data = backup_data
            st.success("✅ 已加载本地备份数据！")
        except Exception as e:
            st.warning(f"⚠️ 本地备份数据损坏：{str(e)}，将使用空数据初始化")
    
    # 加载调整系数备份
    if os.path.exists("daily_hourly_ratios_backup.json"):
        try:
            import json
            with open("daily_hourly_ratios_backup.json", "r", encoding="utf-8") as f:
                st.session_state.daily_hourly_ratios = json.load(f)
            st.success("✅ 已加载按日调整系数备份！")
        except Exception as e:
            st.warning(f"⚠️ 调整系数备份损坏：{str(e)}，将自动初始化")

# ---------------------- 备份数据 ----------------------
def backup_data():
    st.session_state.energy_data.to_csv("energy_data_backup.csv", index=False, encoding="utf-8-sig")
    import json
    with open("daily_hourly_ratios_backup.json", "w", encoding="utf-8") as f:
        json.dump(st.session_state.daily_hourly_ratios, f, ensure_ascii=False, indent=2)

# ---------------------- 数据解析函数 ----------------------
def parse_single_sheet(df, sheet_name):
    required_columns = [
        "日期", "时刻", "日前节点电价(元/MWh)", "实时节点电价(元/MWh)",
        "日前预测出力(MW)", "实时出力(MW)",
        "新能源全省预测(MW)", "新能源全省实测(MW)",
        "非市场化机组预测(MW)", "非市场化机组实测(MW)"
    ]
    
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        st.warning(f"⚠️ 子表[{sheet_name}]缺少字段：{', '.join(missing_cols)}，已跳过")
        return None
    
    sheet_date = standardize_date(sheet_name)
    if not sheet_date:
        return None
    
    df["日期"] = sheet_date
    df["日期"] = df["日期"].apply(standardize_date)
    
    numeric_cols = [col for col in required_columns if col not in ["日期", "时刻"]]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    
    df["时刻"] = df["时刻"].apply(standardize_time)
    
    uploaded_timepts = df["时刻"].tolist()
    missing_timepts = [t for t in FULL_96_TIMEPTS if t not in uploaded_timepts]
    if missing_timepts:
        st.warning(f"⚠️ 子表[{sheet_name}]缺失 {len(missing_timepts)} 个时刻，已自动补0")
        missing_data = pd.DataFrame({
            "日期": [sheet_date]*len(missing_timepts),
            "时刻": missing_timepts,
            "日前节点电价(元/MWh)": [0.0]*len(missing_timepts),
            "实时节点电价(元/MWh)": [0.0]*len(missing_timepts),
            "日前预测出力(MW)": [0.0]*len(missing_timepts),
            "实时出力(MW)": [0.0]*len(missing_timepts),
            "新能源全省预测(MW)": [0.0]*len(missing_timepts),
            "新能源全省实测(MW)": [0.0]*len(missing_timepts),
            "非市场化机组预测(MW)": [0.0]*len(missing_timepts),
            "非市场化机组实测(MW)": [0.0]*len(missing_timepts)
        })
        df = pd.concat([df, missing_data], ignore_index=True)
    
    df = df[df["时刻"].isin(FULL_96_TIMEPTS)]
    df["时刻_order"] = df["时刻"].map({t: i for i, t in enumerate(FULL_96_TIMEPTS)})
    df = df.sort_values("时刻_order").drop(columns=["时刻_order"]).reset_index(drop=True)
    df["日前调整后出力(MW)"] = "未计算"
    
    st.session_state.energy_data = st.session_state.energy_data[st.session_state.energy_data["日期"] != sheet_date].copy()
    
    st.success(f"✅ 子表[{sheet_name}]解析成功：{len(df)} 条数据（96点完整时刻）")
    return df

def parse_multi_sheet_file(file):
    try:
        excel_file = pd.ExcelFile(file)
        sheet_names = [name for name in excel_file.sheet_names if name != "使用说明"]
        if not sheet_names:
            st.error("❌ 未找到有效数据子表，请检查文件结构（子表名需为日期格式：YYYY-MM-DD）")
            return None
        
        st.info(f"ℹ️ 开始解析 {len(sheet_names)} 个子表...")
        all_parsed_data = []
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(file, engine="openpyxl", sheet_name=sheet_name)
                parsed_sheet = parse_single_sheet(df, sheet_name)
                if parsed_sheet is not None and not parsed_sheet.empty:
                    all_parsed_data.append(parsed_sheet)
            except Exception as e:
                st.warning(f"⚠️ 子表[{sheet_name}]解析失败：{str(e)}，已跳过")
        
        if not all_parsed_data:
            st.error("❌ 所有子表均解析失败或无有效数据")
            return None
        
        combined_data = pd.concat(all_parsed_data, ignore_index=True)
        st.success(f"✅ 批量解析完成！共解析 {len(all_parsed_data)} 个日期，合计 {len(combined_data)} 条数据")
        return combined_data
    except Exception as e:
        st.error(f"❌ 文件解析失败：{str(e)}")
        return None

# ---------------------- 调整后出力计算（支持多日批量计算） ----------------------
def calculate_adjusted_output(is_unified=False, unified_ratio=100):
    selected_dates = st.session_state.selected_dates
    if not selected_dates:
        st.warning("⚠️ 请先选择目标日期")
        return
    
    energy_data = st.session_state.energy_data.copy()
    calculated_dates = []
    
    for target_date in selected_dates:
        target_date_str = target_date.strftime("%Y-%m-%d")
        date_mask = energy_data["日期"] == target_date
        
        # 检查该日期是否有数据
        if not energy_data[date_mask].any().any():
            st.warning(f"⚠️ 未找到 {target_date_str} 的数据，已跳过该日期")
            continue
        
        # 获取该日期的系数（无则初始化）
        if target_date_str not in st.session_state.daily_hourly_ratios:
            st.session_state.daily_hourly_ratios[target_date_str] = {hour: 100 for hour in range(24)}
        hourly_ratios = st.session_state.daily_hourly_ratios[target_date_str]
        
        # 补全系数
        for hour in range(24):
            if hour not in hourly_ratios:
                hourly_ratios[hour] = 100
        
        # 计算该日期的调整后出力
        for idx in energy_data[date_mask].index:
            time_point = energy_data.loc[idx, "时刻"]
            original_power = energy_data.loc[idx, "日前预测出力(MW)"]
            
            target_hour = 23 if time_point == "00:00" else int(time_point.split(":")[0])
            ratio = unified_ratio if is_unified else hourly_ratios[target_hour]
            adjusted_power = round(original_power * (ratio / 100), 2)
            energy_data.loc[idx, "日前调整后出力(MW)"] = str(adjusted_power)
        
        calculated_dates.append(target_date_str)
    
    # 更新数据并备份
    st.session_state.energy_data = energy_data
    backup_data()
    
    # 提示结果
    if calculated_dates:
        if is_unified:
            st.success(f"✅ 以下日期全时段统一计算完成！调整系数：{unified_ratio}%\n{', '.join(calculated_dates)}")
        else:
            st.success(f"✅ 以下日期分时段计算完成！\n{', '.join(calculated_dates)}")
    else:
        st.warning("⚠️ 无有效日期完成计算")

# ---------------------- 统计函数（支持多日汇总） ----------------------
def calculate_statistics():
    selected_dates = st.session_state.selected_dates
    if not selected_dates:
        return {"global_stats": {k: 0.0 for k in [
            "avg_day_ahead_price", "avg_real_time_price", "total_day_ahead_power",
            "total_real_time_power", "total_adjusted_power"
        ]}, "daily_stats": pd.DataFrame()}
    
    # 筛选选中日期的数据
    filtered_df = st.session_state.energy_data[st.session_state.energy_data["日期"].isin(selected_dates)].copy()
    if filtered_df.empty:
        return {"global_stats": {k: 0.0 for k in [
            "avg_day_ahead_price", "avg_real_time_price", "total_day_ahead_power",
            "total_real_time_power", "total_adjusted_power"
        ]}, "daily_stats": pd.DataFrame()}
    
    # 处理调整后出力数值
    filtered_df["调整后出力数值"] = pd.to_numeric(filtered_df["日前调整后出力(MW)"], errors="coerce")
    
    # 按日期分组统计
    daily_stats = filtered_df.groupby("日期").agg({
        "日前节点电价(元/MWh)": "mean",
        "实时节点电价(元/MWh)": "mean",
        "日前预测出力(MW)": "sum",
        "实时出力(MW)": "sum",
        "新能源全省预测(MW)": "mean",
        "新能源全省实测(MW)": "mean",
        "非市场化机组预测(MW)": "mean",
        "非市场化机组实测(MW)": "mean",
        "调整后出力数值": "sum"
    }).reset_index()
    
    # 格式化日期和数值
    daily_stats["日期"] = daily_stats["日期"].apply(lambda x: x.strftime("%Y-%m-%d"))
    daily_stats.columns = [
        "日期", "平均日前节点电价(元/MWh)", "平均实时节点电价(元/MWh)",
        "总日前预测出力(MW)", "总实时出力(MW)",
        "平均新能源全省预测(MW)", "平均新能源全省实测(MW)",
        "平均非市场化机组预测(MW)", "平均非市场化机组实测(MW)",
        "总调整后出力(MW)"
    ]
    
    # 保留2位小数
    for col in daily_stats.columns[1:]:
        daily_stats[col] = daily_stats[col].round(2)
    
    # 全局汇总（所有选中日期的合计/平均）
    global_stats = {
        "avg_day_ahead_price": daily_stats["平均日前节点电价(元/MWh)"].mean().round(2),
        "avg_real_time_price": daily_stats["平均实时节点电价(元/MWh)"].mean().round(2),
        "total_day_ahead_power": daily_stats["总日前预测出力(MW)"].sum().round(2),
        "total_real_time_power": daily_stats["总实时出力(MW)"].sum().round(2),
        "total_adjusted_power": daily_stats["总调整后出力(MW)"].sum().round(2)
    }
    
    # 添加总计行
    total_row = pd.DataFrame({
        "日期": ["总计"],
        "平均日前节点电价(元/MWh)": [global_stats["avg_day_ahead_price"]],
        "平均实时节点电价(元/MWh)": [global_stats["avg_real_time_price"]],
        "总日前预测出力(MW)": [global_stats["total_day_ahead_power"]],
        "总实时出力(MW)": [global_stats["total_real_time_power"]],
        "平均新能源全省预测(MW)": [daily_stats["平均新能源全省预测(MW)"].mean().round(2)],
        "平均新能源全省实测(MW)": [daily_stats["平均新能源全省实测(MW)"].mean().round(2)],
        "平均非市场化机组预测(MW)": [daily_stats["平均非市场化机组预测(MW)"].mean().round(2)],
        "平均非市场化机组实测(MW)": [daily_stats["平均非市场化机组实测(MW)"].mean().round(2)],
        "总调整后出力(MW)": [global_stats["total_adjusted_power"]]
    })
    daily_stats = pd.concat([daily_stats, total_row], ignore_index=True)
    
    return {"global_stats": global_stats, "daily_stats": daily_stats}

# ---------------------- 图表函数（支持多日对比） ----------------------
def plot_price_trend():
    selected_dates = st.session_state.selected_dates
    if not selected_dates:
        return go.Figure(layout=go.Layout(title="请先选择目标日期"))
    
    filtered_df = st.session_state.energy_data[st.session_state.energy_data["日期"].isin(selected_dates)].copy()
    if filtered_df.empty:
        return go.Figure(layout=go.Layout(title="所选日期无电价数据"))
    
    df_sorted = filtered_df.sort_values(["日期", "时刻_order" if "时刻_order" in filtered_df.columns else "时刻"])
    df_sorted["时刻_label"] = df_sorted["时刻"]
    df_sorted["日期_str"] = df_sorted["日期"].apply(lambda x: x.strftime("%Y-%m-%d"))
    
    # 收集所有日期的电价数据
    fig = go.Figure()
    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b", "#e377c2", "#7f7f7f"]
    
    for i, (date_str, group) in enumerate(df_sorted.groupby("日期_str")):
        color = colors[i % len(colors)]
        # 日前电价
        fig.add_trace(go.Scatter(
            x=group["时刻_label"], y=group["日前节点电价(元/MWh)"],
            name=f"{date_str} - 日前电价", line=dict(color=color, width=2),
            fill="tozeroy", fillcolor=f"rgba({int(color[1:3],16)}, {int(color[3:5],16)}, {int(color[5:7],16)}, 0.05)"
        ))
        # 实时电价（虚线）
        fig.add_trace(go.Scatter(
            x=group["时刻_label"], y=group["实时节点电价(元/MWh)"],
            name=f"{date_str} - 实时电价", line=dict(color=color, width=2, dash="dash")
        ))
    
    # 计算y轴范围
    price_values = pd.concat([df_sorted["日前节点电价(元/MWh)"], df_sorted["实时节点电价(元/MWh)"]])
    y_range = get_y_axis_range(price_values)
    
    fig.update_layout(
        title=f"节点电价趋势对比（{len(selected_dates)} 个日期）",
        xaxis_title="时刻", yaxis_title="电价（元/MWh）",
        height=350, hovermode="x unified",
        xaxis=dict(tickmode="array", tickvals=df_sorted["时刻_label"].unique()[::8], tickangle=-45),
        yaxis=dict(range=y_range, tickformat=".1f"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5)
    )
    return fig

def plot_power_trend():
    selected_dates = st.session_state.selected_dates
    if not selected_dates:
        return go.Figure(layout=go.Layout(title="请先选择目标日期"))
    
    filtered_df = st.session_state.energy_data[st.session_state.energy_data["日期"].isin(selected_dates)].copy()
    if filtered_df.empty:
        return go.Figure(layout=go.Layout(title="所选日期无出力数据"))
    
    df_sorted = filtered_df.sort_values(["日期", "时刻_order" if "时刻_order" in filtered_df.columns else "时刻"])
    df_sorted["时刻_label"] = df_sorted["时刻"]
    df_sorted["日期_str"] = df_sorted["日期"].apply(lambda x: x.strftime("%Y-%m-%d"))
    df_sorted["调整后出力数值"] = pd.to_numeric(df_sorted["日前调整后出力(MW)"], errors="coerce")
    
    fig = go.Figure()
    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b", "#e377c2", "#7f7f7f"]
    
    for i, (date_str, group) in enumerate(df_sorted.groupby("日期_str")):
        color = colors[i % len(colors)]
        # 原始预测出力
        fig.add_trace(go.Scatter(
            x=group["时刻_label"], y=group["日前预测出力(MW)"],
            name=f"{date_str} - 原始预测", line=dict(color=color, width=2)
        ))
        # 调整后出力（有数据才显示）
        if not group["调整后出力数值"].isna().all():
            fig.add_trace(go.Scatter(
                x=group["时刻_label"], y=group["调整后出力数值"],
                name=f"{date_str} - 调整后", line=dict(color=color, width=2, dash="dash")
            ))
        # 实时出力
        fig.add_trace(go.Scatter(
            x=group["时刻_label"], y=group["实时出力(MW)"],
            name=f"{date_str} - 实时", line=dict(color=color, width=1, dash="dot")
        ))
    
    # 计算y轴范围
    power_values = pd.concat([
        df_sorted["日前预测出力(MW)"],
        df_sorted["调整后出力数值"].dropna(),
        df_sorted["实时出力(MW)"]
    ])
    y_range = get_y_axis_range(power_values)
    
    fig.update_layout(
        title=f"出力趋势对比（{len(selected_dates)} 个日期）",
        xaxis_title="时刻", yaxis_title="出力（MW）",
        height=350, hovermode="x unified",
        xaxis=dict(tickmode="array", tickvals=df_sorted["时刻_label"].unique()[::8], tickangle=-45),
        yaxis=dict(range=y_range, tickformat=".1f"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5)
    )
    return fig

def plot_supply_demand_trend():
    selected_dates = st.session_state.selected_dates
    if not selected_dates:
        return go.Figure(layout=go.Layout(title="请先选择目标日期"))
    
    filtered_df = st.session_state.energy_data[st.session_state.energy_data["日期"].isin(selected_dates)].copy()
    if filtered_df.empty:
        return go.Figure(layout=go.Layout(title="所选日期无供需数据"))
    
    df_sorted = filtered_df.sort_values(["日期", "时刻_order" if "时刻_order" in filtered_df.columns else "时刻"])
    df_sorted["时刻_label"] = df_sorted["时刻"]
    df_sorted["日期_str"] = df_sorted["日期"].apply(lambda x: x.strftime("%Y-%m-%d"))
    
    fig = go.Figure()
    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b", "#e377c2", "#7f7f7f"]
    
    for i, (date_str, group) in enumerate(df_sorted.groupby("日期_str")):
        color = colors[i % len(colors)]
        # 新能源预测
        fig.add_trace(go.Scatter(
            x=group["时刻_label"], y=group["新能源全省预测(MW)"],
            name=f"{date_str} - 新能源预测", line=dict(color=color, width=2)
        ))
        # 新能源实测
        fig.add_trace(go.Scatter(
            x=group["时刻_label"], y=group["新能源全省实测(MW)"],
            name=f"{date_str} - 新能源实测", line=dict(color=color, width=2, dash="dash")
        ))
        # 非市场化预测
        fig.add_trace(go.Scatter(
            x=group["时刻_label"], y=group["非市场化机组预测(MW)"],
            name=f"{date_str} - 非市场化预测", line=dict(color=color, width=1.5, dash="dot")
        ))
        # 非市场化实测
        fig.add_trace(go.Scatter(
            x=group["时刻_label"], y=group["非市场化机组实测(MW)"],
            name=f"{date_str} - 非市场化实测", line=dict(color=color, width=1.5, dash="longdashdot")
        ))
    
    # 计算y轴范围
    supply_values = pd.concat([
        df_sorted["新能源全省预测(MW)"],
        df_sorted["新能源全省实测(MW)"],
        df_sorted["非市场化机组预测(MW)"],
        df_sorted["非市场化机组实测(MW)"]
    ])
    y_range = get_y_axis_range(supply_values)
    
    fig.update_layout(
        title=f"机组出力对比（{len(selected_dates)} 个日期）",
        xaxis_title="时刻", yaxis_title="出力（MW）",
        height=350, hovermode="x unified",
        xaxis=dict(tickmode="array", tickvals=df_sorted["时刻_label"].unique()[::8], tickangle=-45),
        yaxis=dict(range=y_range, tickformat=".0f"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5)
    )
    return fig

# ---------------------- 收益复盘（支持多日汇总） ----------------------
def calculate_revenue():
    selected_dates = st.session_state.selected_dates
    if not selected_dates:
        return {"total": {}, "daily": pd.DataFrame()}
    
    filtered_df = st.session_state.energy_data[st.session_state.energy_data["日期"].isin(selected_dates)].copy()
    revenue_df = filtered_df[filtered_df["日前调整后出力(MW)"] != "未计算"].copy()
    
    if revenue_df.empty:
        return {"total": {}, "daily": pd.DataFrame()}
    
    revenue_df["日前调整后出力数值"] = pd.to_numeric(revenue_df["日前调整后出力(MW)"])
    revenue_df["日期_str"] = revenue_df["日期"].apply(lambda x: x.strftime("%Y-%m-%d"))
    
    # 计算每个15分钟点的收益
    revenue_df["调整后收益(元)"] = (
        revenue_df["日前调整后出力数值"] * revenue_df["日前节点电价(元/MWh)"] +
        (revenue_df["实时出力(MW)"] - revenue_df["日前调整后出力数值"]) * revenue_df["实时节点电价(元/MWh)"]
    )
    revenue_df["调整前收益(元)"] = (
        revenue_df["日前预测出力(MW)"] * revenue_df["日前节点电价(元/MWh)"] +
        (revenue_df["实时出力(MW)"] - revenue_df["日前预测出力(MW)"]) * revenue_df["实时节点电价(元/MWh)"]
    )
    
    # 按日期汇总
    daily_rev = revenue_df.groupby("日期_str").agg({
        "调整前收益(元)": "sum",
        "调整后收益(元)": "sum"
    }).reset_index()
    daily_rev["增收(元)"] = daily_rev["调整后收益(元)"] - daily_rev["调整前收益(元)"]
    daily_rev.rename(columns={"日期_str": "日期"}, inplace=True)
    
    # 添加总计行
    total_row = pd.DataFrame({
        "日期": ["总计"],
        "调整前收益(元)": [daily_rev["调整前收益(元)"].sum()],
        "调整后收益(元)": [daily_rev["调整后收益(元)"].sum()],
        "增收(元)": [daily_rev["增收(元)"].sum()]
    })
    daily_rev = pd.concat([daily_rev, total_row], ignore_index=True)
    
    # 保留2位小数
    for col in daily_rev.columns[1:]:
        daily_rev[col] = daily_rev[col].round(2)
    
    # 总计信息
    total = {
        "调整前总收益(元)": daily_rev.loc[daily_rev["日期"] == "总计", "调整前收益(元)"].iloc[0],
        "调整后总收益(元)": daily_rev.loc[daily_rev["日期"] == "总计", "调整后收益(元)"].iloc[0],
        "总增收(元)": daily_rev.loc[daily_rev["日期"] == "总计", "增收(元)"].iloc[0]
    }
    
    return {"total": total, "daily": daily_rev}

# ---------------------- 导出模板函数 ----------------------
def export_multi_sheet_template():
    template_dates = [date.today() + timedelta(days=i) for i in range(3)]
    sheet_names = [d.strftime("%Y-%m-%d") for d in template_dates]
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for idx, (sheet_name, template_date) in enumerate(zip(sheet_names, template_dates)):
            template_data = pd.DataFrame({
                "日期": [template_date.strftime("%Y-%m-%d")]*96,
                "时刻": FULL_96_TIMEPTS,
                "日前节点电价(元/MWh)": [0.0]*96,
                "实时节点电价(元/MWh)": [0.0]*96,
                "日前预测出力(MW)": [0.0]*96,
                "实时出力(MW)": [0.0]*96,
                "新能源全省预测(MW)": [0.0]*96,
                "新能源全省实测(MW)": [0.0]*96,
                "非市场化机组预测(MW)": [0.0]*96,
                "非市场化机组实测(MW)": [0.0]*96
            })
            template_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        guide_data = pd.DataFrame({
            "使用说明": [
                "1. 子表名格式：必须为日期格式（YYYY-MM-DD），否则无法解析",
                "2. 每个子表对应一天数据，包含96个15分钟时刻（00:15-23:45）",
                "3. 填写说明：",
                "   - 日前节点电价(元/MWh)：当日各时刻日前市场电价",
                "   - 实时节点电价(元/MWh)：当日各时刻实时市场电价",
                "   - 日前预测出力(MW)：当日各时刻初始日前预测出力",
                "   - 实时出力(MW)：当日各时刻实际实时出力",
                "   - 新能源全省预测/实测(MW)：全省新能源出力数据（可选）",
                "   - 非市场化机组预测/实测(MW)：非市场化机组出力数据（可选）",
                "4. 缺失时刻会自动补0，建议完整填写96个时刻",
                "5. 可新增子表（右键→插入），子表名改为目标日期即可"
            ]
        })
        guide_data.to_excel(writer, sheet_name="使用说明", index=False)
    
    output.seek(0)
    return output

# ---------------------- 主函数 ----------------------
def main():
    init_session_state()
    
    # ---------------------- 顶部：多日选择器（核心改进） ----------------------
    st.header("🎯 新能源日前调整策略系统")
    col1, col2 = st.columns([1, 3])
    with col1:
        # 多日选择器（支持按住Ctrl多选）
        selected_dates = st.date_input(
            "选择目标日期（可多选）",
            value=st.session_state.selected_dates,
            min_value=date.today() - timedelta(days=90),
            max_value=date.today() + timedelta(days=30),
            key="date_picker"
        )
        # 处理单日期/多日期格式（st.date_input返回单个date或list）
        if isinstance(selected_dates, date):
            selected_dates = [selected_dates]
        st.session_state.selected_dates = selected_dates
        
        # 显示选中日期
        if selected_dates:
            date_strs = [d.strftime("%Y-%m-%d") for d in selected_dates]
            if len(date_strs) <= 5:
                st.info(f"当前选中日期：\n{', '.join(date_strs)}")
            else:
                st.info(f"当前选中 {len(date_strs)} 个日期（{date_strs[0]} 至 {date_strs[-1]}）")
        else:
            st.warning("⚠️ 请选择至少一个目标日期")
    
    with col2:
        # 原有上传、导出、清空功能
        st.markdown("### 数据上传")
        uploaded_file = st.file_uploader(
            "上传多日期Excel文件（子表名需为日期格式：YYYY-MM-DD）",
            type=["xlsx"],
            key="data_uploader"
        )
        col_export, col_clear = st.columns(2)
        with col_export:
            if st.button("📥 导出标准模板"):
                template_file = export_multi_sheet_template()
                st.download_button(
                    label="下载模板",
                    data=template_file,
                    file_name="新能源调整策略数据模板.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        with col_clear:
            if st.button("🗑️ 清空所有数据", type="secondary", disabled=st.session_state.energy_data.empty):
                if st.checkbox("确认清空（不可恢复）", key="clear_confirm"):
                    st.session_state.energy_data = pd.DataFrame({
                        "日期": [], "时刻": [], "日前节点电价(元/MWh)": [], "实时节点电价(元/MWh)": [],
                        "日前预测出力(MW)": [], "实时出力(MW)": [], "新能源全省预测(MW)": [], "新能源全省实测(MW)": [],
                        "非市场化机组预测(MW)": [], "非市场化机组实测(MW)": [], "日前调整后出力(MW)": pd.Series([], dtype=str)
                    })
                    st.session_state.daily_hourly_ratios = {}
                    if os.path.exists("energy_data_backup.csv"):
                        os.remove("energy_data_backup.csv")
                    if os.path.exists("daily_hourly_ratios_backup.json"):
                        os.remove("daily_hourly_ratios_backup.json")
                    st.success("✅ 所有数据已清空！")
    
    # 处理上传文件
    if uploaded_file is not None:
        batch_data = parse_multi_sheet_file(uploaded_file)
        if batch_data is not None and not batch_data.empty:
            st.session_state.energy_data = pd.concat([st.session_state.energy_data, batch_data], ignore_index=True)
            backup_data()
    
    st.divider()
    
    # ---------------------- 调整后出力计算（支持多日） ----------------------
    selected_dates = st.session_state.selected_dates
    date_title = f"{len(selected_dates)} 个日期" if len(selected_dates) > 1 else selected_dates[0].strftime("%Y-%m-%d") if selected_dates else "目标日期"
    st.subheader(f"📊 {date_title} 调整后出力计算")
    
    # 1. 全时段统一计算（多日共用一个系数）
    st.markdown("### 1. 全时段统一计算")
    st.info("ℹ️ 所有选中日期共用同一个调整系数，批量计算")
    col_unified1, col_unified2 = st.columns([3, 1])
    with col_unified1:
        unified_ratio = st.number_input(
            "统一调整系数（%）",
            min_value=1, max_value=300, value=100, step=1,
            key="unified_ratio"
        )
    with col_unified2:
        st.write("")
        st.write("")
        if st.button("🚀 批量应用统一计算", key="unified_calc_btn", disabled=not selected_dates):
            calculate_adjusted_output(is_unified=True, unified_ratio=unified_ratio)
    
    st.divider()
    
    # 2. 24时段分时段计算（多日选中时，编辑第一个日期的系数）
    st.markdown("### 2. 24时段分时段计算")
    if len(selected_dates) == 0:
        st.warning("⚠️ 请先选择目标日期")
    elif len(selected_dates) > 1:
        # 多日选中时，默认编辑第一个日期的系数，计算时批量应用各日期独立系数
        edit_date = selected_dates[0]
        edit_date_str = edit_date.strftime("%Y-%m-%d")
        st.info(f"ℹ️ 多日选中时，当前编辑 {edit_date_str} 的系数（其他日期系数保持独立）\n计算时将按每个日期的独立系数批量计算")
        
        # 加载第一个日期的系数
        if edit_date_str not in st.session_state.daily_hourly_ratios:
            st.session_state.daily_hourly_ratios[edit_date_str] = {hour: 100 for hour in range(24)}
        hourly_ratios = st.session_state.daily_hourly_ratios[edit_date_str]
    else:
        # 单日选中时，编辑该日期的系数
        edit_date = selected_dates[0]
        edit_date_str = edit_date.strftime("%Y-%m-%d")
        st.info(f"ℹ️ 当前编辑 {edit_date_str} 的系数（仅作用于该日期）")
        
        # 加载该日期的系数
        if edit_date_str not in st.session_state.daily_hourly_ratios:
            st.session_state.daily_hourly_ratios[edit_date_str] = {hour: 100 for hour in range(24)}
        hourly_ratios = st.session_state.daily_hourly_ratios[edit_date_str]
    
    # 补全系数（避免KeyError）
    if selected_dates:
        for hour in range(24):
            if hour not in hourly_ratios:
                hourly_ratios[hour] = 100
        
        # 构建系数表格
        hourly_table_data = []
        for hour in range(24):
            hourly_table_data.append({
                "时段": f"{hour:02d}:00-{hour:02d}:59",
                "对应15分钟点": "、".join(HOUR_TO_TIMEPTS[hour]),
                "调整系数(%)": hourly_ratios[hour]
            })
        
        # 可编辑表格
        edited_hourly_df = st.data_editor(
            hourly_table_data,
            column_config={
                "时段": st.column_config.TextColumn("时段", disabled=True),
                "对应15分钟点": st.column_config.TextColumn("对应15分钟点", disabled=True),
                "调整系数(%)": st.column_config.NumberColumn(
                    "调整系数(%)", min_value=1, max_value=300, step=1, format="%d"
                )
            },
            disabled=False,
            hide_index=True,
            width='stretch',
            height=500,
            key="hourly_ratios_table"
        )
        
        # 保存系数并批量计算所有选中日期
        if st.button("🚀 批量应用分时段计算", key="hourly_calc_btn"):
            # 更新当前编辑日期的系数（核心修复：列表直接用索引访问，去掉.iloc）
            updated_ratios = {hour: edited_hourly_df[hour]["调整系数(%)"] for hour in range(24)}
            st.session_state.daily_hourly_ratios[edit_date_str] = updated_ratios
            # 批量计算所有选中日期（每个日期用自己的系数）
            calculate_adjusted_output(is_unified=False)
    
    st.divider()
    
    # ---------------------- 数据统计（多日汇总） ----------------------
    st.subheader(f"📈 {date_title} 数据统计")
    stats = calculate_statistics()
    global_stats = stats["global_stats"]
    daily_stats = stats["daily_stats"]
    
    # 统计卡片（多日汇总）
    col_stats1, col_stats2, col_stats3, col_stats4 = st.columns(4)
    with col_stats1:
        st.markdown(f"""
        <div style="background:#f0f8ff; border-radius:8px; padding:15px; text-align:center;">
            <h6 style="margin:0 0 8px 0; color:#666;">平均日前电价</h6>
            <p style="margin:0; font-size:20px; font-weight:bold; color:#1f77b4;">{global_stats['avg_day_ahead_price']:.1f} 元/MWh</p>
        </div>
        """, unsafe_allow_html=True)
    with col_stats2:
        st.markdown(f"""
        <div style="background:#f5fafe; border-radius:8px; padding:15px; text-align:center;">
            <h6 style="margin:0 0 8px 0; color:#666;">总预测出力</h6>
            <p style="margin:0; font-size:20px; font-weight:bold; color:#2ca02c;">{global_stats['total_day_ahead_power']:.1f} MW</p>
        </div>
        """, unsafe_allow_html=True)
    with col_stats3:
        st.markdown(f"""
        <div style="background:#fef7fb; border-radius:8px; padding:15px; text-align:center;">
            <h6 style="margin:0 0 8px 0; color:#666;">总调整后出力</h6>
            <p style="margin:0; font-size:20px; font-weight:bold; color:#d62728;">{global_stats['total_adjusted_power']:.1f} MW</p>
        </div>
        """, unsafe_allow_html=True)
    with col_stats4:
        st.markdown(f"""
        <div style="background:#f8f8f8; border-radius:8px; padding:15px; text-align:center;">
            <h6 style="margin:0 0 8px 0; color:#666;">总实时出力</h6>
            <p style="margin:0; font-size:20px; font-weight:bold; color:#9467bd;">{global_stats['total_real_time_power']:.1f} MW</p>
        </div>
        """, unsafe_allow_html=True)
    
    # 详细统计表格（含每日数据+总计）
    st.markdown("#### 每日统计详情（含总计）")
    if not daily_stats.empty:
        st.data_editor(
            daily_stats,
            column_config={
                "日期": st.column_config.TextColumn("日期", disabled=True),
                "平均日前节点电价(元/MWh)": st.column_config.NumberColumn("平均日前节点电价(元/MWh)", format="%.2f"),
                "平均实时节点电价(元/MWh)": st.column_config.NumberColumn("平均实时节点电价(元/MWh)", format="%.2f"),
                "总日前预测出力(MW)": st.column_config.NumberColumn("总日前预测出力(MW)", format="%.2f"),
                "总实时出力(MW)": st.column_config.NumberColumn("总实时出力(MW)", format="%.2f"),
                "平均新能源全省预测(MW)": st.column_config.NumberColumn("平均新能源全省预测(MW)", format="%.2f"),
                "平均新能源全省实测(MW)": st.column_config.NumberColumn("平均新能源全省实测(MW)", format="%.2f"),
                "平均非市场化机组预测(MW)": st.column_config.NumberColumn("平均非市场化机组预测(MW)", format="%.2f"),
                "平均非市场化机组实测(MW)": st.column_config.NumberColumn("平均非市场化机组实测(MW)", format="%.2f"),
                "总调整后出力(MW)": st.column_config.NumberColumn("总调整后出力(MW)", format="%.2f")
            },
            disabled=True,
            hide_index=True,
            width='stretch'
        )
    else:
        st.info("ℹ️ 暂无统计数据，请先上传选中日期的数据并完成调整计算")
    
    st.divider()
    
    # ---------------------- 趋势图表（多日对比） ----------------------
    st.subheader(f"📊 {date_title} 趋势图表")
    st.plotly_chart(plot_price_trend(), width='stretch')
    st.plotly_chart(plot_power_trend(), width='stretch')
    st.plotly_chart(plot_supply_demand_trend(), width='stretch')
    
    st.divider()
    
    # ---------------------- 收益复盘（多日汇总） ----------------------
    st.subheader(f"💰 {date_title} 收益复盘分析")
    if st.button("开始复盘计算", key="rev_calc_btn", disabled=not selected_dates):
        revenue_result = calculate_revenue()
        if not revenue_result["daily"].empty:
            # 总收益对比卡片（多日汇总）
            st.markdown("#### 总收益对比（所有选中日期汇总）")
            col_total1, col_total2, col_total3 = st.columns(3)
            profit_color = "green" if revenue_result["total"]["总增收(元)"] >= 0 else "red"
            
            with col_total1:
                st.markdown(f"""
                <div style="background:#f0f8ff; border-radius:8px; padding:15px; border-left:4px solid #1f77b4;">
                    <h5 style="margin:0 0 10px 0;">调整前总收益</h5>
                    <p style="font-size:18px; font-weight:bold;">{revenue_result['total']['调整前总收益(元)']:.2f} 元</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col_total2:
                st.markdown(f"""
                <div style="background:#f5fafe; border-radius:8px; padding:15px; border-left:4px solid #2ca02c;">
                    <h5 style="margin:0 0 10px 0;">调整后总收益</h5>
                    <p style="font-size:18px; font-weight:bold;">{revenue_result['total']['调整后总收益(元)']:.2f} 元</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col_total3:
                st.markdown(f"""
                <div style="background:#fef7fb; border-radius:8px; padding:15px; border-left:4px solid #{profit_color}; color:{profit_color}">
                    <h5 style="margin:0 0 10px 0;">总增收</h5>
                    <p style="font-size:18px; font-weight:bold;">{revenue_result['total']['总增收(元)']:.2f} 元</p>
                </div>
                """, unsafe_allow_html=True)
            
            # 每日收益详情（含总计）
            st.markdown("#### 每日收益详情（含总计）")
            daily_rev_df = revenue_result["daily"].copy()
            
            # 着色函数（仅对每日数据着色，总计行不着色）
            def color_profit(val):
                if val.name == "增收(元)":
                    colors = []
                    for x in val:
                        if pd.isna(x) or x == "总计":
                            colors.append("")
                        else:
                            colors.append("background-color: lightgreen" if x >= 0 else "background-color: lightcoral")
                    return colors
                return [""] * len(val)
            
            styled_df = daily_rev_df.style.apply(color_profit, axis=0)
            st.data_editor(
                styled_df,
                column_config={
                    "日期": st.column_config.TextColumn("日期", disabled=True),
                    "调整前收益(元)": st.column_config.NumberColumn("调整前收益(元)", format="%.2f"),
                    "调整后收益(元)": st.column_config.NumberColumn("调整后收益(元)", format="%.2f"),
                    "增收(元)": st.column_config.NumberColumn("增收(元)", format="%.2f")
                },
                disabled=True,
                hide_index=True,
                width='stretch'
            )
        else:
            st.warning("⚠️ 无有效复盘数据（可能未完成调整计算，或选中日期无数据）")
    
    st.divider()
    
    # ---------------------- 详细数据表格（多日） ----------------------
    st.subheader(f"📋 {date_title} 详细数据")
    if selected_dates:
        filtered_df = st.session_state.energy_data[st.session_state.energy_data["日期"].isin(selected_dates)].copy()
        if not filtered_df.empty:
            # 按日期+时刻排序
            filtered_df["日期_str"] = filtered_df["日期"].apply(lambda x: x.strftime("%Y-%m-%d"))
            filtered_df["时刻_order"] = filtered_df["时刻"].map({t: i for i, t in enumerate(FULL_96_TIMEPTS)})
            display_df = filtered_df.sort_values(["日期_str", "时刻_order"]).drop(columns=["日期_str", "时刻_order"]).copy()
            
            st.data_editor(
                display_df,
                column_config={
                    "日期": st.column_config.DateColumn("日期", disabled=True, format="YYYY-MM-DD"),
                    "时刻": st.column_config.TextColumn("时刻", disabled=True),
                    "日前节点电价(元/MWh)": st.column_config.NumberColumn("日前节点电价(元/MWh)", format="%.1f", disabled=True),
                    "实时节点电价(元/MWh)": st.column_config.NumberColumn("实时节点电价(元/MWh)", format="%.1f", disabled=True),
                    "日前预测出力(MW)": st.column_config.NumberColumn("日前预测出力(MW)", format="%.1f", disabled=True),
                    "实时出力(MW)": st.column_config.NumberColumn("实时出力(MW)", format="%.1f", disabled=True),
                    "日前调整后出力(MW)": st.column_config.TextColumn("日前调整后出力(MW)", disabled=True),
                    "新能源全省预测(MW)": st.column_config.NumberColumn("新能源全省预测(MW)", format="%.1f", disabled=True),
                    "新能源全省实测(MW)": st.column_config.NumberColumn("新能源全省实测(MW)", format="%.1f", disabled=True),
                    "非市场化机组预测(MW)": st.column_config.NumberColumn("非市场化机组预测(MW)", format="%.1f", disabled=True),
                    "非市场化机组实测(MW)": st.column_config.NumberColumn("非市场化机组实测(MW)", format="%.1f", disabled=True)
                },
                disabled=True,
                hide_index=True,
                width='stretch',
                height=400,
                column_order=[
                    "日期", "时刻",
                    "日前节点电价(元/MWh)", "实时节点电价(元/MWh)",
                    "日前预测出力(MW)", "实时出力(MW)", "日前调整后出力(MW)",
                    "新能源全省预测(MW)", "新能源全省实测(MW)",
                    "非市场化机组预测(MW)", "非市场化机组实测(MW)"
                ]
            )
        else:
            st.info("ℹ️ 所选日期无详细数据，请先上传对应日期数据")
    else:
        st.info("ℹ️ 请先选择目标日期")

if __name__ == "__main__":
    main()