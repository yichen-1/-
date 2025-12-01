import streamlit as st
import pandas as pd
import re
from io import BytesIO
import datetime
import plotly.express as px

# -------------------------- 页面配置 --------------------------
st.set_page_config(
    page_title="光伏/风电功率数据提取工具（24时段汇总版）",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------- 侧边栏配置 --------------------------
st.sidebar.header("⚙️ 配置项")
st.sidebar.subheader("📁 上传Excel文件")
uploaded_files = st.sidebar.file_uploader(
    "选择月度Excel文件（支持多选）",
    type=["xlsx", "xls", "xlsm"],
    accept_multiple_files=True
)

# 核心参数配置
file_keyword = st.sidebar.text_input("文件筛选关键词", value="历史趋势")
time_col_idx = st.sidebar.number_input("时间列索引（E列=4）", value=4, min_value=0)
power_col_idx_wind = st.sidebar.number_input("风电场功率列索引（J列=9）", value=9, min_value=0)
power_col_idx_pv = st.sidebar.number_input("光伏场功率列索引（F列=5）", value=5, min_value=0)
pv_stations = st.sidebar.text_input("光伏场站名单（逗号分隔）", value="浠水渔光,襄北农光")
power_conversion = st.sidebar.number_input("功率转换系数（kW→MW）", value=1000)
skip_rows = st.sidebar.number_input("跳过表头行数", value=1, min_value=0)

# 处理光伏场站名单
pv_stations_list = [s.strip() for s in pv_stations.split(",") if s.strip()]

# -------------------------- 核心工具函数 --------------------------
@st.cache_data(show_spinner="清洗功率数据中...")
def clean_power_data(value):
    """清洗功率列数据：保留含数字的功率值"""
    if pd.isna(value):
        return None
    value_str = str(value).strip()
    if not re.search(r'\d', value_str):
        return None
    num_match = re.search(r'(\d+\.?\d*)', value_str)
    if num_match:
        try:
            return float(num_match.group(1))
        except:
            return None
    return None

def extract_station_name(file_name):
    """从文件名提取场站名"""
    name_without_ext = file_name.split(".")[0]
    station_name = name_without_ext.split("-")[0].strip()
    return station_name

@st.cache_data(show_spinner="提取Excel数据中...")
def extract_excel_data(uploaded_file, time_idx, power_idx, skip_r, conv):
    """提取单个Excel文件数据（强制正序）"""
    try:
        file_name = uploaded_file.name
        suffix = file_name.split(".")[-1].lower()
        engine = "openpyxl" if suffix in ["xlsx", "xlsm"] else "xlrd"
        
        df = pd.read_excel(
            BytesIO(uploaded_file.getvalue()),
            header=None,
            usecols=[time_idx, power_idx],
            skiprows=skip_r,
            engine=engine,
            nrows=None
        )
        
        # 数据清洗
        df.columns = ["时间_原始", "功率_原始"]
        df["功率(kW)"] = df["功率_原始"].apply(clean_power_data)
        df["时间"] = pd.to_datetime(df["时间_原始"], errors="coerce")
        
        # 调试信息
        time_fail = df[df["时间"].isna()]
        power_fail = df[df["功率(kW)"].isna() & df["时间"].notna()]
        if not time_fail.empty:
            st.warning(f"⚠️ {file_name} 时间解析失败{len(time_fail)}条（前5条）：")
            st.dataframe(time_fail[["时间_原始", "功率_原始"]].head(5), use_container_width=True)
        if not power_fail.empty:
            st.warning(f"⚠️ {file_name} 功率清洗失败{len(power_fail)}条（前5条）：")
            st.dataframe(power_fail[["时间", "功率_原始"]].head(5), use_container_width=True)
        
        # 过滤无效数据并正序
        df = df.dropna(subset=["时间", "功率(kW)"])
        if df.empty:
            st.warning(f"⚠️ {file_name} 无有效数据")
            return pd.DataFrame(), file_name
        
        df = df.sort_values("时间").reset_index(drop=True)
        
        # 输出时间范围
        min_time = df["时间"].min()
        max_time = df["时间"].max()
        st.info(f"📄 {file_name} 有效数据：{min_time.strftime('%Y-%m-%d %H:%M')} ~ {max_time.strftime('%Y-%m-%d %H:%M')}（{len(df)}条）")
        
        # 转换单位并整理
        station_name = extract_station_name(file_name)
        df[station_name] = df["功率(kW)"] / conv
        df_result = df[["时间", station_name]].reset_index(drop=True)
        return df_result, file_name
    except Exception as e:
        st.error(f"处理 {file_name} 失败：{str(e)}")
        return pd.DataFrame(), file_name

# -------------------------- 24时段电量汇总核心函数 --------------------------
@st.cache_data(show_spinner="计算24时段电量汇总中...")
def calculate_24h_electricity(merged_df):
    """
    计算各场站24时段月度总上网电量
    逻辑：
    1. 自动识别时间间隔（15/30/60分钟）
    2. 电量 = 功率(MW) × 时间间隔(h) → 单位MWh
    3. 按小时时段（0-23）分组汇总
    """
    # 复制数据避免修改原数据
    df = merged_df.copy()
    
    # 1. 自动计算时间间隔（分钟）
    time_diff = df["时间"].diff().dropna()
    avg_interval_min = time_diff.dt.total_seconds().mean() / 60
    interval_h = avg_interval_min / 60  # 转换为小时
    st.info(f"⏱️ 自动识别数据采集间隔：{avg_interval_min:.0f}分钟（换算系数：{interval_h}小时/条）")
    
    # 2. 提取小时时段（0-23）
    df["小时时段"] = df["时间"].dt.hour
    
    # 3. 定义场站列表（排除时间列）
    stations = [col for col in df.columns if col not in ["时间", "小时时段"]]
    
    # 4. 计算每个时段的电量并汇总
    electricity_data = []
    for hour in range(24):
        hour_df = df[df["小时时段"] == hour].copy()
        row = {"小时时段": f"{hour:02d}:00"}  # 格式化显示（00:00, 01:00...23:00）
        
        for station in stations:
            # 电量 = 功率 × 时间间隔，求和得到该时段总电量
            total_electricity = (hour_df[station] * interval_h).sum()
            row[station] = round(total_electricity, 2)  # 保留2位小数
        
        electricity_data.append(row)
    
    # 转换为DataFrame并填充缺失值为0
    electricity_df = pd.DataFrame(electricity_data)
    electricity_df = electricity_df.fillna(0)
    
    # 计算各场站月度总电量（汇总行）
    total_row = {"小时时段": "月度总计"}
    for station in stations:
        total_row[station] = round(electricity_df[station].sum(), 2)
    electricity_df = pd.concat([electricity_df, pd.DataFrame([total_row])], ignore_index=True)
    
    return electricity_df, interval_h, stations

# -------------------------- 批量处理函数 --------------------------
def batch_extract_data(uploaded_files_list):
    # 筛选文件
    target_files = []
    for file in uploaded_files_list:
        if file_keyword in file.name or file_keyword.lower() in file.name.lower():
            target_files.append(file)
        else:
            st.warning(f"⚠️ {file.name} 不含关键词「{file_keyword}」，已跳过")
    
    if not target_files:
        st.error(f"❌ 未找到包含「{file_keyword}」的文件")
        return None, {}
    
    # 显示待处理文件
    st.info(f"✅ 找到 {len(target_files)} 个待处理文件：")
    file_list = []
    for i, f in enumerate(target_files, 1):
        station = extract_station_name(f.name)
        station_type = "📸 光伏" if station in pv_stations_list else "💨 风电"
        file_list.append(f"{i}. {station_type} {f.name}")
    st.code("\n".join(file_list))
    
    # 批量提取
    all_station_dfs = {}
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, file in enumerate(target_files):
        status_text.text(f"正在处理：{file.name}（{idx+1}/{len(target_files)}）")
        station_name = extract_station_name(file.name)
        power_idx = power_col_idx_pv if station_name in pv_stations_list else power_col_idx_wind
        file_data, file_name = extract_excel_data(file, time_col_idx, power_idx, skip_rows, power_conversion)
        
        if not file_data.empty:
            all_station_dfs[station_name] = file_data
            st.success(f"✅ {station_name}：提取到 {len(file_data)} 条有效数据")
        progress_bar.progress((idx + 1) / len(target_files))
    
    # 合并数据
    status_text.text("处理完成！开始合并数据...")
    if not all_station_dfs:
        st.error("❌ 未提取到任何有效数据")
        return None, {}
    
    df_list = list(all_station_dfs.values())
    merged_df = df_list[0]
    for df in df_list[1:]:
        merged_df = pd.merge(merged_df, df, on="时间", how="outer")
    
    merged_df["时间"] = merged_df["时间"].dt.floor("min")
    merged_df = merged_df.sort_values("时间").reset_index(drop=True)
    
    # 统计信息
    st.success("📊 数据合并完成！")
    st.info(f"""
    合并后统计：
    - 总时间点数：{len(merged_df)}
    - 包含场站：{', '.join(merged_df.columns[1:])}
    - 处理时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    """)
    
    # 各场站有效数据量
    st.subheader("各场站有效数据量")
    stat_data = []
    for station in all_station_dfs.keys():
        valid_count = merged_df[station].notna().sum()
        stat_data.append({"场站名": station, "有效数据条数": valid_count})
    st.dataframe(pd.DataFrame(stat_data), use_container_width=True)
    
    progress_bar.empty()
    status_text.empty()
    
    return merged_df, all_station_dfs

# -------------------------- 下载函数 --------------------------
def to_excel(df, sheet_name="数据"):
    """转换为Excel字节流"""
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    writer.close()
    output.seek(0)
    return output

# -------------------------- 主界面 --------------------------
st.title("📊 光伏/风电功率数据提取工具（24时段汇总版）")
st.markdown("---")

# 使用指引
st.info("""
### 📝 使用指引
1. 上传月度Excel数据文件（支持多选）
2. 确认列索引/光伏场站配置
3. 点击提取数据 → 自动生成24时段电量汇总
4. 预览汇总表格/图表，下载结果文件
""")

# 执行处理
if uploaded_files:
    if st.button("🚀 开始提取并汇总数据", type="primary"):
        with st.spinner("批量处理文件中..."):
            result_df, station_dfs = batch_extract_data(uploaded_files)
            
            if result_df is not None and not result_df.empty:
                # 基础数据预览
                st.markdown("---")
                st.subheader("📈 原始数据预览")
                min_time = result_df["时间"].min().strftime("%Y-%m-%d %H:%M")
                max_time = result_df["时间"].max().strftime("%Y-%m-%d %H:%M")
                st.success(f"✅ 原始数据时间范围：{min_time} ~ {max_time}（共{len(result_df)}条）")
                
                tab1, tab2 = st.tabs(["全部数据", "光伏场站数据"])
                with tab1:
                    st.markdown("**前20条（早期数据）**")
                    st.dataframe(result_df.head(20), use_container_width=True)
                    st.markdown("**后20条（后期数据）**")
                    st.dataframe(result_df.tail(20), use_container_width=True)
                with tab2:
                    pv_cols = [col for col in result_df.columns if col in pv_stations_list]
                    if pv_cols:
                        pv_df = result_df[["时间"] + pv_cols].dropna(subset=pv_cols, how="all").sort_values("时间")
                        st.markdown("**光伏数据前20条**")
                        st.dataframe(pv_df.head(20), use_container_width=True)
                        st.markdown("**光伏数据后20条**")
                        st.dataframe(pv_df.tail(20), use_container_width=True)
                    else:
                        st.info("暂无光伏场站数据")
                
                # 核心功能：24时段电量汇总
                st.markdown("---")
                st.subheader("🔋 24时段月度上网电量汇总（单位：MWh）")
                electricity_df, interval_h, stations = calculate_24h_electricity(result_df)
                
                # 显示汇总表格
                st.dataframe(electricity_df, use_container_width=True)
                
                # 可视化图表
                st.subheader("📊 24时段电量趋势图")
                # 转换为长格式用于绘图
                plot_df = electricity_df[electricity_df["小时时段"] != "月度总计"].copy()
                plot_df_melt = plot_df.melt(
                    id_vars=["小时时段"],
                    value_vars=stations,
                    var_name="场站名称",
                    value_name="上网电量(MWh)"
                )
                
                # 绘制趋势图
                fig = px.line(
                    plot_df_melt,
                    x="小时时段",
                    y="上网电量(MWh)",
                    color="场站名称",
                    title="各场站24时段月度上网电量趋势",
                    markers=True,
                    template="plotly_white"
                )
                fig.update_layout(
                    xaxis_title="小时时段",
                    yaxis_title="上网电量(MWh)",
                    width=1000,
                    height=600
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # 下载功能
                st.markdown("---")
                st.subheader("📥 下载结果")
                current_month = datetime.datetime.now().strftime("%Y%m")
                
                # 下载原始整合数据
                raw_excel = to_excel(result_df, "原始整合数据")
                st.download_button(
                    label="下载原始整合数据（Excel）",
                    data=raw_excel,
                    file_name=f"原始整合数据_{current_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # 下载24时段汇总数据
                electricity_excel = to_excel(electricity_df, "24时段电量汇总")
                st.download_button(
                    label="下载24时段电量汇总（Excel）",
                    data=electricity_excel,
                    file_name=f"24时段电量汇总_{current_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
else:
    st.warning("⚠️ 请先在左侧侧边栏上传Excel数据文件！")

# 侧边栏说明
st.sidebar.markdown("---")
st.sidebar.markdown("### 📝 使用说明")
st.sidebar.markdown("""
1. 上传历史趋势Excel文件（支持多选）
2. 列索引配置（索引从0开始）：
   - 时间列：E列=4
   - 风电功率列：J列=9
   - 光伏功率列：F列=5
3. 点击提取按钮，自动完成：
   - 数据清洗与合并
   - 24时段电量计算（单位：MWh）
   - 生成汇总表格和趋势图
4. 下载结果文件存档
""")

st.sidebar.markdown("### ℹ️ 电量计算逻辑")
st.sidebar.markdown("""
- 自动识别数据采集间隔（15/30/60分钟）
- 单条电量 = 功率(MW) × 间隔小时数
- 时段总电量 = 该时段所有记录电量求和
- 单位：兆瓦时（MWh）
""")
