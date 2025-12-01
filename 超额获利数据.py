import streamlit as st
import pandas as pd
import re
from io import BytesIO
import datetime
import plotly.express as px

# -------------------------- 页面配置 --------------------------
st.set_page_config(
    page_title="光伏/风电数据管理工具（实发+持仓）",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------- 全局变量初始化（用于数据关联） --------------------------
# 存储实发数据（key：场站名，value：月度实发总量MWh）
generated_data = {}
# 存储中长期持仓数据（key：场站名，value：月度净持有电量）
hold_data = {}
# 存储已提取的场站列表
extracted_stations = []

# -------------------------- 侧边栏分类配置（核心优化：分折叠面板归纳） --------------------------
st.sidebar.title("⚙️ 功能配置")

# 1. 场站实发配置（原配置项归纳）
with st.sidebar.expander("📊 场站实发配置", expanded=True):
    st.sidebar.subheader("1. 数据上传")
    uploaded_generated_files = st.sidebar.file_uploader(
        "上传场站实发Excel文件（支持多选）",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="generated_upload"
    )

    st.sidebar.subheader("2. 列索引配置")
    time_col_idx = st.sidebar.number_input("时间列索引（E列=4）", value=4, min_value=0, key="time_idx")
    power_col_idx_wind = st.sidebar.number_input("风电场功率列索引（J列=9）", value=9, min_value=0, key="wind_power_idx")
    power_col_idx_pv = st.sidebar.number_input("光伏场功率列索引（F列=5）", value=5, min_value=0, key="pv_power_idx")

    st.sidebar.subheader("3. 基础参数")
    pv_stations = st.sidebar.text_input("光伏场站名单（逗号分隔）", value="浠水渔光,襄北农光", key="pv_list")
    power_conversion = st.sidebar.number_input("功率转换系数（kW→MW）", value=1000, key="power_conv")
    skip_rows = st.sidebar.number_input("跳过表头行数", value=1, min_value=0, key="skip_rows")
    file_keyword = st.sidebar.text_input("实发文件筛选关键词", value="历史趋势", key="generated_keyword")

# 2. 中长期持仓配置（新增功能）
with st.sidebar.expander("📦 中长期持仓配置", expanded=True):
    st.sidebar.subheader("1. 持仓数据上传")
    uploaded_hold_files = st.sidebar.file_uploader(
        "上传场站持仓Excel文件（支持多选）",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="hold_upload"
    )

    st.sidebar.subheader("2. 持仓数据关联")
    # 下拉选择持仓文件对应的场站（默认显示已提取的实发场站，无则手动输入）
    if extracted_stations:
        selected_station = st.sidebar.selectbox(
            "选择当前持仓文件对应的场站",
            options=[""] + extracted_stations,
            key="hold_station_select"
        )
    else:
        selected_station = st.sidebar.text_input(
            "手动输入持仓文件对应的场站名",
            placeholder="例如：浠水渔光",
            key="hold_station_input"
        )

    st.sidebar.subheader("3. 持仓列配置")
    hold_col_idx = st.sidebar.number_input("净持有电量列索引（D列=3）", value=3, min_value=0, key="hold_col_idx")
    hold_skip_rows = st.sidebar.number_input("持仓表格跳过表头行数", value=1, min_value=0, key="hold_skip_rows")

# 处理光伏场站名单为列表
pv_stations_list = [s.strip() for s in pv_stations.split(",") if s.strip()]

# -------------------------- 核心工具函数 --------------------------
# 1. 实发数据处理函数（保留原有逻辑）
@st.cache_data(show_spinner="清洗实发功率数据中...")
def clean_power_data(value):
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
    name_without_ext = file_name.split(".")[0]
    station_name = name_without_ext.split("-")[0].strip()
    return station_name

@st.cache_data(show_spinner="提取实发Excel数据中...")
def extract_generated_data(uploaded_file, time_idx, power_idx, skip_r, conv):
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
        
        df.columns = ["时间_原始", "功率_原始"]
        df["功率(kW)"] = df["功率_原始"].apply(clean_power_data)
        df["时间"] = pd.to_datetime(df["时间_原始"], errors="coerce")
        
        # 调试信息
        time_fail = df[df["时间"].isna()]
        power_fail = df[df["功率(kW)"].isna() & df["时间"].notna()]
        if not time_fail.empty:
            st.warning(f"⚠️ 实发文件[{file_name}]：时间解析失败{len(time_fail)}条（前5条）")
            st.dataframe(time_fail[["时间_原始", "功率_原始"]].head(5), use_container_width=True)
        if not power_fail.empty:
            st.warning(f"⚠️ 实发文件[{file_name}]：功率清洗失败{len(power_fail)}条（前5条）")
            st.dataframe(power_fail[["时间", "功率_原始"]].head(5), use_container_width=True)
        
        df = df.dropna(subset=["时间", "功率(kW)"])
        if df.empty:
            st.warning(f"⚠️ 实发文件[{file_name}]无有效数据")
            return pd.DataFrame(), file_name
        
        df = df.sort_values("时间").reset_index(drop=True)
        station_name = extract_station_name(file_name)
        df[station_name] = df["功率(kW)"] / conv
        df_result = df[["时间", station_name]].reset_index(drop=True)
        return df_result, file_name, station_name
    except Exception as e:
        st.error(f"处理实发文件[{file_name}]失败：{str(e)}")
        return pd.DataFrame(), file_name, ""

# 2. 中长期持仓数据处理函数（新增）
@st.cache_data(show_spinner="提取持仓Excel数据中...")
def extract_hold_data(uploaded_file, hold_col_idx, skip_r, target_station):
    """提取持仓表格的D列（净持有电量），并关联到指定场站"""
    try:
        file_name = uploaded_file.name
        suffix = file_name.split(".")[-1].lower()
        engine = "openpyxl" if suffix in ["xlsx", "xlsm"] else "xlrd"
        
        # 仅读取净持有电量列（D列=3）
        df = pd.read_excel(
            BytesIO(uploaded_file.getvalue()),
            header=None,
            usecols=[hold_col_idx],
            skiprows=skip_r,
            engine=engine,
            nrows=None
        )
        
        df.columns = ["净持有电量"]
        # 清洗净持有电量（过滤非数值）
        df["净持有电量"] = pd.to_numeric(df["净持有电量"], errors="coerce").fillna(0)
        # 计算当月总净持有电量（求和）
        total_hold = round(df["净持有电量"].sum(), 2)
        
        st.success(f"✅ 持仓文件[{file_name}]处理完成：")
        st.info(f"关联场站：{target_station} | 当月总净持有电量：{total_hold}")
        return total_hold, target_station
    except Exception as e:
        st.error(f"处理持仓文件[{file_name}]失败：{str(e)}")
        return 0.0, target_station

# 3. 24时段实发汇总函数（保留原有逻辑）
@st.cache_data(show_spinner="计算24时段实发汇总中...")
def calculate_24h_generated(merged_df):
    df = merged_df.copy()
    time_diff = df["时间"].diff().dropna()
    avg_interval_min = time_diff.dt.total_seconds().mean() / 60
    interval_h = avg_interval_min / 60
    st.info(f"⏱️ 实发数据采集间隔：{avg_interval_min:.0f}分钟（换算系数：{interval_h}小时/条）")
    
    df["小时时段"] = df["时间"].dt.hour
    stations = [col for col in df.columns if col not in ["时间", "小时时段"]]
    
    generated_data = []
    for hour in range(24):
        hour_df = df[df["小时时段"] == hour].copy()
        row = {"小时时段": f"{hour:02d}:00"}
        for station in stations:
            total_gen = (hour_df[station] * interval_h).sum()
            row[station] = round(total_gen, 2)
        generated_data.append(row)
    
    generated_df = pd.DataFrame(generated_data).fillna(0)
    # 计算各场站月度实发总量（存入全局变量用于关联）
    global generated_data 
    generated_data = {}
    for station in stations:
        total_month_gen = round(generated_df[station].sum(), 2)
        generated_data[station] = total_month_gen
        generated_df.loc[len(generated_df), station] = total_month_gen
    
    generated_df["小时时段"] = generated_df["小时时段"].fillna("月度实发总量")
    return generated_df, interval_h, stations

# -------------------------- 批量处理函数 --------------------------
def batch_process_generated(uploaded_files_list):
    # 筛选实发文件
    target_files = []
    for file in uploaded_files_list:
        if file_keyword in file.name or file_keyword.lower() in file.name.lower():
            target_files.append(file)
        else:
            st.warning(f"⚠️ 实发文件[{file.name}]不含关键词「{file_keyword}」，已跳过")
    
    if not target_files:
        st.error(f"❌ 未找到包含「{file_keyword}」的实发文件")
        return None, []
    
    # 显示待处理实发文件
    st.info(f"✅ 找到 {len(target_files)} 个待处理实发文件：")
    file_list = []
    global extracted_stations 
    extracted_stations = []
    for i, f in enumerate(target_files, 1):
        station = extract_station_name(f.name)
        extracted_stations.append(station)
        station_type = "📸 光伏" if station in pv_stations_list else "💨 风电"
        file_list.append(f"{i}. {station_type} {station}（文件：{f.name}）")
    st.code("\n".join(file_list))
    
    # 批量提取实发数据
    all_station_dfs = {}
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, file in enumerate(target_files):
        status_text.text(f"正在处理实发文件：{file.name}（{idx+1}/{len(target_files)}）")
        station_name = extract_station_name(file.name)
        power_idx = power_col_idx_pv if station_name in pv_stations_list else power_col_idx_wind
        file_data, file_name, station = extract_generated_data(file, time_col_idx, power_idx, skip_rows, power_conversion)
        
        if not file_data.empty and station:
            all_station_dfs[station] = file_data
            st.success(f"✅ 场站[{station}]：提取到 {len(file_data)} 条实发数据")
        progress_bar.progress((idx + 1) / len(target_files))
    
    # 合并实发数据
    status_text.text("实发数据处理完成！开始合并...")
    if not all_station_dfs:
        st.error("❌ 未提取到任何有效实发数据")
        return None, []
    
    df_list = list(all_station_dfs.values())
    merged_df = df_list[0]
    for df in df_list[1:]:
        merged_df = pd.merge(merged_df, df, on="时间", how="outer")
    
    merged_df["时间"] = merged_df["时间"].dt.floor("min")
    merged_df = merged_df.sort_values("时间").reset_index(drop=True)
    
    progress_bar.empty()
    status_text.empty()
    return merged_df, extracted_stations

# -------------------------- 数据关联与展示函数 --------------------------
def show_related_data():
    """展示实发与持仓的关联结果"""
    st.markdown("---")
    st.subheader("🔗 场站实发与中长期持仓关联结果")
    
    # 检查是否有实发和持仓数据
    if not generated_data:
        st.warning("⚠️ 暂未提取到场站实发数据，请先处理「场站实发配置」中的文件")
        return
    if not hold_data:
        st.warning("⚠️ 暂未上传场站持仓数据，请先处理「中长期持仓配置」中的文件")
        return
    
    # 生成关联结果表格
    related_data = []
    for station in generated_data.keys():
        related_data.append({
            "场站名": station,
            "当月实发总量（MWh）": generated_data.get(station, 0),
            "当月中长期持仓（净持有电量）": hold_data.get(station, 0),
            "持仓覆盖度（持仓/实发）": f"{round((hold_data.get(station, 0)/generated_data.get(station, 1))*100, 2)}%"
        })
    
    related_df = pd.DataFrame(related_data)
    st.dataframe(related_df, use_container_width=True)
    
    # 可视化关联结果（双轴图）
    fig = px.bar(
        related_df,
        x="场站名",
        y=["当月实发总量（MWh）", "当月中长期持仓（净持有电量）"],
        barmode="group",
        title="各场站实发总量与中长期持仓对比",
        template="plotly_white"
    )
    fig.update_layout(
        xaxis_title="场站名",
        yaxis_title="电量（MWh）",
        width=1000,
        height=600
    )
    st.plotly_chart(fig, use_container_width=True)

# -------------------------- 下载函数 --------------------------
def to_excel(df, sheet_name="数据"):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    writer.close()
    output.seek(0)
    return output

# -------------------------- 主界面 --------------------------
st.title("📊 光伏/风电数据管理工具（实发+持仓关联版）")
st.markdown("---")

# 1. 实发数据处理流程
if uploaded_generated_files:
    if st.button("🚀 开始处理场站实发数据", type="primary", key="process_generated"):
        with st.spinner("批量处理实发文件中..."):
            merged_gen_df, stations = batch_process_generated(uploaded_generated_files)
            
            if merged_gen_df is not None and not merged_gen_df.empty:
                # 实发数据预览
                st.markdown("---")
                st.subheader("📈 场站实发原始数据预览")
                min_time = merged_gen_df["时间"].min().strftime("%Y-%m-%d %H:%M")
                max_time = merged_gen_df["时间"].max().strftime("%Y-%m-%d %H:%M")
                st.success(f"✅ 实发数据时间范围：{min_time} ~ {max_time}（共{len(merged_gen_df)}条）")
                
                tab1, tab2 = st.tabs(["全部实发数据", "光伏场站实发数据"])
                with tab1:
                    st.markdown("**前20条（早期数据）**")
                    st.dataframe(merged_gen_df.head(20), use_container_width=True)
                    st.markdown("**后20条（后期数据）**")
                    st.dataframe(merged_gen_df.tail(20), use_container_width=True)
                with tab2:
                    pv_cols = [col for col in merged_gen_df.columns if col in pv_stations_list]
                    if pv_cols:
                        pv_df = merged_gen_df[["时间"] + pv_cols].dropna(subset=pv_cols, how="all").sort_values("时间")
                        st.markdown("**光伏数据前20条**")
                        st.dataframe(pv_df.head(20), use_container_width=True)
                        st.markdown("**光伏数据后20条**")
                        st.dataframe(pv_df.tail(20), use_container_width=True)
                    else:
                        st.info("暂无光伏场站实发数据")
                
                # 24时段实发汇总
                st.markdown("---")
                st.subheader("🔋 场站24时段月度实发汇总（单位：MWh）")
                gen_24h_df, interval_h, stations = calculate_24h_generated(merged_gen_df)
                st.dataframe(gen_24h_df, use_container_width=True)
                
                # 实发数据下载
                st.markdown("---")
                st.subheader("📥 实发数据下载")
                current_month = datetime.datetime.now().strftime("%Y%m")
                gen_raw_excel = to_excel(merged_gen_df, "实发原始数据")
                gen_24h_excel = to_excel(gen_24h_df, "24时段实发汇总")
                
                st.download_button(
                    label="下载实发原始整合数据",
                    data=gen_raw_excel,
                    file_name=f"场站实发原始数据_{current_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_gen_raw"
                )
                st.download_button(
                    label="下载24时段实发汇总数据",
                    data=gen_24h_excel,
                    file_name=f"场站24时段实发汇总_{current_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_gen_24h"
                )

# 2. 中长期持仓数据处理流程（新增）
if uploaded_hold_files and selected_station.strip():
    if st.button("📦 开始处理中长期持仓数据", type="primary", key="process_hold"):
        with st.spinner("处理持仓文件并关联场站..."):
            global hold_data 
            for file in uploaded_hold_files:
                total_hold, station = extract_hold_data(file, hold_col_idx, hold_skip_rows, selected_station.strip())
                if station and total_hold > 0:
                    hold_data[station] = total_hold  # 存入全局变量，key为场站名
            st.success("✅ 中长期持仓数据处理完成！已关联到场站")

# 3. 展示关联结果（无论先处理实发还是持仓，都能触发）
if generated_data or hold_data:
    show_related_data()

# 4. 无数据时的提示
if not uploaded_generated_files and not uploaded_hold_files:
    st.warning("⚠️ 请在左侧侧边栏上传「场站实发文件」或「中长期持仓文件」开始处理")

# -------------------------- 侧边栏使用说明 --------------------------
st.sidebar.markdown("---")
st.sidebar.markdown("### 📝 使用流程指引")
st.sidebar.markdown("""
#### 1. 场站实发处理
1. 上传含「历史趋势」关键词的实发Excel文件
2. 确认列索引（时间列=4，风电功率=9，光伏功率=5）
3. 点击「开始处理场站实发数据」→ 生成24时段汇总

#### 2. 中长期持仓处理
1. 上传持仓Excel文件（D列为净持有电量）
2. 选择/输入持仓文件对应的场站名（需与实发场站一致）
3. 点击「开始处理中长期持仓数据」→ 自动关联

#### 3. 数据关联查看
处理完成后，自动显示「实发总量-持仓」关联表和对比图
""")

st.sidebar.markdown("### ℹ️ 关键说明")
st.sidebar.markdown("""
- 实发与持仓关联的核心：**场站名必须完全一致**
- 实发总量单位：MWh（兆瓦时）
- 持仓数据：取D列（净持有电量）当月总和
- 持仓覆盖度=（持仓/实发）×100%
""")
