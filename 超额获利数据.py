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

# -------------------------- 初始化会话状态（核心修正：替代全局变量） --------------------------
if "generated_data" not in st.session_state:
    st.session_state.generated_data = {}  # 实发数据：{场站名: 月度实发总量MWh}
if "hold_data" not in st.session_state:
    st.session_state.hold_data = {}        # 持仓数据：{场站名: 月度净持有电量MWh}
if "extracted_stations" not in st.session_state:
    st.session_state.extracted_stations = []  # 已提取的场站列表
if "merged_gen_df" not in st.session_state:
    st.session_state.merged_gen_df = pd.DataFrame()  # 实发合并数据

# -------------------------- 侧边栏分类配置 --------------------------
st.sidebar.title("⚙️ 功能配置")

# 1. 场站实发配置
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

# 2. 中长期持仓配置
with st.sidebar.expander("📦 中长期持仓配置", expanded=True):
    st.sidebar.subheader("1. 持仓数据上传")
    uploaded_hold_files = st.sidebar.file_uploader(
        "上传场站持仓Excel文件（支持多选）",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="hold_upload"
    )

    st.sidebar.subheader("2. 持仓数据关联")
    # 下拉选择持仓文件对应的场站（支持多选关联）
    selected_stations = st.sidebar.multiselect(
        "选择持仓文件对应的场站（可多选）",
        options=st.session_state.extracted_stations,
        key="hold_station_select"
    )
    # 手动输入未提取的场站
    manual_station = st.sidebar.text_input(
        "手动补充场站名（逗号分隔）",
        placeholder="例如：新场站1,新场站2",
        key="hold_station_manual"
    )
    # 合并选择和手动输入的场站
    target_stations = selected_stations + [s.strip() for s in manual_station.split(",") if s.strip()]

    st.sidebar.subheader("3. 持仓列配置")
    hold_col_idx = st.sidebar.number_input("净持有电量列索引（D列=3）", value=3, min_value=0, key="hold_col_idx")
    hold_skip_rows = st.sidebar.number_input("持仓表格跳过表头行数", value=1, min_value=0, key="hold_skip_rows")

# 处理光伏场站名单
pv_stations_list = [s.strip() for s in pv_stations.split(",") if s.strip()]

# -------------------------- 核心工具函数 --------------------------
# 1. 实发数据清洗
@st.cache_data(show_spinner="清洗实发功率数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
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

# 2. 提取场站名
def extract_station_name(file_name):
    name_without_ext = file_name.split(".")[0]
    station_name = name_without_ext.split("-")[0].strip()
    return station_name

# 3. 提取实发数据（优化缓存key）
@st.cache_data(show_spinner="提取实发Excel数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
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
            return pd.DataFrame(), file_name, ""
        
        df = df.sort_values("时间").reset_index(drop=True)
        station_name = extract_station_name(file_name)
        df[station_name] = df["功率(kW)"] / conv
        df_result = df[["时间", station_name]].reset_index(drop=True)
        return df_result, file_name, station_name
    except Exception as e:
        st.error(f"处理实发文件[{file_name}]失败：{str(e)}")
        return pd.DataFrame(), file_name, ""

# 4. 提取持仓数据（支持批量关联）
@st.cache_data(show_spinner="提取持仓Excel数据中...", hash_funcs={BytesIO: lambda x: x.getvalue()})
def extract_hold_data(uploaded_file, hold_col_idx, skip_r):
    """提取持仓表格的D列（净持有电量），返回该文件的总净持有电量"""
    try:
        file_name = uploaded_file.name
        suffix = file_name.split(".")[-1].lower()
        engine = "openpyxl" if suffix in ["xlsx", "xlsm"] else "xlrd"
        
        df = pd.read_excel(
            BytesIO(uploaded_file.getvalue()),
            header=None,
            usecols=[hold_col_idx],
            skiprows=skip_r,
            engine=engine,
            nrows=None
        )
        
        df.columns = ["净持有电量"]
        df["净持有电量"] = pd.to_numeric(df["净持有电量"], errors="coerce").fillna(0)
        total_hold = round(df["净持有电量"].sum(), 2)
        
        st.success(f"✅ 持仓文件[{file_name}]处理完成：当月总净持有电量={total_hold} MWh")
        return total_hold
    except Exception as e:
        st.error(f"处理持仓文件[{file_name}]失败：{str(e)}")
        return 0.0

# 5. 24时段实发汇总（重构月度总计行）
@st.cache_data(show_spinner="计算24时段实发汇总中...")
def calculate_24h_generated(merged_df):
    df = merged_df.copy()
    # 容错：空数据直接返回空表
    if df.empty:
        st.warning("⚠️ 实发数据为空，无法计算24时段汇总")
        return pd.DataFrame(), 0, []
    
    # 计算时间间隔
    time_diff = df["时间"].diff().dropna()
    avg_interval_min = time_diff.dt.total_seconds().mean() / 60 if not time_diff.empty else 15
    interval_h = avg_interval_min / 60
    st.info(f"⏱️ 实发数据采集间隔：{avg_interval_min:.0f}分钟（换算系数：{interval_h}小时/条）")
    
    # 提取小时时段
    df["小时时段"] = df["时间"].dt.hour
    stations = [col for col in df.columns if col not in ["时间", "小时时段"]]
    
    # 按小时汇总
    generated_rows = []
    for hour in range(24):
        hour_df = df[df["小时时段"] == hour].copy()
        row = {"小时时段": f"{hour:02d}:00"}
        for station in stations:
            total_gen = (hour_df[station] * interval_h).sum()
            row[station] = round(total_gen, 2)
        generated_rows.append(row)
    
    # 生成汇总表
    generated_df = pd.DataFrame(generated_rows).fillna(0)
    
    # 计算月度实发总量并更新会话状态
    st.session_state.generated_data = {}
    total_row = {"小时时段": "月度实发总量"}
    for station in stations:
        total_month_gen = round(generated_df[station].sum(), 2)
        st.session_state.generated_data[station] = total_month_gen
        total_row[station] = total_month_gen
    
    # 追加月度总计行（重构：避免索引混乱）
    generated_df = pd.concat([generated_df, pd.DataFrame([total_row])], ignore_index=True)
    
    return generated_df, interval_h, stations

# -------------------------- 批量处理实发数据 --------------------------
def batch_process_generated(uploaded_files_list):
    # 筛选文件
    target_files = [f for f in uploaded_files_list if file_keyword in f.name or file_keyword.lower() in f.name.lower()]
    if not target_files:
        st.error(f"❌ 未找到包含「{file_keyword}」的实发文件")
        return pd.DataFrame(), []
    
    # 显示待处理文件
    st.info(f"✅ 找到 {len(target_files)} 个待处理实发文件：")
    file_list = []
    extracted_stations = []
    for i, f in enumerate(target_files, 1):
        station = extract_station_name(f.name)
        extracted_stations.append(station)
        station_type = "📸 光伏" if station in pv_stations_list else "💨 风电"
        file_list.append(f"{i}. {station_type} {station}（文件：{f.name}）")
    st.code("\n".join(file_list))
    
    # 批量提取
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
    
    # 合并数据（容错：空数据处理）
    if not all_station_dfs:
        st.error("❌ 未提取到任何有效实发数据")
        return pd.DataFrame(), []
    
    df_list = list(all_station_dfs.values())
    merged_df = df_list[0]
    for df in df_list[1:]:
        merged_df = pd.merge(merged_df, df, on="时间", how="outer")
    
    merged_df["时间"] = merged_df["时间"].dt.floor("min")
    merged_df = merged_df.sort_values("时间").reset_index(drop=True)
    
    # 更新会话状态
    st.session_state.extracted_stations = extracted_stations
    st.session_state.merged_gen_df = merged_df
    
    progress_bar.empty()
    status_text.empty()
    return merged_df, extracted_stations

# -------------------------- 数据关联展示 --------------------------
def show_related_data():
    st.markdown("---")
    st.subheader("🔗 场站实发与中长期持仓关联结果")
    
    # 容错提示
    if not st.session_state.generated_data and not st.session_state.hold_data:
        st.warning("⚠️ 暂无实发或持仓数据，请先处理对应文件")
        return
    
    # 生成关联表
    related_data = []
    # 合并所有涉及的场站
    all_stations = list(set(list(st.session_state.generated_data.keys()) + list(st.session_state.hold_data.keys())))
    
    for station in all_stations:
        gen_total = st.session_state.generated_data.get(station, 0.0)
        hold_total = st.session_state.hold_data.get(station, 0.0)
        # 容错：除以0处理
        coverage = round((hold_total / gen_total * 100) if gen_total > 0 else 0, 2)
        
        related_data.append({
            "场站名": station,
            "当月实发总量（MWh）": gen_total,
            "当月中长期持仓（MWh）": hold_total,
            "持仓覆盖度": f"{coverage}%",
            "实发-持仓差值（MWh）": round(gen_total - hold_total, 2)
        })
    
    related_df = pd.DataFrame(related_data)
    st.dataframe(related_df, use_container_width=True)
    
    # 可视化（容错：空数据不绘图）
    if not related_df.empty:
        fig = px.bar(
            related_df,
            x="场站名",
            y=["当月实发总量（MWh）", "当月中长期持仓（MWh）"],
            barmode="group",
            title="各场站实发总量与中长期持仓对比",
            template="plotly_white",
            color_discrete_map={
                "当月实发总量（MWh）": "#1f77b4",
                "当月中长期持仓（MWh）": "#ff7f0e"
            }
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
    if df.empty:
        st.warning("⚠️ 数据为空，无法生成Excel")
        return BytesIO()
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    writer.close()
    output.seek(0)
    return output

# -------------------------- 主界面 --------------------------
st.title("📊 光伏/风电数据管理工具（实发+持仓关联版）")
st.markdown("---")

# 1. 处理实发数据
if uploaded_generated_files:
    if st.button("🚀 开始处理场站实发数据", type="primary", key="process_generated"):
        with st.spinner("批量处理实发文件中..."):
            merged_gen_df, stations = batch_process_generated(uploaded_generated_files)
            
            if not merged_gen_df.empty:
                # 实发数据预览
                st.markdown("---")
                st.subheader("📈 场站实发原始数据预览")
                min_time = merged_gen_df["时间"].min().strftime("%Y-%m-%d %H:%M") if not merged_gen_df.empty else "无"
                max_time = merged_gen_df["时间"].max().strftime("%Y-%m-%d %H:%M") if not merged_gen_df.empty else "无"
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
                
                # 24时段汇总
                st.markdown("---")
                st.subheader("🔋 场站24时段月度实发汇总（单位：MWh）")
                gen_24h_df, interval_h, stations = calculate_24h_generated(merged_gen_df)
                if not gen_24h_df.empty:
                    st.dataframe(gen_24h_df, use_container_width=True)
                
                # 下载
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
                    key="download_gen_raw",
                    disabled=merged_gen_df.empty
                )
                st.download_button(
                    label="下载24时段实发汇总数据",
                    data=gen_24h_excel,
                    file_name=f"场站24时段实发汇总_{current_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_gen_24h",
                    disabled=gen_24h_df.empty
                )

# 2. 处理持仓数据
if uploaded_hold_files and target_stations:
    if st.button("📦 开始处理中长期持仓数据", type="primary", key="process_hold"):
        with st.spinner("处理持仓文件并关联场站..."):
            # 计算所有持仓文件的总电量
            total_hold_all = 0.0
            for file in uploaded_hold_files:
                total_hold = extract_hold_data(file, hold_col_idx, hold_skip_rows)
                total_hold_all += total_hold
            
            # 按场站分配（均分，或可自定义分配逻辑）
            hold_per_station = round(total_hold_all / len(target_stations), 2) if target_stations else 0.0
            for station in target_stations:
                st.session_state.hold_data[station] = hold_per_station
                st.success(f"✅ 持仓数据关联到场站[{station}]：{hold_per_station} MWh")
            
            st.success(f"✅ 所有持仓文件处理完成！总净持有电量={total_hold_all} MWh，已分配到{len(target_stations)}个场站")

# 3. 展示关联结果
show_related_data()

# 4. 重置数据按钮（新增：解决数据残留问题）
st.markdown("---")
if st.button("🗑️ 重置所有数据（实发+持仓）", type="secondary"):
    st.session_state.generated_data = {}
    st.session_state.hold_data = {}
    st.session_state.extracted_stations = []
    st.session_state.merged_gen_df = pd.DataFrame()
    st.success("✅ 所有数据已重置！")

# -------------------------- 侧边栏说明 --------------------------
st.sidebar.markdown("---")
st.sidebar.markdown("### 📝 使用指引")
st.sidebar.markdown("""
1. 实发处理：上传文件→确认配置→点击处理→生成24时段汇总
2. 持仓处理：上传文件→选择/输入关联场站→点击处理→自动分配数据
3. 关联查看：自动展示实发-持仓对比表+图表
4. 数据重置：若数据异常，点击「重置所有数据」重新处理
""")

st.sidebar.markdown("### ℹ️ 注意事项")
st.sidebar.markdown("""
- 场站名需一致（大小写敏感）
- 持仓数据默认均分至所选场站（可自定义分配逻辑）
- 所有数据存储在会话中，刷新页面不丢失（关闭页面重置）
- 支持.xlsx/.xls/.xlsm格式，建议优先使用.xlsx
""")
