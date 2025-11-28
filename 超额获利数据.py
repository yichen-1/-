import streamlit as st
import pandas as pd
import re
from io import BytesIO
import datetime

# -------------------------- 页面配置 --------------------------
st.set_page_config(
    page_title="光伏/风电功率数据提取工具（导入版）",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------- 侧边栏配置（保留核心参数） --------------------------
st.sidebar.header("⚙️ 配置项")
# 移除文件夹路径，改为文件上传
st.sidebar.subheader("📁 上传Excel文件")
uploaded_files = st.sidebar.file_uploader(
    "选择月度Excel文件（支持多选）",
    type=["xlsx", "xls", "xlsm"],
    accept_multiple_files=True
)

# 核心参数配置（不变）
file_keyword = st.sidebar.text_input("文件筛选关键词（仅显示含该关键词的文件）", value="历史趋势")
time_col_idx = st.sidebar.number_input("时间列索引（E列=4）", value=4, min_value=0)
power_col_idx_wind = st.sidebar.number_input("风电场功率列索引（J列=9）", value=9, min_value=0)
power_col_idx_pv = st.sidebar.number_input("光伏场功率列索引（F列=5）", value=5, min_value=0)
pv_stations = st.sidebar.text_input("光伏场站名单（逗号分隔）", value="浠水渔光,襄北农光")
power_conversion = st.sidebar.number_input("功率转换系数（kW→MW）", value=1000)
skip_rows = st.sidebar.number_input("跳过表头行数", value=1, min_value=0)

# 处理光伏场站名单为列表
pv_stations_list = [s.strip() for s in pv_stations.split(",") if s.strip()]

# -------------------------- 核心工具函数 --------------------------
@st.cache_data(show_spinner="清洗功率数据中...")
def clean_power_data(value):
    """清洗功率列数据：提取数值，过滤文本/特殊字符"""
    if pd.isna(value):
        return None
    value_str = str(value).strip()
    if re.match(r'^[^\d.]+$', value_str):
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
    name_without_ext = file_name.split(".")[0]  # 处理上传文件的文件名（无路径）
    station_name = name_without_ext.split("-")[0].strip()
    return station_name

@st.cache_data(show_spinner="提取Excel数据中...")
def extract_excel_data(uploaded_file, time_idx, power_idx, skip_r, conv):
    """提取单个上传Excel文件的数据（适配BytesIO流）"""
    try:
        # 识别文件格式，选择引擎
        file_name = uploaded_file.name
        suffix = file_name.split(".")[-1].lower()
        engine = "openpyxl" if suffix in ["xlsx", "xlsm"] else "xlrd"
        
        # 读取上传的文件流
        df = pd.read_excel(
            BytesIO(uploaded_file.getvalue()),  # 转换为字节流
            header=None,
            usecols=[time_idx, power_idx],
            skiprows=skip_r,
            engine=engine
        )
        
        # 数据清洗
        df.columns = ["时间", "功率(kW)"]
        df["功率(kW)"] = df["功率(kW)"].apply(clean_power_data)
        df["时间"] = pd.to_datetime(df["时间"], errors="coerce")
        df = df.dropna(subset=["时间", "功率(kW)"])
        
        if df.empty:
            return pd.DataFrame(), file_name
        
        # 提取场站名并转换单位
        station_name = extract_station_name(file_name)
        df[station_name] = df["功率(kW)"] / conv
        df_result = df[["时间", station_name]].reset_index(drop=True)
        return df_result, file_name
    except Exception as e:
        st.error(f"处理 {uploaded_file.name} 失败：{str(e)}")
        return pd.DataFrame(), uploaded_file.name

# -------------------------- 批量处理函数（适配上传文件） --------------------------
def batch_extract_data(uploaded_files_list):
    # 1. 筛选含关键词的文件
    target_files = []
    for file in uploaded_files_list:
        if file_keyword in file.name or file_keyword.lower() in file.name.lower():
            target_files.append(file)
        else:
            st.warning(f"⚠️ {file.name} 不含关键词「{file_keyword}」，已跳过")
    
    if not target_files:
        st.error(f"❌ 未找到包含「{file_keyword}」的上传文件")
        return None, {}
    
    # 2. 显示待处理文件
    st.info(f"✅ 找到 {len(target_files)} 个待处理文件：")
    file_list = []
    for i, f in enumerate(target_files, 1):
        station = extract_station_name(f.name)
        station_type = "📸 光伏" if station in pv_stations_list else "💨 风电"
        file_list.append(f"{i}. {station_type} {f.name}")
    st.code("\n".join(file_list))
    
    # 3. 批量提取
    all_station_dfs = {}
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, file in enumerate(target_files):
        status_text.text(f"正在处理：{file.name}（{idx+1}/{len(target_files)}）")
        station_name = extract_station_name(file.name)
        
        # 选择功率列索引
        if station_name in pv_stations_list:
            power_idx = power_col_idx_pv
        else:
            power_idx = power_col_idx_wind
        
        file_data, file_name = extract_excel_data(file, time_col_idx, power_idx, skip_rows, power_conversion)
        if not file_data.empty:
            all_station_dfs[station_name] = file_data
            st.success(f"✅ {station_name}：提取到 {len(file_data)} 条有效数据")
        progress_bar.progress((idx + 1) / len(target_files))
    
    status_text.text("处理完成！开始合并数据...")
    
    # 4. 合并数据
    if not all_station_dfs:
        st.error("❌ 未提取到任何有效数据")
        return None, {}
    
    df_list = list(all_station_dfs.values())
    merged_df = df_list[0]
    for df in df_list[1:]:
        merged_df = pd.merge(merged_df, df, on="时间", how="outer")
    
    merged_df["时间"] = merged_df["时间"].dt.floor("min")
    merged_df = merged_df.sort_values("时间").reset_index(drop=True)
    
    # 5. 统计数据
    st.success("📊 数据合并完成！")
    st.info(f"""
    统计信息：
    - 总时间点数：{len(merged_df)}
    - 包含场站：{', '.join(merged_df.columns[1:])}
    - 处理时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    """)
    
    # 显示各场站数据量
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
def to_excel(df):
    """将DataFrame转为Excel字节流"""
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    df.to_excel(writer, index=False, sheet_name="功率数据")
    writer.close()
    output.seek(0)
    return output

# -------------------------- 网页主界面 --------------------------
st.title("📊 光伏/风电功率数据提取工具（月度导入版）")
st.markdown("---")

# 提示信息
st.info("""
### 📝 使用指引
1. 在左侧侧边栏上传本月的Excel数据文件（支持多选）
2. 确认列索引/光伏场站等配置（首次配置后无需修改）
3. 点击下方按钮开始提取数据
4. 预览数据后下载整合文件
""")

# 执行按钮（仅当有文件上传时可用）
if uploaded_files:
    if st.button("🚀 开始提取数据", type="primary"):
        with st.spinner("正在批量处理上传的文件..."):
            result_df, station_dfs = batch_extract_data(uploaded_files)
            
            if result_df is not None and not result_df.empty:
                # 数据预览
                st.markdown("---")
                st.subheader("📈 数据预览")
                
                # 切换预览标签
                tab1, tab2 = st.tabs(["全部数据", "光伏场站数据"])
                with tab1:
                    st.dataframe(result_df.head(50), use_container_width=True)
                with tab2:
                    # 筛选光伏场站数据
                    pv_cols = [col for col in result_df.columns if col in pv_stations_list]
                    if pv_cols:
                        pv_df = result_df[["时间"] + pv_cols].dropna(subset=pv_cols, how="all")
                        st.dataframe(pv_df.head(50), use_container_width=True)
                    else:
                        st.info("暂无光伏场站数据")
                
                # 下载按钮
                st.markdown("---")
                st.subheader("📥 下载结果")
                # 生成带年月的文件名（适配月度数据）
                current_month = datetime.datetime.now().strftime("%Y%m")
                excel_data = to_excel(result_df)
                st.download_button(
                    label="下载整合数据（Excel）",
                    data=excel_data,
                    file_name=f"整合数据_历史趋势_{current_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
else:
    st.warning("⚠️ 请先在左侧侧边栏上传Excel数据文件！")

# 侧边栏说明
st.sidebar.markdown("---")
st.sidebar.markdown("### 📝 使用说明")
st.sidebar.markdown("""
1. 上传本月的历史趋势Excel文件（支持多选）
2. 确认列索引配置：
   - 时间列：E列=4（索引从0开始）
   - 风电功率列：J列=9
   - 光伏功率列：F列=5
3. 点击「开始提取数据」
4. 预览数据后下载月度整合文件
""")

st.sidebar.markdown("### ℹ️ 注意事项")
st.sidebar.markdown("""
- 支持.xlsx/.xls/.xlsm格式
- 自动区分光伏/风电场站列索引
- 数据按时间对齐，NaN表示无数据
- 下载文件名自动带年月，方便归档
""")
