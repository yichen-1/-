import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

# --------------------- 页面配置 ---------------------
st.set_page_config(
    page_title="电价相关性分析工具",
    page_icon="📊",
    layout="wide"
)
plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False

st.title("📈 辽宁电力现货电价相关性分析")
st.markdown("### 自动分析：电价 × 负荷 × 联络线 × 新能源 × 竞价空间")

# --------------------- 上传Excel文件 ---------------------
uploaded_file = st.file_uploader("上传运行数据披露文件", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success(f"✅ 数据读取成功！共 {len(df)} 条数据")

    # --------------------- 【定制匹配你的列名】 ---------------------
    key_cols = {
        "电价": None,
        "负荷": None,
        "联络线": None,
        "新能源出力": None,
        "竞价空间": None
    }

    for col in df.columns:
        col_str = str(col)
        if "现货价格" in col_str and "实时价格" in col_str:
            key_cols["电价"] = col
        elif "省内负荷" in col_str and "D-1预测" in col_str:
            key_cols["负荷"] = col
        elif "联络线" in col_str and "D-3预测" in col_str:
            key_cols["联络线"] = col
        elif "非市场机组出力" in col_str and "D-3计划" in col_str:
            key_cols["新能源出力"] = col
        elif "竞价空间" in col_str and "D-1预测" in col_str:
            key_cols["竞价空间"] = col

    # 显示匹配结果
    with st.expander("查看自动匹配的列", expanded=True):
        for k, v in key_cols.items():
            st.write(f"✅ {k} → {v}")

    # --------------------- 数据清洗 ---------------------
    use_cols = list(key_cols.values())
    df_corr = df[use_cols].copy().dropna()
    st.info(f"📊 清洗后有效数据：{len(df_corr)} 条")

    # --------------------- 计算相关性 ---------------------
    price_col = key_cols["电价"]
    corr_result = {}
    for name, col in key_cols.items():
        if col != price_col:
            corr = df_corr[col].corr(df_corr[price_col])
            corr_result[name] = round(corr, 4)

    # 按关联性强弱排序
    sorted_corr = sorted(corr_result.items(), key=lambda x: abs(x[1]), reverse=True)

    # --------------------- 展示结果表格 ---------------------
    st.subheader("📋 电价关联性排序结果")
    result_df = pd.DataFrame([
        {"指标名称": k, "与电价相关系数": v,
         "关联强度": "强相关" if abs(v)>=0.7 else "中等相关" if abs(v)>=0.4 else "弱相关" if abs(v)>=0.1 else "无关",
         "关联方向": "正相关" if v>0 else "负相关"}
        for k, v in sorted_corr
    ])
    st.dataframe(result_df, use_container_width=True)

    # --------------------- 双列展示图表 ---------------------
    st.subheader("📊 可视化分析图表")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 相关性热力图")
        fig1, ax1 = plt.subplots(figsize=(8, 5))
        corr_matrix = df_corr.corr()
        sns.heatmap(corr_matrix, annot=True, cmap="RdBu_r", linewidths=0.5, fmt=".2f", ax=ax1)
        st.pyplot(fig1)

    with col2:
        st.markdown("#### 相关性对比柱状图")
        names = [x[0] for x in sorted_corr]
        values = [x[1] for x in sorted_corr]
        fig2, ax2 = plt.subplots(figsize=(8, 5))
        ax2.bar(names, values, color=["#e74c3c" if x < 0 else "#3498db" for x in values])
        ax2.axhline(y=0, color="black", linewidth=0.8)
        ax2.set_ylabel("相关系数")
        ax2.grid(axis="y", alpha=0.3)
        plt.xticks(rotation=15)
        st.pyplot(fig2)

    # --------------------- 生成Excel并提供下载 ---------------------
    st.subheader("📥 下载分析报告")
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="关联性结果", index=False)
        corr_matrix.to_excel(writer, sheet_name="相关系数矩阵", index=True)

    st.download_button(
        label="下载 Excel 分析报告",
        data=buffer,
        file_name="电价关联性分析报告.xlsx",
        mime="application/vnd.ms-excel"
    )

else:
    st.warning("👆 请上传运行数据披露 Excel 文件")
