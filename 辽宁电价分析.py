import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

# --------------------- 页面配置 ---------------------
st.set_page_config(
    page_title="电价相关性分析",
    page_icon="📊",
    layout="wide"
)
# 中文显示
plt.rcParams["font.sans-serif"] = ["SimHei", "WenQuanYi Micro Hei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

st.title("📈 辽宁电力现货电价相关性分析工具")
st.markdown("支持：负荷、联络线、新能源出力、竞价空间 与电价的关联分析")

# --------------------- 上传文件 ---------------------
uploaded_file = st.file_uploader("上传 运行数据披露.xlsx 文件", type="xlsx")
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success(f"✅ 数据读取成功 | 总数据量：{len(df)} 条")

    # --------------------- 自动匹配你的列名 ---------------------
    key_cols = {
        "电价": None,
        "负荷": None,
        "联络线": None,
        "新能源出力": None,
        "竞价空间": None
    }

    for col in df.columns:
        s = str(col)
        if "现货价格" in s and "实时价格" in s:
            key_cols["电价"] = col
        elif "省内负荷" in s and "D-1预测" in s:
            key_cols["负荷"] = col
        elif "联络线" in s and "D-3预测" in s:
            key_cols["联络线"] = col
        elif "非市场机组出力" in s and "D-3计划" in s:
            key_cols["新能源出力"] = col
        elif "竞价空间" in s and "D-1预测" in s:
            key_cols["竞价空间"] = col

    # 展示匹配结果
    with st.expander("✅ 已自动匹配数据列", expanded=True):
        for k, v in key_cols.items():
            st.write(f"{k} → {v}")

    # 数据清洗
    use_cols = [v for v in key_cols.values() if v is not None]
    df_clean = df[use_cols].dropna()
    st.info(f"📊 有效分析数据：{len(df_clean)} 条")

    # --------------------- 计算相关性 ---------------------
    price_col = key_cols["电价"]
    corr_result = {}
    for name, col in key_cols.items():
        if col != price_col and col is not None:
            corr = df_clean[col].corr(df_clean[price_col])
            corr_result[name] = round(corr, 4)

    # 排序
    sorted_corr = sorted(corr_result.items(), key=lambda x: abs(x[1]), reverse=True)

    # --------------------- 展示结果表 ---------------------
    st.subheader("📋 电价关联性排名")
    result_df = pd.DataFrame([
        {
            "指标名称": k,
            "相关系数": v,
            "关联强度": "强相关" if abs(v)>=0.7 else "中等相关" if abs(v)>=0.4 else "弱相关",
            "方向": "正相关" if v>0 else "负相关"
        } for k, v in sorted_corr
    ])
    st.dataframe(result_df, use_container_width=True)

    # --------------------- 在线生成图表 ---------------------
    st.subheader("📊 可视化图表")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 相关性矩阵")
        corr_matrix = df_clean.corr()
        fig, ax = plt.subplots(figsize=(7, 5))
        im = ax.imshow(corr_matrix.values, cmap="RdBu_r", vmin=-1, vmax=1)
        ax.set_xticks(range(len(corr_matrix.columns)))
        ax.set_yticks(range(len(corr_matrix.columns)))
        ax.set_xticklabels(corr_matrix.columns, rotation=45, fontsize=9)
        ax.set_yticklabels(corr_matrix.columns, fontsize=9)
        # 标注数值
        for i in range(len(corr_matrix.columns)):
            for j in range(len(corr_matrix.columns)):
                ax.text(j, i, f"{corr_matrix.iloc[i,j]:.2f}", ha="center", va="center", fontsize=9)
        plt.colorbar(im, ax=ax)
        plt.tight_layout()
        st.pyplot(fig)

    with col2:
        st.markdown("#### 相关性对比柱状图")
        names = [x[0] for x in sorted_corr]
        values = [x[1] for x in sorted_corr]
        fig2, ax2 = plt.subplots(figsize=(7, 5))
        colors = ["#e74c3c" if x < 0 else "#3498db" for x in values]
        ax2.bar(names, values, color=colors)
        ax2.axhline(0, color="black", linewidth=1)
        ax2.grid(axis='y', alpha=0.3)
        plt.xticks(rotation=15)
        st.pyplot(fig2)

    # --------------------- 下载Excel报告 ---------------------
    st.subheader("📥 导出分析报告")
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as f:
        result_df.to_excel(f, sheet_name="相关性结果", index=False)
        corr_matrix.to_excel(f, sheet_name="系数矩阵", index=True)

    st.download_button(
        label="下载 Excel 分析报告",
        data=buffer.getvalue(),
        file_name="电价相关性分析报告.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("👆 请上传你的运行数据文件")
