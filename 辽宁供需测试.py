# 新能源出力预测-实时平衡区间分析工具（Windows专用，无固定路径）
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 解决Windows中文乱码
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# ---------------------- 页面设置 ----------------------
st.set_page_config(page_title="新能源出力平衡区间分析", layout="wide")
st.title("⚡ 新能源出力预测 vs 实时数据 平衡区间分析工具")
st.markdown("### 功能：自动找平衡点 | 预测越大实际越大 | 预测越小实际更小")
st.divider()

# ---------------------- 上传文件 ----------------------
uploaded_file = st.file_uploader("上传你的Excel数据文件", type=["xlsx", "xls"])

if uploaded_file is not None:
    # 读取数据
    df = pd.read_excel(uploaded_file)
    st.success("✅ 文件上传成功！")

    # 数据概览
    with st.expander("📊 数据概览（点击展开）"):
        st.dataframe(df.head(10), use_container_width=True)
        st.write(f"数据行数：{len(df)} | 数据列数：{len(df.columns)}")

    st.divider()

    # ---------------------- 核心字段识别 ----------------------
    st.subheader("🔍 选择分析字段")
    # 手动选择预测列和实时列（最稳定）
    col1, col2 = st.columns(2)
    with col1:
        selected_pred = st.multiselect("选择预测出力列（D-3/D-2/D-1）", df.columns.tolist())
    with col2:
        real_col = st.selectbox("选择实时出力列", df.columns.tolist())

    st.divider()

    # ---------------------- 平衡区间分析 ----------------------
    st.subheader("📈 核心分析：平衡区间 & 偏差规律")
    if selected_pred and real_col:
        # 数据清洗（去除空值）
        df_analysis = df.dropna(subset=selected_pred + [real_col]).copy()

        # 1. 计算偏差：实际值 - 预测值
        for col in selected_pred:
            df_analysis[f"{col}_偏差"] = df_analysis[real_col] - df_analysis[col]

        # 2. 分区间统计
        interval_num = st.slider("划分区间数量", min_value=5, max_value=20, value=10)
        target_pred = st.selectbox("选择要分析的预测版本", selected_pred, index=0)

        # 划分区间
        df_analysis['区间'] = pd.cut(df_analysis[target_pred], bins=interval_num)
        # 区间统计
        stat_df = df_analysis.groupby('区间').agg({
            f"{target_pred}_偏差": ['mean', 'min', 'max', 'count']
        }).round(2)
        stat_df.columns = ['平均偏差', '最小偏差', '最大偏差', '数据量']
        stat_df = stat_df.reset_index()

        # 判定平衡区间
        stat_df['区间判定'] = np.where(stat_df['平均偏差'] > 0, "✅ 实际 > 预测",
                                      np.where(stat_df['平均偏差'] < 0, "❌ 实际 < 预测", "⚖️ 平衡"))

        # 展示结果
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown(f"#### 【{target_pred}】区间统计结果")
            st.dataframe(stat_df, use_container_width=True)
        with col_b:
            # 可视化：预测区间 vs 平均偏差
            fig, ax = plt.subplots(figsize=(8, 5))
            ax.bar(stat_df['区间'].astype(str), stat_df['平均偏差'], color=['green' if x>0 else 'red' for x in stat_df['平均偏差']])
            ax.axhline(y=0, color='black', linestyle='--', linewidth=2)
            plt.xticks(rotation=45)
            plt.title(f'{target_pred} 预测区间 → 实际出力偏差')
            plt.ylabel('偏差（实际-预测）')
            plt.tight_layout()
            st.pyplot(fig)

        st.divider()

        # ---------------------- 规律验证 ----------------------
        st.subheader("🔎 规律验证：预测越大→实际越大")
        corr = df_analysis[target_pred].corr(df_analysis[real_col])
        st.markdown(f"**预测值与实际值 相关系数：{cor:.2f}**")
        if corr > 0.8:
            st.success("✅ 强相关：预测越大，实际越大，规律成立！")
        elif corr > 0.6:
            st.info("⚠️ 中等相关：规律基本成立")
        else:
            st.warning("❌ 弱相关：规律不明显")

        # 散点图
        fig2, ax2 = plt.subplots(figsize=(10, 5))
        ax2.scatter(df_analysis[target_pred], df_analysis[real_col], alpha=0.6)
        ax2.plot([df_analysis[target_pred].min(), df_analysis[target_pred].max()],
                 [df_analysis[target_pred].min(), df_analysis[target_pred].max()], 'r--', label='平衡线')
        plt.xlabel(f'{target_pred} 预测出力')
        plt.ylabel(f'{real_col} 实际出力')
        plt.title('预测出力 vs 实际出力 散点图')
        plt.legend()
        st.pyplot(fig2)

        st.divider()

        # ---------------------- 导出分析结果 ----------------------
        st.subheader("💾 导出分析报告")
        if st.button("生成并下载Excel分析报告"):
            with pd.ExcelWriter("新能源出力平衡区间分析报告.xlsx") as writer:
                df_analysis.to_excel(writer, sheet_name="原始数据+偏差", index=False)
                stat_df.to_excel(writer, sheet_name="区间统计结果", index=False)
            st.success("✅ 分析报告已生成！保存在代码同一文件夹里")

    else:
        st.warning("请选择预测列和实时列！")

else:
    st.info("👆 请上传你的Excel数据文件开始分析")
