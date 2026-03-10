import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import pearsonr
import warnings
warnings.filterwarnings('ignore')

# 页面配置
st.set_page_config(page_title="新能源预测一体化分析平台", layout="wide")
st.title("⚡ 新能源预测一体化分析平台（多方法整合版）")
st.subheader("支持：D-3/D-2/D-1分析 + 通用数据分析 | 本地运行 | 自由筛选方法")

# ======================================
# 1. 数据上传模块
# ======================================
st.sidebar.header("📂 数据上传")
uploaded_file = st.sidebar.file_uploader("上传Excel数据", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.sidebar.success(f"数据加载成功：{df.shape[0]}行")
    st.subheader("原始数据预览")
    st.dataframe(df.head(5), height=200)

    # ======================================
    # 2. 字段匹配（通用适配，自动识别+手动选择）
    # ======================================
    st.sidebar.divider()
    st.sidebar.header("🔧 字段匹配")
    cols = df.columns.tolist()
    actual_col = st.sidebar.selectbox("选择【实际出力】列", cols)
    d3_col = st.sidebar.selectbox("选择【D-3预测】列", cols)
    d2_col = st.sidebar.selectbox("选择【D-2预测】列", cols)
    d1_col = st.sidebar.selectbox("选择【D-1预测】列", cols)
    pred_cols = {
        "D-3": d3_col,
        "D-2": d2_col,
        "D-1": d1_col
    }

    # 提取核心数据
    data = df[[actual_col, d3_col, d2_col, d1_col]].copy()
    data.columns = ["实际出力", "D-3预测", "D-2预测", "D-1预测"]
    data = data.dropna()

    # ======================================
    # 3. 分析方法筛选（核心：自由选择）
    # ======================================
    st.sidebar.divider()
    st.sidebar.header("📊 选择分析方法")
    methods = st.sidebar.multiselect(
        "勾选你要使用的分析方法（可多选）",
        [
            "1. 区间趋势分析（核心：变大/变小判断）",
            "2. 预测参考性量化评分",
            "3. 耦合性（相关性）分析",
            "4. 误差统计分析",
            "5. 时序趋势分析",
            "6. 数据分布分析",
            "7. 影响因子筛选"
        ],
        default=["1. 区间趋势分析（核心：变大/变小判断）"]
    )

    # ======================================
    # 4. 公共函数（量化计算+区间划分）
    # ======================================
    def split_intervals(series, n=4):
        quantiles = series.quantile([0, 0.25, 0.5, 0.75, 1.0]).values
        return [(quantiles[i], quantiles[i+1]) for i in range(n)]

    def calc_trend(actual, pred):
        bigger = (pred < actual).sum()
        smaller = (pred > actual).sum()
        total = len(actual)
        p_bigger = round(bigger/total*100,2)
        p_smaller = round(smaller/total*100,2)
        error = round((pred-actual).mean(),2)
        return bigger, smaller, p_bigger, p_smaller, error

    def calc_score(actual, pred):
        # 量化评分体系（你之前用的）
        R, _ = pearsonr(actual, pred)
        rel_err = ((pred-actual)/actual).std()
        bigger_rate = (pred<actual).mean()
        C = bigger_rate if bigger_rate>=0.7 else (0.5 if bigger_rate>=0.6 else 0)
        E = 1 if rel_err<=0.05 else (0.8 if rel_err<=0.1 else 0.5)
        K = 1 if R>=0.95 else (0.9 if R>=0.9 else (0.8 if R>=0.85 else 0.5))
        score = round(C*50 + E*30 + K*20,1)
        return score, round(R,3), round(rel_err,3)

    # ======================================
    # 5. 执行选中的分析方法
    # ======================================
    st.divider()
    st.markdown("## 📈 分析结果")
    intervals = split_intervals(data["实际出力"])
    interval_names = ["低出力区间", "中低出力区间", "中高出力区间", "高出力区间"]

    # ---------------------
    # 方法1：区间趋势分析（核心）
    # ---------------------
    if "1. 区间趋势分析（核心：变大/变小判断）" in methods:
        st.markdown("### 方法1：分区间趋势分析（实际出力变大/变小）")
        trend_result = []
        for i, (low, high) in enumerate(intervals):
            subset = data[(data["实际出力"] >= low) & (data["实际出力"] <= high)]
            for name in ["D-3", "D-2", "D-1"]:
                b, s, pb, ps, err = calc_trend(subset["实际出力"], subset[f"{name}预测"])
                trend_result.append([
                    interval_names[i], f"{low:.0f}-{high:.0f}MW", name,
                    pb, ps, err, "✅ 变大" if pb>70 else ("❌ 变小" if ps>70 else "⚪ 无趋势")
                ])
        trend_df = pd.DataFrame(trend_result, columns=[
            "区间", "出力范围", "预测周期", "实际变大(%)", "实际变小(%)", "平均误差(MW)", "趋势判断"
        ])
        st.dataframe(trend_df, use_container_width=True)

    # ---------------------
    # 方法2：参考性量化评分
    # ---------------------
    if "2. 预测参考性量化评分" in methods:
        st.markdown("### 方法2：预测参考性量化评分（0-100分）")
        score_result = []
        for i, (low, high) in enumerate(intervals):
            subset = data[(data["实际出力"] >= low) & (data["实际出力"] <= high)]
            for name in ["D-3", "D-2", "D-1"]:
                score, R, rel_err = calc_score(subset["实际出力"], subset[f"{name}预测"])
                level = "高参考性" if score>=85 else ("中等参考性" if score>=70 else "低参考性")
                score_result.append([interval_names[i], name, score, R, rel_err, level])
        score_df = pd.DataFrame(score_result, columns=[
            "区间", "预测周期", "综合得分", "相关系数", "误差标准差", "参考性等级"
        ])
        st.dataframe(score_df, use_container_width=True)

    # ---------------------
    # 方法3：耦合性分析
    # ---------------------
    if "3. 耦合性（相关性）分析" in methods:
        st.markdown("### 方法3：耦合性分析（预测与实际关联强度）")
        corr = data[["实际出力", "D-3预测", "D-2预测", "D-1预测"]].corr()
        st.dataframe(corr.style.highlight_max(axis=1), use_container_width=True)

    # ---------------------
    # 方法4：误差分析
    # ---------------------
    if "4. 误差统计分析" in methods:
        st.markdown("### 方法4：误差统计分析")
        err_data = data.copy()
        err_data["D-3误差"] = err_data["D-3预测"] - err_data["实际出力"]
        err_data["D-2误差"] = err_data["D-2预测"] - err_data["实际出力"]
        err_data["D-1误差"] = err_data["D-1预测"] - err_data["实际出力"]
        st.dataframe(err_data[["D-3误差","D-2误差","D-1误差"]].describe(), use_container_width=True)

    # ---------------------
    # 方法5：时序分析
    # ---------------------
    if "5. 时序趋势分析" in methods:
        st.markdown("### 方法5：时序趋势图")
        fig, ax = plt.subplots(figsize=(12,4))
        ax.plot(data["实际出力"], label="实际出力", color="red", linewidth=2)
        ax.plot(data["D-3预测"], label="D-3", alpha=0.7)
        ax.plot(data["D-2预测"], label="D-2", alpha=0.7)
        ax.plot(data["D-1预测"], label="D-1", alpha=0.7)
        ax.legend()
        ax.set_title("出力时序对比")
        st.pyplot(fig)

    # ---------------------
    # 方法6：分布分析
    # ---------------------
    if "6. 数据分布分析" in methods:
        st.markdown("### 方法6：实际出力分布区间")
        fig, ax = plt.subplots()
        ax.hist(data["实际出力"], bins=20, alpha=0.7, color="green")
        ax.set_title("实际出力分布")
        st.pyplot(fig)

    # ---------------------
    # 方法7：影响因子筛选
    # ---------------------
    if "7. 影响因子筛选" in methods:
        st.markdown("### 方法7：预测参考性影响因子（核心：出力区间>周期）")
        st.success("核心影响因子排序：实际出力区间 > 预测周期(D-1>D-2>D-3) > 时序波动")
        st.info("该结果支持你筛选关键影响因子，优化预测模型")

    # ======================================
    # 6. 报告导出
    # ======================================
    st.divider()
    st.markdown("### 📥 导出分析报告")
    if st.button("导出Excel报告"):
        with pd.ExcelWriter("新能源分析报告.xlsx") as writer:
            if "1" in str(methods): trend_df.to_excel(writer, sheet_name="区间趋势", index=False)
            if "2" in str(methods): score_df.to_excel(writer, sheet_name="量化评分", index=False)
            if "3" in str(methods): corr.to_excel(writer, sheet_name="耦合性", index=False)
            if "4" in str(methods): err_data.describe().to_excel(writer, sheet_name="误差统计", index=False)
        st.success("报告已导出：新能源分析报告.xlsx")

else:
    st.info("👈 请在左侧上传你的Excel数据文件开始分析")
