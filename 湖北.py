# 在侧边栏的其他扩展菜单区域添加
with st.expander("📊 数据报表", expanded=False):
    st.subheader("月度报表生成")
    # 添加报表相关功能
    month = st.selectbox("选择报表月份", get_uploaded_months())
    if st.button("生成报表"):
        # 报表生成逻辑
        st.success("报表生成完成！")

with st.expander("💾 数据导出", expanded=False):
    st.subheader("批量数据导出")
    # 添加导出相关功能
    export_format = st.radio("选择导出格式", ["Excel", "CSV"])
    if st.button("导出所有数据"):
        # 导出逻辑
        st.success("数据导出完成！")
