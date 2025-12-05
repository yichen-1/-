# -------------------------- 方案电量手动调增调减（新增模块） --------------------------
st.divider()
st.header("✏️ 方案电量手动调增调减（总量保持不变）")

if st.session_state.calculated and st.session_state.selected_months:
    # 1. 选择调整的月份和方案
    col_adj1, col_adj2 = st.columns(2)
    with col_adj1:
        adj_month = st.selectbox(
            "选择要调整的月份",
            st.session_state.selected_months,
            key="adj_month_select"
        )
    with col_adj2:
        adj_scheme = st.selectbox(
            "选择要调整的方案",
            ["方案一（典型曲线）", "方案二（套利/直线曲线）"],
            key="adj_scheme_select"
        )

    # 2. 获取对应方案的数据和原始发电权重（核心：用原始平均发电量做权重依据）
    # 方案数据（方案一/方案二）
    if adj_scheme == "方案一（典型曲线）":
        scheme_df = st.session_state.trade_power_typical.get(adj_month, None)
        scheme_col = "方案一月度电量(MWh)"
    else:
        scheme_df = st.session_state.trade_power_arbitrage.get(adj_month, None)
        scheme_col = "方案二月度电量(MWh)"
    # 原始平均发电量（来自月度基础数据，不随调整变化，保证权重稳定）
    base_df = st.session_state.monthly_data.get(adj_month, None)

    if not scheme_df or not base_df:
        st.warning("⚠️ 请先生成该月份的方案数据")
    else:
        # 提取原始平均发电量（权重依据）
        avg_gen_list = base_df["平均发电量(MWh)"].tolist()
        avg_gen_total = sum(avg_gen_list)
        
        # 校验权重有效性
        if avg_gen_total <= 0:
            st.error("❌ 该月份原始平均发电量总和为0，无法按权重分摊调整量")
        else:
            # 保存调整前的原始数据（用于计算变化量）
            old_scheme_df = scheme_df.copy()
            total_fixed = old_scheme_df[scheme_col].sum()  # 固定总量（调整后不变）

            # 3. 显示可编辑的电量表格（仅开放方案电量列编辑）
            st.write(f"### {adj_scheme} - {adj_month}月电量调整（固定总量：{total_fixed:.2f} MWh）")
            edit_df = st.data_editor(
                scheme_df[["时段", "平均发电量(MWh)", "时段比重(%)", scheme_col]],
                column_config={
                    "时段": st.column_config.NumberColumn("时段", disabled=True),
                    "平均发电量(MWh)": st.column_config.NumberColumn("原始平均发电量(MWh)", disabled=True),
                    "时段比重(%)": st.column_config.NumberColumn("时段比重(%)", disabled=True),
                    scheme_col: st.column_config.NumberColumn(
                        f"{scheme_col}（可编辑）",
                        min_value=0.0,  # 禁止负电量
                        step=0.1,
                        format="%.2f",
                        help="修改后其他时段按「原始平均发电量权重」自动分摊调整量，总量不变"
                    )
                },
                use_container_width=True,
                num_rows="fixed",
                key=f"edit_adjust_scheme_{adj_month}_{adj_scheme}"
            )

            # 4. 检测表格修改，自动计算并分摊调整量
            if not edit_df.equals(old_scheme_df):
                # 计算每个时段的变化量，找到修改的时段（仅支持单时段修改，避免冲突）
                delta_series = edit_df[scheme_col] - old_scheme_df[scheme_col]
                modified_indices = delta_series[delta_series != 0].index.tolist()

                if len(modified_indices) > 1:
                    st.warning("⚠️ 暂支持单次修改1个时段，请保存当前调整后再修改其他时段！")
                    # 恢复原始数据，避免多时段修改导致总量混乱
                    if adj_scheme == "方案一（典型曲线）":
                        st.session_state.trade_power_typical[adj_month] = old_scheme_df
                    else:
                        st.session_state.trade_power_arbitrage[adj_month] = old_scheme_df
                elif len(modified_indices) == 1:
                    # 获取修改的时段索引和变化量
                    mod_idx = modified_indices[0]  # DataFrame行索引（对应时段1-24）
                    mod_hour = edit_df.loc[mod_idx, "时段"]  # 修改的时段（1-24）
                    delta = delta_series.iloc[0]  # 变化量（新值-旧值）

                    # 计算其他时段的分摊权重（排除修改的时段）
                    other_indices = [idx for idx in range(24) if idx != mod_idx]
                    other_avg_gen = [avg_gen_list[idx] for idx in other_indices]
                    other_avg_total = sum(other_avg_gen)

                    if other_avg_total <= 0:
                        st.error("❌ 其他时段原始平均发电量总和为0，无法分摊调整量！")
                    else:
                        # 5. 按权重分摊调整量（其他时段 = 原调整后值 + 分摊量，分摊量=-delta×权重占比）
                        adjusted_df = edit_df.copy()
                        for idx in other_indices:
                            # 该时段权重占比 = 该时段原始平均发电量 / 其他时段原始平均发电总量
                            weight_ratio = avg_gen_list[idx] / other_avg_total
                            # 分摊调整量（负delta：调增则其他减，调减则其他加）
                            share_amount = -delta * weight_ratio
                            # 新值 = 编辑后的值 + 分摊量（保证总量不变）
                            new_val = adjusted_df.loc[idx, scheme_col] + share_amount
                            # 边界保护：不能小于0
                            adjusted_df.loc[idx, scheme_col] = max(round(new_val, 2), 0.0)

                        # 6. 修正计算误差（确保总量严格等于原始总量，避免浮点数精度问题）
                        current_total = adjusted_df[scheme_col].sum()
                        if not np.isclose(current_total, total_fixed, atol=0.01):
                            # 最后一个其他时段兜底修正（不影响修改的时段）
                            last_other_idx = other_indices[-1]
                            correction = total_fixed - current_total
                            adjusted_df.loc[last_other_idx, scheme_col] = max(
                                round(adjusted_df.loc[last_other_idx, scheme_col] + correction, 2),
                                0.0
                            )

                        # 7. 更新时段比重（按新电量重新计算）
                        adjusted_df["时段比重(%)"] = round(adjusted_df[scheme_col] / total_fixed * 100, 4)

                        # 8. 保存调整后的数据到Session State（覆盖原方案数据）
                        if adj_scheme == "方案一（典型曲线）":
                            st.session_state.trade_power_typical[adj_month] = adjusted_df
                        else:
                            st.session_state.trade_power_arbitrage[adj_month] = adjusted_df

                        # 9. 显示调整结果反馈
                        st.success(
                            f"✅ 调整成功！\n"
                            f"- 修改时段：{mod_hour}点\n"
                            f"- 电量变化：{delta:.2f} MWh（原：{old_scheme_df.loc[mod_idx, scheme_col]:.2f} → 新：{adjusted_df.loc[mod_idx, scheme_col]:.2f}）\n"
                            f"- 其他时段按「原始平均发电量权重」自动分摊，总量保持 {total_fixed:.2f} MWh"
                        )
                else:
                    st.info("ℹ️ 未检测到有效修改（请直接编辑「可编辑」列的电量值）")
else:
    st.warning("⚠️ 请先生成年度方案后再进行电量调整")
