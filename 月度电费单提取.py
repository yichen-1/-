import streamlit as st
import pandas as pd
import os
from datetime import datetime
from io import BytesIO

# -------------------------- 基础配置（保留映射表） --------------------------
plant_name_mapping = {
    "襄北聚合光伏": "襄阳聚合光伏",
    "圣境山": "荆门协合圣境山风电",
    "栗溪": "荆门协合栗溪风电",
    "风储一期": "三王（协合襄北）风电",
    "风储二期": "襄州协合三王风光储能电站风电二期",
    "峪山一期": "襄阳协合峪山泉水风电",
    "峪山二期": "襄阳峪龙峪山风电",
    "浠水光伏": "浠水聚合关口光伏",
    "北风垭风电": "北风垭风电",
    "南岭风电": "南岭风电",
    "牛庄风电": "牛庄风电",
    "中节能伙牌汤岗风电场": "中节能伙牌汤岗风电场",
    "中节能五峰牛庄风电场二期": "中节能五峰牛庄风电场二期"
}

# 目标科目定义（完整保留）
TARGET_AUX_SERVICES = ["省间调峰购入分摊退补", "省内调频辅助服务退补"]
TARGET_TWO_RULES = ["两个细则考核费用（新）清算", "两个细则补偿费用清算", "两个细则分摊费用清算", "两个细则考核费用（新）退补", "两个细则返还费用退补", "两个细则补偿费用退补", "两个细则分摊费用退补"]
TARGET_PROFIT_RECOVERY = "中长期超额获利回收电费（现货）"
TARGET_STORAGE_TWO_RULES = [
    "配建储能两个细则考核费用（新）退补",
    "配建储能两个细则返还费用退补",
    "配建储能两个细则补偿费用退补",
    "配建储能两个细则分摊费用退补"
]

NEW_TARGETS = {
    "省内现货交易": {
        "power_field": "省内现货偏差电量（万千瓦时）",
        "fee_field": "省内现货电费（万元）",
        "power_col_index": 3
    },
    "省间现货交易": {
        "power_field": "省间现货电量（万千瓦时）",
        "fee_field": "省间现货电费（万元）",
        "power_col_index": 3
    },
    "中长期交易": {
        "power_field": None,
        "fee_field": "中长期电费（万元）",
        "power_col_index": None
    },
    "其他优先发购电量": {
        "power_field": None,
        "fee_field": "保障性电费（万元）",
        "power_col_index": None
    }
}

TARGET_MECHANISM = ["机制电量差价结算费用", "机制电量差价结算费用退补"]
MECHANISM_POWER_COL_INDEX = 3
amount_col_index = 6

# 必需列定义
required_columns = [
    '电厂名称', '月份', '考核金额', '省间现货电量（万千瓦时）', 
    '是否有偏差考核', '上网电量（万千瓦时）', '基础电量/优先发电量（万千瓦时）',
    '两个细则（元）', '辅助服务（元）', '两个细则电费（万元）', '辅助服务电费(万元)',
    '省内现货电费（万元）', '中长期电费（万元）', '省内现货偏差电量（万千瓦时）', 
    '结算电费（万元）', '不含绿电中长期交易结算电量（万千瓦时）', 
    '交易电量（万千瓦时）', '交易电量占比（%）', 
    '不含辅助服务与两个细则结算电费（万元）',
    '不含辅助服务与两个细则结算平均电价(元/千瓦时)',
    '省间现货电费（万元）',
    '保障性电费（万元）',
    '机制电量（万kwh）',
    '机制电费（元）',
    '中长期超额获利回收电费（元）',
    '配储两个细则（元）',
    '配储两个细则电费（万元）'
]

# -------------------------- 工具函数（完整保留） --------------------------
def clean_data(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, str):
        cleaned = val.strip().replace(',', '').replace(' ', '')
        if cleaned in ['/', '无', 'None', '']:
            return 0.0
        try:
            return float(cleaned)
        except:
            return 0.0
    return float(val)

# -------------------------- Streamlit 界面配置 --------------------------
st.set_page_config(
    page_title="湖北协合结算数据处理工具",
    page_icon="📊",
    layout="wide"
)

st.title("📊 湖北协合结算数据处理工具")
st.markdown("---")

# 侧边栏配置
with st.sidebar:
    st.header("⚙️ 配置选项")
    # 月份选择
    year = st.selectbox("选择年份", options=range(2023, 2030), index=2)  # 默认2025
    month = st.selectbox("选择月份", options=range(1, 13), index=10)     # 默认11月
    month_str = f"{month:02d}"
    date_str = f'{year}-{month_str}-01'
    
    st.markdown("---")
    st.header("📤 文件上传")
    
    # 上传模板文件（主表格）
    template_file = st.file_uploader(
        "上传主表格模板（湖北每月数据更新.xlsx）",
        type=['xlsx', 'xls'],
        accept_multiple_files=False
    )
    
    # 上传当月结算文件（多个）
    settlement_files = st.file_uploader(
        "上传当月结算文件（Excel格式）",
        type=['xlsx', 'xls', 'XLSX'],
        accept_multiple_files=True
    )
    
    st.markdown("---")
    st.info("""
    📝 使用说明：
    1. 选择处理年份和月份
    2. 上传主表格模板（可选，无则创建新表格）
    3. 上传所有电厂的当月结算文件
    4. 点击下方【开始处理】按钮
    5. 处理完成后可下载结果文件
    """)

# 主界面
col1, col2 = st.columns(2)
with col1:
    st.subheader("🔧 当前配置")
    st.write(f"📅 处理月份：{year}年{month}月")
    st.write(f"📁 已上传结算文件数：{len(settlement_files)}")
    if template_file:
        st.write(f"✅ 已上传模板文件：{template_file.name}")
    else:
        st.write("ℹ️ 未上传模板文件，将创建新表格")

with col2:
    st.subheader("📋 功能说明")
    st.markdown("""
    - ✅ 支持批量处理多个结算文件
    - ✅ 自动提取辅助服务、两个细则等费用
    - ✅ 包含配储两个细则和超额获利回收提取
    - ✅ 自动计算交易电量、占比等衍生指标
    - ✅ 支持结果文件下载（Excel格式）
    """)

st.markdown("---")

# -------------------------- 数据处理逻辑 --------------------------
if st.button("🚀 开始处理", type="primary"):
    if not settlement_files:
        st.error("❌ 请先上传当月结算文件！")
    else:
        with st.spinner("⏳ 正在初始化数据..."):
            # 初始化主数据框
            if template_file:
                try:
                    df = pd.read_excel(template_file, sheet_name='Sheet1', engine='openpyxl')
                    st.success(f"✅ 成功读取模板文件（{len(df)}行数据）")
                except Exception as e:
                    st.warning(f"⚠️ 读取模板文件失败：{str(e)}，创建新表格")
                    df = pd.DataFrame()
            else:
                st.info("ℹ️ 未提供模板文件，创建新表格")
                df = pd.DataFrame()
            
            # 补全所有列
            for col in required_columns:
                if col not in df.columns:
                    df[col] = 0.0 if any(key in col for key in ['（元）', '（万元）', '（%）', '万千瓦时', '万kwh']) else ""
            
            # 设置月份
            df['月份'] = month
            st.success(f"✅ 月份统一设置为：{month}月")
        
        # 存储上传的文件信息（文件名→文件对象）
        settlement_file_dict = {}
        for file in settlement_files:
            # 提取文件名（不含后缀）用于匹配电厂名称
            file_name = os.path.splitext(file.name)[0]
            settlement_file_dict[file_name] = file
        st.success(f"✅ 已加载 {len(settlement_file_dict)} 个结算文件")
        
        # 开始处理每个电厂
        st.markdown("---")
        st.subheader("📊 处理进度")
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        result_container = st.container()
        
        total_plants = len(df) if not df.empty else len(plant_name_mapping)
        processed_count = 0
        
        for index, row in df.iterrows():
            plant_name = row['电厂名称']
            if pd.isna(plant_name) or str(plant_name).strip() == "":
                with result_container:
                    st.warning(f"⚠️ 行{index+1}：电厂名称为空，跳过处理")
                processed_count += 1
                progress_bar.progress(processed_count / total_plants)
                continue
            
            status_text.text(f"🔧 正在处理：{plant_name}（行{index+1}）")
            
            with result_container:
                st.markdown(f"### 🔍 处理电厂：{plant_name}")
                
                if plant_name in plant_name_mapping:
                    base_name = plant_name_mapping[plant_name]
                    target_file_name = f'{base_name}{date_str}'
                    
                    # 查找匹配的文件
                    matched_file = None
                    for file_name, file_obj in settlement_file_dict.items():
                        if target_file_name in file_name:
                            matched_file = file_obj
                            break
                    
                    if not matched_file:
                        st.error(f"❌ 未找到对应的结算文件：{target_file_name}.xlsx")
                        processed_count += 1
                        progress_bar.progress(processed_count / total_plants)
                        continue
                    
                    try:
                        # 读取结算文件
                        try:
                            target_df = pd.read_excel(matched_file, sheet_name='sheet1', header=4, engine='openpyxl')
                        except:
                            target_df = pd.read_excel(matched_file, sheet_name='sheet1', header=4, engine='xlrd')
                        
                        st.success(f"✅ 成功读取结算文件：{matched_file.name}（数据形状：{target_df.shape}）")
                        
                        # 初始化提取变量
                        aux_service_sum = 0.0
                        two_rules_sum = 0.0
                        storage_two_rules_sum = 0.0
                        profit_recovery = 0.0
                        new_target_results = {
                            "省内现货偏差电量（万千瓦时）": 0.0,
                            "省内现货电费（万元）": 0.0,
                            "省间现货电量（万千瓦时）": 0.0,
                            "省间现货电费（万元）": 0.0,
                            "中长期电费（万元）": 0.0,
                            "保障性电费（万元）": 0.0
                        }
                        mechanism_power_sum = 0.0
                        mechanism_fee_sum = 0.0
                        
                        # 遍历结算文件行
                        for row_idx in range(len(target_df)):
                            row_data = target_df.iloc[row_idx]
                            row_str = str(row_data).strip()
                            
                            # 1. 辅助服务提取
                            if any(aux_sub in row_str for aux_sub in TARGET_AUX_SERVICES):
                                if len(target_df.columns) > amount_col_index:
                                    amount_val = target_df.iloc[row_idx, amount_col_index]
                                    amount = clean_data(amount_val)
                                    aux_service_sum += amount
                                    matched_aux = [sub for sub in TARGET_AUX_SERVICES if sub in row_str][0]
                                    st.write(f"  ✅ 辅助服务：{matched_aux} → {amount:.2f}元")
                            
                            # 2. 普通两个细则提取
                            if any(two_sub in row_str for two_sub in TARGET_TWO_RULES):
                                if len(target_df.columns) > amount_col_index:
                                    amount_val = target_df.iloc[row_idx, amount_col_index]
                                    amount = clean_data(amount_val)
                                    two_rules_sum += amount
                                    matched_two = [sub for sub in TARGET_TWO_RULES if sub in row_str][0]
                                    st.write(f"  ✅ 普通两个细则：{matched_two} → {amount:.2f}元")
                            
                            # 3. 配储两个细则提取
                            if any(storage_sub in row_str for storage_sub in TARGET_STORAGE_TWO_RULES):
                                if len(target_df.columns) > amount_col_index:
                                    amount_val = target_df.iloc[row_idx, amount_col_index]
                                    amount = clean_data(amount_val)
                                    storage_two_rules_sum += amount
                                    matched_storage = [sub for sub in TARGET_STORAGE_TWO_RULES if sub in row_str][0]
                                    st.write(f"  ✅ 配储两个细则：{matched_storage} → {amount:.2f}元")
                            
                            # 4. 超额获利回收提取
                            if TARGET_PROFIT_RECOVERY in row_str:
                                if len(target_df.columns) > amount_col_index:
                                    amount_val = target_df.iloc[row_idx, amount_col_index]
                                    profit_recovery = clean_data(amount_val)
                                    st.write(f"  ✅ 超额获利回收：{TARGET_PROFIT_RECOVERY} → {profit_recovery:.2f}元")
                            
                            # 5. 新增科目提取
                            for target_subject, mapping in NEW_TARGETS.items():
                                if target_subject in row_str:
                                    st.write(f"  🔍 匹配新增科目：{target_subject}")
                                    # 提取电费
                                    if len(target_df.columns) > amount_col_index:
                                        fee_val = target_df.iloc[row_idx, amount_col_index]
                                        fee = clean_data(fee_val) / 10000
                                        new_target_results[mapping["fee_field"]] = round(fee, 2)
                                        st.write(f"    ✅ 电费：{fee:.2f}万元")
                                    # 提取电量
                                    if mapping["power_field"] and mapping["power_col_index"] is not None:
                                        power_col = mapping["power_col_index"]
                                        if len(target_df.columns) > power_col:
                                            power_val = target_df.iloc[row_idx, power_col]
                                            power = clean_data(power_val) / 10
                                            new_target_results[mapping["power_field"]] = round(power, 2)
                                            st.write(f"    ✅ 电量：{power:.2f}万千瓦时")
                            
                            # 6. 机制电量相关提取
                            if any(mech_sub in row_str for mech_sub in TARGET_MECHANISM):
                                matched_mech = [sub for sub in TARGET_MECHANISM if sub in row_str][0]
                                st.write(f"  🔍 匹配机制科目：{matched_mech}")
                                # 提取电量
                                if len(target_df.columns) > MECHANISM_POWER_COL_INDEX:
                                    power_val = target_df.iloc[row_idx, MECHANISM_POWER_COL_INDEX]
                                    power = clean_data(power_val) / 10
                                    mechanism_power_sum += power
                                    st.write(f"    ✅ 机制电量：{power:.2f}万kwh")
                                # 提取电费
                                if len(target_df.columns) > amount_col_index:
                                    fee_val = target_df.iloc[row_idx, amount_col_index]
                                    fee = clean_data(fee_val)
                                    mechanism_fee_sum += fee
                                    st.write(f"    ✅ 机制电费：{fee:.2f}元")
                        
                        # 赋值到主表格
                        df.at[index, '辅助服务（元）'] = round(aux_service_sum, 2)
                        df.at[index, '辅助服务电费(万元)'] = round(aux_service_sum / 10000, 4)
                        df.at[index, '两个细则（元）'] = round(two_rules_sum, 2)
                        df.at[index, '两个细则电费（万元）'] = round(two_rules_sum / 10000, 4)
                        df.at[index, '配储两个细则（元）'] = round(storage_two_rules_sum, 2)
                        df.at[index, '配储两个细则电费（万元）'] = round(storage_two_rules_sum / 10000, 4)
                        df.at[index, '中长期超额获利回收电费（元）'] = round(profit_recovery, 2)
                        
                        # 新增字段赋值
                        for field, value in new_target_results.items():
                            df.at[index, field] = value
                        
                        # 机制字段赋值
                        df.at[index, '机制电量（万kwh）'] = round(mechanism_power_sum, 2)
                        df.at[index, '机制电费（元）'] = round(mechanism_fee_sum, 2)
                        
                        # 其他指标提取
                        # 上网电量
                        if '实际上网电量' in target_df.columns and len(target_df) > 0:
                            try:
                                actual_power = clean_data(target_df['实际上网电量'].iloc[0]) / 10
                                df.at[index, '上网电量（万千瓦时）'] = round(actual_power, 2)
                                st.write(f"📊 上网电量：{actual_power:.2f}万千瓦时")
                            except:
                                st.warning("⚠️ 上网电量提取失败")
                        
                        # 基础电量
                        base_power_row = 11
                        base_power_col = 3
                        if len(target_df) > base_power_row and len(target_df.columns) > base_power_col:
                            try:
                                base_power_val = target_df.iloc[base_power_row, base_power_col]
                                base_power = clean_data(base_power_val) / 10
                                df.at[index, '基础电量/优先发电量（万千瓦时）'] = round(base_power, 2)
                                st.write(f"📊 基础电量：{base_power:.2f}万千瓦时")
                            except:
                                st.warning("⚠️ 基础电量提取失败")
                        
                        # 考核金额
                        assessment_row = 168
                        if len(target_df) > assessment_row and len(target_df.columns) > amount_col_index:
                            try:
                                assess_amt_val = target_df.iloc[assessment_row, amount_col_index]
                                assess_amt = clean_data(assess_amt_val) / 10000
                                df.at[index, '考核金额'] = round(assess_amt, 2)
                                df.at[index, '是否有偏差考核'] = '是' if assess_amt != 0 else '否'
                                st.write(f"📊 考核金额：{assess_amt:.2f}万元，偏差考核：{df.at[index, '是否有偏差考核']}")
                            except:
                                df.at[index, '是否有偏差考核'] = '否'
                                st.warning("⚠️ 考核金额提取失败")
                        else:
                            df.at[index, '是否有偏差考核'] = '否'
                            st.warning("⚠️ 考核金额行/列不存在")
                        
                        # 结算电费
                        if len(target_df) > 0 and len(target_df.columns) > amount_col_index:
                            try:
                                settle_fee_val = target_df.iloc[0, amount_col_index]
                                settle_fee = clean_data(settle_fee_val) / 10000
                                df.at[index, '结算电费（万元）'] = round(settle_fee, 2)
                                st.write(f"📊 结算电费：{settle_fee:.2f}万元")
                            except:
                                st.warning("⚠️ 结算电费提取失败")
                        
                        # 衍生计算
                        online_power = df.at[index, '上网电量（万千瓦时）']
                        base_power = df.at[index, '基础电量/优先发电量（万千瓦时）']
                        if isinstance(online_power, (int, float)) and isinstance(base_power, (int, float)):
                            trade_power = online_power - base_power
                            df.at[index, '交易电量（万千瓦时）'] = round(trade_power, 2)
                            if online_power != 0:
                                trade_ratio = (trade_power / online_power) * 100
                                df.at[index, '交易电量占比（%）'] = round(trade_ratio, 2)
                            st.write(f"📊 交易电量：{trade_power:.2f}万千瓦时，占比：{df.at[index, '交易电量占比（%）']:.2f}%")
                        
                        settle_fee = df.at[index, '结算电费（万元）']
                        total_deduct = (aux_service_sum + two_rules_sum + storage_two_rules_sum) / 10000
                        if isinstance(settle_fee, (int, float)) and online_power != 0:
                            net_fee = settle_fee - total_deduct
                            df.at[index, '不含辅助服务与两个细则结算电费（万元）'] = round(net_fee, 2)
                            net_price = (net_fee * 10000) / (online_power * 10000)
                            df.at[index, '不含辅助服务与两个细则结算平均电价(元/千瓦时)'] = round(net_price, 4)
                            st.write(f"📊 净结算电费：{net_fee:.2f}万元，平均电价：{net_price:.4f}元/千瓦时")
                        
                        st.success(f"✅ {plant_name} 处理完成！")
                        st.markdown("---")
                        
                    except Exception as e:
                        st.error(f"❌ 处理失败：{str(e)}")
                        st.markdown("---")
                else:
                    st.error(f"❌ 电厂名称 {plant_name} 未在映射表中")
                    st.markdown("---")
            
            processed_count += 1
            progress_bar.progress(processed_count / total_plants)
        
        # 处理完成
        status_text.text("✅ 所有电厂处理完成！")
        progress_bar.progress(1.0)
        
        st.markdown("---")
        st.subheader("🎉 处理完成！")
        
        # 生成下载文件
        current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'湖北每月数据更新_新版_{year}{month_str}_{current_time}.xlsx'
        
        # 保存到BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        
        output.seek(0)
        
        # 显示结果概览
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("总电厂数", len(df))
        with col2:
            st.metric("总列数", len(df.columns))
        with col3:
            st.metric("处理完成率", f"{processed_count/total_plants:.0%}")
        
        # 下载按钮
        st.download_button(
            label="📥 下载结果文件",
            data=output,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        # 显示前5行预览
        st.markdown("---")
        st.subheader("📋 数据预览（前5行）")
        st.dataframe(df.head(), use_container_width=True)
