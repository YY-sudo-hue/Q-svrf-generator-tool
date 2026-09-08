import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="SVRF 自动生成系统", page_icon="📄")
st.title("📄 SVRF 质量文件自动生成器")
st.markdown("输入客户与零件信息，系统将自动从数据库匹配并生成标准 SVRF 文件。")

db_file = "SVRF 信息表.xlsx"

# 全表查询面板
with st.expander("🔍 点击查看基础数据库 (完整信息表)"):
    col_db1, col_db2 = st.columns(2)
    with col_db1:
        if st.button("📊 查看客户名称信息"):
            try:
                st.dataframe(pd.read_excel(db_file, sheet_name=0))
            except Exception as e:
                st.error(f"读取失败，请检查文件: {e}")
    with col_db2:
        if st.button("📊 查看 CCLSCP 信息"):
            try:
                st.dataframe(pd.read_excel(db_file, sheet_name=1))
            except Exception as e:
                st.error(f"读取失败，请检查文件: {e}")

st.divider() 

# 输入区域
customer = st.text_input("客户名称 (例如: 奇瑞)")
part = st.text_input("子零件名称 (例如: foam)")

col_query, col_generate = st.columns(2)

with col_query:
    query_btn = st.button("🔎 查询当前匹配信息")
with col_generate:
    generate_btn = st.button("✨ 生成 SVRF 文件", type="primary")

# 查询逻辑
if query_btn:
    if not customer or not part:
        st.warning("⚠️ 请输入完整的客户名称和子零件名称！")
    else:
        try:
            df_warranty = pd.read_excel(db_file, sheet_name=0)
            df_part = pd.read_excel(db_file, sheet_name=1)
            
            df_warranty.columns = df_warranty.columns.str.strip()
            df_part.columns = df_part.columns.str.strip()
            
            match_customer = df_warranty[df_warranty['客户名称'] == customer]
            match_part = df_part[df_part['子零件名称'] == part]
            
            if match_customer.empty:
                st.error(f"❌ 客户名称表中未找到客户：{customer}")
            else:
                st.success(f"✅ 找到 {customer} 的客户名称信息：")
                st.dataframe(match_customer)
                
            if match_part.empty:
                st.error(f"❌ CCLSCP表中未找到零件：{part}")
            else:
                st.success(f"✅ 找到 {part} 的 CCLSCP 信息：")
                st.dataframe(match_part)
        except Exception as e:
            st.error(f"❌ 读取数据出错。详细错误: {e}")

# 生成文件逻辑
if generate_btn:
    if not customer or not part:
        st.warning("⚠️ 请输入完整的客户名称和子零件名称！")
    else:
        try:
            df_warranty = pd.read_excel(db_file, sheet_name=0)
            df_part = pd.read_excel(db_file, sheet_name=1)
            
            df_warranty.columns = df_warranty.columns.str.strip()
            df_part.columns = df_part.columns.str.strip()
            
            match_customer = df_warranty[df_warranty['客户名称'] == customer]
            match_part = df_part[df_part['子零件名称'] == part]
            
            if match_customer.empty or match_part.empty:
                st.error("❌ 数据库中未找到该客户或零件，请先点击左侧'查询当前匹配信息'检查拼写。")
            else:
                data_customer = match_customer.iloc[0].fillna("")
                data_part = match_part.iloc[0].fillna("")
                
                wb = openpyxl.load_workbook("20250112_SVRF_Quality.xlsx")
                sheet = wb.active
                
                replace_dict = {
                    "<1>": str(data_customer.get("质保年限", "")),
                    "<2>": str(data_customer.get("质保公里数", "")),
                    "<3>": str(data_customer.get("设计寿命", "")),
                    "<4>": str(data_part.get("CCL", "")),
                    "<5>": str(data_part.get("SCP", ""))
                }
                
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            for key, val in replace_dict.items():
                                if key in cell.value:
                                    cell.value = cell.value.replace(key, val)
                
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                output_name = f"{customer}_{part}_Quality_20250112.xlsx"
                
                st.success("✅ 文件生成成功！请点击下方按钮下载：")
                st.download_button(
                    label="📥 下载生成的文件",
                    data=output,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ 运行出错。详细错误: {e}")
