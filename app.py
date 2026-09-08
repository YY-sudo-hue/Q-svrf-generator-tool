import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# 设置网页标签页的标题和图标
st.set_page_config(page_title="SVRF 自动生成系统", page_icon="📄")
st.title("📄 SVRF 质量文件自动生成器")
st.markdown("输入客户与零件信息，系统将自动从数据库匹配并生成标准 SVRF 文件。")

# 创建输入框
customer = st.text_input("客户名称 (例如: 奇瑞)")
part = st.text_input("子零件名称 (例如: foam)")

# 创建生成按钮
if st.button("生成 SVRF 文件", type="primary"):
    if not customer or not part:
        st.warning("⚠️ 请输入完整的客户名称和子零件名称！")
    else:
        try:
            # 读取基础信息表 (请确保表名和 Sheet 顺序与实际一致)
           url_warranty = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSU0pfEGkRcFe8ehE9nNUolB1u0nciqR_e6hzWgzeAKk-KkXfLVcM4zkbssEbzgqGtWUhbjSnj8ybNs/pubhtml?gid=1613958416&single=true"
url_part = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSU0pfEGkRcFe8ehE9nNUolB1u0nciqR_e6hzWgzeAKk-KkXfLVcM4zkbssEbzgqGtWUhbjSnj8ybNs/pubhtml?gid=1738470985&single=true"
df_warranty = pd.read_csv(url_warranty)
df_part = pd.read_csv(url_part)
            # 清理列名空格，防止匹配失败
            df_warranty.columns = df_warranty.columns.str.strip()
            df_part.columns = df_part.columns.str.strip()
            
            # 匹配数据
            match_customer = df_warranty[df_warranty['客户名称'] == customer]
            match_part = df_part[df_part['子零件名称'] == part]
            
            if match_customer.empty or match_part.empty:
                st.error("❌ 数据库中未找到该客户或零件，请检查拼写。")
            else:
                data_customer = match_customer.iloc[0].fillna("")
                data_part = match_part.iloc[0].fillna("")
                
                # 读取模板并替换占位符
                wb = openpyxl.load_workbook("20250112_SVRF_Quality.xlsx")
                sheet = wb.active
                
                replace_dict = {
                    "<1>": str(data_customer.get("质保年限", "")),
                    "<2>": str(data_customer.get("质保公里数", "")),
                    "<3>": str(data_customer.get("设计寿命", "")),
                    "<4>": str(data_part.get("CCL", "")),
                    "<5>": str(data_part.get("SCP", ""))
                }
                
                # 遍历单元格替换
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            for key, val in replace_dict.items():
                                if key in cell.value:
                                    cell.value = cell.value.replace(key, val)
                
                # 将生成的文件保存到内存中（适配网页下载机制）
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                output_name = f"{customer}_{part}_Quality_20250112.xlsx"
                
                st.success("✅ 文件生成成功！请点击下方按钮下载：")
                # 提供下载按钮
                st.download_button(
                    label="📥 下载生成的文件",
                    data=output,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ 运行出错，请确保模板和信息表文件名正确无误。详细错误: {e}")
