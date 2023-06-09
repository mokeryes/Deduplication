import streamlit as st
import pandas as pd

from ExcelOperate import ExcelOperate


st.set_page_config(page_title="Perking文件去重服务", layout="wide")

entire_excel = ""
dest_excel = ""

summary_excel_file = st.file_uploader(label="上传总表", type=["xls", "xlsx"])

if summary_excel_file:
    col1, col2 = st.columns(2)

    summary_excel_file_name = summary_excel_file.name

    with col1:
        st.write("总表")
        entire_excel = ExcelOperate(summary_excel_file_name, summary_excel_file).df
        st.dataframe(entire_excel, use_container_width=True)

    with col2:
        st.write("去重后的总表")
        st.dataframe(ExcelOperate().deduplicate(entire_excel), use_container_width=True)

st.divider()

dest_excel_file = st.file_uploader(label="上传目标表", type=["xls", "xlsx"])

if dest_excel_file:
    col1, col2 = st.columns(2)

    dest_excel_file_name = dest_excel_file.name

    with col1:
        st.write("目标表")
        dest_excel = ExcelOperate(dest_excel_file_name, dest_excel_file).df
        st.dataframe(dest_excel, use_container_width=True)

    with col2:
        st.write("去重后的目标表")
        st.dataframe(ExcelOperate().deduplicate(dest_excel), use_container_width=True)

    st.divider()

    st.write("对比去重后的表")
    deduplicated_excel = ExcelOperate().deduplicates(df_compare=entire_excel, df_dest=dest_excel)
    
    # 保存去重后的文件到本地
    deduplicated_excel.to_excel('deduplicated_excel.xlsx')

    st.dataframe(deduplicated_excel, use_container_width=True)
    with open("deduplicated_excel.xlsx", "rb") as file:
        btn = st.download_button(
                label="下载deduplicated_excel.xlsx",
                data=file,
                file_name="deduplicated_excel.xlsx",
                mime="Excel/xlsx"
              )
