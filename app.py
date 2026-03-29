import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Dashboard bán hàng", layout="wide")

st.title("📊 Phân tích dữ liệu bán hàng")

uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=13)
    df = df.dropna(how='all')
    df.columns = df.columns.str.strip()

    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")

    df["Doanh thu"] = df["Thành tiền bán"]
    df["Lợi nhuận"] = df["Thành tiền bán"] - df["Thành tiền vốn"]

    # KPI
    col1, col2 = st.columns(2)
    col1.metric("Doanh thu", f"{df['Doanh thu'].sum():,.0f}")
    col2.metric("Lợi nhuận", f"{df['Lợi nhuận'].sum():,.0f}")

    # Top khách hàng
    top_kh = df.groupby("Tên khách hàng")["Doanh thu"].sum().reset_index().sort_values(by="Doanh thu", ascending=False).head(10)
    st.bar_chart(top_kh.set_index("Tên khách hàng"))

    # Theo thời gian
    dt = df.groupby("Ngày chứng từ")["Doanh thu"].sum().reset_index()
    st.line_chart(dt.set_index("Ngày chứng từ"))

    # Cảnh báo
    st.subheader("⚠️ Cảnh báo")

    xe = df.groupby("Số xe")["Tên khách hàng"].nunique().reset_index()
    xe_risk = xe[xe["Tên khách hàng"] > 1]

    if not xe_risk.empty:
        st.write("Xe nhiều khách:")
        st.dataframe(xe_risk)

    lo = df[df["Lợi nhuận"] < 0]
    if not lo.empty:
        st.write("Đơn hàng lỗ:")
        st.dataframe(lo)

    st.dataframe(df)