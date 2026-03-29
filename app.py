import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Phân tích đơn hàng", layout="wide")

st.title("📊 PHÂN TÍCH KHÁCH HÀNG & BIẾN ĐỘNG ĐƠN HÀNG")

uploaded_file = st.file_uploader("📂 Upload file Excel", type=["xlsx"])

if uploaded_file:

    # ===== ĐỌC FILE =====
    try:
        df = pd.read_excel(uploaded_file, header=13)
    except:
        df = pd.read_excel(uploaded_file)

    df = df.dropna(how='all')
    df.columns = df.columns.str.strip()

    # ===== KIỂM TRA CỘT =====
    required_cols = ["Ngày chứng từ", "Tên khách hàng", "Mã hàng", "Tên hàng", "Thành tiền bán"]
    for col in required_cols:
        if col not in df.columns:
            st.error(f"❌ Thiếu cột: {col}")
            st.stop()

    # ===== XỬ LÝ =====
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Quý"] = df["Ngày chứng từ"].dt.to_period("Q").astype(str)

    # ===== SIDEBAR FILTER =====
    st.sidebar.header("🔎 Bộ lọc")

    # Khách hàng
    kh = st.sidebar.multiselect("Khách hàng", df["Tên khách hàng"].dropna().unique())
    if kh:
        df = df[df["Tên khách hàng"].isin(kh)]

    # Mã hàng
    mh = st.sidebar.multiselect("Mã hàng", df["Mã hàng"].dropna().unique())
    if mh:
        df = df[df["Mã hàng"].isin(mh)]

    # Vận chuyển
    if "Phương thức vận chuyển" in df.columns:
        vc = st.sidebar.multiselect("Vận chuyển", df["Phương thức vận chuyển"].dropna().unique())
        if vc:
            df = df[df["Phương thức vận chuyển"].isin(vc)]

    # Biển số xe
    if "Số xe" in df.columns:
        xe = st.sidebar.multiselect("Biển số xe", df["Số xe"].dropna().unique())
        if xe:
            df = df[df["Số xe"].isin(xe)]

    # ===== KPI =====
    col1, col2 = st.columns(2)

    col1.metric("💰 Tổng doanh thu", f"{df['Thành tiền bán'].sum():,.0f}")
    col2.metric("📦 Số đơn", len(df))

    st.markdown("---")

    # ===== PHÂN TÍCH THEO KHÁCH HÀNG =====
    st.subheader("📊 Doanh thu theo khách hàng")

    doanhthu_kh = df.groupby("Tên khách hàng")["Thành tiền bán"].sum().reset_index().sort_values(by="Thành tiền bán", ascending=False)

    fig1 = px.bar(doanhthu_kh, x="Tên khách hàng", y="Thành tiền bán", title="Doanh thu theo khách hàng")
    st.plotly_chart(fig1, use_container_width=True)

    # ===== BIẾN ĐỘNG THEO QUÝ =====
    st.subheader("📈 Biến động đơn hàng theo quý")

    pivot = df.groupby(["Quý", "Tên khách hàng"])["Thành tiền bán"].sum().reset_index()

    fig2 = px.line(
        pivot,
        x="Quý",
        y="Thành tiền bán",
        color="Tên khách hàng",
        markers=True,
        title="So sánh doanh thu theo quý giữa các khách hàng"
    )

    st.plotly_chart(fig2, use_container_width=True)

    # ===== PHÂN TÍCH MÃ HÀNG =====
    st.subheader("📦 Mã hàng khách mua")

    hang = df.groupby(["Tên khách hàng", "Mã hàng", "Tên hàng"])["Thành tiền bán"].sum().reset_index()

    st.dataframe(hang)

    # ===== BẢNG CHI TIẾT =====
    st.subheader("📋 Dữ liệu chi tiết")

    st.dataframe(df)