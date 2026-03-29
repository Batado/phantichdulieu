import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="KSKD PRO MAX", layout="wide")

st.title("🚀 DASHBOARD KIỂM SOÁT KINH DOANH - PRO MAX")

uploaded_file = st.file_uploader("📂 Upload file Excel", type=["xlsx"])

if uploaded_file:

    # ===== READ FILE =====
    try:
        df = pd.read_excel(uploaded_file, header=13)
    except:
        df = pd.read_excel(uploaded_file)

    df = df.dropna(how='all')
    df.columns = df.columns.str.strip()

    # ===== CHECK =====
    required = ["Ngày chứng từ","Tên khách hàng","Thành tiền bán","Thành tiền vốn"]
    for col in required:
        if col not in df.columns:
            st.error(f"Thiếu cột: {col}")
            st.stop()

    # ===== PROCESS =====
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Doanh thu"] = df["Thành tiền bán"]
    df["Lợi nhuận"] = df["Thành tiền bán"] - df["Thành tiền vốn"]

    # ===== SIDEBAR =====
    st.sidebar.header("🔎 Bộ lọc")

    # Date filter
    if df["Ngày chứng từ"].notna().any():
        min_date = df["Ngày chứng từ"].min()
        max_date = df["Ngày chứng từ"].max()
        date_range = st.sidebar.date_input("Chọn ngày", [min_date, max_date])

        if len(date_range) == 2:
            df = df[(df["Ngày chứng từ"] >= pd.to_datetime(date_range[0])) &
                    (df["Ngày chứng từ"] <= pd.to_datetime(date_range[1]))]

    # KH filter
    kh = st.sidebar.multiselect("Khách hàng", df["Tên khách hàng"].dropna().unique())
    if kh:
        df = df[df["Tên khách hàng"].isin(kh)]

    # Xe filter
    if "Số xe" in df.columns:
        xe = st.sidebar.multiselect("Biển số xe", df["Số xe"].dropna().unique())
        if xe:
            df = df[df["Số xe"].isin(xe)]

    # ===== KPI =====
    col1, col2, col3, col4 = st.columns(4)

    col1.metric("💰 Doanh thu", f"{df['Doanh thu'].sum():,.0f}")
    col2.metric("📈 Lợi nhuận", f"{df['Lợi nhuận'].sum():,.0f}")
    col3.metric("📦 Sản lượng", f"{df['Số lượng'].sum():,.0f}" if "Số lượng" in df.columns else "N/A")

    if "Tên khách hàng" in df.columns:
        col4.metric("👥 Số KH", df["Tên khách hàng"].nunique())

    st.markdown("---")

    # ===== CHART =====
    tab1, tab2, tab3 = st.tabs(["📊 Tổng quan", "📦 Sản phẩm", "🚚 Vận chuyển"])

    with tab1:
        col1, col2 = st.columns(2)

        top_kh = df.groupby("Tên khách hàng")["Doanh thu"].sum().reset_index().sort_values(by="Doanh thu", ascending=False).head(10)
        fig1 = px.bar(top_kh, x="Tên khách hàng", y="Doanh thu", title="Top khách hàng")
        col1.plotly_chart(fig1, use_container_width=True)

        bottom_kh = df.groupby("Tên khách hàng")["Doanh thu"].sum().reset_index().sort_values(by="Doanh thu", ascending=True).head(10)
        fig2 = px.bar(bottom_kh, x="Tên khách hàng", y="Doanh thu", title="Khách hàng thấp nhất")
        col2.plotly_chart(fig2, use_container_width=True)

        dt = df.groupby("Ngày chứng từ")["Doanh thu"].sum().reset_index()
        fig3 = px.line(dt, x="Ngày chứng từ", y="Doanh thu", title="Doanh thu theo thời gian")
        st.plotly_chart(fig3, use_container_width=True)

    with tab2:
        if "Mã hàng" in df.columns:
            hang = df.groupby("Mã hàng")["Doanh thu"].sum().reset_index().sort_values(by="Doanh thu", ascending=False).head(10)
            fig4 = px.bar(hang, x="Mã hàng", y="Doanh thu", title="Top mã hàng")
            st.plotly_chart(fig4, use_container_width=True)

    with tab3:
        if "Phương thức vận chuyển" in df.columns:
            vc = df.groupby("Phương thức vận chuyển")["Doanh thu"].sum().reset_index()
            fig5 = px.pie(vc, names="Phương thức vận chuyển", values="Doanh thu", title="Cơ cấu vận chuyển")
            st.plotly_chart(fig5, use_container_width=True)

    # ===== RỦI RO =====
    st.subheader("⚠️ CẢNH BÁO RỦI RO")

    col1, col2 = st.columns(2)

    # Lỗ
    lo = df[df["Lợi nhuận"] < 0]
    col1.metric("Đơn hàng lỗ", len(lo))

    # Xe nhiều KH
    if "Số xe" in df.columns:
        xe_kh = df.groupby("Số xe")["Tên khách hàng"].nunique().reset_index()
        xe_risk = xe_kh[xe_kh["Tên khách hàng"] > 1]
        col2.metric("Xe nhiều KH", len(xe_risk))

    if not lo.empty:
        st.write("💸 Chi tiết đơn lỗ")
        st.dataframe(lo)

    if "Số xe" in df.columns and not xe_risk.empty:
        st.write("🚚 Xe rủi ro")
        st.dataframe(xe_risk)

    # ===== EXPORT =====
    st.subheader("📤 Xuất dữ liệu")

    st.download_button(
        label="📥 Tải file Excel",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="data_phan_tich.csv",
        mime="text/csv"
    )

    # ===== DATA =====
    st.subheader("📋 Dữ liệu chi tiết")
    st.dataframe(df)