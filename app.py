import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="KSKD Dashboard", layout="wide")

st.title("📊 PHÂN TÍCH KHÁCH HÀNG (KSKD)")

file = st.file_uploader("📂 Upload Excel", type=["xlsx"])

if file:
    # ===== LOAD NHẸ (TRÁNH LAG) =====
    try:
        df = pd.read_excel(file, header=13)
    except:
        df = pd.read_excel(file)

    df = df.dropna(how='all')
    df.columns = df.columns.str.strip()

    # ===== CHUẨN HÓA =====
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Tháng"] = df["Ngày chứng từ"].dt.to_period("M").astype(str)
    df["Quý"] = df["Ngày chứng từ"].dt.to_period("Q").astype(str)

    # ===== LỢI NHUẬN =====
    if "Thành tiền vốn" in df.columns:
        df["Lợi nhuận"] = df["Thành tiền bán"] - df["Thành tiền vốn"]
    else:
        df["Lợi nhuận"] = 0

    # ===== CHỌN KH =====
    kh = st.selectbox("🔎 Chọn khách hàng", df["Tên khách hàng"].dropna().unique())
    df_kh = df[df["Tên khách hàng"] == kh]

    st.markdown("---")

    # ===== KPI =====
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Doanh thu", f"{df_kh['Thành tiền bán'].sum():,.0f}")
    col2.metric("📦 Sản lượng", f"{df_kh['Số lượng'].sum():,.0f}")
    col3.metric("⚖️ Khối lượng", f"{df_kh.get('Khối lượng',0).sum():,.0f}")
    col4.metric("📈 Lợi nhuận", f"{df_kh['Lợi nhuận'].sum():,.0f}")

    # ===== 1. THÓI QUEN MUA =====
    st.subheader("🧠 Thói quen mua hàng")

    hang = df_kh.groupby(["Tên hàng"])["Thành tiền bán"].sum().reset_index().sort_values(by="Thành tiền bán", ascending=False)

    fig1 = px.bar(hang.head(10), x="Tên hàng", y="Thành tiền bán", title="Top hàng mua")
    st.plotly_chart(fig1, use_container_width=True)

    st.write(f"👉 KH tập trung mua: **{hang.iloc[0]['Tên hàng']}**")

    # ===== 2. BIẾN ĐỘNG =====
    st.subheader("📈 Biến động theo thời gian")

    time = df_kh.groupby("Tháng")[["Thành tiền bán","Số lượng","Lợi nhuận"]].sum().reset_index()

    fig2 = px.line(time, x="Tháng", y=["Thành tiền bán","Số lượng","Lợi nhuận"], markers=True)
    st.plotly_chart(fig2, use_container_width=True)

    # ===== 3. GIAO HÀNG =====
    st.subheader("🚚 Giao hàng")

    col1, col2 = st.columns(2)

    if "Phương thức vận chuyển" in df.columns:
        vc = df_kh["Phương thức vận chuyển"].value_counts()
        col1.write(vc)

    dia = df_kh["Nơi giao hàng"].value_counts()
    col2.write(dia)

    if dia.nunique() > 1:
        st.warning("⚠️ KH thay đổi nơi giao hàng → cần kiểm soát")

    # ===== 4. THANH TOÁN =====
    st.subheader("💳 Thanh toán (BCCN)")

    if "Ngày thanh toán" in df.columns:
        df_kh["Ngày thanh toán"] = pd.to_datetime(df_kh["Ngày thanh toán"], errors="coerce")
        df_kh["Số ngày TT"] = (df_kh["Ngày thanh toán"] - df_kh["Ngày chứng từ"]).dt.days

        avg_days = df_kh["Số ngày TT"].mean()
        st.metric("⏳ TB ngày thanh toán", f"{avg_days:.0f} ngày")

        if avg_days > 30:
            st.error("❌ Thanh toán chậm → rủi ro cao")
        else:
            st.success("✅ Thanh toán ổn định")

    # ===== 5. THỊ PHẦN =====
    st.subheader("🏆 Thị phần")

    market = df.groupby("Tên khách hàng")["Thành tiền bán"].sum().reset_index()
    market = market.sort_values(by="Thành tiền bán", ascending=False)

    rank = market[market["Tên khách hàng"] == kh].index[0] + 1

    st.write(f"👉 KH đang đứng TOP {rank} doanh thu")

    fig3 = px.bar(market.head(10), x="Tên khách hàng", y="Thành tiền bán", title="Top khách hàng")
    st.plotly_chart(fig3, use_container_width=True)

    # ===== 6. RỦI RO =====
    st.subheader("⚠️ Đánh giá rủi ro")

    risk = []

    if (df_kh["Lợi nhuận"] < 0).sum() > 0:
        risk.append("Có đơn hàng lỗ")

    if df_kh["Nơi giao hàng"].nunique() > 2:
        risk.append("Giao nhiều nơi bất thường")

    if "Số ngày TT" in df_kh and df_kh["Số ngày TT"].mean() > 30:
        risk.append("Thanh toán chậm")

    if len(risk) == 0:
        st.success("✅ KH an toàn")
    else:
        for r in risk:
            st.warning(f"⚠️ {r}")

    # ===== DATA =====
    st.subheader("📋 Chi tiết")
    st.dataframe(df_kh)