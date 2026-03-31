import streamlit as st
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="Power BI KSKD", layout="wide")

# ===== STYLE =====
st.markdown("""
<style>
body {
    background-color: #f5f6fa;
}
.metric-box {
    background: white;
    padding: 15px;
    border-radius: 12px;
    box-shadow: 0px 2px 8px rgba(0,0,0,0.05);
}
</style>
""", unsafe_allow_html=True)

# ===== SIDEBAR =====
st.sidebar.title("📊 KSKD Dashboard")
file = st.sidebar.file_uploader("📂 Upload Excel", type=["xlsx"])

# ===== CLEAN ADDRESS =====
def clean_address(addr):
    if pd.isna(addr):
        return ""
    addr = addr.lower()
    addr = re.sub(r"tphcm|tp hcm|hồ chí minh", "hcm", addr)
    addr = re.sub(r"hà nội|hn", "hanoi", addr)
    addr = re.sub(r"[^a-z0-9 ]", "", addr)
    return addr.strip()

if file:
    # ===== LOAD =====
    try:
        df = pd.read_excel(file, header=13)
    except:
        df = pd.read_excel(file)

    df = df.dropna(how='all')
    df.columns = df.columns.str.strip()

    # ===== PROCESS =====
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Tháng"] = df["Ngày chứng từ"].dt.to_period("M").astype(str)
    df["Quý"] = df["Ngày chứng từ"].dt.to_period("Q").astype(str)

    df["Địa chỉ chuẩn"] = df["Nơi giao hàng"].apply(clean_address)

    if "Thành tiền vốn" in df.columns:
        df["Lợi nhuận"] = df["Thành tiền bán"] - df["Thành tiền vốn"]
    else:
        df["Lợi nhuận"] = 0

    # ===== FILTER =====
    kh_list = df["Tên khách hàng"].dropna().unique()
    kh = st.sidebar.selectbox("Khách hàng", kh_list)

    df_kh = df[df["Tên khách hàng"] == kh]

    # ===== TABS =====
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Tổng quan",
        "🧠 Thói quen mua",
        "🚚 Giao hàng",
        "⚠️ Thanh toán & Rủi ro"
    ])

    # ===== TAB 1 =====
    with tab1:
        st.subheader("📊 Tổng quan")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("💰 Doanh thu", f"{df_kh['Thành tiền bán'].sum():,.0f}")
        c2.metric("📦 Số đơn", len(df_kh))
        c3.metric("📍 Điểm giao", df_kh["Địa chỉ chuẩn"].nunique())
        c4.metric("📈 Lợi nhuận", f"{df_kh['Lợi nhuận'].sum():,.0f}")

        # Biểu đồ quý
        q = df_kh.groupby("Quý")["Thành tiền bán"].sum().reset_index()
        fig = px.bar(q, x="Quý", y="Thành tiền bán", text="Thành tiền bán")
        st.plotly_chart(fig, use_container_width=True)

    # ===== TAB 2 =====
    with tab2:
        st.subheader("🧠 Thói quen mua hàng")

        # Tần suất
        freq = df_kh.groupby("Tháng").size().reset_index(name="Số đơn")
        fig1 = px.bar(freq, x="Tháng", y="Số đơn", text="Số đơn")
        st.plotly_chart(fig1, use_container_width=True)

        # Sản phẩm
        hang = df_kh.groupby("Tên hàng")["Thành tiền bán"].sum().reset_index().sort_values(by="Thành tiền bán", ascending=False)
        fig2 = px.bar(hang.head(10), x="Tên hàng", y="Thành tiền bán", text="Thành tiền bán")
        st.plotly_chart(fig2, use_container_width=True)

        st.write(f"👉 KH tập trung: **{hang.iloc[0]['Tên hàng']}**")

    # ===== TAB 3 =====
    with tab3:
        st.subheader("🚚 Giao hàng")

        col1, col2 = st.columns(2)

        if "Phương thức vận chuyển" in df.columns:
            vc = df_kh["Phương thức vận chuyển"].value_counts().reset_index()
            vc.columns = ["Hình thức", "Số lần"]
            fig3 = px.pie(vc, names="Hình thức", values="Số lần")
            col1.plotly_chart(fig3, use_container_width=True)

        dia = df_kh.groupby("Địa chỉ chuẩn").size().reset_index(name="Số lần")
        fig4 = px.bar(dia, x="Địa chỉ chuẩn", y="Số lần", text="Số lần")
        col2.plotly_chart(fig4, use_container_width=True)

    # ===== TAB 4 =====
    with tab4:
        st.subheader("⚠️ Thanh toán & Rủi ro")

        risk = []

        # Thanh toán
        if "Ngày thanh toán" in df.columns:
            df_kh["Ngày thanh toán"] = pd.to_datetime(df_kh["Ngày thanh toán"], errors="coerce")
            df_kh["Số ngày TT"] = (df_kh["Ngày thanh toán"] - df_kh["Ngày chứng từ"]).dt.days

            avg = df_kh["Số ngày TT"].mean()
            st.metric("⏳ TB ngày thanh toán", f"{avg:.0f} ngày")

            if avg > 30:
                risk.append("Thanh toán chậm")

        # Lỗ
        if (df_kh["Lợi nhuận"] < 0).sum() > 0:
            risk.append("Có đơn hàng lỗ")

        # Giao nhiều nơi
        if df_kh["Địa chỉ chuẩn"].nunique() > 2:
            risk.append("Giao nhiều nơi")

        # Hiển thị
        if len(risk) == 0:
            st.success("✅ KH an toàn")
        else:
            for r in risk:
                st.warning(f"⚠️ {r}")

else:
    st.info("👈 Upload file bên trái để bắt đầu")