import streamlit as st
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="KSKD Dashboard Pro", layout="wide")

# ===== STYLE =====
st.markdown("""
<style>
body {background-color: #f5f6fa;}
.block-container {padding-top: 1rem;}
</style>
""", unsafe_allow_html=True)

st.sidebar.title("📊 Dashboard KSKD")
file = st.sidebar.file_uploader("📂 Upload Excel", type=["xlsx"])

# ===== HÀM GỘP ĐỊA CHỈ =====
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

    # ===== XỬ LÝ =====
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Tháng"] = df["Ngày chứng từ"].dt.to_period("M").astype(str)
    df["Quý"] = df["Ngày chứng từ"].dt.to_period("Q").astype(str)

    df["Địa chỉ chuẩn"] = df["Nơi giao hàng"].apply(clean_address)

    if "Thành tiền vốn" in df.columns:
        df["Lợi nhuận"] = df["Thành tiền bán"] - df["Thành tiền vốn"]
    else:
        df["Lợi nhuận"] = 0

    # ===== FILTER =====
    kh = st.sidebar.selectbox("Chọn khách hàng", df["Tên khách hàng"].dropna().unique())
    df_kh = df[df["Tên khách hàng"] == kh]

    # ===== TABS =====
    tab1, tab2, tab3 = st.tabs([
        "📊 Tổng quan",
        "🧠 Thói quen mua",
        "🚚 Giao hàng"
    ])

    # =========================
    # TAB 1: TỔNG QUAN
    # =========================
    with tab1:
        st.subheader("📊 Tổng quan")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("💰 Doanh thu", f"{df_kh['Thành tiền bán'].sum():,.0f}")
        c2.metric("📦 Số đơn", len(df_kh))
        c3.metric("📍 Điểm giao", df_kh["Địa chỉ chuẩn"].nunique())
        c4.metric("📈 Lợi nhuận", f"{df_kh['Lợi nhuận'].sum():,.0f}")

        # Doanh thu theo quý
        q = df_kh.groupby("Quý")["Thành tiền bán"].sum().reset_index()

        fig_q = px.bar(q, x="Quý", y="Thành tiền bán",
                       text="Thành tiền bán",
                       title="Doanh thu theo quý")

        fig_q.update_traces(textposition="outside")
        st.plotly_chart(fig_q, use_container_width=True)

    # =========================
    # TAB 2: THÓI QUEN MUA
    # =========================
    with tab2:
        st.subheader("📅 Tần suất mua theo tháng")

        freq = df_kh.groupby("Tháng").size().reset_index(name="Số đơn")

        fig1 = px.bar(freq, x="Tháng", y="Số đơn",
                      text="Số đơn",
                      title="Số lần mua theo tháng")

        fig1.update_traces(textposition="outside")
        st.plotly_chart(fig1, use_container_width=True)

        # ===== TẦN SUẤT MÃ HÀNG =====
        st.subheader("📦 Tần suất mua theo mã hàng")

        freq_product = df_kh.groupby("Tên hàng").size().reset_index(name="Số lần")
        freq_product = freq_product.sort_values(by="Số lần", ascending=False)

        fig2 = px.bar(
            freq_product.head(10),
            x="Tên hàng",
            y="Số lần",
            text="Số lần",
            title="Top mã hàng theo số lần mua"
        )

        fig2.update_traces(textposition="outside")
        fig2.update_layout(xaxis_tickangle=-30)

        st.plotly_chart(fig2, use_container_width=True)

        # ===== HEATMAP =====
        st.subheader("🔥 Heatmap thói quen mua")

        heat = df_kh.groupby(["Tháng","Tên hàng"]).size().reset_index(name="Số lần")

        fig_heat = px.density_heatmap(
            heat,
            x="Tháng",
            y="Tên hàng",
            z="Số lần",
            color_continuous_scale="Blues"
        )

        st.plotly_chart(fig_heat, use_container_width=True)

    # =========================
    # TAB 3: GIAO HÀNG
    # =========================
    with tab3:
        st.subheader("📍 Phân bố nơi giao hàng")

        dia = df_kh.groupby("Địa chỉ chuẩn").size().reset_index(name="Số lần")
        dia = dia.sort_values(by="Số lần", ascending=False)

        # Top 5 + Others
        top = dia.head(5)
        others = pd.DataFrame({
            "Địa chỉ chuẩn": ["Others"],
            "Số lần": [dia["Số lần"][5:].sum()]
        })

        dia_plot = pd.concat([top, others]) if len(dia) > 5 else dia

        dia_plot["%"] = dia_plot["Số lần"] / dia_plot["Số lần"].sum()

        fig3 = px.bar(
            dia_plot,
            x="Địa chỉ chuẩn",
            y="Số lần",
            text=dia_plot["Số lần"].astype(str) + " (" + (dia_plot["%"]*100).round(1).astype(str) + "%)",
            title="Top khu vực giao hàng"
        )

        fig3.update_traces(textposition="outside")
        fig3.update_layout(xaxis_tickangle=-30)

        st.plotly_chart(fig3, use_container_width=True)

        # ===== VẬN CHUYỂN =====
        if "Phương thức vận chuyển" in df.columns:
            st.subheader("🚚 Hình thức vận chuyển")

            vc = df_kh["Phương thức vận chuyển"].value_counts().reset_index()
            vc.columns = ["Hình thức", "Số lần"]

            fig4 = px.pie(vc, names="Hình thức", values="Số lần")
            st.plotly_chart(fig4, use_container_width=True)

else:
    st.info("👈 Upload file bên trái để bắt đầu")