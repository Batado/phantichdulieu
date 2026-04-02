import streamlit as st
import pandas as pd
import plotly.express as px
import re
import os

st.set_page_config(page_title="KSKD PRO MAX", layout="wide")

st.sidebar.title("📊 Dashboard KSKD PRO")

# ===== LƯU FILE =====
if not os.path.exists("data"):
    os.makedirs("data")

file = st.sidebar.file_uploader("📂 Upload Excel", type=["xlsx"])

if file:
    filepath = f"data/{file.name}"
    with open(filepath, "wb") as f:
        f.write(file.getbuffer())

# ===== LOAD HISTORY =====
all_files = os.listdir("data")

df_list = []
for f in all_files:
    try:
        df_temp = pd.read_excel(f"data/{f}", header=13)
        df_list.append(df_temp)
    except:
        pass

if len(df_list) > 0:
    df = pd.concat(df_list, ignore_index=True)
    df = df.dropna(how='all')
    df.columns = df.columns.str.strip()

    # ===== XỬ LÝ =====
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Tháng"] = df["Ngày chứng từ"].dt.to_period("M").astype(str)

    # ===== TÁCH TỈNH =====
    def extract_province(addr):
        if pd.isna(addr):
            return "Không xác định"
        addr = str(addr).lower()

        if "hcm" in addr:
            return "TP HCM"
        if "hà nội" in addr:
            return "Hà Nội"
        if "bình dương" in addr:
            return "Bình Dương"
        if "đồng nai" in addr:
            return "Đồng Nai"

        return "Khác"

    df["Tỉnh"] = df["Nơi giao hàng"].apply(extract_province)

    # ===== BIỂN SỐ XE =====
    def clean_vehicle(x):
        if pd.isna(x):
            return None
        x = str(x).replace(" ", "").upper()
        if x in ["GK", ""]:
            return None
        if 7 <= len(x) <= 9:
            return x
        return None

    df["Xe"] = df["Biển số xe"].apply(clean_vehicle)

    # ===== LỢI NHUẬN =====
    if "Thành tiền vốn" in df.columns:
        df["Lợi nhuận"] = df["Thành tiền bán"] - df["Thành tiền vốn"]
    else:
        df["Lợi nhuận"] = 0

    # ===== CHỌN KH =====
    kh = st.sidebar.selectbox("Khách hàng", df["Tên khách hàng"].dropna().unique())
    df_kh = df[df["Tên khách hàng"] == kh]

    # ===== TABS =====
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Tổng quan",
        "🧠 Thói quen",
        "🗺️ Giao hàng",
        "🤖 Insight AI"
    ])

    # ================= TAB 1 =================
    with tab1:
        st.subheader("📊 Doanh thu theo tháng")

        dt = df_kh.groupby("Tháng")["Thành tiền bán"].sum().reset_index()

        fig = px.line(dt, x="Tháng", y="Thành tiền bán", markers=True,
                      text="Thành tiền bán")

        fig.update_traces(textposition="top center")
        st.plotly_chart(fig, use_container_width=True)

    # ================= TAB 2 =================
    with tab2:
        st.subheader("📦 Tần suất mua hàng")

        freq = df_kh.groupby("Tháng").size().reset_index(name="Số đơn")

        fig = px.bar(freq, x="Tháng", y="Số đơn", text="Số đơn")
        fig.update_traces(textposition="outside")

        st.plotly_chart(fig, use_container_width=True)

    # ================= TAB 3 =================
    with tab3:
        st.subheader("🗺️ Phân tích theo tỉnh")

        tinh = df_kh.groupby("Tỉnh")["Thành tiền bán"].sum().reset_index()

        fig = px.bar(tinh, x="Tỉnh", y="Thành tiền bán",
                     text="Thành tiền bán",
                     title="Doanh thu theo tỉnh")

        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

        # ===== MAP =====
        st.subheader("🌍 Bản đồ giao hàng")

        map_data = pd.DataFrame({
            "Tỉnh": ["TP HCM","Hà Nội","Bình Dương","Đồng Nai"],
            "lat": [10.82,21.02,11.07,10.95],
            "lon": [106.63,105.85,106.67,106.82]
        })

        map_merge = pd.merge(map_data, tinh, on="Tỉnh", how="left").fillna(0)

        st.map(map_merge.rename(columns={"lat":"latitude","lon":"longitude"}))

    # ================= TAB 4 =================
    with tab4:
        st.subheader("🤖 Insight tự động")

        insights = []

        # THÁNG MUA NHIỀU NHẤT
        if not dt.empty:
            top_month = dt.sort_values(by="Thành tiền bán", ascending=False).iloc[0]
            insights.append(f"📈 Mua nhiều nhất vào {top_month['Tháng']} ({top_month['Thành tiền bán']:,.0f})")

        # TẦN SUẤT
        avg_freq = freq["Số đơn"].mean()
        insights.append(f"🔁 Trung bình {avg_freq:.1f} đơn/tháng")

        # TỈNH CHÍNH
        top_tinh = tinh.sort_values(by="Thành tiền bán", ascending=False).iloc[0]
        insights.append(f"📍 Tập trung giao tại {top_tinh['Tỉnh']}")

        # RỦI RO
        if df_kh["Xe"].nunique() > 5:
            insights.append("⚠️ Dùng nhiều xe → rủi ro vận chuyển")

        if df_kh["Tỉnh"].nunique() > 3:
            insights.append("⚠️ Giao nhiều tỉnh → phân tán")

        for i in insights:
            st.write("👉", i)

else:
    st.info("👈 Upload file để bắt đầu")