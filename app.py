import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="KSKD PRO MAX", layout="wide")

st.title("📊 Dashboard KSKD (FAST VERSION)")

# ================= CACHE =================
@st.cache_data
def load_all_data():
    if not os.path.exists("data"):
        return pd.DataFrame()

    files = os.listdir("data")
    df_list = []

    for f in files:
        try:
            df = pd.read_excel(f"data/{f}", header=13, engine="openpyxl")
            df_list.append(df)
        except:
            continue

    if len(df_list) == 0:
        return pd.DataFrame()

    df = pd.concat(df_list, ignore_index=True)
    df.columns = df.columns.str.strip()

    # ===== CHỈ GIỮ CỘT CẦN =====
    cols = [
        "Ngày chứng từ", "Tên khách hàng",
        "Nơi giao hàng", "Biển số xe",
        "Thành tiền bán", "Thành tiền vốn"
    ]
    df = df[[c for c in cols if c in df.columns]]

    # ===== XỬ LÝ NHANH =====
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Tháng"] = df["Ngày chứng từ"].dt.strftime("%Y-%m")

    # ===== TỈNH (vector hóa) =====
    df["Tỉnh"] = "Khác"
    df.loc[df["Nơi giao hàng"].str.contains("hcm", case=False, na=False), "Tỉnh"] = "TP HCM"
    df.loc[df["Nơi giao hàng"].str.contains("hà nội", case=False, na=False), "Tỉnh"] = "Hà Nội"
    df.loc[df["Nơi giao hàng"].str.contains("bình dương", case=False, na=False), "Tỉnh"] = "Bình Dương"
    df.loc[df["Nơi giao hàng"].str.contains("đồng nai", case=False, na=False), "Tỉnh"] = "Đồng Nai"

    # ===== XE =====
    df["Biển số xe"] = df["Biển số xe"].astype(str).str.replace(" ", "").str.upper()
    df = df[~df["Biển số xe"].isin(["GK", "NAN", ""])]
    df = df[df["Biển số xe"].str.len().between(7,9)]

    # ===== LỢI NHUẬN =====
    if "Thành tiền vốn" in df.columns:
        df["Lợi nhuận"] = df["Thành tiền bán"] - df["Thành tiền vốn"]
    else:
        df["Lợi nhuận"] = 0

    return df


# ================= UPLOAD =================
if not os.path.exists("data"):
    os.makedirs("data")

file = st.sidebar.file_uploader("📂 Upload Excel", type=["xlsx"])

if file:
    with open(f"data/{file.name}", "wb") as f:
        f.write(file.getbuffer())
    st.cache_data.clear()  # clear cache để load file mới
    st.success("Upload thành công!")

# ================= LOAD =================
df = load_all_data()

if df.empty:
    st.warning("👉 Upload file để bắt đầu")
    st.stop()

# ================= FILTER =================
kh_list = df["Tên khách hàng"].dropna().unique()
kh = st.sidebar.selectbox("Chọn khách hàng", kh_list)

df_kh = df[df["Tên khách hàng"] == kh]

# ================= TABS =================
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Doanh thu",
    "📦 Tần suất",
    "🗺️ Giao hàng",
    "🤖 Insight"
])

# ================= TAB 1 =================
with tab1:
    dt = df_kh.groupby("Tháng")["Thành tiền bán"].sum().reset_index()

    fig = px.line(dt, x="Tháng", y="Thành tiền bán", markers=True)
    st.plotly_chart(fig, use_container_width=True)

# ================= TAB 2 =================
with tab2:
    freq = df_kh.groupby("Tháng").size().reset_index(name="Số đơn")

    fig = px.bar(freq, x="Tháng", y="Số đơn")
    st.plotly_chart(fig, use_container_width=True)

# ================= TAB 3 =================
with tab3:
    tinh = df_kh.groupby("Tỉnh")["Thành tiền bán"].sum().reset_index()

    fig = px.bar(tinh, x="Tỉnh", y="Thành tiền bán")
    st.plotly_chart(fig, use_container_width=True)

    # MAP NHẸ
    map_data = pd.DataFrame({
        "Tỉnh": ["TP HCM","Hà Nội","Bình Dương","Đồng Nai"],
        "lat": [10.82,21.02,11.07,10.95],
        "lon": [106.63,105.85,106.67,106.82]
    })

    map_merge = map_data.merge(tinh, on="Tỉnh", how="left").fillna(0)

    st.map(map_merge.rename(columns={"lat":"latitude","lon":"longitude"}))

# ================= TAB 4 =================
with tab4:
    insights = []

    if not dt.empty:
        top = dt.sort_values("Thành tiền bán", ascending=False).iloc[0]
        insights.append(f"🔥 Mua nhiều nhất: {top['Tháng']}")

    insights.append(f"📦 TB {round(freq['Số đơn'].mean(),1)} đơn/tháng")

    top_tinh = tinh.sort_values("Thành tiền bán", ascending=False).iloc[0]
    insights.append(f"📍 Giao nhiều nhất: {top_tinh['Tỉnh']}")

    if df_kh["Biển số xe"].nunique() > 5:
        insights.append("⚠️ Nhiều xe → rủi ro")

    for i in insights:
        st.write("👉", i)