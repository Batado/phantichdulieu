import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(layout="wide")
st.title("📊 Dashboard Phân tích khách hàng")

# ===== TẠO THƯ MỤC =====
DATA_PATH = "data"
os.makedirs(DATA_PATH, exist_ok=True)

# ===== UPLOAD =====
file = st.file_uploader("📂 Upload file Excel", type=["xlsx"])

if file is not None:
    path = os.path.join(DATA_PATH, file.name)
    with open(path, "wb") as f:
        f.write(file.getbuffer())

    st.success("✅ Upload thành công")
    st.cache_data.clear()
    st.rerun()

# ===== LOAD DATA =====
@st.cache_data
def load_data():
    files = os.listdir(DATA_PATH)
    if not files:
        return pd.DataFrame()

    all_df = []

    for f in files:
        try:
            path = os.path.join(DATA_PATH, f)

            # 🔥 FIX 1: ĐỌC ĐÚNG SHEET
            xl = pd.ExcelFile(path)
            sheet = xl.sheet_names[0]

            # 🔥 FIX 2: ĐỌC HEADER ĐÚNG (ERP)
            df = pd.read_excel(path, sheet_name=sheet, header=13)

            all_df.append(df)

        except Exception as e:
            st.error(f"Lỗi file {f}: {e}")

    if not all_df:
        return pd.DataFrame()

    df = pd.concat(all_df, ignore_index=True)

    # ===== CLEAN COLUMN =====
    df.columns = df.columns.astype(str).str.strip()

    # ===== DEBUG =====
    st.write("📌 Columns:", df.columns.tolist())
    st.write("📊 Shape:", df.shape)

    # ===== RENAME =====
    df.rename(columns={
        "Tên KH": "Tên khách hàng",
        "Khách hàng": "Tên khách hàng",
        "Ngày CT": "Ngày chứng từ",
        "Địa chỉ giao hàng": "Nơi giao hàng"
    }, inplace=True)

    # ===== CHECK =====
    if "Tên khách hàng" not in df.columns:
        st.error("❌ Sai format file (không có cột KH)")
        return pd.DataFrame()

    # ===== XỬ LÝ =====
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Tháng"] = df["Ngày chứng từ"].dt.strftime("%Y-%m")

    df["Thành tiền bán"] = pd.to_numeric(df.get("Thành tiền bán", 0), errors="coerce")

    df["Nơi giao hàng"] = df.get("Nơi giao hàng", "").astype(str)

    # ===== GỘP TỈNH =====
    df["Tỉnh"] = "Khác"
    df.loc[df["Nơi giao hàng"].str.contains("hcm", case=False, na=False), "Tỉnh"] = "TP HCM"
    df.loc[df["Nơi giao hàng"].str.contains("hà nội", case=False, na=False), "Tỉnh"] = "Hà Nội"
    df.loc[df["Nơi giao hàng"].str.contains("bình dương", case=False, na=False), "Tỉnh"] = "Bình Dương"
    df.loc[df["Nơi giao hàng"].str.contains("đồng nai", case=False, na=False), "Tỉnh"] = "Đồng Nai"

    # ===== XE =====
    df["Biển số xe"] = df.get("Biển số xe", "").astype(str)
    df["Biển số xe"] = df["Biển số xe"].str.replace(" ", "").str.upper()
    df = df[df["Biển số xe"].str.len().between(7, 9)]

    return df


df = load_data()

# ===== STOP NẾU RỖNG =====
if df.empty:
    st.warning("⚠️ Không có dữ liệu sau khi load")
    st.stop()

# ===== FILTER KH =====
kh_list = df["Tên khách hàng"].dropna().unique()

if len(kh_list) == 0:
    st.error("❌ Không có khách hàng")
    st.stop()

kh = st.selectbox("👤 Chọn khách hàng", kh_list)
df_kh = df[df["Tên khách hàng"] == kh]

# ===== KPI =====
c1, c2, c3 = st.columns(3)
c1.metric("💰 Doanh thu", f"{df_kh['Thành tiền bán'].sum():,.0f}")
c2.metric("📦 Số đơn", len(df_kh))
c3.metric("🚚 Số xe", df_kh["Biển số xe"].nunique())

# ===== DOANH THU =====
st.subheader("📈 Doanh thu theo tháng")
dt = df_kh.groupby("Tháng")["Thành tiền bán"].sum().reset_index()

fig1 = px.line(dt, x="Tháng", y="Thành tiền bán", markers=True)
st.plotly_chart(fig1, use_container_width=True)

# ===== TẦN SUẤT =====
st.subheader("🔁 Tần suất mua")
freq = df_kh.groupby("Tháng").size().reset_index(name="Số đơn")

fig2 = px.bar(freq, x="Tháng", y="Số đơn")
st.plotly_chart(fig2, use_container_width=True)

# ===== GIAO HÀNG =====
st.subheader("📍 Khu vực giao hàng")
tinh = df_kh.groupby("Tỉnh")["Thành tiền bán"].sum().reset_index()

fig3 = px.pie(tinh, names="Tỉnh", values="Thành tiền bán")
st.plotly_chart(fig3, use_container_width=True)

# ===== INSIGHT =====
st.subheader("🧠 Insight")

if not dt.empty:
    top = dt.sort_values("Thành tiền bán", ascending=False).iloc[0]
    st.success(f"🔥 Mua nhiều nhất tháng {top['Tháng']}")

st.info(f"📦 TB {round(freq['Số đơn'].mean(),1)} đơn/tháng")