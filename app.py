import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="Phân tích kinh doanh", layout="wide")

st.title("📊 Dashboard Phân tích khách hàng")

# ===== TẠO FOLDER DATA =====
if not os.path.exists("data"):
    os.makedirs("data")

# ===== UPLOAD FILE =====
file = st.file_uploader("📂 Upload file Excel", type=["xlsx"])

if file is not None:
    with open(f"data/{file.name}", "wb") as f:
        f.write(file.getbuffer())

    st.success("✅ Upload thành công")
    st.cache_data.clear()
    st.rerun()

# ===== LOAD DATA =====
@st.cache_data
def load_data():
    files = os.listdir("data")

    if len(files) == 0:
        return pd.DataFrame()

    df_list = []

    for f in files:
        try:
            path = os.path.join("data", f)

            # 🔥 FIX HEADER ERP
            df = pd.read_excel(path, header=13, engine="openpyxl")

            df_list.append(df)

        except Exception as e:
            st.error(f"Lỗi file {f}: {e}")

    if not df_list:
        return pd.DataFrame()

    df = pd.concat(df_list, ignore_index=True)

    # ===== CLEAN CỘT =====
    df.columns = df.columns.astype(str).str.strip()

    # ===== RENAME =====
    df = df.rename(columns={
        "Tên KH": "Tên khách hàng",
        "Khách hàng": "Tên khách hàng",
        "Ngày CT": "Ngày chứng từ",
        "Địa chỉ giao hàng": "Nơi giao hàng"
    })

    if "Tên khách hàng" not in df.columns:
        st.error("❌ Không đọc được file - sai format")
        return pd.DataFrame()

    # ===== XỬ LÝ DATA =====
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Tháng"] = df["Ngày chứng từ"].dt.strftime("%Y-%m")

    df["Thành tiền bán"] = pd.to_numeric(df.get("Thành tiền bán", 0), errors="coerce")
    df["Thành tiền vốn"] = pd.to_numeric(df.get("Thành tiền vốn", 0), errors="coerce")

    df["Nơi giao hàng"] = df.get("Nơi giao hàng", "Không xác định").astype(str)

    # ===== GỘP ĐỊA CHỈ =====
    df["Tỉnh"] = df["Nơi giao hàng"].str.extract(r'(Hà Nội|Hồ Chí Minh|Bắc Ninh|Hải Phòng|Đà Nẵng|Bình Dương|Đồng Nai)', expand=False)
    df["Tỉnh"] = df["Tỉnh"].fillna("Khác")

    # ===== XE =====
    df["Biển số xe"] = df.get("Biển số xe", "").astype(str)
    df["Biển số xe"] = df["Biển số xe"].str.replace(" ", "").str.upper()
    df = df[df["Biển số xe"].str.len().between(7, 9)]

    return df


df = load_data()

if df.empty:
    st.warning("⚠️ Chưa có dữ liệu")
    st.stop()

# ===== FILTER =====
kh = st.selectbox("👤 Chọn khách hàng", df["Tên khách hàng"].dropna().unique())

df_kh = df[df["Tên khách hàng"] == kh]

# ===== KPI =====
col1, col2, col3 = st.columns(3)

col1.metric("💰 Doanh thu", f"{df_kh['Thành tiền bán'].sum():,.0f}")
col2.metric("📦 Số đơn", len(df_kh))
col3.metric("🚚 Số xe", df_kh["Biển số xe"].nunique())

# ===== BIỂU ĐỒ DOANH THU =====
st.subheader("📈 Doanh thu theo tháng")

df_thang = df_kh.groupby("Tháng")["Thành tiền bán"].sum().reset_index()

fig1 = px.line(
    df_thang,
    x="Tháng",
    y="Thành tiền bán",
    markers=True
)

st.plotly_chart(fig1, use_container_width=True)

# ===== TẦN SUẤT MUA =====
st.subheader("🔁 Tần suất mua hàng")

df_freq = df_kh.groupby("Tháng").size().reset_index(name="Số đơn")

fig2 = px.bar(
    df_freq,
    x="Tháng",
    y="Số đơn",
)

st.plotly_chart(fig2, use_container_width=True)

# ===== SẢN PHẨM =====
st.subheader("📦 Top sản phẩm")

df_sp = df_kh.groupby("Tên mã hàng")["Thành tiền bán"].sum().reset_index()

df_sp = df_sp.sort_values(by="Thành tiền bán", ascending=False).head(10)

fig3 = px.bar(
    df_sp,
    x="Tên mã hàng",
    y="Thành tiền bán"
)

st.plotly_chart(fig3, use_container_width=True)

# ===== NƠI GIAO HÀNG =====
st.subheader("📍 Nơi giao hàng")

df_map = df_kh.groupby("Tỉnh")["Thành tiền bán"].sum().reset_index()

fig4 = px.pie(
    df_map,
    names="Tỉnh",
    values="Thành tiền bán"
)

st.plotly_chart(fig4, use_container_width=True)

# ===== INSIGHT AI =====
st.subheader("🧠 Insight tự động")

if not df_thang.empty:
    top_month = df_thang.loc[df_thang["Thành tiền bán"].idxmax()]

    st.info(
        f"""
        📌 Khách hàng mua nhiều nhất vào tháng: {top_month['Tháng']}  
        💰 Doanh thu cao nhất: {top_month['Thành tiền bán']:,.0f}
        """
    )