import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(layout="wide")
st.title("📊 Dashboard Phân tích khách hàng")

DATA_PATH = "data"
os.makedirs(DATA_PATH, exist_ok=True)

# ================= UPLOAD =================
file = st.file_uploader("📂 Upload Excel", type=["xlsx"])

if file is not None:
    path = os.path.join(DATA_PATH, file.name)
    with open(path, "wb") as f:
        f.write(file.getbuffer())

    st.success("✅ Upload thành công")
    st.cache_data.clear()
    st.rerun()


# ================= AUTO READ =================
@st.cache_data
def read_excel_smart(path):
    xl = pd.ExcelFile(path)

    for sheet in xl.sheet_names:
        try:
            df_raw = pd.read_excel(path, sheet_name=sheet, header=None)

            # tìm dòng chứa header thật
            for i in range(0, 30):
                row_text = " ".join(df_raw.iloc[i].astype(str)).lower()

                if "kh" in row_text or "khách" in row_text:
                    df = pd.read_excel(path, sheet_name=sheet, header=i)

                    # nếu có dữ liệu thật thì return luôn
                    if df.shape[1] > 5:
                        return df

        except:
            continue

    return pd.DataFrame()


# ================= LOAD =================
@st.cache_data
def load_data():
    files = os.listdir(DATA_PATH)

    if not files:
        return pd.DataFrame()

    all_df = []

    for f in files:
        try:
            path = os.path.join(DATA_PATH, f)

            df = read_excel_smart(path)

            if not df.empty:
                all_df.append(df)

        except Exception as e:
            st.error(f"Lỗi file {f}: {e}")

    if not all_df:
        return pd.DataFrame()

    df = pd.concat(all_df, ignore_index=True)

    # ===== CLEAN =====
    df.columns = df.columns.astype(str).str.strip()

    # ===== DEBUG =====
    st.write("📌 Columns:", df.columns.tolist())
    st.write("📊 Shape:", df.shape)

    # ===== RENAME =====
    df.rename(columns={
        "Tên KH": "Tên khách hàng",
        "Khách hàng": "Tên khách hàng",
        "Tên khách": "Tên khách hàng",
        "Ngày CT": "Ngày chứng từ",
        "Ngày": "Ngày chứng từ",
        "Địa chỉ giao hàng": "Nơi giao hàng"
    }, inplace=True)

    # ===== CHECK =====
    if "Tên khách hàng" not in df.columns:
        st.error("❌ Không nhận diện được file")
        return pd.DataFrame()

    # ===== FIX DATA =====
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

if df.empty:
    st.error("❌ Không đọc được dữ liệu từ file")
    st.stop()

# ================= FILTER =================
kh_list = df["Tên khách hàng"].dropna().unique()

if len(kh_list) == 0:
    st.error("❌ Không có khách hàng")
    st.stop()

kh = st.selectbox("👤 Chọn khách hàng", kh_list)
df_kh = df[df["Tên khách hàng"] == kh]

# ================= KPI =================
c1, c2, c3 = st.columns(3)
c1.metric("💰 Doanh thu", f"{df_kh['Thành tiền bán'].sum():,.0f}")
c2.metric("📦 Số đơn", len(df_kh))
c3.metric("🚚 Số xe", df_kh["Biển số xe"].nunique())

# ================= BIỂU ĐỒ =================
st.subheader("📈 Doanh thu theo tháng")
dt = df_kh.groupby("Tháng")["Thành tiền bán"].sum().reset_index()
st.line_chart(dt.set_index("Tháng"))

st.subheader("🔁 Tần suất mua")
freq = df_kh.groupby("Tháng").size().reset_index(name="Số đơn")
st.bar_chart(freq.set_index("Tháng"))

st.subheader("📍 Giao hàng")
tinh = df_kh.groupby("Tỉnh")["Thành tiền bán"].sum().reset_index()
st.bar_chart(tinh.set_index("Tỉnh"))