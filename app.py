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

    # lưu lịch sử (không ghi đè)
    if not os.path.exists(path):
        with open(path, "wb") as f:
            f.write(file.getbuffer())

    st.success("✅ Upload thành công (đã lưu lịch sử)")
    st.cache_data.clear()
    st.rerun()


# ================= ĐỌC FILE THÔNG MINH =================
@st.cache_data
def read_excel_smart(path):
    xl = pd.ExcelFile(path)

    for sheet in xl.sheet_names:
        try:
            df_raw = pd.read_excel(path, sheet_name=sheet, header=None)

            for i in range(0, 30):
                row = " ".join(df_raw.iloc[i].astype(str)).lower()

                if "kh" in row or "khách" in row:
                    df = pd.read_excel(path, sheet_name=sheet, header=i)

                    if df.shape[1] > 5:
                        return df

        except:
            continue

    return pd.DataFrame()


# ================= LOAD DATA =================
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
                df["Nguồn file"] = f  # lưu lịch sử file
                all_df.append(df)

        except:
            continue

    if not all_df:
        return pd.DataFrame()

    df = pd.concat(all_df, ignore_index=True)

    # ===== CLEAN =====
    df.columns = df.columns.astype(str).str.strip()

    # ===== RENAME =====
    df.rename(columns={
        "Tên KH": "Tên khách hàng",
        "Khách hàng": "Tên khách hàng",
        "Tên khách": "Tên khách hàng",
        "Ngày CT": "Ngày chứng từ",
        "Ngày": "Ngày chứng từ",
        "Địa chỉ giao hàng": "Nơi giao hàng",
        "Tên mã hàng": "Tên mã hàng"
    }, inplace=True)

    if "Tên khách hàng" not in df.columns:
        return pd.DataFrame()

    # ===== XỬ LÝ =====
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Tháng"] = df["Ngày chứng từ"].dt.strftime("%Y-%m")

    df["Thành tiền bán"] = pd.to_numeric(df.get("Thành tiền bán", 0), errors="coerce")
    df["Tên mã hàng"] = df.get("Tên mã hàng", "Không rõ")

    return df


df = load_data()

if df.empty:
    st.warning("⚠️ Chưa có dữ liệu")
    st.stop()

# ================= FILTER =================
kh_list = df["Tên khách hàng"].dropna().unique()

kh = st.selectbox("👤 Chọn khách hàng", kh_list)
df_kh = df[df["Tên khách hàng"] == kh]

# ================= KPI =================
c1, c2, c3 = st.columns(3)

c1.metric("💰 Doanh thu", f"{df_kh['Thành tiền bán'].sum():,.0f}")
c2.metric("📦 Số đơn", len(df_kh))
c3.metric("📁 Số file", df_kh["Nguồn file"].nunique())

# ================= DOANH THU THEO THÁNG =================
st.subheader("📈 Doanh thu theo tháng")

dt = df_kh.groupby("Tháng")["Thành tiền bán"].sum().reset_index()

fig1 = px.line(dt, x="Tháng", y="Thành tiền bán", markers=True)
st.plotly_chart(fig1, use_container_width=True)

# ================= TẦN SUẤT =================
st.subheader("🔁 Tần suất mua hàng")

freq = df_kh.groupby("Tháng").size().reset_index(name="Số đơn")

fig2 = px.bar(freq, x="Tháng", y="Số đơn")
st.plotly_chart(fig2, use_container_width=True)

# ================= MÃ HÀNG =================
st.subheader("📦 Mã hàng mua nhiều")

sp = df_kh.groupby("Tên mã hàng").size().reset_index(name="Số lần")

sp = sp.sort_values(by="Số lần", ascending=False).head(10)

fig3 = px.bar(sp, x="Tên mã hàng", y="Số lần")
st.plotly_chart(fig3, use_container_width=True)

# ================= AI INSIGHT =================
st.subheader("🧠 Insight tự động")

if not dt.empty:
    top_month = dt.loc[dt["Thành tiền bán"].idxmax()]
    st.success(f"🔥 Tháng mua nhiều nhất: {top_month['Tháng']}")

if not sp.empty:
    top_sp = sp.iloc[0]
    st.success(f"📦 Mã hàng mua nhiều nhất: {top_sp['Tên mã hàng']} ({top_sp['Số lần']} lần)")

if not freq.empty:
    top_freq = freq.loc[freq["Số đơn"].idxmax()]
    st.info(f"🔁 Tháng mua nhiều đơn nhất: {top_freq['Tháng']} ({top_freq['Số đơn']} đơn)")