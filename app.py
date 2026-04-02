import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Cấu hình trang
st.set_page_config(page_title="Phân tích Kinh Doanh KH", layout="wide")
st.title("📊 Dashboard Phân tích khách hàng")

DATA_PATH = "data"
os.makedirs(DATA_PATH, exist_ok=True)

# Upload file và lưu vào thư mục DATA_PATH
file = st.sidebar.file_uploader("📂 Upload Excel", type=["xlsx"])
if file is not None:
    path = os.path.join(DATA_PATH, file.name)
    if not os.path.exists(path):
        with open(path, "wb") as f:
            f.write(file.getbuffer())
    st.success(f"✅ Upload thành công: {file.name}")
    # Clear cache và rerun để load file mới
    st.cache_data.clear()
    st.rerun()

# Hàm tự động đọc file Excel (tìm sheet và header phù hợp)
@st.cache_data
def read_excel_smart(path):
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return pd.DataFrame()
    for sheet in xl.sheet_names:
        try:
            df_raw = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
        except Exception:
            continue
        # Tìm dòng header có chứa từ "khách hàng" hoặc "số chứng từ"
        for i in range(0, min(30, df_raw.shape[0])):
            row_text = " ".join(df_raw.iloc[i].astype(str).tolist()).lower()
            if "khách hàng" in row_text or "tên kh" in row_text or "số chứng từ" in row_text:
                try:
                    df = pd.read_excel(path, sheet_name=sheet, header=i, engine="openpyxl")
                except Exception:
                    continue
                return df
    return pd.DataFrame()

# Hàm load toàn bộ data từ các file đã lưu
@st.cache_data
def load_data():
    files = os.listdir(DATA_PATH)
    all_df = []
    for f in files:
        file_path = os.path.join(DATA_PATH, f)
        df = read_excel_smart(file_path)
        if not df.empty:
            df["Nguồn file"] = f  # lưu tên file để tracking (nếu cần)
            all_df.append(df)
    if not all_df:
        return pd.DataFrame()
    df = pd.concat(all_df, ignore_index=True)
    # Làm sạch tên cột
    df.columns = df.columns.astype(str).str.strip()
    df.rename(columns={
        # Chuẩn hóa tên cột thường gặp
        "Tên KH": "Tên khách hàng",
        "Khách hàng": "Tên khách hàng",
        "Tên khách": "Tên khách hàng",
        "Ngày CT": "Ngày chứng từ",
        "Nơi giao hàng": "Nơi giao hàng",
        "Địa chỉ giao hàng": "Nơi giao hàng",
        "Tên hàng": "Tên mã hàng",
        "Biển số xe": "Biển số xe"
    }, inplace=True)
    # Nếu thiếu cột bắt buộc, trả DF rỗng
    if "Tên khách hàng" not in df.columns or "Ngày chứng từ" not in df.columns:
        st.error("❌ Sai định dạng file: thiếu cột bắt buộc (Tên khách hoặc Ngày chứng từ).")
        return pd.DataFrame()
    # Chuyển kiểu cột
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Thành tiền bán"] = pd.to_numeric(df.get("Thành tiền bán", 0), errors="coerce")
    df["Thành tiền vốn"] = pd.to_numeric(df.get("Thành tiền vốn", 0), errors="coerce")
    # Thêm cột Tháng (YYYY-MM)
    df["Tháng"] = df["Ngày chứng từ"].dt.strftime("%Y-%m")
    # Gom nhóm tỉnh từ Nơi giao hàng
    df["Nơi giao hàng"] = df.get("Nơi giao hàng", "").astype(str)
    df["Tỉnh"] = "Khác"
    df.loc[df["Nơi giao hàng"].str.contains("hcm", case=False, na=False), "Tỉnh"] = "TP HCM"
    df.loc[df["Nơi giao hàng"].str.contains("hà nội", case=False, na=False), "Tỉnh"] = "Hà Nội"
    df.loc[df["Nơi giao hàng"].str.contains("bình dương", case=False, na=False), "Tỉnh"] = "Bình Dương"
    df.loc[df["Nơi giao hàng"].str.contains("đồng nai", case=False, na=False), "Tỉnh"] = "Đồng Nai"
    # Số xe (biển số) lọc ký tự và độ dài hợp lệ
    df["Biển số xe"] = df.get("Biển số xe", "").astype(str).str.replace(" ", "").str.replace("-", "").str.upper()
    df = df[df["Biển số xe"].str.len().between(7, 9)]
    df = df[~df["Biển số xe"].isin(["GK", "", "NONE", "nan"])]
    return df

df = load_data()
if df.empty:
    st.warning("⚠️ Chưa có dữ liệu hợp lệ. Vui lòng upload file Excel.")
    st.stop()

# Lọc theo khách hàng
kh_list = df["Tên khách hàng"].dropna().unique()
kh = st.sidebar.selectbox("👤 Chọn khách hàng", kh_list)
df_kh = df[df["Tên khách hàng"] == kh]

# KPI tổng quan
c1, c2, c3 = st.columns(3)
tong_doanhthu = df_kh["Thành tiền bán"].sum()
c1.metric("💰 Doanh thu (VNĐ)", f"{tong_doanhthu:,.0f}")
c2.metric("📦 Số đơn", len(df_kh))
c3.metric("📁 Số file phân tích", df_kh["Nguồn file"].nunique())

# Biểu đồ Doanh thu theo tháng
st.subheader("📈 Doanh thu theo tháng")
df_thang = df_kh.groupby("Tháng")["Thành tiền bán"].sum().reset_index()
fig1 = px.line(df_thang, x="Tháng", y="Thành tiền bán", markers=True,
               labels={"Tháng":"Tháng", "Thành tiền bán":"Doanh thu"})
fig1.update_traces(textposition="top center")
st.plotly_chart(fig1, use_container_width=True)

# Biểu đồ Tần suất mua hàng theo tháng
st.subheader("🔁 Tần suất mua hàng")
df_freq = df_kh.groupby("Tháng").size().reset_index(name="Số đơn")
fig2 = px.bar(df_freq, x="Tháng", y="Số đơn",
              labels={"Tháng":"Tháng", "Số đơn":"Số đơn"})
fig2.update_traces(textposition="outside")
st.plotly_chart(fig2, use_container_width=True)

# Biểu đồ Top 10 sản phẩm mua nhiều nhất
st.subheader("📦 Top sản phẩm (số lần mua)")
df_prod = df_kh.groupby("Tên mã hàng").size().reset_index(name="Số lần")
df_prod = df_prod.sort_values(by="Số lần", ascending=False).head(10)
fig3 = px.bar(df_prod, x="Tên mã hàng", y="Số lần",
              labels={"Tên mã hàng":"Mã hàng", "Số lần":"Số lần mua"})
fig3.update_traces(textposition="outside")
fig3.update_layout(xaxis_tickangle=-30)
st.plotly_chart(fig3, use_container_width=True)

# Biểu đồ khu vực giao hàng
st.subheader("📍 Doanh thu theo tỉnh giao hàng")
df_province = df_kh.groupby("Tỉnh")["Thành tiền bán"].sum().reset_index()
fig4 = px.pie(df_province, names="Tỉnh", values="Thành tiền bán",
              title="Tỷ trọng doanh thu theo tỉnh")
st.plotly_chart(fig4, use_container_width=True)

# Bản đồ đơn giản (các tỉnh chính)
map_data = pd.DataFrame({
    "Tỉnh": ["TP HCM","Hà Nội","Bình Dương","Đồng Nai"],
    "latitude": [10.82, 21.02, 11.06, 10.95],
    "longitude": [106.63, 105.85, 106.67, 106.83]
})
map_merge = pd.merge(map_data, df_province, on="Tỉnh", how="left").fillna(0)
st.map(map_merge)

# Insight tự động
st.subheader("🧠 Insight tự động")
if not df_thang.empty:
    top_thang = df_thang.loc[df_thang["Thành tiền bán"].idxmax()]
    st.success(f"🔥 Khách mua nhiều nhất vào tháng: {top_thang['Tháng']} (Doanh thu {int(top_thang['Thành tiền bán']):,d} VNĐ)")
if not df_prod.empty:
    top_prod = df_prod.iloc[0]
    st.success(f"📦 Sản phẩm mua nhiều nhất: {top_prod['Tên mã hàng']} ({top_prod['Số lần']} lần)")
if not df_freq.empty:
    top_month_orders = df_freq.loc[df_freq["Số đơn"].idxmax()]
    st.info(f"🔁 Tháng có nhiều đơn nhất: {top_month_orders['Tháng']} ({top_month_orders['Số đơn']} đơn)")
