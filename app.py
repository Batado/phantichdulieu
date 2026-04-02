import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="Phân tích Kinh Doanh KH", layout="wide")
st.title("📊 Dashboard Phân tích khách hàng")

DATA_PATH = "data"
os.makedirs(DATA_PATH, exist_ok=True)

file = st.sidebar.file_uploader("📂 Upload Excel", type=["xlsx"])
if file is not None:
    path = os.path.join(DATA_PATH, file.name)
    with open(path, "wb") as f:
        f.write(file.getbuffer())
    st.success(f"✅ Upload thành công: {file.name}")
    st.cache_data.clear()
    st.rerun()


@st.cache_data
def read_excel_smart(path):
    """
    Đọc file Excel, tự động tìm dòng header.
    File báo cáo OM_RPT_055 có header ở row 13 (0-indexed).
    Hỗ trợ cả file có header ở các vị trí khác.
    """
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return pd.DataFrame()

    for sheet in xl.sheet_names:
        try:
            df_raw = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
        except Exception:
            continue

        # Tìm dòng header: so sánh không phân biệt hoa thường, bỏ dấu accent
        header_row = None
        for i in range(min(30, df_raw.shape[0])):
            row_vals = df_raw.iloc[i].astype(str).tolist()
            row_text = " ".join(row_vals).lower()
            # Kiểm tra các từ khoá không dấu và có dấu
            keywords = ["khach hang", "khách hàng", "ten kh", "tên kh",
                        "so chung tu", "số chứng từ", "ma khach hang", "mã khách hàng"]
            if any(kw in row_text for kw in keywords):
                header_row = i
                break

        if header_row is None:
            continue

        try:
            df = pd.read_excel(path, sheet_name=sheet, header=header_row, engine="openpyxl")
            return df
        except Exception:
            continue

    return pd.DataFrame()


@st.cache_data
def load_data():
    files = [f for f in os.listdir(DATA_PATH) if f.endswith(".xlsx")]
    all_df = []

    for f in files:
        file_path = os.path.join(DATA_PATH, f)
        df = read_excel_smart(file_path)
        if not df.empty:
            df["Nguồn file"] = f
            all_df.append(df)

    if not all_df:
        return pd.DataFrame()

    df = pd.concat(all_df, ignore_index=True)
    df.columns = df.columns.astype(str).str.strip()

    # Chuẩn hoá tên cột — thêm "Số xe" → "Biển số xe" (tên thực tế trong file)
    df.rename(columns={
        "Tên KH": "Tên khách hàng",
        "Khách hàng": "Tên khách hàng",
        "Tên khách": "Tên khách hàng",
        "Ngày CT": "Ngày chứng từ",
        "Địa chỉ giao hàng": "Nơi giao hàng",
        "Tên hàng": "Tên mã hàng",        # file thực tế dùng "Tên hàng"
        "Số xe": "Biển số xe",             # ← FIX: file dùng "Số xe", không phải "Biển số xe"
    }, inplace=True)

    # Kiểm tra cột bắt buộc
    required = ["Tên khách hàng", "Ngày chứng từ"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"❌ Thiếu cột bắt buộc: {missing}. Vui lòng kiểm tra file.")
        return pd.DataFrame()

    # Chuyển kiểu dữ liệu
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")

    # FIX: dùng `in df.columns` thay vì df.get() — pandas DataFrame không hỗ trợ .get()
    if "Thành tiền bán" in df.columns:
        df["Thành tiền bán"] = pd.to_numeric(df["Thành tiền bán"], errors="coerce").fillna(0)
    else:
        df["Thành tiền bán"] = 0

    if "Thành tiền vốn" in df.columns:
        df["Thành tiền vốn"] = pd.to_numeric(df["Thành tiền vốn"], errors="coerce").fillna(0)
    else:
        df["Thành tiền vốn"] = 0

    df["Tháng"] = df["Ngày chứng từ"].dt.strftime("%Y-%m")

    # Phân tỉnh từ Nơi giao hàng
    if "Nơi giao hàng" not in df.columns:
        df["Nơi giao hàng"] = ""
    df["Nơi giao hàng"] = df["Nơi giao hàng"].astype(str)
    df["Tỉnh"] = "Khác"
    province_map = {
        "TP HCM":       ["hồ chí minh", "hcm", "tp hcm", "tphcm"],
        "Hà Nội":       ["hà nội", "hanoi"],
        "Bình Dương":   ["bình dương", "binh duong"],
        "Đồng Nai":     ["đồng nai", "dong nai"],
        "Long An":      ["long an"],
        "Bà Rịa - VT":  ["bà rịa", "vũng tàu", "ba ria"],
    }
    col_lower = df["Nơi giao hàng"].str.lower()
    for tinh, keywords in province_map.items():
        mask = col_lower.str.contains("|".join(keywords), na=False)
        df.loc[mask, "Tỉnh"] = tinh

    # Xử lý biển số xe — FIX: bỏ filter độ dài cứng nhắc gây mất data
    if "Biển số xe" in df.columns:
        df["Biển số xe"] = (
            df["Biển số xe"].astype(str)
            .str.replace(" ", "", regex=False)
            .str.replace("-", "", regex=False)
            .str.upper()
            .str.strip()
        )
        # Chỉ loại bỏ giá trị rõ ràng là rác, giữ lại biển số hợp lệ (6–10 ký tự)
        invalid = {"NAN", "NONE", "", "GK", "N/A"}
        df = df[~df["Biển số xe"].isin(invalid)]
        df = df[df["Biển số xe"].str.len().between(6, 10)]
    else:
        df["Biển số xe"] = "N/A"

    return df


# ─── Load data ───────────────────────────────────────────────
df = load_data()

if df.empty:
    st.warning("⚠️ Chưa có dữ liệu hợp lệ. Vui lòng upload file Excel.")
    st.stop()

# ─── Sidebar: filter ─────────────────────────────────────────
kh_list = sorted(df["Tên khách hàng"].dropna().unique())
kh = st.sidebar.selectbox("👤 Chọn khách hàng", kh_list)
df_kh = df[df["Tên khách hàng"] == kh]

# ─── KPI ─────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
tong_dt = df_kh["Thành tiền bán"].sum()
tong_von = df_kh["Thành tiền vốn"].sum()
loi_nhuan = tong_dt - tong_von
so_don = df_kh["Số chứng từ"].nunique() if "Số chứng từ" in df_kh.columns else len(df_kh)

c1.metric("💰 Doanh thu (VNĐ)", f"{tong_dt:,.0f}")
c2.metric("📦 Số chứng từ", f"{so_don:,}")
c3.metric("💹 Lợi nhuận (VNĐ)", f"{loi_nhuan:,.0f}")
c4.metric("📁 Số file", df_kh["Nguồn file"].nunique())

# ─── Biểu đồ doanh thu theo tháng ────────────────────────────
st.subheader("📈 Doanh thu theo tháng")
df_thang = df_kh.groupby("Tháng")["Thành tiền bán"].sum().reset_index()
fig1 = px.line(df_thang, x="Tháng", y="Thành tiền bán", markers=True,
               labels={"Tháng": "Tháng", "Thành tiền bán": "Doanh thu (VNĐ)"})
st.plotly_chart(fig1, use_container_width=True)

# ─── Tần suất mua hàng ───────────────────────────────────────
st.subheader("🔁 Tần suất đơn hàng theo tháng")
freq_col = "Số chứng từ" if "Số chứng từ" in df_kh.columns else "Tháng"
if "Số chứng từ" in df_kh.columns:
    df_freq = df_kh.groupby("Tháng")["Số chứng từ"].nunique().reset_index(name="Số đơn")
else:
    df_freq = df_kh.groupby("Tháng").size().reset_index(name="Số đơn")
fig2 = px.bar(df_freq, x="Tháng", y="Số đơn",
              labels={"Tháng": "Tháng", "Số đơn": "Số đơn hàng"})
st.plotly_chart(fig2, use_container_width=True)

# ─── Top sản phẩm ────────────────────────────────────────────
if "Tên mã hàng" in df_kh.columns:
    st.subheader("📦 Top 10 sản phẩm mua nhiều nhất")
    df_prod = (df_kh.groupby("Tên mã hàng")["Thành tiền bán"]
               .sum().reset_index(name="Doanh thu")
               .sort_values("Doanh thu", ascending=False).head(10))
    fig3 = px.bar(df_prod, x="Tên mã hàng", y="Doanh thu",
                  labels={"Tên mã hàng": "Sản phẩm", "Doanh thu": "Doanh thu (VNĐ)"})
    fig3.update_layout(xaxis_tickangle=-30)
    st.plotly_chart(fig3, use_container_width=True)
else:
    df_prod = pd.DataFrame()

# ─── Doanh thu theo tỉnh ─────────────────────────────────────
st.subheader("📍 Doanh thu theo tỉnh giao hàng")
df_province = df_kh.groupby("Tỉnh")["Thành tiền bán"].sum().reset_index()
fig4 = px.pie(df_province, names="Tỉnh", values="Thành tiền bán",
              title="Tỷ trọng doanh thu theo tỉnh")
st.plotly_chart(fig4, use_container_width=True)

# ─── Bản đồ ──────────────────────────────────────────────────
map_data = pd.DataFrame({
    "Tỉnh":      ["TP HCM", "Hà Nội", "Bình Dương", "Đồng Nai", "Long An", "Bà Rịa - VT"],
    "latitude":  [10.82,    21.02,     11.06,         10.95,      10.54,      10.58],
    "longitude": [106.63,   105.85,    106.67,        106.83,     106.41,     107.17],
})
map_merge = pd.merge(map_data, df_province, on="Tỉnh", how="left").fillna(0)
st.map(map_merge[map_merge["Thành tiền bán"] > 0])

# ─── Khu vực bán hàng ────────────────────────────────────────
if "Khu vực" in df_kh.columns:
    st.subheader("🗺️ Doanh thu theo Khu vực")
    df_kv = df_kh.groupby("Khu vực")["Thành tiền bán"].sum().reset_index()
    fig5 = px.bar(df_kv, x="Khu vực", y="Thành tiền bán",
                  labels={"Khu vực": "Khu vực", "Thành tiền bán": "Doanh thu (VNĐ)"})
    st.plotly_chart(fig5, use_container_width=True)

# ─── Insight tự động ─────────────────────────────────────────
st.subheader("🧠 Insight tự động")
if not df_thang.empty:
    top_thang = df_thang.loc[df_thang["Thành tiền bán"].idxmax()]
    st.success(f"🔥 Tháng doanh thu cao nhất: {top_thang['Tháng']} "
               f"({int(top_thang['Thành tiền bán']):,} VNĐ)")
if not df_prod.empty:
    top_prod = df_prod.iloc[0]
    st.success(f"📦 Sản phẩm doanh thu cao nhất: {top_prod['Tên mã hàng']} "
               f"({int(top_prod['Doanh thu']):,} VNĐ)")
if not df_freq.empty:
    top_freq = df_freq.loc[df_freq["Số đơn"].idxmax()]
    st.info(f"🔁 Tháng nhiều đơn nhất: {top_freq['Tháng']} ({top_freq['Số đơn']} đơn)")
if loi_nhuan > 0:
    margin = loi_nhuan / tong_dt * 100 if tong_dt else 0
    st.info(f"💹 Biên lợi nhuận: {margin:.1f}%")