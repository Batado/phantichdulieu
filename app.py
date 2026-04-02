import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Phân tích Kinh Doanh KH", layout="wide")
st.title("📊 Dashboard Phân tích khách hàng")

# ─── Upload trực tiếp vào memory, KHÔNG lưu disk ─────────────
uploaded_files = st.sidebar.file_uploader(
    "📂 Upload file Excel (có thể chọn nhiều file)",
    type=["xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("👈 Vui lòng upload file Excel báo cáo bán hàng ở sidebar bên trái.")
    st.stop()


# ─── Đọc và parse file ───────────────────────────────────────
def find_header_row(path_or_buffer):
    """Tìm dòng header tự động trong file báo cáo."""
    df_raw = pd.read_excel(path_or_buffer, header=None, engine="openpyxl", nrows=35)
    keywords = ["khách hàng", "khach hang", "tên kh", "số chứng từ", "so chung tu", "mã khách hàng"]
    for i in range(df_raw.shape[0]):
        # FIX: fillna("") trước khi join để tránh lỗi "expected str, got float" khi có ô NaN
        row_vals = [str(v) if not (isinstance(v, float) and pd.isna(v)) else "" 
                    for v in df_raw.iloc[i].tolist()]
        row_text = " ".join(row_vals).lower()
        if any(kw in row_text for kw in keywords):
            return i
    return 0  # fallback: dùng row đầu tiên


def parse_file(uploaded_file):
    """Đọc một file Excel, trả về DataFrame đã chuẩn hoá."""
    try:
        # Phải đọc bytes trước vì Streamlit buffer chỉ đọc được 1 lần
        file_bytes = uploaded_file.read()
        import io
        buf = io.BytesIO(file_bytes)

        header_row = find_header_row(io.BytesIO(file_bytes))
        buf.seek(0)
        df = pd.read_excel(buf, header=header_row, engine="openpyxl")
        df.columns = df.columns.astype(str).str.strip()

        # Chuẩn hoá tên cột
        df.rename(columns={
            "Tên KH": "Tên khách hàng",
            "Khách hàng": "Tên khách hàng",
            "Tên khách": "Tên khách hàng",
            "Ngày CT": "Ngày chứng từ",
            "Địa chỉ giao hàng": "Nơi giao hàng",
            "Tên hàng": "Tên mã hàng",
            "Số xe": "Biển số xe",
            "Biển số xe": "Biển số xe",
        }, inplace=True)

        df["Nguồn file"] = uploaded_file.name
        return df

    except Exception as e:
        st.warning(f"⚠️ Không đọc được file `{uploaded_file.name}`: {e}")
        return pd.DataFrame()


def process_data(df):
    """Chuẩn hoá kiểu dữ liệu và thêm cột phân tích."""
    required = ["Tên khách hàng", "Ngày chứng từ"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"❌ Thiếu cột: {missing}. Kiểm tra lại file.")
        return pd.DataFrame()

    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Tháng"] = df["Ngày chứng từ"].dt.strftime("%Y-%m")

    for col in ["Thành tiền bán", "Thành tiền vốn"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0

    # Phân tỉnh
    if "Nơi giao hàng" not in df.columns:
        df["Nơi giao hàng"] = ""
    df["Nơi giao hàng"] = df["Nơi giao hàng"].astype(str)
    df["Tỉnh"] = "Khác"
    province_map = {
        "TP HCM":      ["hồ chí minh", "hcm", "tphcm"],
        "Hà Nội":      ["hà nội", "hanoi"],
        "Bình Dương":  ["bình dương"],
        "Đồng Nai":    ["đồng nai"],
        "Long An":     ["long an"],
        "Bà Rịa - VT": ["bà rịa", "vũng tàu"],
    }
    col_lower = df["Nơi giao hàng"].str.lower()
    for tinh, kws in province_map.items():
        df.loc[col_lower.str.contains("|".join(kws), na=False), "Tỉnh"] = tinh

    return df


# ─── Gộp tất cả file ─────────────────────────────────────────
all_frames = []
for uf in uploaded_files:
    parsed = parse_file(uf)
    if not parsed.empty:
        all_frames.append(parsed)

if not all_frames:
    st.error("❌ Không đọc được dữ liệu từ file nào. Kiểm tra định dạng file.")
    st.stop()

df_all = pd.concat(all_frames, ignore_index=True)
df_all = process_data(df_all)

if df_all.empty:
    st.stop()

# ─── Sidebar filters ─────────────────────────────────────────
kh_list = sorted(df_all["Tên khách hàng"].dropna().unique())
kh = st.sidebar.selectbox("👤 Chọn khách hàng", kh_list)
df_kh = df_all[df_all["Tên khách hàng"] == kh].copy()

# Lọc tháng (tuỳ chọn)
thang_list = sorted(df_kh["Tháng"].dropna().unique())
if len(thang_list) > 1:
    thang_chon = st.sidebar.multiselect("📅 Lọc theo tháng", thang_list, default=thang_list)
    df_kh = df_kh[df_kh["Tháng"].isin(thang_chon)]

# ─── KPI ─────────────────────────────────────────────────────
tong_dt  = df_kh["Thành tiền bán"].sum()
tong_von = df_kh["Thành tiền vốn"].sum()
loi_nhuan = tong_dt - tong_von
so_chung_tu = df_kh["Số chứng từ"].nunique() if "Số chứng từ" in df_kh.columns else len(df_kh)
margin_pct  = (loi_nhuan / tong_dt * 100) if tong_dt else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Doanh thu",    f"{tong_dt:,.0f} đ")
col2.metric("📦 Số chứng từ", f"{so_chung_tu:,}")
col3.metric("💹 Lợi nhuận",   f"{loi_nhuan:,.0f} đ")
col4.metric("📊 Biên LN",     f"{margin_pct:.1f}%")

st.divider()

# ─── Biểu đồ doanh thu theo tháng ────────────────────────────
st.subheader("📈 Doanh thu theo tháng")
df_thang = df_kh.groupby("Tháng")["Thành tiền bán"].sum().reset_index()
if not df_thang.empty:
    fig1 = px.line(df_thang, x="Tháng", y="Thành tiền bán", markers=True,
                   labels={"Thành tiền bán": "Doanh thu (VNĐ)"})
    st.plotly_chart(fig1, use_container_width=True)

# ─── Tần suất đơn hàng ───────────────────────────────────────
st.subheader("🔁 Số chứng từ theo tháng")
if "Số chứng từ" in df_kh.columns:
    df_freq = df_kh.groupby("Tháng")["Số chứng từ"].nunique().reset_index(name="Số đơn")
else:
    df_freq = df_kh.groupby("Tháng").size().reset_index(name="Số đơn")
if not df_freq.empty:
    fig2 = px.bar(df_freq, x="Tháng", y="Số đơn",
                  labels={"Số đơn": "Số chứng từ"}, text_auto=True)
    st.plotly_chart(fig2, use_container_width=True)

# ─── Top sản phẩm ────────────────────────────────────────────
if "Tên mã hàng" in df_kh.columns:
    st.subheader("📦 Top 10 sản phẩm theo doanh thu")
    df_prod = (df_kh.groupby("Tên mã hàng")["Thành tiền bán"]
               .sum().reset_index()
               .sort_values("Thành tiền bán", ascending=False).head(10))
    if not df_prod.empty:
        fig3 = px.bar(df_prod, x="Tên mã hàng", y="Thành tiền bán",
                      labels={"Tên mã hàng": "Sản phẩm", "Thành tiền bán": "Doanh thu (VNĐ)"},
                      text_auto=True)
        fig3.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(fig3, use_container_width=True)
else:
    df_prod = pd.DataFrame()

# ─── Khu vực ─────────────────────────────────────────────────
col_left, col_right = st.columns(2)

with col_left:
    st.subheader("📍 Doanh thu theo tỉnh")
    df_prov = df_kh.groupby("Tỉnh")["Thành tiền bán"].sum().reset_index()
    if not df_prov.empty:
        fig4 = px.pie(df_prov, names="Tỉnh", values="Thành tiền bán")
        st.plotly_chart(fig4, use_container_width=True)

with col_right:
    if "Khu vực" in df_kh.columns:
        st.subheader("🗺️ Doanh thu theo khu vực")
        df_kv = df_kh.groupby("Khu vực")["Thành tiền bán"].sum().reset_index()
        if not df_kv.empty:
            fig5 = px.bar(df_kv, x="Khu vực", y="Thành tiền bán",
                          labels={"Thành tiền bán": "Doanh thu (VNĐ)"}, text_auto=True)
            st.plotly_chart(fig5, use_container_width=True)

# ─── Bảng chi tiết ───────────────────────────────────────────
with st.expander("📋 Xem dữ liệu chi tiết"):
    show_cols = [c for c in ["Số chứng từ", "Ngày chứng từ", "Tên mã hàng",
                              "Số lượng", "Thành tiền bán", "Thành tiền vốn",
                              "Tỉnh", "Nguồn file"] if c in df_kh.columns]
    st.dataframe(df_kh[show_cols], use_container_width=True)

# ─── Insight ─────────────────────────────────────────────────
st.subheader("🧠 Insight tự động")
if not df_thang.empty:
    top = df_thang.loc[df_thang["Thành tiền bán"].idxmax()]
    st.success(f"🔥 Tháng doanh thu cao nhất: **{top['Tháng']}** — {int(top['Thành tiền bán']):,} VNĐ")
if not df_prod.empty:
    p = df_prod.iloc[0]
    st.success(f"📦 Sản phẩm bán chạy nhất: **{p['Tên mã hàng']}** — {int(p['Thành tiền bán']):,} VNĐ")
if not df_freq.empty:
    f = df_freq.loc[df_freq["Số đơn"].idxmax()]
    st.info(f"🔁 Tháng nhiều đơn nhất: **{f['Tháng']}** — {f['Số đơn']} đơn")
st.info(f"💹 Biên lợi nhuận tổng: **{margin_pct:.1f}%**")
