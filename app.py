import io
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(page_title="Phân tích KH – Hoa Sen", layout="wide", page_icon="📊")

st.markdown("""
<style>
.risk-high   {background:#4a1010;border-left:4px solid #e74c3c;padding:10px 14px;border-radius:6px;margin:5px 0;color:#fff;}
.risk-medium {background:#3d2e10;border-left:4px solid #f39c12;padding:10px 14px;border-radius:6px;margin:5px 0;color:#fff;}
.risk-low    {background:#0f3020;border-left:4px solid #26c281;padding:10px 14px;border-radius:6px;margin:5px 0;color:#fff;}
.risk-info   {background:#0d1f35;border-left:4px solid #4e79d4;padding:10px 14px;border-radius:6px;margin:5px 0;color:#fff;}
.section-title{font-size:16px;font-weight:700;color:#e0e0e0;margin:18px 0 8px 0;
               padding-bottom:5px;border-bottom:1px solid #2e3350;}
.kpi-card    {background:#1a2035;border-radius:8px;padding:14px 16px;text-align:center;}
.kpi-val     {font-size:22px;font-weight:800;color:#fff;}
.kpi-lab     {font-size:11px;color:#9aa0b0;margin-top:4px;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
COL_ALIASES = {
    "Ten khach hang":     ["Tên KH","Khách hàng","Tên khách","customer_name"],
    "Ma khach hang":      ["Mã KH","customer_code"],
    "Ngay chung tu":      ["Ngày CT","Ngày","date"],
    "So chung tu":        ["Số CT","voucher_no"],
    "Ma nhom KH":         ["Nhóm KH","sales_group"],
    "Ten nhom KH":        ["Tên nhóm","sales_group_name"],
    "Ten hang":           ["Tên mã hàng","Tên SP","product_name"],
    "Ma hang":            ["Mã SP","product_code"],
    "Ma nhom hang":       ["Nhóm hàng","product_group"],
    "Khoi luong":         ["KL","weight"],
    "So luong":           ["SL","qty"],
    "Thanh tien ban":     ["Doanh thu","revenue","amount"],
    "Thanh tien von":     ["Vốn","cost_total"],
    "Loi nhuan":          ["Profit","profit"],
    "Gia ban":            ["Đơn giá","unit_price"],
    "Gia von":            ["cost_price"],
    "Don gia quy doi":    ["converted_price"],
    "Don gia van chuyen": ["freight_unit"],
    "Noi giao hang":      ["Địa chỉ giao hàng","delivery_address"],
    "Freight Terms":      ["Điều kiện giao hàng","freight"],
    "Shipping method":    ["Phương tiện","shipping"],
    "Bien so xe":         ["Số xe","plate_no"],
    "Tai xe":             ["Tài xế","Driver"],
    "Ten DVVC":           ["Đơn vị vận chuyển","carrier"],
    "Ghi chu":            ["Note","notes"],
    "Khu vuc":            ["Region"],
    "Ma kho":             ["warehouse"],
    "Loai don hang":      ["order_type"],
}

# Map từ alias → tên chuẩn trong file (tiếng Việt)
STANDARD_NAMES = {
    "Ten khach hang":     "Tên khách hàng",
    "Ma khach hang":      "Mã khách hàng",
    "Ngay chung tu":      "Ngày chứng từ",
    "So chung tu":        "Số chứng từ",
    "Ma nhom KH":         "Mã nhóm KH",
    "Ten nhom KH":        "Tên nhóm KH",
    "Ten hang":           "Tên hàng",
    "Ma hang":            "Mã hàng",
    "Ma nhom hang":       "Mã nhóm hàng",
    "Khoi luong":         "Khối lượng",
    "So luong":           "Số lượng",
    "Thanh tien ban":     "Thành tiền bán",
    "Thanh tien von":     "Thành tiền vốn",
    "Loi nhuan":          "Lợi nhuận",
    "Gia ban":            "Giá bán",
    "Gia von":            "Giá vốn",
    "Don gia quy doi":    "Đơn giá quy đổi",
    "Don gia van chuyen": "Đơn giá vận chuyển",
    "Noi giao hang":      "Nơi giao hàng",
    "Freight Terms":      "Freight Terms",
    "Shipping method":    "Shipping method",
    "Bien so xe":         "Biển số xe",
    "Tai xe":             "Tài Xế",
    "Ten DVVC":           "Tên ĐVVC",
    "Ghi chu":            "Ghi chú",
    "Khu vuc":            "Khu vực",
    "Ma kho":             "Mã kho",
    "Loai don hang":      "Loại đơn hàng",
}

NHOM_SP = [
    ("Ống HDPE",        r"HDPE"),
    ("Ống PVC nước",    r"PVC.*(?:nước|nong dài|nong trơn|thoát)"),
    ("Ống PVC bơm cát", r"PVC.*(?:cát|bơm cát)"),
    ("Ống PPR",         r"PPR"),
    ("Lõi PVC",         r"(?:Lơi|Lõi|lori)"),
    ("Phụ kiện & Keo",  r"(?:Nối|Co |Tê |Van |Keo |Măng|Bít|Y PVC|Y PPR)"),
]

def normalize_columns(df):
    """Map alias → tên chuẩn tiếng Việt. Không đổi nếu cột đã đúng."""
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    rename = {}
    for key, aliases in COL_ALIASES.items():
        std = STANDARD_NAMES[key]
        if std in df.columns:
            continue
        for a in aliases:
            if a in df.columns:
                rename[a] = std
                break
            if a.strip().lower() in cols_lower:
                rename[cols_lower[a.strip().lower()]] = std
                break
    df.rename(columns=rename, inplace=True)
    return df

def find_header_row(fb):
    try:
        raw = pd.read_excel(io.BytesIO(fb), header=None, engine="openpyxl", nrows=40)
    except Exception:
        return 0
    kws = ["khách hàng","khach hang","tên kh","số chứng từ","mã khách","ngày chứng từ","tên hàng"]
    for i in range(raw.shape[0]):
        vals = ["" if (isinstance(v, float) and pd.isna(v)) else str(v) for v in raw.iloc[i].tolist()]
        txt  = " ".join(vals).lower()
        if sum(1 for k in kws if k in txt) >= 2:
            return i
    return 0

def get_col(df, col, default=0):
    if col in df.columns:
        return pd.to_numeric(df[col], errors="coerce").fillna(0)
    return pd.Series(default, index=df.index)

def fmt(v):
    try:
        return f"{float(v):,.0f}"
    except Exception:
        return str(v)

# ══════════════════════════════════════════════════════════════
#  UPLOAD
# ══════════════════════════════════════════════════════════════
st.sidebar.markdown("## 📂 Upload dữ liệu")
uploaded = st.sidebar.file_uploader("File Excel báo cáo bán hàng",
                                     type=["xlsx"], accept_multiple_files=True)
if not uploaded:
    st.markdown("## 👈 Upload file Excel để bắt đầu")
    st.info("Hỗ trợ báo cáo OM_RPT_055 – header tự động nhận diện.")
    st.stop()

@st.cache_data(show_spinner="Đang xử lý dữ liệu…")
def load_all(file_data):
    frames, errs = [], []
    for name, fb in file_data:
        try:
            hr = find_header_row(fb)
            df = pd.read_excel(io.BytesIO(fb), header=hr, engine="openpyxl")
            df.columns = [str(c).strip().replace("\n"," ").replace("\r","") for c in df.columns]
            df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
            df.dropna(how="all", inplace=True)
            normalize_columns(df)
            df["Nguon file"] = name
            frames.append(df)
        except Exception as e:
            errs.append(f"`{name}`: {e}")
    for e in errs:
        st.warning(f"⚠️ {e}")
    if not frames:
        return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True)

    # Kiểm tra cột bắt buộc
    for col in ["Ngày chứng từ", "Tên khách hàng"]:
        if col not in df.columns:
            st.error(f"❌ Không tìm thấy cột '{col}'. Kiểm tra file.")
            return pd.DataFrame()

    # Số hoá
    num_cols = ["Thành tiền bán","Thành tiền vốn","Lợi nhuận",
                "Khối lượng","Số lượng","Giá bán","Giá vốn",
                "Đơn giá vận chuyển","Đơn giá quy đổi"]
    for c in num_cols:
        df[c] = get_col(df, c)

    # Ngày
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], dayfirst=True, errors="coerce")
    df = df[df["Ngày chứng từ"].notna()].copy()
    df["Nam"]   = df["Ngày chứng từ"].dt.year.astype(str)
    df["Quy"]   = df["Ngày chứng từ"].dt.to_period("Q").astype(str)
    df["Thang"] = df["Ngày chứng từ"].dt.to_period("M").astype(str)
    df["Ngay_trong_thang"] = df["Ngày chứng từ"].dt.day
    df["Cuoi_thang"] = df["Ngay_trong_thang"] >= 28

    # Loại GD
    gc = df["Ghi chú"].astype(str).str.upper() if "Ghi chú" in df.columns else pd.Series("", index=df.index)
    loai = df["Loại đơn hàng"].astype(str).str.upper() if "Loại đơn hàng" in df.columns else pd.Series("", index=df.index)
    df["Loai_GD"] = "Xuất bán"
    df.loc[gc.str.contains(r"NHẬP TRẢ|TRẢ HÀNG", regex=True, na=False), "Loai_GD"] = "Trả hàng"
    df.loc[loai.str.contains(r"TRA HANG|HUY HD", regex=True, na=False), "Loai_GD"] = "Trả hàng"
    df.loc[gc.str.contains(r"BỔ SUNG|THAY THẾ", regex=True, na=False), "Loai_GD"] = "Xuất bổ sung"

    # Nhóm SP
    ten = df["Tên hàng"].astype(str) if "Tên hàng" in df.columns else pd.Series("", index=df.index)
    df["Nhom_SP"] = "Khác"
    for label, pat in NHOM_SP:
        df.loc[ten.str.contains(pat, case=False, regex=True, na=False), "Nhom_SP"] = label

    # Biên LN
    df["Bien_LN"] = np.where(df["Thành tiền bán"] != 0,
                              df["Lợi nhuận"] / df["Thành tiền bán"] * 100, 0)

    # Mã kho text
    if "Mã kho" in df.columns:
        df["Mã kho"] = df["Mã kho"].astype(str).str.replace(".0","",regex=False).str.strip()

    return df

file_data = [(u.name, u.read()) for u in uploaded]
df_all = load_all(file_data)
if df_all.empty:
    st.error("Không có dữ liệu hợp lệ.")
    st.stop()

# ══════════════════════════════════════════════════════════════
#  SIDEBAR FILTERS
# ══════════════════════════════════════════════════════════════
st.sidebar.markdown("---")
st.sidebar.markdown("## 🔍 Bộ lọc")

view_mode = st.sidebar.radio("📌 Chế độ xem", ["Theo Khách hàng", "Theo Phòng KD"])

if view_mode == "Theo Phòng KD":
    pkd_opts = sorted(df_all["Mã nhóm KH"].dropna().astype(str).unique()) if "Mã nhóm KH" in df_all.columns else []
    pkd_chon = st.sidebar.selectbox("🏢 Phòng KD", ["Tất cả"] + pkd_opts)
    if pkd_chon != "Tất cả" and "Mã nhóm KH" in df_all.columns:
        df_scope = df_all[df_all["Mã nhóm KH"].astype(str) == pkd_chon].copy()
    else:
        df_scope = df_all.copy()
    scope_label = f"Phòng KD: {pkd_chon}"
    single_kh = False
else:
    kh_list = sorted(df_all["Tên khách hàng"].dropna().astype(str).unique())
    kh = st.sidebar.selectbox("👤 Khách hàng", kh_list)
    df_scope = df_all[df_all["Tên khách hàng"].astype(str) == kh].copy()
    scope_label = kh
    single_kh = True

# Filter thời gian
st.sidebar.markdown("**📅 Cụm thời gian**")
time_col_map = {"Năm": "Nam", "Quý": "Quy", "Tháng": "Thang"}
time_level_label = st.sidebar.radio("Xem theo", ["Năm", "Quý", "Tháng"])
time_col = time_col_map[time_level_label]
time_vals = sorted(df_scope[time_col].dropna().unique())
time_chon = st.sidebar.multiselect(f"Chọn {time_level_label}", time_vals, default=time_vals)
df_scope  = df_scope[df_scope[time_col].isin(time_chon)].copy()
df_ban    = df_scope[df_scope["Loai_GD"] == "Xuất bán"].copy()

# ══════════════════════════════════════════════════════════════
#  HEADER KPI
# ══════════════════════════════════════════════════════════════
st.markdown(f"# 📊 {scope_label}")
ngay_min = df_scope["Ngày chứng từ"].min()
ngay_max = df_scope["Ngày chứng từ"].max()
st.markdown(f"*{ngay_min.strftime('%d/%m/%Y') if pd.notna(ngay_min) else '?'} → "
            f"{ngay_max.strftime('%d/%m/%Y') if pd.notna(ngay_max) else '?'} "
            f"| Cụm: **{time_level_label}** | {', '.join(time_chon[:4])}{'…' if len(time_chon)>4 else ''}*")

if df_ban.empty:
    st.warning("Không có dữ liệu xuất bán cho bộ lọc đã chọn.")
    st.stop()

tong_dt  = df_ban["Thành tiền bán"].sum()
tong_ln  = df_ban["Lợi nhuận"].sum()
tong_kl  = df_ban["Khối lượng"].sum() / 1000
bien_ln  = (tong_ln / tong_dt * 100) if tong_dt else 0
n_ct     = df_ban["Số chứng từ"].nunique() if "Số chứng từ" in df_ban.columns else len(df_ban)
n_kh     = df_ban["Tên khách hàng"].nunique()

c1,c2,c3,c4,c5 = st.columns(5)
for col, val, lab in [
    (c1, f"{fmt(tong_dt)} đ",       "💰 Doanh thu"),
    (c2, f"{fmt(tong_ln)} đ",       "💹 Lợi nhuận"),
    (c3, f"{bien_ln:.1f}%",          "📊 Biên LN"),
    (c4, f"{tong_kl:,.2f} tấn",     "⚖️ Khối lượng"),
    (c5, f"{n_ct} CT | {n_kh} KH", "📋 Chứng từ / KH"),
]:
    col.markdown(f'<div class="kpi-card"><div class="kpi-val">{val}</div>'
                 f'<div class="kpi-lab">{lab}</div></div>', unsafe_allow_html=True)
st.markdown("")

# ══════════════════════════════════════════════════════════════
#  TABS
# ══════════════════════════════════════════════════════════════
tab_pkd, tab_time, tab_sp, tab_ln_tab, tab_giao, tab_freq, tab_risk = st.tabs([
    "🏢 Phòng KD & KH",
    "📅 Cụm thời gian",
    "📦 Sản phẩm",
    "💹 Lợi nhuận & Chính sách",
    "🚚 Giao hàng",
    "🔁 Tần suất mua",
    "⚠️ Rủi ro & Bất thường",
])

# ══════════════════════════════════════════════════════════════
#  TAB 1 – PHÒNG KD & KH
# ══════════════════════════════════════════════════════════════
with tab_pkd:
    has_pkd = "Mã nhóm KH" in df_ban.columns and "Tên nhóm KH" in df_ban.columns

    st.markdown('<div class="section-title">🏢 Cơ cấu Doanh thu: Phòng KD → Khách hàng</div>', unsafe_allow_html=True)

    if has_pkd:
        df_pkd = (df_ban.groupby(["Mã nhóm KH","Tên nhóm KH"])
                  .agg(DT=("Thành tiền bán","sum"),
                       LN=("Lợi nhuận","sum"),
                       KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                       N_KH=("Tên khách hàng","nunique"),
                       N_CT=("Số chứng từ","nunique") if "Số chứng từ" in df_ban.columns else ("Thành tiền bán","count"))
                  .reset_index().sort_values("DT", ascending=False))
        df_pkd["Bien"] = (df_pkd["LN"]/df_pkd["DT"]*100).round(1)

        # Sunburst
        df_sun = (df_ban.groupby(["Tên nhóm KH","Tên khách hàng"])
                  .agg(DT=("Thành tiền bán","sum")).reset_index())
        fig_sun = px.sunburst(df_sun, path=["Tên nhóm KH","Tên khách hàng"],
                               values="DT", title="Sunburst: Phòng KD → KH",
                               color="DT", color_continuous_scale="Blues")
        fig_sun.update_layout(height=480)
        st.plotly_chart(fig_sun, use_container_width=True)

        # Bảng PKD
        st.markdown('<div class="section-title">📋 Tổng hợp Phòng KD</div>', unsafe_allow_html=True)
        df_pkd_s = df_pkd.copy()
        df_pkd_s["DT"] = df_pkd_s["DT"].map(fmt)
        df_pkd_s["LN"] = df_pkd_s["LN"].map(fmt)
        df_pkd_s.columns = ["Mã PKD","Tên Phòng KD","DT (VNĐ)","LN (VNĐ)","KL (tấn)","Số KH","Số CT","Biên (%)"]
        st.dataframe(df_pkd_s, use_container_width=True, hide_index=True)
    else:
        st.info("File không có cột Mã nhóm KH / Tên nhóm KH.")

    # Bảng KH hiệu quả
    st.markdown('<div class="section-title">👥 Khách hàng & Hiệu quả kinh doanh</div>', unsafe_allow_html=True)
    grp = ["Mã nhóm KH","Tên nhóm KH","Tên khách hàng"] if has_pkd else ["Tên khách hàng"]
    df_kh_tbl = (df_ban.groupby(grp)
                 .agg(DT=("Thành tiền bán","sum"),
                      LN=("Lợi nhuận","sum"),
                      KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                      N_CT=("Số chứng từ","nunique") if "Số chứng từ" in df_ban.columns else ("Thành tiền bán","count"),
                      N_SP=("Tên hàng","nunique") if "Tên hàng" in df_ban.columns else ("Thành tiền bán","count"))
                 .reset_index().sort_values("DT", ascending=False))
    df_kh_tbl["Bien"] = (df_kh_tbl["LN"]/df_kh_tbl["DT"]*100).round(1)
    df_kh_tbl["DT"]   = df_kh_tbl["DT"].map(fmt)
    df_kh_tbl["LN"]   = df_kh_tbl["LN"].map(fmt)
    st.dataframe(df_kh_tbl, use_container_width=True, hide_index=True)

    # Bar top KH (màu = biên LN)
    df_bar = (df_ban.groupby("Tên khách hàng")
              .agg(DT=("Thành tiền bán","sum"), LN=("Lợi nhuận","sum"))
              .reset_index().sort_values("DT",ascending=False).head(15))
    df_bar["Bien"] = (df_bar["LN"]/df_bar["DT"]*100).round(1)
    fig_bar = px.bar(df_bar, x="Tên khách hàng", y="DT", color="Bien",
                     color_continuous_scale="RdYlGn", text_auto=".3s",
                     title="Top 15 KH – Doanh thu (màu = Biên LN%)",
                     labels={"DT":"DT (VNĐ)","Bien":"Biên LN %"})
    fig_bar.update_layout(xaxis_tickangle=-30, height=420)
    st.plotly_chart(fig_bar, use_container_width=True)

    # Khu vực
    if "Khu vực" in df_ban.columns:
        st.markdown('<div class="section-title">🗺️ Doanh thu theo Khu vực</div>', unsafe_allow_html=True)
        df_kv = (df_ban.groupby("Khu vực")
                 .agg(DT=("Thành tiền bán","sum"), LN=("Lợi nhuận","sum"),
                      N_KH=("Tên khách hàng","nunique"))
                 .reset_index())
        df_kv["Bien"] = (df_kv["LN"]/df_kv["DT"]*100).round(1)
        c1,c2 = st.columns(2)
        with c1:
            fig_kv = px.pie(df_kv, names="Khu vực", values="DT", hole=0.4,
                            title="Cơ cấu DT theo Khu vực")
            st.plotly_chart(fig_kv, use_container_width=True)
        with c2:
            df_kv2 = df_kv.copy()
            df_kv2["DT"] = df_kv2["DT"].map(fmt)
            df_kv2["LN"] = df_kv2["LN"].map(fmt)
            st.dataframe(df_kv2, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
#  TAB 2 – CỤM THỜI GIAN
# ══════════════════════════════════════════════════════════════
with tab_time:
    st.markdown(f'<div class="section-title">📅 Phân tích theo: {time_level_label}</div>', unsafe_allow_html=True)

    df_t = (df_ban.groupby(time_col)
            .agg(DT=("Thành tiền bán","sum"),
                 LN=("Lợi nhuận","sum"),
                 KL=("Khối lượng", lambda x: round(x.sum()/1000,3)),
                 SL=("Số lượng","sum"),
                 N_CT=("Số chứng từ","nunique") if "Số chứng từ" in df_ban.columns else ("Thành tiền bán","count"),
                 N_KH=("Tên khách hàng","nunique"))
            .reset_index().sort_values(time_col))
    df_t["Bien"] = (df_t["LN"]/df_t["DT"].replace(0, np.nan)*100).round(1).fillna(0)
    df_t["Tang_truong"] = (df_t["DT"].pct_change()*100).round(1)

    fig = make_subplots(rows=3, cols=1, shared_xaxes=True,
                        subplot_titles=("Doanh thu & Lợi nhuận (VNĐ)",
                                        "Khối lượng (tấn) & Số lượng (cái)",
                                        "Biên LN (%) & Tăng trưởng DT (%)"),
                        vertical_spacing=0.08)
    fig.add_trace(go.Bar(x=df_t[time_col], y=df_t["DT"], name="Doanh thu",
                         marker_color="#4e79d4", opacity=0.85), row=1, col=1)
    fig.add_trace(go.Scatter(x=df_t[time_col], y=df_t["LN"], name="Lợi nhuận",
                             mode="lines+markers", line=dict(color="#26c281",width=2)), row=1, col=1)
    fig.add_trace(go.Bar(x=df_t[time_col], y=df_t["KL"], name="KL (tấn)",
                         marker_color="#9b59b6", opacity=0.8), row=2, col=1)
    fig.add_trace(go.Scatter(x=df_t[time_col], y=df_t["SL"], name="SL (cái)",
                             mode="lines+markers", line=dict(color="#f0a500",width=2)), row=2, col=1)
    fig.add_trace(go.Scatter(x=df_t[time_col], y=df_t["Bien"], name="Biên LN (%)",
                             mode="lines+markers+text",
                             text=[f"{v:.1f}%" for v in df_t["Bien"]],
                             textposition="top center",
                             line=dict(color="#e74c3c",width=2)), row=3, col=1)
    fig.add_trace(go.Bar(x=df_t[time_col], y=df_t["Tang_truong"],
                         name="Tăng trưởng (%)", marker_color="#4e79d4", opacity=0.4), row=3, col=1)
    fig.update_layout(height=680, legend=dict(orientation="h", y=-0.08))
    st.plotly_chart(fig, use_container_width=True)

    # Bảng
    df_t_s = df_t.copy()
    df_t_s["DT"] = df_t_s["DT"].map(fmt)
    df_t_s["LN"] = df_t_s["LN"].map(fmt)
    df_t_s = df_t_s.drop(columns=["Tang_truong"], errors="ignore")
    df_t_s.columns = [time_level_label,"DT (VNĐ)","LN (VNĐ)","KL (tấn)","SL (cái)","Số CT","Số KH","Biên (%)"]
    st.dataframe(df_t_s, use_container_width=True, hide_index=True)

    # Drill-down tháng khi đang xem Năm/Quý
    if time_level_label in ["Năm","Quý"] and single_kh:
        st.markdown('<div class="section-title">🔎 Chi tiết từng tháng</div>', unsafe_allow_html=True)
        df_m = (df_ban.groupby("Thang")
                .agg(DT=("Thành tiền bán","sum"), LN=("Lợi nhuận","sum"),
                     KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                     N_CT=("Số chứng từ","nunique") if "Số chứng từ" in df_ban.columns else ("Thành tiền bán","count"))
                .reset_index().sort_values("Thang"))
        df_m["Bien"] = (df_m["LN"]/df_m["DT"].replace(0,np.nan)*100).round(1).fillna(0)
        df_m["DT"]   = df_m["DT"].map(fmt)
        df_m["LN"]   = df_m["LN"].map(fmt)
        df_m.columns = ["Tháng","DT (VNĐ)","LN (VNĐ)","KL (tấn)","Số CT","Biên (%)"]
        st.dataframe(df_m, use_container_width=True, hide_index=True)

    # Heatmap KH × Tháng khi xem theo PKD
    if not single_kh:
        st.markdown('<div class="section-title">🗓️ Heatmap DT: KH × Tháng</div>', unsafe_allow_html=True)
        df_hm = (df_ban.groupby(["Tên khách hàng","Thang"])["Thành tiền bán"]
                 .sum().reset_index())
        df_piv = df_hm.pivot(index="Tên khách hàng", columns="Thang",
                              values="Thành tiền bán").fillna(0)
        fig_hm = px.imshow(df_piv, color_continuous_scale="Blues", aspect="auto",
                            title="Doanh thu (VNĐ): KH × Tháng")
        fig_hm.update_layout(height=max(350, len(df_piv)*26))
        st.plotly_chart(fig_hm, use_container_width=True)

# ══════════════════════════════════════════════════════════════
#  TAB 3 – SẢN PHẨM
# ══════════════════════════════════════════════════════════════
with tab_sp:
    st.markdown('<div class="section-title">📦 Cơ cấu sản phẩm & Thói quen mua</div>', unsafe_allow_html=True)

    df_nhom = (df_ban.groupby("Nhom_SP")
               .agg(N=("Nhom_SP","count"),
                    KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                    DT=("Thành tiền bán","sum"))
               .reset_index().sort_values("DT", ascending=False))

    c1,c2 = st.columns(2)
    with c1:
        fig1 = px.bar(df_nhom, x="Nhom_SP", y="DT", color="Nhom_SP", text_auto=".3s",
                      title="DT theo nhóm SP", labels={"DT":"DT (VNĐ)","Nhom_SP":""})
        fig1.update_layout(showlegend=False, height=360)
        st.plotly_chart(fig1, use_container_width=True)
    with c2:
        fig2 = px.pie(df_nhom, names="Nhom_SP", values="N", hole=0.45,
                      title="Tỷ trọng số lần mua")
        fig2.update_layout(height=360)
        st.plotly_chart(fig2, use_container_width=True)

    if "Tên hàng" in df_ban.columns:
        st.markdown('<div class="section-title">🏆 Top 15 sản phẩm & xu hướng</div>', unsafe_allow_html=True)
        df_top = (df_ban.groupby("Tên hàng")
                  .agg(N=("Tên hàng","count"),
                       KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                       DT=("Thành tiền bán","sum"))
                  .reset_index().sort_values("DT",ascending=False).head(15))

        fig3 = px.bar(df_top, y="Tên hàng", x="DT", orientation="h",
                      text_auto=".3s", color="DT", color_continuous_scale="Blues",
                      title="Top 15 sản phẩm theo DT")
        fig3.update_layout(height=480, yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig3, use_container_width=True)

        top5 = df_top["Tên hàng"].head(5).tolist()
        df_tr = (df_ban[df_ban["Tên hàng"].isin(top5)]
                 .groupby(["Tên hàng", time_col])["Thành tiền bán"].sum().reset_index())
        if not df_tr.empty:
            fig_tr = px.line(df_tr, x=time_col, y="Thành tiền bán", color="Tên hàng",
                             markers=True, title=f"Xu hướng Top 5 SP theo {time_level_label}",
                             labels={"Thành tiền bán":"DT (VNĐ)", time_col: time_level_label})
            fig_tr.update_layout(height=380)
            st.plotly_chart(fig_tr, use_container_width=True)

    st.markdown('<div class="section-title">🎯 Nhận định mục đích sử dụng</div>', unsafe_allow_html=True)
    nhom_set = set(df_nhom["Nhom_SP"].tolist())
    for k, (lvl, msg) in {
        "Ống HDPE":        ("info",   "🔵 <b>Ống HDPE</b>: Dự án hạ tầng cấp thoát nước, thuỷ lợi công trình lớn."),
        "Ống PVC nước":    ("low",    "🟢 <b>Ống PVC nước</b>: Xây dựng dân dụng, công nghiệp, khu công nghiệp."),
        "Ống PVC bơm cát": ("low",    "🟡 <b>Ống PVC bơm cát</b>: Thuỷ lợi, nông nghiệp, nuôi trồng thuỷ sản."),
        "Ống PPR":         ("info",   "🟠 <b>Ống PPR</b>: Hệ thống nước nóng/lạnh nội thất."),
        "Lõi PVC":         ("medium", "⚪ <b>Lõi PVC</b>: Có thể là đại lý hoặc nhà SX thứ cấp."),
        "Phụ kiện & Keo":  ("low",    "🔴 <b>Phụ kiện & Keo</b>: Tự thi công hoặc bán lại trọn gói."),
    }.items():
        if k in nhom_set:
            st.markdown(f'<div class="risk-{lvl}">{msg}</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  TAB 4 – LỢI NHUẬN
# ══════════════════════════════════════════════════════════════
with tab_ln_tab:
    st.markdown('<div class="section-title">💹 Lợi nhuận & Phát hiện Chính sách</div>', unsafe_allow_html=True)

    df_ln_t = (df_ban.groupby(time_col)
               .agg(DT=("Thành tiền bán","sum"),
                    Von=("Thành tiền vốn","sum"),
                    LN=("Lợi nhuận","sum"))
               .reset_index().sort_values(time_col))
    df_ln_t["Bien"] = (df_ln_t["LN"]/df_ln_t["DT"].replace(0,np.nan)*100).round(2).fillna(0)

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                        subplot_titles=("DT / Vốn / LN (VNĐ)", f"Biên LN (%) theo {time_level_label}"),
                        vertical_spacing=0.12)
    fig.add_trace(go.Bar(x=df_ln_t[time_col], y=df_ln_t["DT"],  name="Doanh thu",
                         marker_color="#4e79d4"), row=1, col=1)
    fig.add_trace(go.Bar(x=df_ln_t[time_col], y=df_ln_t["Von"], name="Giá vốn",
                         marker_color="#e05c5c"), row=1, col=1)
    fig.add_trace(go.Scatter(x=df_ln_t[time_col], y=df_ln_t["LN"], name="LN",
                             mode="lines+markers", line=dict(color="#26c281",width=2.5)), row=1, col=1)
    fig.add_trace(go.Scatter(x=df_ln_t[time_col], y=df_ln_t["Bien"], name="Biên LN (%)",
                             mode="lines+markers+text",
                             text=[f"{v:.1f}%" for v in df_ln_t["Bien"]],
                             textposition="top center",
                             line=dict(color="#f0a500",width=2.5)), row=2, col=1)
    fig.update_layout(height=520, barmode="group")
    st.plotly_chart(fig, use_container_width=True)

    # Phát hiện bất thường
    if len(df_ln_t) >= 2:
        mb = df_ln_t["Bien"].mean()
        sb = df_ln_t["Bien"].std() or 1
        st.markdown('<div class="section-title">🔍 Kỳ nghi có Chiết khấu / Chính sách đặc biệt</div>', unsafe_allow_html=True)
        for _, r in df_ln_t.iterrows():
            diff = r["Bien"] - mb
            if diff < -sb * 1.5:
                st.markdown(f'<div class="risk-high">🔴 <b>{r[time_col]}</b>: Biên LN = {r["Bien"]:.1f}% (TB: {mb:.1f}%) → Nghi chiết khấu lớn / giá đặc biệt</div>', unsafe_allow_html=True)
            elif diff < -sb:
                st.markdown(f'<div class="risk-medium">🟡 <b>{r[time_col]}</b>: Biên LN = {r["Bien"]:.1f}% (TB: {mb:.1f}%) → Cần kiểm tra chính sách</div>', unsafe_allow_html=True)

    # Giá bán < giá vốn
    st.markdown('<div class="section-title">❗ Dòng Giá bán &lt; Giá vốn (Bán lỗ)</div>', unsafe_allow_html=True)
    if "Giá bán" in df_ban.columns and "Giá vốn" in df_ban.columns:
        df_lo = df_ban[(df_ban["Giá bán"]>0)&(df_ban["Giá vốn"]>0)&(df_ban["Giá bán"]<df_ban["Giá vốn"])]
        if not df_lo.empty:
            st.error(f"⚠️ {len(df_lo)} dòng có giá bán < giá vốn")
            cols_lo = [c for c in ["Ngày chứng từ","Tên khách hàng","Tên hàng",
                                    "Giá bán","Giá vốn","Lợi nhuận","Ghi chú"] if c in df_lo.columns]
            st.dataframe(df_lo[cols_lo], use_container_width=True, hide_index=True)
        else:
            st.markdown('<div class="risk-low">✅ Không có dòng nào giá bán thấp hơn giá vốn.</div>', unsafe_allow_html=True)

    # Hàng trả lại
    df_tra = df_scope[df_scope["Loai_GD"] == "Trả hàng"]
    if not df_tra.empty:
        st.markdown('<div class="section-title">↩️ Hàng trả lại</div>', unsafe_allow_html=True)
        cols_tr = [c for c in ["Số chứng từ","Ngày chứng từ","Tên khách hàng",
                                "Tên hàng","Thành tiền bán","Ghi chú"] if c in df_tra.columns]
        df_tra_s = df_tra[cols_tr].copy()
        if "Thành tiền bán" in df_tra_s.columns:
            df_tra_s["Thành tiền bán"] = df_tra_s["Thành tiền bán"].map(fmt)
        st.error(f"Tổng GT trả hàng: {fmt(abs(df_tra['Thành tiền bán'].sum()))} VNĐ")
        st.dataframe(df_tra_s, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
#  TAB 5 – GIAO HÀNG
# ══════════════════════════════════════════════════════════════
with tab_giao:
    st.markdown('<div class="section-title">🚚 Hình thức & Địa điểm giao hàng</div>', unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1:
        if "Freight Terms" in df_ban.columns:
            ft = df_ban["Freight Terms"].value_counts().reset_index()
            ft.columns = ["Hình thức","Số lần"]
            st.plotly_chart(px.pie(ft, names="Hình thức", values="Số lần",
                                   title="Điều kiện giao hàng", hole=0.4),
                            use_container_width=True)
    with c2:
        if "Shipping method" in df_ban.columns:
            sm = df_ban["Shipping method"].value_counts().reset_index()
            sm.columns = ["Phương tiện","Số lần"]
            st.plotly_chart(px.pie(sm, names="Phương tiện", values="Số lần",
                                   title="Phương tiện vận chuyển", hole=0.4),
                            use_container_width=True)

    if "Nơi giao hàng" in df_ban.columns:
        st.markdown('<div class="section-title">📍 Địa điểm giao hàng</div>', unsafe_allow_html=True)
        df_noi = (df_ban.groupby("Nơi giao hàng")
                  .agg(N=("Nơi giao hàng","count"),
                       KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                       DT=("Thành tiền bán","sum"))
                  .reset_index().sort_values("DT",ascending=False))
        df_noi["DT"] = df_noi["DT"].map(fmt)
        df_noi.columns = ["Địa điểm","Số lần","KL (tấn)","DT (VNĐ)"]
        st.dataframe(df_noi, use_container_width=True, hide_index=True)

    if "Mã kho" in df_ban.columns:
        st.markdown('<div class="section-title">🏭 Kho xuất hàng</div>', unsafe_allow_html=True)
        df_kho = (df_ban.groupby("Mã kho")
                  .agg(N=("Mã kho","count"), DT=("Thành tiền bán","sum"),
                       KL=("Khối lượng", lambda x: round(x.sum()/1000,2)))
                  .reset_index().sort_values("DT",ascending=False))
        df_kho["DT"] = df_kho["DT"].map(fmt)
        st.dataframe(df_kho, use_container_width=True, hide_index=True)
        if df_ban["Mã kho"].nunique() >= 2:
            st.markdown('<div class="risk-medium">🟡 KH nhận hàng từ nhiều kho – kiểm tra giá xuất kho & tính nhất quán.</div>', unsafe_allow_html=True)

    with st.expander("🚛 Danh sách xe & tài xế"):
        xe_cols = [c for c in ["Biển số xe","Tài Xế","Tên ĐVVC","Shipping method",
                                "Ngày chứng từ","Nơi giao hàng"] if c in df_ban.columns]
        if xe_cols:
            st.dataframe(df_ban[xe_cols].drop_duplicates().sort_values("Ngày chứng từ"),
                         use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
#  TAB 6 – TẦN SUẤT
# ══════════════════════════════════════════════════════════════
with tab_freq:
    st.markdown('<div class="section-title">🔁 Heatmap tần suất: Sản phẩm × Tháng</div>', unsafe_allow_html=True)
    if "Tên hàng" in df_ban.columns:
        df_fr = df_ban.groupby(["Tên hàng","Thang"]).size().reset_index(name="N")
        df_piv = df_fr.pivot(index="Tên hàng", columns="Thang", values="N").fillna(0)
        top_n = df_piv.sum(axis=1).nlargest(min(25, len(df_piv))).index
        fig_h = px.imshow(df_piv.loc[top_n],
                          labels=dict(x="Tháng",y="Sản phẩm",color="Số lần"),
                          title="Heatmap tần suất – Top SP × Tháng",
                          color_continuous_scale="YlOrRd", aspect="auto")
        fig_h.update_layout(height=max(400, len(top_n)*22))
        st.plotly_chart(fig_h, use_container_width=True)

        st.markdown('<div class="section-title">📋 Chi tiết theo Quý – từng tháng</div>', unsafe_allow_html=True)
        for q in sorted(df_ban["Quy"].dropna().unique()):
            df_q = df_ban[df_ban["Quy"] == q]
            months = sorted(df_q["Thang"].dropna().unique())
            with st.expander(f"📅 {q}  |  {', '.join(months)}"):
                agg = (df_q.groupby(["Tên hàng","Thang"])
                       .agg(N=("Tên hàng","count"),
                            KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                            DT=("Thành tiền bán","sum"))
                       .reset_index().sort_values("DT",ascending=False))
                agg["DT"] = agg["DT"].map(fmt)
                agg.columns = ["Sản phẩm","Tháng","Số lần","KL (tấn)","DT (VNĐ)"]
                st.dataframe(agg, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
#  TAB 7 – RỦI RO
# ══════════════════════════════════════════════════════════════
with tab_risk:
    st.markdown('<div class="section-title">⚠️ Nhận diện Rủi ro & Dấu hiệu Bất thường (12 nhóm)</div>', unsafe_allow_html=True)

    risks  = []   # (level, category, message, detail_df_or_None)
    score  = 0
    tong_dt_ban = df_ban["Thành tiền bán"].sum()
    tong_ln_ban = df_ban["Lợi nhuận"].sum()
    bien_tb     = (tong_ln_ban/tong_dt_ban*100) if tong_dt_ban else 0

    # R1 – Hàng trả lại
    df_tra = df_scope[df_scope["Loai_GD"] == "Trả hàng"]
    tong_tra = abs(df_tra["Thành tiền bán"].sum())
    tl_tra   = (tong_tra/tong_dt_ban*100) if tong_dt_ban else 0
    if tl_tra > 10:
        score += 30
        risks.append(("high","↩️ Trả hàng",
                       f"Tỷ lệ trả hàng <b>{tl_tra:.1f}%</b> ({fmt(tong_tra)} VNĐ) — Rủi ro tranh chấp chất lượng / hủy HĐ cao", df_tra))
    elif tl_tra > 3:
        score += 15
        risks.append(("medium","↩️ Trả hàng", f"Tỷ lệ trả hàng <b>{tl_tra:.1f}%</b> — Cần theo dõi", df_tra))
    else:
        risks.append(("low","↩️ Trả hàng", f"Tỷ lệ trả hàng thấp: {tl_tra:.1f}% ✓", None))

    # R2 – Giá bán < Giá vốn
    if "Giá bán" in df_ban.columns and "Giá vốn" in df_ban.columns:
        df_lo = df_ban[(df_ban["Giá bán"]>0)&(df_ban["Giá vốn"]>0)&(df_ban["Giá bán"]<df_ban["Giá vốn"])]
        if not df_lo.empty:
            score += 25
            risks.append(("high","💸 Bán lỗ trực tiếp",
                           f"<b>{len(df_lo)}</b> dòng giá bán &lt; giá vốn — bán lỗ hoặc sai dữ liệu giá", df_lo))
        else:
            risks.append(("low","💸 Giá bán vs Giá vốn","Không có dòng nào giá bán < giá vốn ✓", None))

    # R3 – Lợi nhuận âm (không phải trả hàng)
    df_ln_am = df_ban[df_ban["Lợi nhuận"] < 0]
    if not df_ln_am.empty:
        tong_ln_am = abs(df_ln_am["Lợi nhuận"].sum())
        score += 20
        risks.append(("high","📉 Lợi nhuận âm",
                       f"<b>{len(df_ln_am)}</b> dòng xuất bán có LN âm (tổng -{fmt(tong_ln_am)} VNĐ) — chiết khấu lớn hoặc sai giá vốn", df_ln_am))
    else:
        risks.append(("low","📉 Lợi nhuận âm","Không có dòng xuất bán nào bị lỗ ✓", None))

    # R4 – Biên LN bất thường theo kỳ
    if len(df_ln_t) >= 2:
        mb2 = df_ln_t["Bien"].mean()
        sb2 = df_ln_t["Bien"].std() or 1
        anom_kys = df_ln_t[df_ln_t["Bien"] < mb2 - sb2*1.5]
        if not anom_kys.empty:
            ky_str = ", ".join(anom_kys[time_col].tolist())
            score += 10
            risks.append(("medium","📊 Biên LN bất thường",
                           f"Kỳ <b>{ky_str}</b> có biên LN thấp bất thường → nghi có chiết khấu / chương trình đặc biệt", None))
        else:
            risks.append(("low","📊 Biên LN","Biên LN ổn định qua các kỳ ✓", None))

    # R5 – Tập trung đơn hàng cuối tháng
    df_cuoi   = df_ban[df_ban["Cuoi_thang"] == True]
    dt_cuoi   = df_cuoi["Thành tiền bán"].sum()
    tl_dt_cuoi = (dt_cuoi/tong_dt_ban*100) if tong_dt_ban else 0
    if tl_dt_cuoi > 40:
        score += 15
        risks.append(("high","📅 Tập trung cuối tháng",
                       f"<b>{tl_dt_cuoi:.1f}%</b> DT ({fmt(dt_cuoi)} VNĐ) tập trung ngày 28–31 — nghi đẩy doanh số / ghi nhận ảo", None))
    elif tl_dt_cuoi > 25:
        score += 8
        risks.append(("medium","📅 Tập trung cuối tháng",
                       f"{tl_dt_cuoi:.1f}% DT tập trung cuối tháng — kiểm tra thực chất giao nhận", None))
    else:
        risks.append(("low","📅 Phân bổ thời gian", f"Đơn hàng phân bổ đều trong tháng ✓ (cuối tháng: {tl_dt_cuoi:.1f}%)", None))

    # R6 – Đơn hàng outlier
    q1v  = df_ban["Thành tiền bán"].quantile(0.25)
    q3v  = df_ban["Thành tiền bán"].quantile(0.75)
    iqr  = q3v - q1v
    if iqr > 0:
        df_out = df_ban[df_ban["Thành tiền bán"] > q3v + 3*iqr]
        if not df_out.empty:
            score += 10
            risks.append(("medium","📦 Đơn hàng giá trị lớn bất thường",
                           f"<b>{len(df_out)}</b> đơn vượt 3×IQR — kiểm soát hạn mức tín dụng", df_out))
        else:
            risks.append(("low","📦 Quy mô đơn hàng","Không có đơn bất thường về giá trị ✓", None))

    # R7 – Xuất bổ sung / thay thế
    df_bs = df_scope[df_scope["Loai_GD"] == "Xuất bổ sung"]
    if not df_bs.empty:
        score += 12
        risks.append(("medium","🔄 Xuất bổ sung/thay thế",
                       f"<b>{len(df_bs)}</b> dòng xuất bổ sung/thay thế — giao nhầm, giao thiếu hoặc hàng lỗi", df_bs))
    else:
        risks.append(("low","🔄 Xuất bổ sung","Không có đơn xuất bổ sung/thay thế ✓", None))

    # R8 – Địa điểm giao hàng phân tán
    if "Nơi giao hàng" in df_ban.columns:
        n_noi = df_ban["Nơi giao hàng"].nunique()
        if n_noi >= 6:
            score += 10
            risks.append(("medium","📍 Giao hàng phân tán",
                           f"Giao tới <b>{n_noi}</b> địa điểm — nhà thầu nhiều dự án, rủi ro phân tán công nợ", None))
        else:
            risks.append(("low","📍 Địa điểm giao hàng", f"{n_noi} địa điểm — bình thường ✓", None))

    # R9 – Nhiều kho
    if "Mã kho" in df_ban.columns and df_ban["Mã kho"].nunique() >= 2:
        score += 5
        risks.append(("medium","🏭 Nhiều kho xuất hàng",
                       f"KH nhận từ <b>{df_ban['Mã kho'].nunique()}</b> kho — kiểm tra giá xuất kho & nhất quán", None))
    else:
        risks.append(("low","🏭 Kho xuất hàng","Xuất từ 1 kho — nhất quán ✓", None))

    # R10 – Tần suất mua không đều
    n_thang_mua   = df_ban["Thang"].nunique()
    delta_days    = (df_ban["Ngày chứng từ"].max()-df_ban["Ngày chứng từ"].min()).days
    n_thang_range = max(delta_days//30, 1)
    if n_thang_mua/n_thang_range < 0.5:
        score += 10
        risks.append(("medium","🔁 Mua hàng không đều",
                       f"Chỉ mua <b>{n_thang_mua}/{n_thang_range}</b> tháng — phụ thuộc dự án, dòng tiền bất ổn", None))
    else:
        risks.append(("low","🔁 Tần suất mua", f"Mua đều đặn: {n_thang_mua} tháng có GD ✓", None))

    # R11 – Biên LN tổng thấp
    if bien_tb < 5:
        score += 20
        risks.append(("high","💹 Biên LN tổng thấp",
                       f"Biên LN tổng = <b>{bien_tb:.1f}%</b> — nguy cơ bán phá giá hoặc chiết khấu quá mức", None))
    elif bien_tb < 15:
        score += 8
        risks.append(("medium","💹 Biên LN tổng", f"Biên LN = <b>{bien_tb:.1f}%</b> — ở mức thấp, cần theo dõi", None))
    else:
        risks.append(("low","💹 Biên LN tổng", f"Biên LN = {bien_tb:.1f}% — bình thường ✓", None))

    # R12 – Đơn PO/Dự án (BCCN)
    if "Ghi chú" in df_scope.columns:
        gc2    = df_scope["Ghi chú"].astype(str).str.upper()
        df_po  = df_scope[gc2.str.contains(r"PO|B[0-9]{3}|HỢP ĐỒNG", regex=True, na=False)]
        if not df_po.empty:
            gt_po = abs(df_po["Thành tiền bán"].sum())
            score += 8
            n_po  = df_po["Số chứng từ"].nunique() if "Số chứng từ" in df_po.columns else len(df_po)
            risks.append(("medium","📄 Đơn PO/Dự án (BCCN)",
                           f"<b>{n_po}</b> chứng từ có PO/HĐ ({fmt(gt_po)} VNĐ) — thanh toán thường chậm NET 30–90 ngày, rủi ro công nợ dự án", df_po))

    # ══ TỔNG ĐIỂM ══
    st.markdown("---")
    n_high   = sum(1 for l,_,_,_ in risks if l=="high")
    n_medium = sum(1 for l,_,_,_ in risks if l=="medium")
    n_low    = sum(1 for l,_,_,_ in risks if l=="low")

    if score >= 60:
        color, label = "#e74c3c", "🔴 RỦI RO CAO"
    elif score >= 30:
        color, label = "#f39c12", "🟡 RỦI RO TRUNG BÌNH"
    else:
        color, label = "#26c281", "🟢 RỦI RO THẤP"

    st.markdown(f"""
    <div style='background:#1a2035;border-radius:10px;padding:20px;text-align:center;margin:10px 0 18px 0;'>
        <div style='font-size:32px;font-weight:900;color:{color};'>{label}</div>
        <div style='font-size:15px;color:#9aa0b0;margin-top:8px;'>
            Điểm rủi ro: <b style='color:{color}'>{score}/130</b>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            🔴 {n_high} cao &nbsp; 🟡 {n_medium} trung bình &nbsp; 🟢 {n_low} thấp
        </div>
    </div>
    """, unsafe_allow_html=True)

    for lvl, cat, msg, detail in risks:
        st.markdown(f'<div class="risk-{lvl}"><b>[{cat}]</b>&nbsp; {msg}</div>', unsafe_allow_html=True)
        if detail is not None and not detail.empty and lvl in ["high","medium"]:
            show_c = [c for c in ["Số chứng từ","Ngày chứng từ","Tên khách hàng","Tên hàng",
                                   "Thành tiền bán","Lợi nhuận","Giá bán","Giá vốn","Ghi chú"] if c in detail.columns]
            with st.expander(f"📋 Chi tiết – {cat}"):
                d2 = detail[show_c].head(50).copy()
                for c in ["Thành tiền bán","Lợi nhuận","Giá bán","Giá vốn"]:
                    if c in d2.columns:
                        d2[c] = d2[c].map(fmt)
                st.dataframe(d2, use_container_width=True, hide_index=True)

    # BCCN Note
    st.markdown("---")
    st.markdown('<div class="section-title">📄 BCCN – Lưu ý Thanh toán & Công nợ</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="risk-info">
    ⚠️ File OM_RPT_055 <b>không chứa ngày thanh toán thực tế</b>. Để phân tích BCCN đầy đủ cần bổ sung:<br>
    &nbsp;&nbsp;• <b>Sổ AR Aging</b> → số ngày tồn đọng, quá hạn, hạn mức tín dụng<br>
    &nbsp;&nbsp;• <b>Lịch sử thanh toán</b> → thói quen NET 30/60/90<br><br>
    <b>Dấu hiệu nhận biết từ dữ liệu hiện có:</b><br>
    &nbsp;&nbsp;• Ghi chú <b>B-xxx / PO</b> → đơn dự án → TT chậm, phụ thuộc chủ đầu tư<br>
    &nbsp;&nbsp;• <b>Trả hàng</b> → tranh chấp → kéo dài & phức tạp hoá công nợ<br>
    &nbsp;&nbsp;• <b>Nhiều địa điểm giao</b> → công nợ phân tán nhiều công trình<br>
    &nbsp;&nbsp;• <b>Đơn cuối tháng giá trị lớn</b> → kiểm tra giao nhận thực tế trước khi ghi nhận DT
    </div>
    """, unsafe_allow_html=True)