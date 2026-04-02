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
.info-box    {background:#1a2035;border-radius:8px;padding:14px;margin:8px 0;color:#ccc;font-size:14px;}
.section-title{font-size:16px;font-weight:700;color:#e0e0e0;margin:18px 0 8px 0;
               padding-bottom:5px;border-bottom:1px solid #2e3350;}
.badge-pkd   {background:#1e3a5f;color:#7ec8f7;padding:3px 10px;border-radius:12px;
               font-size:12px;font-weight:600;margin-right:4px;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════════
COL_ALIASES = {
    "Tên khách hàng":     ["Tên khách hàng","Tên KH","Khách hàng","Tên khách"],
    "Mã khách hàng":      ["Mã khách hàng","Mã KH","customer_code"],
    "Mã nhóm KH":         ["Mã nhóm KH","Nhóm KH","dept_code"],
    "Tên nhóm KH":        ["Tên nhóm KH","Phòng KD","dept_name"],
    "Ngày chứng từ":      ["Ngày chứng từ","Ngày CT","date"],
    "Số chứng từ":        ["Số chứng từ","Số CT","voucher_no"],
    "Số ĐH":              ["Số ĐH","Order No","order_no"],
    "Tên hàng":           ["Tên hàng","Tên mã hàng","Tên SP"],
    "Mã hàng":            ["Mã hàng","Mã SP"],
    "Mã nhóm hàng":       ["Mã nhóm hàng","Nhóm hàng"],
    "Khối lượng":         ["Khối lượng","KL","weight"],
    "Số lượng":           ["Số lượng","SL","qty"],
    "Thành tiền bán":     ["Thành tiền bán","Doanh thu","revenue"],
    "Thành tiền vốn":     ["Thành tiền vốn","Giá vốn hàng bán","cost_total"],
    "Lợi nhuận":          ["Lợi nhuận","Profit"],
    "Giá bán":            ["Giá bán","Đơn giá","unit_price"],
    "Giá vốn":            ["Giá vốn","cost_price"],
    "Đơn giá vận chuyển": ["Đơn giá vận chuyển","Chi phí VC"],
    "Đơn giá quy đổi":    ["Đơn giá quy đổi","converted_price"],
    "Nơi giao hàng":      ["Nơi giao hàng","Địa chỉ giao hàng"],
    "Freight Terms":      ["Freight Terms","Điều kiện giao hàng"],
    "Shipping method":    ["Shipping method","Phương tiện"],
    "Biển số xe":         ["Số xe","Biển số xe","Biển xe"],
    "Tài Xế":             ["Tài Xế","Tài xế","Driver"],
    "Tên ĐVVC":           ["Tên ĐVVC","Đơn vị vận chuyển"],
    "Ghi chú":            ["Ghi chú","Note","Ghi Chú"],
    "Khu vực":            ["Khu vực","Region"],
    "Loại đơn hàng":      ["Loại đơn hàng","Order Type"],
}

NUM_COLS = ["Thành tiền bán","Thành tiền vốn","Lợi nhuận",
            "Khối lượng","Số lượng","Giá bán","Giá vốn",
            "Đơn giá vận chuyển","Đơn giá quy đổi"]

SP_MAPPING = [
    ("Ống HDPE",        r"HDPE"),
    ("Ống PVC nước",    r"PVC.*(?:nước|nong dài|nong trơn|thoát)"),
    ("Ống PVC bơm cát", r"PVC.*(?:cát|bơm cát)"),
    ("Ống PPR",         r"PPR"),
    ("Lõi PVC",         r"(?:Lơi|Lõi|lori)"),
    ("Phụ kiện & Keo",  r"(?:Nối|Co |Tê |Van |Keo |Măng|Bít)"),
]

# ══════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════
def normalize_columns(df):
    cols_lower = {c.strip().lower(): c for c in df.columns}
    rename = {}
    for std, aliases in COL_ALIASES.items():
        if std in df.columns:
            continue
        for a in aliases:
            if a in df.columns:
                rename[a] = std; break
            if a.strip().lower() in cols_lower:
                rename[cols_lower[a.strip().lower()]] = std; break
    df.rename(columns=rename, inplace=True)
    return df

def find_header_row(fb):
    try:
        raw = pd.read_excel(io.BytesIO(fb), header=None, engine="openpyxl", nrows=40)
    except Exception:
        return 0
    kws = ["khách hàng","tên kh","số chứng từ","ngày chứng từ","tên hàng","mã khách"]
    for i in range(raw.shape[0]):
        vals = ["" if (isinstance(v,float) and pd.isna(v)) else str(v) for v in raw.iloc[i]]
        txt  = " ".join(vals).lower()
        if sum(1 for k in kws if k in txt) >= 2:
            return i
    return 0

def col(df, name, default=0):
    if name in df.columns:
        return pd.to_numeric(df[name], errors="coerce").fillna(default)
    return pd.Series([default]*len(df), index=df.index)

def fmt(n): return f"{n:,.0f}"

# ══════════════════════════════════════════════════════════════
#  LOAD
# ══════════════════════════════════════════════════════════════
@st.cache_data(show_spinner="Đang xử lý dữ liệu…")
def load_all(file_data):
    frames = []
    for name, fb in file_data:
        try:
            hr = find_header_row(fb)
            df = pd.read_excel(io.BytesIO(fb), header=hr, engine="openpyxl")
            df.columns = [str(c).strip().replace("\n"," ") for c in df.columns]
            df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
            df.dropna(how="all", inplace=True)
            normalize_columns(df)
            df["Nguồn file"] = name
            frames.append(df)
        except Exception as e:
            st.warning(f"⚠️ Lỗi đọc `{name}`: {e}")
    if not frames:
        return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True)

    for req in ["Ngày chứng từ","Tên khách hàng"]:
        if req not in df.columns:
            st.error(f"❌ Không tìm thấy cột '{req}'. Kiểm tra lại file.")
            return pd.DataFrame()

    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], dayfirst=True, errors="coerce")
    df = df[df["Ngày chứng từ"].notna()].copy()

    # Thời gian
    df["Năm"]  = df["Ngày chứng từ"].dt.year.astype(str)
    df["Quý"]  = df["Ngày chứng từ"].dt.to_period("Q").astype(str)
    df["Tháng"]= df["Ngày chứng từ"].dt.to_period("M").astype(str)
    df["Tuần"] = df["Ngày chứng từ"].dt.isocalendar().week.astype(str)
    df["Ngày"] = df["Ngày chứng từ"].dt.date

    # Số
    for c in NUM_COLS:
        df[c] = col(df, c)

    # Loại GD
    gc = df["Ghi chú"].astype(str).str.upper() if "Ghi chú" in df.columns else pd.Series([""]*len(df))
    df["Loại GD"] = "Xuất bán"
    df.loc[gc.str.contains(r"NHẬP TRẢ|TRẢ HÀNG", regex=True, na=False), "Loại GD"] = "Trả hàng"
    df.loc[gc.str.contains(r"BỔ SUNG|THAY THẾ", regex=True, na=False), "Loại GD"]  = "Xuất bổ sung"

    # Nhóm SP
    ten = df["Tên hàng"].astype(str) if "Tên hàng" in df.columns else pd.Series([""]*len(df))
    df["Nhóm SP"] = "Khác"
    for label, pat in SP_MAPPING:
        df.loc[ten.str.contains(pat, case=False, regex=True, na=False), "Nhóm SP"] = label

    # PKD display
    if "Mã nhóm KH" in df.columns and "Tên nhóm KH" in df.columns:
        df["PKD"] = df["Mã nhóm KH"].astype(str) + " – " + df["Tên nhóm KH"].astype(str)
    elif "Mã nhóm KH" in df.columns:
        df["PKD"] = df["Mã nhóm KH"].astype(str)
    else:
        df["PKD"] = "Không xác định"

    return df

# ══════════════════════════════════════════════════════════════
#  UPLOAD
# ══════════════════════════════════════════════════════════════
st.sidebar.markdown("## 📂 Upload dữ liệu")
uploaded = st.sidebar.file_uploader("File Excel báo cáo bán hàng",
                                     type=["xlsx"], accept_multiple_files=True)
if not uploaded:
    st.title("📊 Phân tích Kinh Doanh – Hoa Sen")
    st.info("👈 Upload file Excel báo cáo bán hàng để bắt đầu.")
    st.stop()

file_data = [(u.name, u.read()) for u in uploaded]
df_all = load_all(file_data)
if df_all.empty:
    st.error("Không có dữ liệu hợp lệ."); st.stop()

# ══════════════════════════════════════════════════════════════
#  SIDEBAR FILTERS
# ══════════════════════════════════════════════════════════════
st.sidebar.markdown("---")
st.sidebar.markdown("## 🔍 Bộ lọc")

# Chế độ xem: toàn phòng KD hoặc 1 KH
view_mode = st.sidebar.radio("Chế độ xem", ["📋 Tổng quan Phòng KD", "👤 Chi tiết Khách hàng"])

pkd_list = sorted(df_all["PKD"].dropna().unique())
pkd_sel  = st.sidebar.multiselect("🏢 Phòng KD", pkd_list, default=pkd_list)

# Bộ lọc thời gian theo cụm
time_level = st.sidebar.selectbox("⏱ Cụm thời gian", ["Tháng","Quý","Năm"])

nam_list  = sorted(df_all["Năm"].dropna().unique())
quy_list  = sorted(df_all["Quý"].dropna().unique())
thang_list= sorted(df_all["Tháng"].dropna().unique())

if time_level == "Năm":
    sel_time = st.sidebar.multiselect("Chọn Năm",  nam_list,   default=nam_list)
    time_col = "Năm"
elif time_level == "Quý":
    sel_time = st.sidebar.multiselect("Chọn Quý",  quy_list,   default=quy_list)
    time_col = "Quý"
else:
    sel_time = st.sidebar.multiselect("Chọn Tháng",thang_list, default=thang_list)
    time_col = "Tháng"

df_pkd = df_all[df_all["PKD"].isin(pkd_sel) & df_all[time_col].isin(sel_time)].copy()

if view_mode == "👤 Chi tiết Khách hàng":
    kh_list = sorted(df_pkd["Tên khách hàng"].dropna().astype(str).unique())
    kh_sel  = st.sidebar.selectbox("👤 Khách hàng", kh_list)
    df_view = df_pkd[df_pkd["Tên khách hàng"].astype(str) == kh_sel].copy()
    page_title = f"👤 {kh_sel}"
else:
    kh_sel  = None
    df_view = df_pkd.copy()
    page_title = "📋 Tổng quan Phòng Kinh Doanh"

df_ban  = df_view[df_view["Loại GD"] == "Xuất bán"].copy()
df_tra  = df_view[df_view["Loại GD"] == "Trả hàng"].copy()
df_bs   = df_view[df_view["Loại GD"] == "Xuất bổ sung"].copy()

# ══════════════════════════════════════════════════════════════
#  HEADER KPIs
# ══════════════════════════════════════════════════════════════
st.title(page_title)

tong_dt  = df_ban["Thành tiền bán"].sum()
tong_von = df_ban["Thành tiền vốn"].sum()
tong_ln  = df_ban["Lợi nhuận"].sum()
bien_ln  = tong_ln/tong_dt*100 if tong_dt else 0
tong_kl  = df_ban["Khối lượng"].sum()/1000
n_ct     = df_ban["Số chứng từ"].nunique() if "Số chứng từ" in df_ban.columns else len(df_ban)
n_kh     = df_ban["Tên khách hàng"].nunique()

c1,c2,c3,c4,c5,c6 = st.columns(6)
c1.metric("💰 Doanh thu",    fmt(tong_dt)+" đ")
c2.metric("📦 Số chứng từ",  fmt(n_ct))
c3.metric("💹 Lợi nhuận",    fmt(tong_ln)+" đ")
c4.metric("📊 Biên LN",      f"{bien_ln:.1f}%")
c5.metric("⚖️ KL (tấn)",     f"{tong_kl:,.2f}")
c6.metric("👥 Số KH",        str(n_kh))

st.divider()

# ══════════════════════════════════════════════════════════════
#  TABS
# ══════════════════════════════════════════════════════════════
tabs = st.tabs([
    "🏢 Phòng KD & KH",
    "⏱ Phân tích Thời gian",
    "📦 Sản phẩm & Thói quen",
    "💹 Lợi nhuận & Chính sách",
    "🚚 Giao hàng",
    "🔁 Tần suất mua hàng",
    "⚠️ Rủi ro & Bất thường",
])
tab_pkd, tab_time, tab_sp, tab_ln, tab_gh, tab_freq, tab_risk = tabs

# ══════════════════════════════════════════════════════════════
#  TAB 1 – PHÒNG KD & KH
# ══════════════════════════════════════════════════════════════
with tab_pkd:
    st.markdown('<div class="section-title">🏢 Doanh thu theo Phòng Kinh Doanh</div>', unsafe_allow_html=True)

    df_pkd_sum = (df_ban.groupby("PKD")
                  .agg(DT=("Thành tiền bán","sum"),
                       LN=("Lợi nhuận","sum"),
                       KL=("Khối lượng", lambda x: x.sum()/1000),
                       n_KH=("Tên khách hàng","nunique"),
                       n_CT=("Số chứng từ","nunique") if "Số chứng từ" in df_ban.columns else ("Thành tiền bán","count"))
                  .reset_index().sort_values("DT", ascending=False))
    df_pkd_sum["Biên LN (%)"] = (df_pkd_sum["LN"]/df_pkd_sum["DT"].replace(0,np.nan)*100).round(1).fillna(0)

    col_l, col_r = st.columns([1.5,1])
    with col_l:
        fig = px.bar(df_pkd_sum, x="PKD", y="DT", color="PKD", text_auto=".3s",
                     title="Doanh thu theo Phòng KD",
                     labels={"DT":"Doanh thu (VNĐ)","PKD":""})
        fig.update_layout(showlegend=False, height=340)
        st.plotly_chart(fig, use_container_width=True)
    with col_r:
        fig2 = px.scatter(df_pkd_sum, x="DT", y="Biên LN (%)", size="n_KH",
                          color="PKD", text="PKD", title="DT vs Biên LN (bong bóng = số KH)",
                          labels={"DT":"Doanh thu"})
        fig2.update_traces(textposition="top center")
        fig2.update_layout(height=340, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    df_pkd_show = df_pkd_sum.copy()
    df_pkd_show["DT"] = df_pkd_show["DT"].map(fmt)
    df_pkd_show["LN"] = df_pkd_show["LN"].map(fmt)
    df_pkd_show["KL"] = df_pkd_show["KL"].round(2)
    df_pkd_show.columns = ["Phòng KD","Doanh thu","Lợi nhuận","KL (tấn)","Số KH","Số CT","Biên LN (%)"]
    st.dataframe(df_pkd_show, use_container_width=True, hide_index=True)

    # KH theo từng PKD
    st.markdown('<div class="section-title">👥 Danh sách khách hàng theo Phòng KD</div>', unsafe_allow_html=True)

    for pkd_val in sorted(df_ban["PKD"].unique()):
        df_p = df_ban[df_ban["PKD"] == pkd_val]
        df_kh_sum = (df_p.groupby("Tên khách hàng")
                     .agg(DT=("Thành tiền bán","sum"),
                          LN=("Lợi nhuận","sum"),
                          KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                          n_CT=("Số chứng từ","nunique") if "Số chứng từ" in df_p.columns else ("Thành tiền bán","count"),
                          n_SP=("Tên hàng","nunique") if "Tên hàng" in df_p.columns else ("Thành tiền bán","count"))
                     .reset_index().sort_values("DT", ascending=False))
        df_kh_sum["Biên LN (%)"] = (df_kh_sum["LN"]/df_kh_sum["DT"].replace(0,np.nan)*100).round(1).fillna(0)
        df_kh_sum["DT"] = df_kh_sum["DT"].map(fmt)
        df_kh_sum["LN"] = df_kh_sum["LN"].map(fmt)
        df_kh_sum.columns = ["Khách hàng","DT (VNĐ)","LN (VNĐ)","KL (tấn)","Số CT","Số SP","Biên LN (%)"]

        with st.expander(f"🏢 {pkd_val}  —  {len(df_kh_sum)} KH", expanded=True):
            fig_kh = px.bar(df_kh_sum.head(15), x="Khách hàng", y="DT (VNĐ)",
                            color="Biên LN (%)", color_continuous_scale="RdYlGn",
                            title=f"Top KH – {pkd_val}", text_auto=True)
            fig_kh.update_layout(xaxis_tickangle=-30, height=320, showlegend=False)
            st.plotly_chart(fig_kh, use_container_width=True)
            st.dataframe(df_kh_sum, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
#  TAB 2 – PHÂN TÍCH THỜI GIAN
# ══════════════════════════════════════════════════════════════
with tab_time:
    st.markdown(f'<div class="section-title">⏱ Phân tích theo {time_level} – Drill-down Năm → Quý → Tháng</div>', unsafe_allow_html=True)

    # Drill-down 3 cấp
    for level, grp_col in [("Năm","Năm"),("Quý","Quý"),("Tháng","Tháng")]:
        df_t = (df_ban.groupby(grp_col)
                .agg(DT=("Thành tiền bán","sum"),
                     LN=("Lợi nhuận","sum"),
                     KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                     n_CT=("Số chứng từ","nunique") if "Số chứng từ" in df_ban.columns else ("Thành tiền bán","count"),
                     n_KH=("Tên khách hàng","nunique"))
                .reset_index().sort_values(grp_col))
        df_t["Biên (%)"] = (df_t["LN"]/df_t["DT"].replace(0,np.nan)*100).round(1).fillna(0)
        df_t["Tăng trưởng DT (%)"] = df_t["DT"].pct_change()*100

        if df_t.empty or len(df_t) < 1:
            continue

        with st.expander(f"📅 Theo {level}", expanded=(level==time_level)):
            fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                                subplot_titles=(f"Doanh thu & KL (tấn) theo {level}",
                                                f"Biên LN (%) & Tăng trưởng (%)"),
                                vertical_spacing=0.12)
            fig.add_trace(go.Bar(x=df_t[grp_col], y=df_t["DT"], name="Doanh thu",
                                 marker_color="#4e79d4", opacity=0.85), row=1, col=1)
            fig.add_trace(go.Scatter(x=df_t[grp_col], y=df_t["KL"], name="KL (tấn)",
                                     mode="lines+markers", line=dict(color="#f0a500",width=2)), row=1, col=1)
            fig.add_trace(go.Scatter(x=df_t[grp_col], y=df_t["Biên (%)"], name="Biên LN (%)",
                                     mode="lines+markers+text",
                                     text=[f"{v:.1f}%" for v in df_t["Biên (%)"]],
                                     textposition="top center",
                                     line=dict(color="#26c281",width=2)), row=2, col=1)
            fig.add_trace(go.Bar(x=df_t[grp_col], y=df_t["Tăng trưởng DT (%)"],
                                 name="Tăng trưởng (%)",
                                 marker_color=["#26c281" if v>=0 else "#e74c3c"
                                               for v in df_t["Tăng trưởng DT (%)"].fillna(0)],
                                 opacity=0.7), row=2, col=1)
            fig.update_layout(height=480, barmode="overlay",
                              legend=dict(orientation="h", y=-0.12))
            st.plotly_chart(fig, use_container_width=True)

            show = df_t.copy()
            show["DT"] = show["DT"].map(fmt)
            show["LN"] = show["LN"].map(fmt)
            show["Tăng trưởng DT (%)"] = show["Tăng trưởng DT (%)"].round(1)
            show.columns = [grp_col,"DT (VNĐ)","LN (VNĐ)","KL (tấn)","Số CT","Số KH","Biên (%)","Tăng trưởng (%)"]
            st.dataframe(show, use_container_width=True, hide_index=True)

    # Heatmap KH × Tháng
    st.markdown('<div class="section-title">🗓 Heatmap Doanh thu – KH × Tháng</div>', unsafe_allow_html=True)
    df_hm = (df_ban.groupby(["Tên khách hàng","Tháng"])["Thành tiền bán"]
             .sum().reset_index())
    df_pivot = df_hm.pivot(index="Tên khách hàng", columns="Tháng", values="Thành tiền bán").fillna(0)
    top_kh_hm = df_pivot.sum(axis=1).nlargest(min(20,len(df_pivot))).index
    fig_hm = px.imshow(df_pivot.loc[top_kh_hm]/1e6,
                       labels=dict(x="Tháng",y="Khách hàng",color="DT (triệu VNĐ)"),
                       color_continuous_scale="Blues", aspect="auto",
                       title="Doanh thu (triệu VNĐ) – Top KH × Tháng")
    fig_hm.update_layout(height=max(350, len(top_kh_hm)*26))
    st.plotly_chart(fig_hm, use_container_width=True)

# ══════════════════════════════════════════════════════════════
#  TAB 3 – SẢN PHẨM & THÓI QUEN
# ══════════════════════════════════════════════════════════════
with tab_sp:
    st.markdown('<div class="section-title">📦 Cơ cấu sản phẩm & thói quen mua hàng</div>', unsafe_allow_html=True)

    c1,c2 = st.columns(2)
    df_nhom = (df_ban.groupby("Nhóm SP")
               .agg(So_lan=("Nhóm SP","count"),
                    KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                    DT=("Thành tiền bán","sum"))
               .reset_index().sort_values("DT", ascending=False))
    with c1:
        fig = px.bar(df_nhom, x="Nhóm SP", y="DT", color="Nhóm SP",
                     text_auto=".3s", title="Doanh thu theo nhóm SP",
                     labels={"DT":"Doanh thu (VNĐ)","Nhóm SP":""})
        fig.update_layout(showlegend=False, height=340)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig2 = px.pie(df_nhom, names="Nhóm SP", values="So_lan",
                      title="Tỷ trọng số lần mua", hole=0.45)
        fig2.update_layout(height=340)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown('<div class="section-title">🏆 Top 15 sản phẩm</div>', unsafe_allow_html=True)
    if "Tên hàng" in df_ban.columns:
        df_top = (df_ban.groupby("Tên hàng")
                  .agg(n=("Tên hàng","count"),
                       KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                       SL=("Số lượng","sum"),
                       DT=("Thành tiền bán","sum"))
                  .reset_index().sort_values("DT", ascending=False).head(15))
        df_top["DT"] = df_top["DT"].map(fmt)
        df_top.columns = ["Sản phẩm","Số lần","KL (tấn)","SL (cái)","DT (VNĐ)"]
        st.dataframe(df_top, use_container_width=True, hide_index=True)

    # Thói quen: nhóm SP theo tháng
    st.markdown('<div class="section-title">📅 Thói quen mua nhóm SP theo thời gian</div>', unsafe_allow_html=True)
    df_sp_t = (df_ban.groupby(["Nhóm SP", time_col])["Thành tiền bán"]
               .sum().reset_index())
    fig3 = px.line(df_sp_t, x=time_col, y="Thành tiền bán",
                   color="Nhóm SP", markers=True,
                   title=f"Doanh thu nhóm SP theo {time_level}",
                   labels={"Thành tiền bán":"Doanh thu (VNĐ)"})
    fig3.update_layout(height=360)
    st.plotly_chart(fig3, use_container_width=True)

    # Mục đích sử dụng
    st.markdown('<div class="section-title">🎯 Nhận định mục đích sử dụng</div>', unsafe_allow_html=True)
    nhom_set = set(df_nhom["Nhóm SP"])
    insight_map = {
        "Ống HDPE":        ("🔵","Dự án hạ tầng kỹ thuật, cấp thoát nước công trình lớn, dự án nhà nước"),
        "Ống PVC nước":    ("🟢","Xây dựng dân dụng, công nghiệp, nông nghiệp"),
        "Ống PVC bơm cát": ("🟡","Thuỷ lợi, nông nghiệp, nuôi trồng thuỷ sản"),
        "Ống PPR":         ("🟠","Hệ thống nước nóng/lạnh nội thất, chung cư, văn phòng"),
        "Lõi PVC":         ("⚪","Đại lý / nhà sản xuất thứ cấp – mua nguyên liệu thô"),
        "Phụ kiện & Keo":  ("🔴","Tự thi công hoặc bán lại trọn gói – nhà thầu cơ điện"),
    }
    for k,(icon,desc) in insight_map.items():
        if k in nhom_set:
            st.markdown(f'<div class="risk-low">{icon} <b>{k}</b>: {desc}</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  TAB 4 – LỢI NHUẬN & CHÍNH SÁCH
# ══════════════════════════════════════════════════════════════
with tab_ln:
    st.markdown('<div class="section-title">💹 Biến động lợi nhuận & nhận diện chính sách</div>', unsafe_allow_html=True)

    df_ln = (df_ban.groupby(time_col)
             .agg(DT=("Thành tiền bán","sum"),
                  Von=("Thành tiền vốn","sum"),
                  LN=("Lợi nhuận","sum"))
             .reset_index().sort_values(time_col))
    df_ln["Biên (%)"] = (df_ln["LN"]/df_ln["DT"].replace(0,np.nan)*100).round(2).fillna(0)
    df_ln["Δ Biên"]   = df_ln["Biên (%)"].diff().round(2)

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                        subplot_titles=("DT / Vốn / LN (VNĐ)", f"Biên LN (%) theo {time_level}"),
                        vertical_spacing=0.12)
    fig.add_trace(go.Bar(x=df_ln[time_col], y=df_ln["DT"],  name="Doanh thu", marker_color="#4e79d4"), row=1, col=1)
    fig.add_trace(go.Bar(x=df_ln[time_col], y=df_ln["Von"], name="Giá vốn",   marker_color="#c0392b"), row=1, col=1)
    fig.add_trace(go.Scatter(x=df_ln[time_col], y=df_ln["LN"], name="LN",
                             mode="lines+markers", line=dict(color="#26c281",width=2)), row=1, col=1)
    fig.add_trace(go.Scatter(x=df_ln[time_col], y=df_ln["Biên (%)"], name="Biên LN (%)",
                             mode="lines+markers+text",
                             text=[f"{v:.1f}%" for v in df_ln["Biên (%)"]],
                             textposition="top center",
                             line=dict(color="#f0a500",width=2)), row=2, col=1)
    fig.update_layout(height=500, barmode="group", legend=dict(orientation="h",y=-0.12))
    st.plotly_chart(fig, use_container_width=True)

    # Phát hiện tháng bất thường
    if len(df_ln) >= 2:
        mean_b = df_ln["Biên (%)"].mean()
        std_b  = df_ln["Biên (%)"].std() or 1
        low_b  = df_ln[df_ln["Biên (%)"] < mean_b - std_b]
        high_b = df_ln[df_ln["Biên (%)"] > mean_b + std_b*1.5]

        st.markdown('<div class="section-title">🔍 Phát hiện kỳ có chính sách đặc biệt</div>', unsafe_allow_html=True)
        for _, r in low_b.iterrows():
            st.markdown(f'<div class="risk-medium">⚠️ {time_level} <b>{r[time_col]}</b>: Biên LN = <b>{r["Biên (%)"]:.1f}%</b> (TB={mean_b:.1f}%) → khả năng có <b>chiết khấu, khuyến mãi, hoặc ép giá</b></div>', unsafe_allow_html=True)
        for _, r in high_b.iterrows():
            st.markdown(f'<div class="risk-low">✅ {time_level} <b>{r[time_col]}</b>: Biên LN = <b>{r["Biên (%)"]:.1f}%</b> → cao bất thường, kiểm tra lại đơn giá</div>', unsafe_allow_html=True)
        if low_b.empty and high_b.empty:
            st.markdown('<div class="risk-low">✅ Biên LN ổn định, không phát hiện kỳ bất thường.</div>', unsafe_allow_html=True)

    # Trả hàng
    if not df_tra.empty:
        st.markdown('<div class="section-title">↩️ Đơn hàng trả lại</div>', unsafe_allow_html=True)
        cols_tra = [c for c in ["Số chứng từ","Ngày chứng từ","Tên khách hàng",
                                 "Tên hàng","Khối lượng","Thành tiền bán","Ghi chú"] if c in df_tra.columns]
        df_tra_s = df_tra[cols_tra].copy()
        if "Thành tiền bán" in df_tra_s.columns:
            df_tra_s["Thành tiền bán"] = df_tra_s["Thành tiền bán"].map(fmt)
        st.dataframe(df_tra_s, use_container_width=True, hide_index=True)
        st.error(f"Tổng trả hàng: **{fmt(abs(df_tra['Thành tiền bán'].sum()))} VNĐ**")

# ══════════════════════════════════════════════════════════════
#  TAB 5 – GIAO HÀNG
# ══════════════════════════════════════════════════════════════
with tab_gh:
    st.markdown('<div class="section-title">🚚 Hình thức & địa điểm giao hàng</div>', unsafe_allow_html=True)

    c1,c2 = st.columns(2)
    with c1:
        if "Freight Terms" in df_ban.columns:
            df_ft = df_ban["Freight Terms"].value_counts().reset_index()
            df_ft.columns = ["Hình thức","Số lần"]
            fig_ft = px.pie(df_ft, names="Hình thức", values="Số lần",
                            title="Điều kiện giao hàng", hole=0.4)
            st.plotly_chart(fig_ft, use_container_width=True)
    with c2:
        if "Shipping method" in df_ban.columns:
            df_sm = df_ban["Shipping method"].value_counts().reset_index()
            df_sm.columns = ["Phương tiện","Số lần"]
            fig_sm = px.pie(df_sm, names="Phương tiện", values="Số lần",
                            title="Phương tiện vận chuyển", hole=0.4)
            st.plotly_chart(fig_sm, use_container_width=True)

    if "Nơi giao hàng" in df_ban.columns:
        st.markdown('<div class="section-title">📍 Địa điểm giao hàng</div>', unsafe_allow_html=True)
        df_noi = (df_ban.groupby("Nơi giao hàng")
                  .agg(n=("Nơi giao hàng","count"),
                       KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                       DT=("Thành tiền bán","sum"))
                  .reset_index().sort_values("DT", ascending=False))
        df_noi["DT"] = df_noi["DT"].map(fmt)
        df_noi.columns = ["Địa điểm","Số lần","KL (tấn)","DT (VNĐ)"]
        st.dataframe(df_noi, use_container_width=True, hide_index=True)

    with st.expander("🚛 Xe & tài xế"):
        xe_cols = [c for c in ["Biển số xe","Tài Xế","Tên ĐVVC","Shipping method",
                                "Ngày chứng từ","Nơi giao hàng"] if c in df_ban.columns]
        if xe_cols:
            st.dataframe(df_ban[xe_cols].drop_duplicates().sort_values("Ngày chứng từ"),
                         use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
#  TAB 6 – TẦN SUẤT
# ══════════════════════════════════════════════════════════════
with tab_freq:
    st.markdown('<div class="section-title">🔁 Tần suất mua hàng – Heatmap SP × Tháng</div>', unsafe_allow_html=True)

    if "Tên hàng" in df_ban.columns:
        df_freq_data = df_ban.groupby(["Tên hàng","Tháng"]).size().reset_index(name="n")
        df_piv = df_freq_data.pivot(index="Tên hàng", columns="Tháng", values="n").fillna(0)
        top20  = df_piv.sum(axis=1).nlargest(min(20,len(df_piv))).index
        fig_h = px.imshow(df_piv.loc[top20],
                          labels=dict(x="Tháng",y="Sản phẩm",color="Số lần"),
                          color_continuous_scale="Blues", aspect="auto",
                          title="Top 20 SP – Tần suất mua × Tháng")
        fig_h.update_layout(height=max(380, len(top20)*24))
        st.plotly_chart(fig_h, use_container_width=True)

        st.markdown('<div class="section-title">📋 Chi tiết tần suất theo Quý (ghi rõ từng tháng)</div>', unsafe_allow_html=True)
        for qval in sorted(df_ban["Quý"].dropna().unique()):
            dq = df_ban[df_ban["Quý"]==qval]
            agg = (dq.groupby(["Tên hàng","Tháng"])
                   .agg(n=("Tên hàng","count"),
                        KL=("Khối lượng", lambda x: round(x.sum()/1000,2)),
                        DT=("Thành tiền bán","sum"))
                   .reset_index().sort_values("DT",ascending=False))
            agg["DT"] = agg["DT"].map(fmt)
            agg.columns = ["Sản phẩm","Tháng","Số lần","KL (tấn)","DT (VNĐ)"]
            with st.expander(f"📅 {qval}"):
                st.dataframe(agg, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
#  TAB 7 – RỦI RO & BẤT THƯỜNG
# ══════════════════════════════════════════════════════════════
with tab_risk:
    st.markdown('<div class="section-title">⚠️ Nhận diện rủi ro & dấu hiệu bất thường</div>', unsafe_allow_html=True)

    risks  = []  # (level, title, detail)
    score  = 0

    tong_ban_v = df_ban["Thành tiền bán"].sum()
    tong_tra_v = abs(df_tra["Thành tiền bán"].sum()) if not df_tra.empty else 0
    tong_ln_v  = df_ban["Lợi nhuận"].sum()
    bien_v     = tong_ln_v/tong_ban_v*100 if tong_ban_v else 0

    # ── R1: Trả hàng ──────────────────────────────────────────
    tl_tra = tong_tra_v/tong_ban_v*100 if tong_ban_v else 0
    if tl_tra > 15:
        score += 35
        risks.append(("high","↩️ Tỷ lệ hàng trả rất cao",
                       f"{tl_tra:.1f}% ({fmt(tong_tra_v)} VNĐ) — tranh chấp chất lượng nghiêm trọng hoặc giao sai đơn hàng"))
    elif tl_tra > 5:
        score += 20
        risks.append(("medium","↩️ Hàng trả lại đáng chú ý",
                       f"{tl_tra:.1f}% — cần xác minh nguyên nhân (chất lượng, giao sai, thay đổi dự án)"))
    else:
        risks.append(("low","✅ Tỷ lệ trả hàng thấp", f"{tl_tra:.1f}%"))

    # ── R2: Biên LN tổng ─────────────────────────────────────
    if bien_v < 0:
        score += 40
        risks.append(("high","💸 Lỗ ròng",
                       f"Biên LN = {bien_v:.1f}% — bán dưới giá vốn, có thể do chiết khấu quá lớn hoặc trả hàng"))
    elif bien_v < 5:
        score += 25
        risks.append(("high","💹 Biên LN rất thấp",
                       f"{bien_v:.1f}% — cần kiểm tra chính sách giá, chiết khấu đặc biệt"))
    elif bien_v < 15:
        score += 10
        risks.append(("medium","💹 Biên LN thấp hơn kỳ vọng", f"{bien_v:.1f}%"))
    else:
        risks.append(("low","✅ Biên LN bình thường", f"{bien_v:.1f}%"))

    # ── R3: Xuất bổ sung / thay thế ──────────────────────────
    if not df_bs.empty:
        pct_bs = len(df_bs)/len(df_ban)*100 if len(df_ban) else 0
        if pct_bs > 10:
            score += 20
            risks.append(("high","🔄 Tỷ lệ xuất bổ sung cao",
                           f"{pct_bs:.1f}% dòng là bổ sung/thay thế — nhiều đơn giao nhầm, thiếu hàng, tranh chấp"))
        else:
            score += 10
            risks.append(("medium","🔄 Có đơn xuất bổ sung/thay thế",
                           f"{len(df_bs)} dòng ({pct_bs:.1f}%) — cần theo dõi"))

    # ── R4: Đơn hàng outlier (giá trị cực lớn) ───────────────
    if len(df_ban) >= 4:
        q75 = df_ban["Thành tiền bán"].quantile(0.75)
        q25 = df_ban["Thành tiền bán"].quantile(0.25)
        iqr = q75 - q25
        if iqr > 0:
            outs = df_ban[df_ban["Thành tiền bán"] > q75 + 3*iqr]
            if not outs.empty:
                score += 15
                risks.append(("medium","📦 Đơn hàng giá trị bất thường (outlier)",
                               f"{len(outs)} đơn vượt ngưỡng thống kê — kiểm soát hạn mức tín dụng"))

    # ── R5: Biên LN âm trên từng giao dịch ───────────────────
    don_lo = df_ban[df_ban["Lợi nhuận"] < 0]
    if not don_lo.empty:
        pct_lo = len(don_lo)/len(df_ban)*100
        score += min(20, int(pct_lo))
        risks.append(("high" if pct_lo>10 else "medium",
                      "🔴 Có đơn hàng lỗ từng dòng",
                      f"{len(don_lo)} dòng ({pct_lo:.1f}%) có Lợi nhuận < 0 — bán dưới giá vốn"))

    # ── R6: Giá bán biến động lớn cùng 1 mã hàng ─────────────
    if "Giá bán" in df_ban.columns and "Tên hàng" in df_ban.columns:
        gb_cv = (df_ban[df_ban["Giá bán"]>0]
                 .groupby("Tên hàng")["Giá bán"]
                 .agg(["mean","std"])
                 .dropna())
        gb_cv["CV"] = gb_cv["std"]/gb_cv["mean"]*100
        gb_cv_high = gb_cv[gb_cv["CV"]>20]
        if not gb_cv_high.empty:
            score += 10
            sp_list = ", ".join(gb_cv_high.index[:3].tolist())
            risks.append(("medium","💲 Giá bán biến động lớn cùng mã hàng",
                           f"{len(gb_cv_high)} SP có CV>20%: {sp_list}... — thiếu nhất quán chính sách giá"))

    # ── R7: Địa điểm giao phân tán ───────────────────────────
    if "Nơi giao hàng" in df_ban.columns:
        n_noi = df_ban["Nơi giao hàng"].nunique()
        if n_noi >= 6:
            score += 10
            risks.append(("medium","📍 Giao hàng phân tán nhiều địa điểm",
                           f"{n_noi} địa điểm — nhà thầu nhiều dự án, rủi ro công nợ phân tán, khó thu hồi"))
        else:
            risks.append(("low","✅ Số địa điểm giao hàng", f"{n_noi} địa điểm (bình thường)"))

    # ── R8: Mua không đều (theo tháng) ───────────────────────
    n_thang_mua = df_ban["Tháng"].nunique()
    delta_days  = max((df_ban["Ngày chứng từ"].max()-df_ban["Ngày chứng từ"].min()).days, 1)
    n_thang_khoang = max(delta_days//30, 1)
    if n_thang_khoang >= 3 and n_thang_mua/n_thang_khoang < 0.5:
        score += 10
        risks.append(("medium","🔁 Tần suất mua không đều",
                       f"Chỉ mua {n_thang_mua}/{n_thang_khoang} tháng — phụ thuộc dự án, dòng tiền không ổn định"))
    else:
        risks.append(("low","✅ Tần suất mua hàng đều đặn",
                       f"{n_thang_mua} tháng có giao dịch"))

    # ── R9: Đơn hàng dự án (PO/B-xxx) → rủi ro thanh toán chậm
    gc_s = df["Ghi chú"].astype(str).str.upper() if "Ghi chú" in df.columns else pd.Series([""]*len(df))
    df_po = df[gc_s.str.contains(r"PO|B[0-9]{3}|HỢP ĐỒNG", regex=True, na=False)]
    if not df_po.empty:
        pct_po = len(df_po)/len(df)*100
        if pct_po > 30:
            score += 15
            risks.append(("medium","📋 Tỷ trọng đơn hàng dự án cao",
                           f"{pct_po:.0f}% đơn có PO/B-code/Hợp đồng — thanh toán thường NET 30–90 ngày, rủi ro công nợ"))
        else:
            risks.append(("low","✅ Tỷ trọng đơn dự án", f"{pct_po:.0f}%"))

    # ── R10: Khu vực khác nhau so với PKD ────────────────────
    if "Khu vực" in df_ban.columns and "Mã nhóm KH" in df_ban.columns:
        kv_unique = df_ban["Khu vực"].dropna().unique()
        if len(kv_unique) > 1:
            score += 5
            risks.append(("medium","🗺️ Giao hàng nhiều khu vực khác nhau",
                           f"{', '.join(kv_unique)} — chi phí vận chuyển cao, rủi ro giao nhầm"))

    # ── R11: Đơn hàng cuối tháng tập trung ───────────────────
    if len(df_ban) >= 10:
        df_ban["Ngày trong tháng"] = df_ban["Ngày chứng từ"].dt.day
        cuoi_thang = df_ban[df_ban["Ngày trong tháng"] >= 25]
        pct_cuoi = len(cuoi_thang)/len(df_ban)*100
        if pct_cuoi > 40:
            score += 10
            risks.append(("medium","📆 Đơn hàng tập trung cuối tháng",
                           f"{pct_cuoi:.0f}% đơn xuất trong ngày 25–31 — dấu hiệu đẩy doanh số cuối kỳ, tăng rủi ro công nợ"))

    # ── R12: Chênh lệch đơn giá vận chuyển bất thường ────────
    if "Đơn giá vận chuyển" in df_ban.columns:
        vc = df_ban[df_ban["Đơn giá vận chuyển"]>0]["Đơn giá vận chuyển"]
        if len(vc) >= 5:
            vc_mean = vc.mean()
            vc_hi   = vc[vc > vc_mean * 3]
            if not vc_hi.empty:
                score += 5
                risks.append(("medium","🚚 Đơn giá vận chuyển bất thường",
                               f"{len(vc_hi)} đơn có đơn giá VC > 3× trung bình ({vc_mean:.0f} VNĐ/kg)"))

    # ── Hiển thị điểm rủi ro ─────────────────────────────────
    if score >= 60:
        color, label = "#e74c3c", "🔴 RỦI RO CAO"
    elif score >= 30:
        color, label = "#f39c12", "🟡 RỦI RO TRUNG BÌNH"
    else:
        color, label = "#26c281", "🟢 RỦI RO THẤP"

    st.markdown(f"""
    <div style='background:#1a2035;border-radius:12px;padding:22px;text-align:center;margin:10px 0 18px 0;'>
        <div style='font-size:38px;font-weight:900;color:{color};'>{label}</div>
        <div style='font-size:16px;color:#9aa0b0;margin-top:6px;'>
            Điểm rủi ro tổng hợp: <b style='color:{color};font-size:20px;'>{score}/100</b>
        </div>
    </div>
    """, unsafe_allow_html=True)

    for lvl, title, detail in risks:
        st.markdown(
            f'<div class="risk-{lvl}"><b>{title}</b><br><span style="font-size:13px;opacity:0.85;">{detail}</span></div>',
            unsafe_allow_html=True)

    # ── BCCN ─────────────────────────────────────────────────
    st.markdown('<div class="section-title">📄 BCCN – Phân tích công nợ & thanh toán</div>', unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📋 Đơn có PO/HĐ",    df_po["Số chứng từ"].nunique() if "Số chứng từ" in df_po.columns and not df_po.empty else 0)
    c2.metric("↩️ Phiếu trả hàng", df_tra["Số chứng từ"].nunique() if "Số chứng từ" in df_tra.columns and not df_tra.empty else 0)
    c3.metric("🔄 Xuất bổ sung",   df_bs["Số chứng từ"].nunique()  if "Số chứng từ" in df_bs.columns  and not df_bs.empty  else 0)
    c4.metric("💰 GT trả hàng",    fmt(tong_tra_v)+" VNĐ")

    st.markdown("""
    <div class="info-box">
    <b>⚠️ File OM_RPT_055 không chứa ngày thanh toán thực tế.</b> Để phân tích BCCN đầy đủ cần bổ sung:<br>
    &nbsp;• <b>Sổ AR Aging</b> → số ngày tồn đọng, hạn mức tín dụng<br>
    &nbsp;• <b>Lịch sử thanh toán</b> → thói quen TT (NET 30/60/90 ngày)<br><br>
    <b>Dấu hiệu từ dữ liệu hiện có:</b><br>
    &nbsp;• <b>B-xxx / PO</b> trong ghi chú → đơn dự án → TT chậm, NET 30–90<br>
    &nbsp;• <b>Trả hàng</b> → tranh chấp → kéo dài công nợ<br>
    &nbsp;• <b>Nhiều địa điểm giao</b> → nhà thầu nhiều công trình → phân tán công nợ
    </div>
    """, unsafe_allow_html=True)

    if not df_po.empty:
        with st.expander(f"📋 Chi tiết {len(df_po)} đơn có PO / mã dự án"):
            po_cols = [c for c in ["Số chứng từ","Ngày chứng từ","Tên khách hàng",
                                    "Tên hàng","Thành tiền bán","Ghi chú"] if c in df_po.columns]
            df_po_s = df_po[po_cols].drop_duplicates().copy()
            if "Thành tiền bán" in df_po_s.columns:
                df_po_s["Thành tiền bán"] = df_po_s["Thành tiền bán"].map(fmt)
            st.dataframe(df_po_s, use_container_width=True, hide_index=True)

    if not don_lo.empty:
        with st.expander(f"🔴 Chi tiết {len(don_lo)} đơn hàng lỗ"):
            lo_cols = [c for c in ["Số chứng từ","Ngày chứng từ","Tên khách hàng",
                                    "Tên hàng","Giá bán","Giá vốn",
                                    "Thành tiền bán","Lợi nhuận","Ghi chú"] if c in don_lo.columns]
            df_lo_s = don_lo[lo_cols].copy()
            for cc in ["Thành tiền bán","Lợi nhuận","Giá bán","Giá vốn"]:
                if cc in df_lo_s.columns:
                    df_lo_s[cc] = pd.to_numeric(df_lo_s[cc], errors="coerce").map(
                        lambda x: fmt(x) if pd.notna(x) else "")
            st.dataframe(df_lo_s, use_container_width=True, hide_index=True)