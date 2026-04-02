import io
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(page_title="Phân tích KH – Hoa Sen", layout="wide", page_icon="📊")

# ══════════════════════════════════════════════════════════════
#  CSS
# ══════════════════════════════════════════════════════════════
st.markdown("""
<style>
.kpi-box { background:#1e2130; border-radius:10px; padding:16px 20px; margin-bottom:8px; }
.kpi-label { color:#9aa0b0; font-size:12px; margin-bottom:4px; }
.kpi-value { color:#ffffff; font-size:22px; font-weight:700; }
.kpi-delta-pos { color:#26c281; font-size:12px; }
.kpi-delta-neg { color:#e74c3c; font-size:12px; }
.risk-high   { background:#4a1010; border-left:4px solid #e74c3c; padding:10px 14px; border-radius:6px; margin:6px 0; }
.risk-medium { background:#3d2e10; border-left:4px solid #f39c12; padding:10px 14px; border-radius:6px; margin:6px 0; }
.risk-low    { background:#0f3020; border-left:4px solid #26c281; padding:10px 14px; border-radius:6px; margin:6px 0; }
.section-title { font-size:17px; font-weight:700; color:#e0e0e0; margin:20px 0 10px 0; padding-bottom:6px; border-bottom:1px solid #2e3350; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  UPLOAD & PARSE
# ══════════════════════════════════════════════════════════════
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/6/6e/Hoa_Sen_Group_logo.svg/200px-Hoa_Sen_Group_logo.svg.png", width=140)
st.sidebar.markdown("## 📂 Upload dữ liệu")
uploaded_files = st.sidebar.file_uploader(
    "Upload file Excel báo cáo bán hàng",
    type=["xlsx"], accept_multiple_files=True
)

if not uploaded_files:
    st.markdown("## 👈 Upload file Excel để bắt đầu phân tích")
    st.info("File cần có cấu trúc báo cáo OM_RPT_055 (header tự động nhận diện).")
    st.stop()


def find_header_row(file_bytes):
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl", nrows=35)
    kws = ["khách hàng", "khach hang", "tên kh", "số chứng từ", "mã khách hàng"]
    for i in range(df_raw.shape[0]):
        row_vals = ["" if (isinstance(v, float) and pd.isna(v)) else str(v) for v in df_raw.iloc[i].tolist()]
        if any(kw in " ".join(row_vals).lower() for kw in kws):
            return i
    return 0


def parse_file(uf):
    try:
        fb = uf.read()
        hr = find_header_row(fb)
        df = pd.read_excel(io.BytesIO(fb), header=hr, engine="openpyxl")
        df.columns = df.columns.astype(str).str.strip()
        df.rename(columns={"Tên KH": "Tên khách hàng", "Tên hàng": "Tên hàng",
                            "Số xe": "Biển số xe"}, inplace=True)
        df["Nguồn file"] = uf.name
        return df
    except Exception as e:
        st.warning(f"Lỗi đọc `{uf.name}`: {e}")
        return pd.DataFrame()


@st.cache_data(show_spinner="Đang xử lý dữ liệu…")
def load_all(file_data):  # file_data: list of (name, bytes)
    frames = []
    for name, fb in file_data:
        hr = find_header_row(fb)
        df = pd.read_excel(io.BytesIO(fb), header=hr, engine="openpyxl")
        df.columns = df.columns.astype(str).str.strip()
        df.rename(columns={"Tên KH": "Tên khách hàng", "Tên hàng": "Tên hàng",
                            "Số xe": "Biển số xe"}, inplace=True)
        df["Nguồn file"] = name
        frames.append(df)

    df = pd.concat(frames, ignore_index=True)

    # Ngày
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], dayfirst=True, errors="coerce")
    df["Tháng"]  = df["Ngày chứng từ"].dt.to_period("M").astype(str)
    df["Quý"]    = df["Ngày chứng từ"].dt.to_period("Q").astype(str)
    df["Tuần"]   = df["Ngày chứng từ"].dt.isocalendar().week.astype(str)

    # Số
    for col in ["Thành tiền bán", "Thành tiền vốn", "Lợi nhuận",
                "Khối lượng", "Số lượng", "Giá bán", "Giá vốn",
                "Đơn giá vận chuyển", "Đơn giá quy đổi"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Loại giao dịch
    ghi_chu = df["Ghi chú"].astype(str).str.upper()
    df["Loại GD"] = "Xuất bán"
    df.loc[ghi_chu.str.contains("NHẬP TRẢ|TRẢ HÀNG", na=False), "Loại GD"] = "Trả hàng"
    df.loc[ghi_chu.str.contains("BỔ SUNG|THAY THẾ", na=False), "Loại GD"] = "Xuất bổ sung/thay thế"

    # Nhóm hàng đơn giản hoá
    ten_hang = df["Tên hàng"].astype(str)
    df["Nhóm sản phẩm"] = "Khác"
    df.loc[ten_hang.str.contains("HDPE|hdpe", case=False), "Nhóm sản phẩm"] = "Ống HDPE"
    df.loc[ten_hang.str.contains("PVC.*nước|nong dài|nong trơn", case=False), "Nhóm sản phẩm"] = "Ống PVC nước"
    df.loc[ten_hang.str.contains("PVC.*cát|bơm cát", case=False), "Nhóm sản phẩm"] = "Ống PVC bơm cát"
    df.loc[ten_hang.str.contains("PPR|ppr", case=False), "Nhóm sản phẩm"] = "Ống PPR"
    df.loc[ten_hang.str.contains("Lơi|Lõi|lori", case=False), "Nhóm sản phẩm"] = "Lõi PVC"
    df.loc[ten_hang.str.contains("Nối|Co|Tê|Van|Keo", case=False), "Nhóm sản phẩm"] = "Phụ kiện & Keo"

    return df


# Build cache key from file content
file_data = [(uf.name, uf.read()) for uf in uploaded_files]
df_all = load_all(file_data)

# ══════════════════════════════════════════════════════════════
#  SIDEBAR FILTERS
# ══════════════════════════════════════════════════════════════
st.sidebar.markdown("---")
st.sidebar.markdown("## 🔍 Bộ lọc")

kh_list = sorted(df_all["Tên khách hàng"].dropna().unique())
kh = st.sidebar.selectbox("👤 Khách hàng", kh_list)

quy_list = sorted(df_all["Quý"].dropna().unique())
quy_chon = st.sidebar.multiselect("📅 Quý", quy_list, default=quy_list)

df = df_all[
    (df_all["Tên khách hàng"] == kh) &
    (df_all["Quý"].isin(quy_chon))
].copy()

df_ban = df[df["Loại GD"] == "Xuất bán"]

st.markdown(f"# 📊 Phân tích khách hàng: **{kh}**")
st.markdown(f"*Dữ liệu từ: {df['Ngày chứng từ'].min().strftime('%d/%m/%Y') if not df.empty else '—'} → {df['Ngày chứng từ'].max().strftime('%d/%m/%Y') if not df.empty else '—'}*")

if df.empty:
    st.warning("Không có dữ liệu cho khách hàng và kỳ đã chọn.")
    st.stop()

# ══════════════════════════════════════════════════════════════
#  TAB LAYOUT
# ══════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📦 Thói quen & Sản phẩm",
    "📈 Doanh thu & Sản lượng",
    "💹 Lợi nhuận & Chính sách",
    "🚚 Giao hàng",
    "🔁 Tần suất mua hàng",
    "⚠️ Rủi ro KH & Đơn hàng"
])

# ══════════════════════════════════════════════════════════════
#  TAB 1 – THÓI QUEN MUA HÀNG
# ══════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-title">📦 Thói quen mua hàng theo sản phẩm</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([1.2, 1])

    # Tần suất theo nhóm sản phẩm
    with col1:
        df_nhom = (df_ban.groupby("Nhóm sản phẩm")
                   .agg(So_lan=("Số chứng từ", "count"),
                        KL_tong=("Khối lượng", "sum"),
                        DT_tong=("Thành tiền bán", "sum"))
                   .reset_index()
                   .sort_values("DT_tong", ascending=False))
        fig = px.bar(df_nhom, x="Nhóm sản phẩm", y="DT_tong", color="Nhóm sản phẩm",
                     text_auto=".3s", title="Doanh thu theo nhóm sản phẩm",
                     labels={"DT_tong": "Doanh thu (VNĐ)", "Nhóm sản phẩm": ""})
        fig.update_layout(showlegend=False, height=350)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = px.pie(df_nhom, names="Nhóm sản phẩm", values="So_lan",
                      title="Tỷ trọng lần mua theo nhóm SP",
                      hole=0.45)
        fig2.update_layout(height=350)
        st.plotly_chart(fig2, use_container_width=True)

    # Top 15 sản phẩm cụ thể
    st.markdown('<div class="section-title">🏆 Top 15 mã hàng mua nhiều nhất</div>', unsafe_allow_html=True)
    df_top = (df_ban.groupby(["Tên hàng", "Mã hàng"])
              .agg(So_lan=("Số chứng từ", "count"),
                   KL_tong=("Khối lượng", "sum"),
                   SL_tong=("Số lượng", "sum"),
                   DT_tong=("Thành tiền bán", "sum"))
              .reset_index()
              .sort_values("DT_tong", ascending=False).head(15))
    df_top.columns = ["Tên hàng", "Mã hàng", "Số lần mua", "KL (kg)", "SL (cái)", "Doanh thu (VNĐ)"]
    df_top["KL (tấn)"] = (df_top["KL (kg)"] / 1000).round(2)
    df_top["Doanh thu (VNĐ)"] = df_top["Doanh thu (VNĐ)"].map("{:,.0f}".format)
    st.dataframe(df_top[["Tên hàng", "Mã hàng", "Số lần mua", "KL (tấn)", "SL (cái)", "Doanh thu (VNĐ)"]],
                 use_container_width=True, hide_index=True)

    # Mục đích sử dụng suy luận
    st.markdown('<div class="section-title">🎯 Nhận định mục đích sử dụng</div>', unsafe_allow_html=True)
    nhom_counts = df_ban["Nhóm sản phẩm"].value_counts()
    insights = []
    if "Ống HDPE" in nhom_counts and nhom_counts["Ống HDPE"] > 0:
        insights.append("🔵 **Ống HDPE** (đường kính lớn): khả năng cao dùng cho **dự án hạ tầng kỹ thuật, cấp nước, thoát nước công trình lớn**.")
    if "Ống PVC nước" in nhom_counts:
        insights.append("🟢 **Ống PVC cấp thoát nước**: dùng cho **dự án xây dựng dân dụng, công nghiệp, nông nghiệp**.")
    if "Ống PVC bơm cát" in nhom_counts:
        insights.append("🟡 **Ống PVC bơm cát**: phục vụ **công trình thủy lợi, nông nghiệp, nuôi trồng thủy sản**.")
    if "Phụ kiện & Keo" in nhom_counts:
        insights.append("🔴 **Phụ kiện & Keo**: mua kèm → KH **tự thi công hoặc bán lại trọn gói**.")
    if "Lõi PVC" in nhom_counts:
        insights.append("⚪ **Lõi PVC**: KH có thể là **đại lý/nhà sản xuất thứ cấp**, mua nguyên liệu thô.")
    for ins in insights:
        st.markdown(f'<div class="risk-low">{ins}</div>', unsafe_allow_html=True)

    if not insights:
        st.info("Không đủ dữ liệu để suy luận mục đích sử dụng.")

# ══════════════════════════════════════════════════════════════
#  TAB 2 – DOANH THU & SẢN LƯỢNG
# ══════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">📈 Biến động doanh thu, sản lượng & khối lượng theo tháng</div>', unsafe_allow_html=True)

    df_month = (df_ban.groupby("Tháng")
                .agg(DT=("Thành tiền bán", "sum"),
                     KL_tan=("Khối lượng", lambda x: x.sum() / 1000),
                     SL=("Số lượng", "sum"),
                     So_CT=("Số chứng từ", "nunique"))
                .reset_index().sort_values("Tháng"))

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                        subplot_titles=("Doanh thu (VNĐ) & Khối lượng (tấn)", "Số lượng (cái) & Số chứng từ"),
                        vertical_spacing=0.12)

    fig.add_trace(go.Bar(x=df_month["Tháng"], y=df_month["DT"],
                         name="Doanh thu", marker_color="#4e79d4", opacity=0.8), row=1, col=1)
    fig.add_trace(go.Scatter(x=df_month["Tháng"], y=df_month["KL_tan"],
                             name="KL (tấn)", mode="lines+markers",
                             line=dict(color="#f0a500", width=2),
                             yaxis="y2"), row=1, col=1)

    fig.add_trace(go.Bar(x=df_month["Tháng"], y=df_month["SL"],
                         name="Số lượng", marker_color="#26c281", opacity=0.8), row=2, col=1)
    fig.add_trace(go.Scatter(x=df_month["Tháng"], y=df_month["So_CT"],
                             name="Số CT", mode="lines+markers",
                             line=dict(color="#e74c3c", width=2)), row=2, col=1)

    fig.update_layout(height=520, legend=dict(orientation="h", y=-0.1))
    st.plotly_chart(fig, use_container_width=True)

    # Bảng tổng hợp tháng
    df_month_show = df_month.copy()
    df_month_show["KL_tan"] = df_month_show["KL_tan"].round(2)
    df_month_show["DT"] = df_month_show["DT"].map("{:,.0f}".format)
    df_month_show.columns = ["Tháng", "Doanh thu (VNĐ)", "KL (tấn)", "SL (cái)", "Số CT"]
    st.dataframe(df_month_show, use_container_width=True, hide_index=True)

    # Theo quý
    st.markdown('<div class="section-title">📊 Tổng hợp theo Quý</div>', unsafe_allow_html=True)
    df_quy = (df_ban.groupby("Quý")
              .agg(DT=("Thành tiền bán", "sum"),
                   KL_tan=("Khối lượng", lambda x: round(x.sum() / 1000, 2)),
                   LN=("Lợi nhuận", "sum"),
                   So_CT=("Số chứng từ", "nunique"))
              .reset_index())
    df_quy["Biên LN (%)"] = (df_quy["LN"] / df_quy["DT"] * 100).round(1)
    df_quy["DT"] = df_quy["DT"].map("{:,.0f}".format)
    df_quy["LN"] = df_quy["LN"].map("{:,.0f}".format)
    df_quy.columns = ["Quý", "Doanh thu (VNĐ)", "KL (tấn)", "Lợi nhuận (VNĐ)", "Số CT", "Biên LN (%)"]
    st.dataframe(df_quy, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
#  TAB 3 – LỢI NHUẬN & CHÍNH SÁCH
# ══════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">💹 Biến động lợi nhuận & phân tích chính sách</div>', unsafe_allow_html=True)

    df_ln = (df_ban.groupby("Tháng")
             .agg(DT=("Thành tiền bán", "sum"),
                  Von=("Thành tiền vốn", "sum"),
                  LN=("Lợi nhuận", "sum"))
             .reset_index().sort_values("Tháng"))
    df_ln["Biên LN (%)"] = (df_ln["LN"] / df_ln["DT"].replace(0, float("nan")) * 100).round(2)
    df_ln["Giá bán bình quân"] = (df_ln["DT"] / df_ban.groupby("Tháng")["Khối lượng"].sum().values)

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                        subplot_titles=("Doanh thu / Vốn / Lợi nhuận (VNĐ)", "Biên lợi nhuận (%)"),
                        vertical_spacing=0.12)
    fig.add_trace(go.Bar(x=df_ln["Tháng"], y=df_ln["DT"], name="Doanh thu", marker_color="#4e79d4"), row=1, col=1)
    fig.add_trace(go.Bar(x=df_ln["Tháng"], y=df_ln["Von"], name="Giá vốn", marker_color="#e05c5c"), row=1, col=1)
    fig.add_trace(go.Scatter(x=df_ln["Tháng"], y=df_ln["LN"], name="Lợi nhuận",
                             mode="lines+markers", line=dict(color="#26c281", width=2)), row=1, col=1)
    fig.add_trace(go.Scatter(x=df_ln["Tháng"], y=df_ln["Biên LN (%)"], name="Biên LN (%)",
                             mode="lines+markers+text", text=df_ln["Biên LN (%)"].astype(str) + "%",
                             textposition="top center", line=dict(color="#f0a500", width=2)), row=2, col=1)
    fig.update_layout(height=520, barmode="group")
    st.plotly_chart(fig, use_container_width=True)

    # Phát hiện tháng biên LN thấp bất thường
    mean_bien = df_ln["Biên LN (%)"].mean()
    std_bien  = df_ln["Biên LN (%)"].std()
    anomaly   = df_ln[df_ln["Biên LN (%)"] < mean_bien - std_bien]

    st.markdown('<div class="section-title">🔍 Phát hiện tháng có chiết khấu/chính sách đặc biệt</div>', unsafe_allow_html=True)
    if not anomaly.empty:
        for _, row in anomaly.iterrows():
            st.markdown(
                f'<div class="risk-medium">⚠️ Tháng <b>{row["Tháng"]}</b>: Biên LN = <b>{row["Biên LN (%)"]:.1f}%</b> '
                f'(thấp hơn TB {mean_bien:.1f}% → có thể có chiết khấu, khuyến mãi, hoặc giá bán đặc biệt)</div>',
                unsafe_allow_html=True)
    else:
        st.markdown('<div class="risk-low">✅ Không phát hiện tháng bất thường về biên lợi nhuận.</div>', unsafe_allow_html=True)

    # Bảng đơn hàng trả lại
    df_tra = df[df["Loại GD"] == "Trả hàng"]
    if not df_tra.empty:
        st.markdown('<div class="section-title">↩️ Đơn hàng trả lại / nhập trả</div>', unsafe_allow_html=True)
        cols_show = [c for c in ["Số chứng từ", "Ngày chứng từ", "Tên hàng", "Khối lượng",
                                  "Thành tiền bán", "Lợi nhuận", "Ghi chú"] if c in df_tra.columns]
        df_tra_show = df_tra[cols_show].copy()
        df_tra_show["Thành tiền bán"] = df_tra_show["Thành tiền bán"].map("{:,.0f}".format)
        st.dataframe(df_tra_show, use_container_width=True, hide_index=True)
        total_tra = df[df["Loại GD"] == "Trả hàng"]["Thành tiền bán"].sum()
        st.error(f"Tổng giá trị hàng trả: **{total_tra:,.0f} VNĐ**")

    # Chi tiết giá bán từng dòng
    with st.expander("📋 Chi tiết giá bán / đơn giá từng giao dịch"):
        cols_price = [c for c in ["Số chứng từ", "Ngày chứng từ", "Tên hàng",
                                   "Giá bán", "Đơn giá quy đổi", "Đơn giá vận chuyển",
                                   "Thành tiền bán", "Lợi nhuận", "Ghi chú"] if c in df_ban.columns]
        st.dataframe(df_ban[cols_price].sort_values("Ngày chứng từ"), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
#  TAB 4 – GIAO HÀNG
# ══════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">🚚 Hình thức & địa điểm giao hàng</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        # Freight Terms
        df_ft = df_ban["Freight Terms"].value_counts().reset_index()
        df_ft.columns = ["Hình thức", "Số lần"]
        fig_ft = px.pie(df_ft, names="Hình thức", values="Số lần",
                        title="Điều kiện giao hàng (Freight Terms)", hole=0.4)
        st.plotly_chart(fig_ft, use_container_width=True)

    with col2:
        # Shipping method
        df_sm = df_ban["Shipping method"].value_counts().reset_index()
        df_sm.columns = ["Phương tiện", "Số lần"]
        fig_sm = px.pie(df_sm, names="Phương tiện", values="Số lần",
                        title="Phương tiện vận chuyển", hole=0.4)
        st.plotly_chart(fig_sm, use_container_width=True)

    # Nơi giao hàng
    st.markdown('<div class="section-title">📍 Địa điểm giao hàng</div>', unsafe_allow_html=True)
    df_ngh = df_ban.groupby("Nơi giao hàng").agg(
        So_lan=("Số chứng từ", "count"),
        KL_tan=("Khối lượng", lambda x: round(x.sum() / 1000, 2)),
        DT=("Thành tiền bán", "sum")
    ).reset_index().sort_values("DT", ascending=False)
    df_ngh["DT"] = df_ngh["DT"].map("{:,.0f}".format)
    df_ngh.columns = ["Nơi giao hàng", "Số lần giao", "KL (tấn)", "Doanh thu (VNĐ)"]
    st.dataframe(df_ngh, use_container_width=True, hide_index=True)

    # Chi phí vận chuyển
    if "Đơn giá vận chuyển" in df_ban.columns:
        st.markdown('<div class="section-title">💰 Chi phí vận chuyển theo tháng</div>', unsafe_allow_html=True)
        df_vc = (df_ban.groupby("Tháng")
                 .agg(CP_VC=("Đơn giá vận chuyển", lambda x: (x * df_ban.loc[x.index, "Khối lượng"]).sum() / 1000))
                 .reset_index())
        fig_vc = px.bar(df_vc, x="Tháng", y="CP_VC", text_auto=".3s",
                        labels={"CP_VC": "Chi phí VC (VNĐ/tấn)"}, title="Đơn giá vận chuyển theo tháng")
        st.plotly_chart(fig_vc, use_container_width=True)

    # Biển số xe & tài xế
    with st.expander("🚛 Danh sách xe & tài xế giao hàng"):
        cols_xe = [c for c in ["Biển số xe", "Tài Xế", "Tên ĐVVC", "Shipping method",
                                "Ngày chứng từ", "Nơi giao hàng"] if c in df_ban.columns]
        st.dataframe(df_ban[cols_xe].drop_duplicates().sort_values("Ngày chứng từ"),
                     use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
#  TAB 5 – TẦN SUẤT MUA HÀNG
# ══════════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="section-title">🔁 Tần suất mua hàng theo Quý – chi tiết từng tháng</div>', unsafe_allow_html=True)

    # Heatmap: mã hàng × tháng
    df_freq = (df_ban.groupby(["Tên hàng", "Tháng"])["Số chứng từ"]
               .count().reset_index(name="Số lần"))
    df_pivot = df_freq.pivot(index="Tên hàng", columns="Tháng", values="Số lần").fillna(0)

    # Lấy top 20 mã hàng (theo tổng)
    top20 = df_pivot.sum(axis=1).nlargest(20).index
    df_pivot_top = df_pivot.loc[top20]

    fig_heat = px.imshow(df_pivot_top,
                         labels=dict(x="Tháng", y="Tên hàng", color="Số lần"),
                         title="Heatmap tần suất mua – Top 20 mã hàng × Tháng",
                         color_continuous_scale="Blues", aspect="auto")
    fig_heat.update_layout(height=max(400, len(top20) * 22))
    st.plotly_chart(fig_heat, use_container_width=True)

    # Bảng tần suất theo quý
    st.markdown('<div class="section-title">📋 Chi tiết tần suất theo Quý & Tháng</div>', unsafe_allow_html=True)
    for quy_val in sorted(df_ban["Quý"].dropna().unique()):
        df_q = df_ban[df_ban["Quý"] == quy_val]
        st.markdown(f"**📅 {quy_val}**")
        summary = (df_q.groupby(["Tên hàng", "Tháng"])
                   .agg(Lan=("Số chứng từ", "count"),
                        KL=("Khối lượng", lambda x: round(x.sum() / 1000, 2)),
                        DT=("Thành tiền bán", "sum"))
                   .reset_index()
                   .sort_values(["DT"], ascending=False))
        summary["DT"] = summary["DT"].map("{:,.0f}".format)
        summary.columns = ["Tên hàng", "Tháng", "Số lần", "KL (tấn)", "Doanh thu (VNĐ)"]
        st.dataframe(summary, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
#  TAB 6 – RỦI RO
# ══════════════════════════════════════════════════════════════
with tab6:
    st.markdown('<div class="section-title">⚠️ Đánh giá rủi ro khách hàng & đơn hàng</div>', unsafe_allow_html=True)

    risks = []
    score = 0  # điểm rủi ro tổng hợp (0–100)

    # 1. Hàng trả lại
    df_tra = df[df["Loại GD"] == "Trả hàng"]
    tong_tra  = abs(df_tra["Thành tiền bán"].sum())
    tong_ban  = df_ban["Thành tiền bán"].sum()
    ty_le_tra = (tong_tra / tong_ban * 100) if tong_ban else 0
    if ty_le_tra > 10:
        score += 30
        risks.append(("high",   f"↩️ Tỷ lệ hàng trả lại cao: **{ty_le_tra:.1f}%** ({tong_tra:,.0f} VNĐ) — nguy cơ tranh chấp/chất lượng"))
    elif ty_le_tra > 3:
        score += 15
        risks.append(("medium", f"↩️ Hàng trả lại: **{ty_le_tra:.1f}%** — cần theo dõi"))
    else:
        risks.append(("low",    f"✅ Tỷ lệ hàng trả lại thấp: {ty_le_tra:.1f}%"))

    # 2. Biên lợi nhuận
    tong_ln  = df_ban["Lợi nhuận"].sum()
    bien_ln  = (tong_ln / tong_ban * 100) if tong_ban else 0
    if bien_ln < 5:
        score += 25
        risks.append(("high",   f"💹 Biên lợi nhuận rất thấp: **{bien_ln:.1f}%** — có thể đang bán dưới giá hoặc chiết khấu lớn"))
    elif bien_ln < 15:
        score += 10
        risks.append(("medium", f"💹 Biên lợi nhuận ở mức thấp: **{bien_ln:.1f}%**"))
    else:
        risks.append(("low",    f"✅ Biên lợi nhuận bình thường: {bien_ln:.1f}%"))

    # 3. Xuất bổ sung / thay thế
    df_bs = df[df["Loại GD"] == "Xuất bổ sung/thay thế"]
    if not df_bs.empty:
        score += 15
        risks.append(("medium", f"🔄 Có **{len(df_bs)}** dòng xuất bổ sung/thay thế — có thể do giao nhầm, giao thiếu hoặc tranh chấp chất lượng"))

    # 4. Đơn hàng lớn bất thường (outlier)
    q3   = df_ban["Thành tiền bán"].quantile(0.75)
    iqr  = df_ban["Thành tiền bán"].quantile(0.75) - df_ban["Thành tiền bán"].quantile(0.25)
    outliers = df_ban[df_ban["Thành tiền bán"] > q3 + 3 * iqr]
    if not outliers.empty:
        score += 10
        risks.append(("medium", f"📦 Có **{len(outliers)}** đơn hàng giá trị rất lớn (outlier) — cần kiểm soát hạn mức tín dụng"))

    # 5. Giao hàng nhiều địa điểm
    n_noi = df_ban["Nơi giao hàng"].nunique()
    if n_noi >= 5:
        score += 10
        risks.append(("medium", f"📍 Giao hàng tới **{n_noi} địa điểm** khác nhau — KH có thể là nhà thầu nhiều dự án song song, rủi ro thu hồi công nợ phân tán"))
    else:
        risks.append(("low",    f"✅ Số địa điểm giao hàng: {n_noi} (bình thường)"))

    # 6. Tần suất mua không đều
    thang_mua = df_ban["Tháng"].nunique()
    thang_total = max((df_ban["Ngày chứng từ"].max() - df_ban["Ngày chứng từ"].min()).days // 30, 1)
    ty_le_active = thang_mua / thang_total if thang_total else 1
    if ty_le_active < 0.5:
        score += 10
        risks.append(("medium", f"🔁 KH chỉ mua **{thang_mua}/{thang_total} tháng** — mua không đều, phụ thuộc dự án, dòng tiền không ổn định"))
    else:
        risks.append(("low",    f"✅ Mua đều đặn: {thang_mua} tháng có giao dịch"))

    # 7. BCCN – nhận diện từ ghi chú & loại đơn hàng
    st.markdown('<div class="section-title">📄 BCCN – Phân tích thanh toán & công nợ</div>', unsafe_allow_html=True)

    # Phát hiện từ khoá liên quan thanh toán trong ghi chú
    ghi_chu_all = df["Ghi chú"].astype(str).str.upper()
    po_orders = df[ghi_chu_all.str.contains("PO|PURCHASE ORDER|HỢP ĐỒNG", na=False)]
    da_tra    = df[ghi_chu_all.str.contains("ĐÃ THANH TOÁN|ĐTTK|TT", na=False)]
    tra_hang  = df[df["Loại GD"] == "Trả hàng"]

    col1, col2, col3 = st.columns(3)
    col1.metric("📋 Đơn hàng có PO/Hợp đồng", len(po_orders["Số chứng từ"].unique()) if not po_orders.empty else 0,
                help="Đơn hàng ghi nhận PO hoặc hợp đồng trong Ghi chú")
    col2.metric("↩️ Số phiếu trả hàng", len(tra_hang["Số chứng từ"].unique()) if not tra_hang.empty else 0)
    col3.metric("💰 Giá trị trả hàng (VNĐ)", f"{abs(tra_hang['Thành tiền bán'].sum()):,.0f}" if not tra_hang.empty else "0")

    bccn_html = """
    <div style='background:#1a2035;border-radius:8px;padding:16px;margin:10px 0;'>
    <b>⚠️ Lưu ý về phân tích BCCN:</b><br>
    File báo cáo bán hàng (OM_RPT_055) <b>không chứa thông tin ngày thanh toán thực tế</b> của KH.<br>
    Để đánh giá đầy đủ BCCN cần bổ sung báo cáo:<br>
    &nbsp;• <b>Sổ công nợ phải thu</b> (AR Aging) – để tính số ngày tồn đọng<br>
    &nbsp;• <b>Lịch sử thanh toán</b> – để xác định thói quen TT (30/60/90 ngày)<br>
    &nbsp;• <b>Hạn mức tín dụng</b> – để so sánh dư nợ thực tế<br><br>
    <b>Dấu hiệu nhận biết từ dữ liệu hiện có:</b><br>
    &nbsp;• Đơn hàng có ký hiệu <b>B-xxx</b> trong ghi chú → khả năng là đơn theo lô/dự án → TT theo giai đoạn<br>
    &nbsp;• Ghi chú <b>"PO"</b> → đơn hàng theo hợp đồng → thường TT chậm hơn (NET 30–90)<br>
    &nbsp;• Đơn <b>trả hàng</b> → nguy cơ tranh chấp, kéo dài công nợ<br>
    </div>
    """
    st.markdown(bccn_html, unsafe_allow_html=True)

    # Ghi chú có mã lô/PO
    df_po = df[ghi_chu_all.str.contains("PO|B[0-9]{3}", na=False)][
        [c for c in ["Số chứng từ", "Ngày chứng từ", "Thành tiền bán", "Ghi chú"] if c in df.columns]
    ].drop_duplicates()
    if not df_po.empty:
        with st.expander(f"📋 {len(df_po)} đơn hàng có PO / mã lô dự án"):
            df_po["Thành tiền bán"] = df_po["Thành tiền bán"].map("{:,.0f}".format)
            st.dataframe(df_po, use_container_width=True, hide_index=True)

    # ── Tổng điểm rủi ro ──
    st.markdown('<div class="section-title">🎯 Tổng hợp điểm rủi ro</div>', unsafe_allow_html=True)

    if score >= 50:
        level, color, label = "HIGH",   "#e74c3c", "🔴 RỦI RO CAO"
    elif score >= 25:
        level, color, label = "MEDIUM", "#f39c12", "🟡 RỦI RO TRUNG BÌNH"
    else:
        level, color, label = "LOW",    "#26c281", "🟢 RỦI RO THẤP"

    st.markdown(f"""
    <div style='background:#1a2035;border-radius:10px;padding:20px;text-align:center;margin-bottom:16px;'>
        <div style='font-size:36px;font-weight:900;color:{color};'>{label}</div>
        <div style='font-size:18px;color:#9aa0b0;margin-top:6px;'>Điểm rủi ro tổng hợp: <b style='color:{color}'>{score}/100</b></div>
    </div>
    """, unsafe_allow_html=True)

    for level_r, msg in risks:
        st.markdown(f'<div class="risk-{level_r}">{msg}</div>', unsafe_allow_html=True)