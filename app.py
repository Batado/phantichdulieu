import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Phân tích hành vi khách hàng", layout="wide")

st.title("📊 PHÂN TÍCH THÓI QUEN MUA HÀNG KHÁCH HÀNG")

uploaded_file = st.file_uploader("📂 Upload file Excel", type=["xlsx"])

if uploaded_file:

    # ===== LOAD =====
    try:
        df = pd.read_excel(uploaded_file, header=13)
    except:
        df = pd.read_excel(uploaded_file)

    df = df.dropna(how='all')
    df.columns = df.columns.str.strip()

    # ===== CHECK =====
    required = ["Ngày chứng từ","Tên khách hàng","Mã hàng","Tên hàng","Nơi giao hàng","Thành tiền bán"]
    for col in required:
        if col not in df.columns:
            st.error(f"Thiếu cột: {col}")
            st.stop()

    # ===== XỬ LÝ =====
    df["Ngày chứng từ"] = pd.to_datetime(df["Ngày chứng từ"], errors="coerce")
    df["Tháng"] = df["Ngày chứng từ"].dt.to_period("M").astype(str)

    # ===== CHỌN KHÁCH =====
    kh_list = df["Tên khách hàng"].dropna().unique()
    kh = st.selectbox("🔎 Chọn khách hàng", kh_list)

    df_kh = df[df["Tên khách hàng"] == kh]

    st.markdown("---")

    # ===== KPI =====
    col1, col2, col3 = st.columns(3)

    col1.metric("💰 Tổng mua", f"{df_kh['Thành tiền bán'].sum():,.0f}")
    col2.metric("📦 Số đơn", len(df_kh))
    col3.metric("📍 Số nơi giao", df_kh["Nơi giao hàng"].nunique())

    st.markdown("---")

    # ===== 1. THỜI ĐIỂM MUA =====
    st.subheader("📅 Thời điểm mua hàng")

    time_data = df_kh.groupby("Tháng")["Thành tiền bán"].sum().reset_index()

    fig1 = px.line(time_data, x="Tháng", y="Thành tiền bán", markers=True,
                   title="Xu hướng mua hàng theo thời gian")

    st.plotly_chart(fig1, use_container_width=True)

    # ===== 2. MÃ HÀNG =====
    st.subheader("📦 Mã hàng khách mua")

    hang = df_kh.groupby(["Mã hàng","Tên hàng"])["Thành tiền bán"].sum().reset_index().sort_values(by="Thành tiền bán", ascending=False)

    fig2 = px.bar(hang.head(10), x="Mã hàng", y="Thành tiền bán",
                  title="Top mã hàng khách mua")

    st.plotly_chart(fig2, use_container_width=True)

    st.dataframe(hang)

    # ===== 3. NƠI GIAO HÀNG =====
    st.subheader("📍 Nơi giao hàng")

    dia_diem = df_kh.groupby("Nơi giao hàng")["Thành tiền bán"].sum().reset_index()

    fig3 = px.pie(dia_diem, names="Nơi giao hàng", values="Thành tiền bán",
                  title="Phân bố nơi giao hàng")

    st.plotly_chart(fig3, use_container_width=True)

    # ===== 4. KIỂM TRA THAY ĐỔI =====
    st.subheader("🔄 Phân tích thay đổi nơi giao hàng")

    dia_diem_list = df_kh["Nơi giao hàng"].unique()

    if len(dia_diem_list) == 1:
        st.success("✅ Khách hàng chỉ giao 1 địa điểm → ổn định")
    else:
        st.warning(f"⚠️ Khách hàng thay đổi {len(dia_diem_list)} nơi giao hàng")

        change_data = df_kh[["Ngày chứng từ","Nơi giao hàng"]].sort_values("Ngày chứng từ")
        st.dataframe(change_data)

    # ===== 5. THÓI QUEN =====
    st.subheader("🧠 Nhận định thói quen")

    insights = []

    # Tháng mua nhiều nhất
    top_month = time_data.sort_values(by="Thành tiền bán", ascending=False).iloc[0]
    insights.append(f"📈 Mua nhiều nhất vào: {top_month['Tháng']}")

    # Mã hàng chính
    top_product = hang.iloc[0]
    insights.append(f"🏆 Mã hàng mua nhiều nhất: {top_product['Mã hàng']}")

    # Tần suất
    freq = len(df_kh) / df_kh["Tháng"].nunique()
    insights.append(f"🔁 Tần suất mua: ~{freq:.1f} đơn/tháng")

    for i in insights:
        st.write(i)

    # ===== DATA =====
    st.subheader("📋 Dữ liệu chi tiết")
    st.dataframe(df_kh)