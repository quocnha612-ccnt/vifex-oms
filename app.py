import streamlit as st
import pandas as pd
from datetime import datetime

# Cấu hình trang
st.set_page_config(page_title="VIFEX | Lên đơn hàng", layout="wide", page_icon="📦")

# Tùy biến giao diện CSS chuẩn nhận diện VIFEX
st.markdown("""
<style>
    .header-box {
        background-color: #0E4D34;
        padding: 16px 24px;
        border-radius: 12px;
        margin-bottom: 20px;
    }
    .header-sub {
        color: #B2D8C8;
        font-size: 13px;
        font-weight: 600;
        letter-spacing: 1px;
    }
    .header-main {
        color: #FFFFFF;
        font-size: 24px;
        font-weight: bold;
        margin-top: 2px;
    }
    .total-box {
        background-color: #E2ECE6;
        padding: 15px 20px;
        border-radius: 10px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border: 1px solid #C4DAD0;
        margin-top: 25px;
    }
    .total-label {
        font-size: 15px;
        font-weight: 700;
        color: #1B4D3E;
    }
    .total-value {
        font-size: 24px;
        font-weight: 800;
        color: #C0392B;
    }
    div.stButton > button {
        border-radius: 8px;
        border: 1px solid #1B4D3E;
        color: #1B4D3E;
        background-color: #FFFFFF;
        font-weight: 600;
        width: 100%;
        padding: 8px 16px;
    }
    div.stButton > button:hover {
        background-color: #EBF3EF;
        border-color: #0E4D34;
        color: #0E4D34;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# 1. LOAD DỮ LIỆU TỪ GOOGLE SHEETS (TỰ CẬP NHẬT SAU 60 GIÂY)
# -------------------------------------------------------------
# Dán link CSV đã xuất bản của các sheet tương ứng vào đây:
URL_SAN_PHAM = "THAY_LINK_CSV_SHEET_SAN_PHAM_TAI_DAY"
URL_KHACH_HANG = "THAY_LINK_CSV_SHEET_KHACH_HANG_TAI_DAY"

@st.cache_data(ttl=60)  # Tự xóa cache và cập nhật sau mỗi 60 giây
def load_data():
    try:
        df_sp = pd.read_csv(URL_SAN_PHAM)
        df_sp.columns = [c.strip() for c in df_sp.columns]
        
        # LỌC SẢN PHẨM: Loại bỏ khoảng trắng và chỉ lấy các món "Đang hoạt động"
        if 'Trang_thai' in df_sp.columns:
            df_sp['Trang_thai'] = df_sp['Trang_thai'].astype(str).str.strip()
            df_sp = df_sp[df_sp['Trang_thai'] == 'Đang hoạt động']
            
        try:
            df_kh = pd.read_csv(URL_KHACH_HANG)
            df_kh.columns = [c.strip() for c in df_kh.columns]
        except Exception:
            df_kh = pd.DataFrame({'Ten_KH': ['MINH PHÚC - NAM ĐỊNH', 'LAN THÀNH', 'NPP KHÁC']})
            
        return df_sp, df_kh
    except Exception as e:
        st.error(f"Lỗi kết nối dữ liệu: {e}")
        return pd.DataFrame(), pd.DataFrame()

df_sanpham, df_khachhang = load_data()

# -------------------------------------------------------------
# 2. HEADER
# -------------------------------------------------------------
st.markdown("""
<div class="header-box">
    <div class="header-sub">▌ VIFEX</div>
    <div class="header-main">Lên đơn hàng</div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# 3. CHỌN KHÁCH HÀNG & SALE
# -------------------------------------------------------------
col_kh, col_sale = st.columns([3, 1])

with col_kh:
    kh_options = df_khachhang['Ten_KH'].dropna().unique().tolist() if not df_khachhang.empty and 'Ten_KH' in df_khachhang.columns else ['MINH PHÚC - NAM ĐỊNH']
    khach_hang_chon = st.selectbox("Khách hàng (NPP)", options=kh_options)

with col_sale:
    sale_options = ["NV001", "NV002", "NV003"]
    sale_phu_trach = st.selectbox("Sale phụ trách", options=sale_options)

st.divider()

# -------------------------------------------------------------
# 4. DANH SÁCH MẶT HÀNG (ĐÃ LỌC BỎ CÁC HÀNG NGỪNG BÁN)
# -------------------------------------------------------------
if "items_count" not in st.session_state:
    st.session_state.items_count = 1

if not df_sanpham.empty and 'Ten_SP' in df_sanpham.columns:
    sp_options = df_sanpham['Ten_SP'].dropna().tolist()
else:
    sp_options = ["Không có sản phẩm nào khả dụng"]

tong_tien = 0

for i in range(st.session_state.items_count):
    st.markdown(f"**Sản phẩm #{i+1}**")
    sp_chon = st.selectbox(
        label=f"sp_{i}",
        options=sp_options,
        key=f"sp_select_{i}",
        label_visibility="collapsed"
    )
    
    c1, c2, c3 = st.columns(3)
    with c1:
        sl = st.number_input("SL đặt", min_value=0, value=0, step=1, key=f"sl_{i}")
    with c2:
        tang = st.number_input("Tặng", min_value=0, value=0, step=1, key=f"tang_{i}")
    with c3:
        ck = st.number_input("CK (đ)", min_value=0, value=0, step=1000, key=f"ck_{i}")
    
    don_gia = 420000
    thanh_tien = max(0, (sl * don_gia) - ck)
    tong_tien += thanh_tien
    st.write("")

if st.button("➕ Thêm sản phẩm"):
    st.session_state.items_count += 1
    st.rerun()

# -------------------------------------------------------------
# 5. TỔNG TIỀN VÀ XÁC NHẬN
# -------------------------------------------------------------
st.markdown(f"""
<div class="total-box">
    <div class="total-label">TỔNG GIÁ TRỊ ĐƠN HÀNG (DỰ KIẾN):</div>
    <div class="total-value">{tong_tien:,.0f}đ</div>
</div>
""", unsafe_allow_html=True)

st.write("")
if st.button("🚀 Xác nhận tạo đơn hàng", use_container_width=True):
    st.success("Đã ghi nhận đơn hàng thành công!")
