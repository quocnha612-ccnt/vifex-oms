import streamlit as st
import pandas as pd
from datetime import datetime

# Cấu hình trang
st.set_page_config(page_title="VIFEX | Lên đơn hàng", layout="wide", page_icon="📦")

# Tùy biến giao diện CSS chuẩn màu VIFEX
st.markdown("""
<style>
    .header-box {
        background-color: #0E4D34;
        padding: 18px 25px;
        border-radius: 12px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 25px;
    }
    .header-title-sub {
        color: #E0E0E0;
        font-size: 14px;
        font-weight: 600;
        letter-spacing: 1px;
        margin-bottom: 2px;
    }
    .header-title-main {
        color: #FFFFFF;
        font-size: 26px;
        font-weight: bold;
    }
    .total-box {
        background-color: #E2ECE6;
        padding: 16px 20px;
        border-radius: 10px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border: 1px solid #C4DAD0;
        margin-top: 20px;
    }
    .total-label {
        font-size: 15px;
        font-weight: 700;
        color: #1B4D3E;
        letter-spacing: 0.5px;
    }
    .total-value {
        font-size: 22px;
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
# 1. LOAD DỮ LIỆU TỪ GOOGLE SHEETS
# -------------------------------------------------------------
# Thay đường link CSV tương ứng từ Google Sheet (VIFEX_Database_Goc_n1)
URL_SAN_PHAM = "https://docs.google.com/spreadsheets/d/e/2PACX-.../pub?gid=...&single=true&output=csv"
URL_KHACH_HANG = "https://docs.google.com/spreadsheets/d/e/2PACX-.../pub?gid=...&single=true&output=csv"
URL_GIA = "https://docs.google.com/spreadsheets/d/e/2PACX-.../pub?gid=...&single=true&output=csv"

@st.cache_data(ttl=60)  # Tự động làm mới cache sau 60 giây
def load_data():
    try:
        # Đọc dữ liệu sản phẩm
        df_sp = pd.read_csv(URL_SAN_PHAM)
        
        # LÀM SẠCH KHOẢNG TRẮNG CÁC CỘT
        df_sp.columns = [c.strip() for c in df_sp.columns]
        if 'Trang_thai' in df_sp.columns:
            df_sp['Trang_thai'] = df_sp['Trang_thai'].astype(str).str.strip()
            # QUAN TRỌNG: CHỈ LỌC CÁC SẢN PHẨM ĐANG HOẠT ĐỘNG
            df_sp = df_sp[df_sp['Trang_thai'] == 'Đang hoạt động']
        
        # Đọc danh sách khách hàng
        try:
            df_kh = pd.read_csv(URL_KHACH_HANG)
            df_kh.columns = [c.strip() for c in df_kh.columns]
        except Exception:
            df_kh = pd.DataFrame({'Ten_KH': ['MINH PHÚC - NAM ĐỊNH', 'LAN THÀNH', 'NPP KHÁC']})
            
        return df_sp, df_kh
    except Exception as e:
        st.error(f"Lỗi khi tải dữ liệu: {e}")
        return pd.DataFrame(), pd.DataFrame()

df_sanpham, df_khachhang = load_data()

# -------------------------------------------------------------
# 2. GIAO DIỆN HEADER VIFEX
# -------------------------------------------------------------
st.markdown("""
<div class="header-box">
    <div>
        <div class="header-title-sub">▌ VIFEX</div>
        <div class="header-title-main">Lên đơn hàng</div>
    </div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# 3. THÔNG TIN CHUNG ĐƠN HÀNG
# -------------------------------------------------------------
col_kh, col_sale = st.columns([3, 1])

with col_kh:
    kh_list = df_khachhang['Ten_KH'].dropna().unique().tolist() if not df_khachhang.empty and 'Ten_KH' in df_khachhang.columns else ['MINH PHÚC - NAM ĐỊNH']
    khach_hang_chon = st.selectbox("Khách hàng (NPP)", options=kh_list)

with col_sale:
    sale_list = ["NV001", "NV002", "NV003"]
    sale_phu_trach = st.selectbox("Sale phụ trách", options=sale_list)

st.markdown("<hr style='margin: 15px 0; border: none; border-top: 1px solid #e0e0e0;'>", unsafe_allow_html=True)

# -------------------------------------------------------------
# 4. DANH SÁCH MẶT HÀNG TRONG ĐƠN
# -------------------------------------------------------------
if "items_count" not in st.session_state:
    st.session_state.items_count = 1

# Danh sách tên sản phẩm (ĐÃ ĐƯỢC LỌC CHỈ CÒN SẢN PHẨM ĐANG HOẠT ĐỘNG)
if not df_sanpham.empty and 'Ten_SP' in df_sanpham.columns:
    sp_options = df_sanpham['Ten_SP'].dropna().tolist()
else:
    sp_options = ["Không có sản phẩm nào khả dụng"]

tong_tien_don_hang = 0

for i in range(st.session_state.items_count):
    st.markdown(f"**Sản phẩm #{i+1}**")
    sp_chon = st.selectbox(
        label="Tên sản phẩm",
        options=sp_options,
        key=f"sp_{i}",
        label_visibility="collapsed"
    )
    
    c1, c2, c3 = st.columns(3)
    with c1:
        sl_dat = st.number_input("SL đặt", min_value=0, value=0, step=1, key=f"sl_{i}")
    with c2:
        tang = st.number_input("Tặng", min_value=0, value=0, step=1, key=f"tang_{i}")
    with c3:
        ck = st.number_input("CK (đ)", min_value=0, value=0, step=1000, key=f"ck_{i}")
    
    # Giả định đơn giá mẫu (có thể lookup từ bảng Lich_su_gia)
    don_gia = 420000 
    thanh_tien = max(0, (sl_dat * don_gia) - ck)
    tong_tien_don_hang += thanh_tien
    st.write("")

# Nút thêm dòng sản phẩm
if st.button("➕ Thêm sản phẩm"):
    st.session_state.items_count += 1
    st.rerun()

# -------------------------------------------------------------
# 5. TỔNG GIÁ TRỊ & NÚT XÁC NHẬN
# -------------------------------------------------------------
st.markdown(f"""
<div class="total-box">
    <div class="total-label">TỔNG GIÁ TRỊ ĐƠN HÀNG (DỰ KIẾN):</div>
    <div class="total-value">{tong_tien_don_hang:,.0f}đ</div>
</div>
""", unsafe_allow_html=True)

st.write("")
if st.button("🚀 Xác nhận tạo đơn hàng", use_container_width=True):
    st.success("Đã ghi nhận đơn hàng thành công!")
