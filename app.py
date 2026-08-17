import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date

st.set_page_config(page_title="VIFEX - Lên đơn hàng", page_icon="📦", layout="centered")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

# ---------- Kết nối Google Sheets ----------
@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(creds)

@st.cache_resource
def get_spreadsheet():
    client = get_client()
    return client.open_by_key(st.secrets["SPREADSHEET_ID"])

def read_sheet(name):
    ws = get_spreadsheet().worksheet(name)
    data = ws.get_all_records()
    return pd.DataFrame(data), ws

def next_code(ws, col_index, prefix, width):
    """Sinh mã tiếp theo dạng PREFIX00001 dựa trên cột mã hiện có."""
    col = ws.col_values(col_index)[1:]  # bỏ header
    nums = []
    for v in col:
        v = str(v).replace(prefix, "").strip()
        if v.isdigit():
            nums.append(int(v))
    n = (max(nums) + 1) if nums else 1
    return f"{prefix}{str(n).zfill(width)}"

def lookup_gia(ma_sp, ngay, lich_su_gia_df):
    df = lich_su_gia_df[lich_su_gia_df["Ma_SP"] == ma_sp].copy()
    if df.empty:
        return 0
    df["Ngay_bat_dau"] = pd.to_datetime(df["Ngay_bat_dau"], errors="coerce")
    df["Ngay_ket_thuc"] = pd.to_datetime(df["Ngay_ket_thuc"], errors="coerce")
    ngay = pd.to_datetime(ngay)
    valid = df[(df["Ngay_bat_dau"] <= ngay) &
               ((df["Ngay_ket_thuc"].isna()) | (df["Ngay_ket_thuc"] >= ngay))]
    if valid.empty:
        return 0
    return int(valid.sort_values("Ngay_bat_dau").iloc[-1]["Gia_ap_dung"])

# ---------- Load dữ liệu danh mục ----------
st.title("📦 Lên đơn hàng — VIFEX")

try:
    khach_hang_df, _ = read_sheet("Khach_hang")
    san_pham_df, _ = read_sheet("San_pham")
    lich_su_gia_df, _ = read_sheet("Lich_su_gia")
    _, don_hang_ws = read_sheet("Don_hang")
    _, ctdh_ws = read_sheet("Chi_tiet_don_hang")
except Exception as e:
    st.error(f"Không kết nối được Google Sheets. Kiểm tra lại Secrets / quyền chia sẻ. Chi tiết: {e}")
    st.stop()

if khach_hang_df.empty or san_pham_df.empty:
    st.warning("Chưa có dữ liệu Khách hàng hoặc Sản phẩm — vào Sheets nhập trước đã.")
    st.stop()

# ---------- Form lên đơn ----------
with st.form("form_don_hang", clear_on_submit=False):
    ten_npp = st.selectbox("Khách hàng (NPP)", khach_hang_df["Ten_NPP"].tolist())
    kh_row = khach_hang_df[khach_hang_df["Ten_NPP"] == ten_npp].iloc[0]
    ma_kh = kh_row["Ma_KH"]
    st.caption(f"Sale phụ trách: **{kh_row['Sale_phu_trach']}**")

    ngay_len_don = st.date_input("Ngày lên đơn", value=date.today())
    hinh_thuc_tt = st.selectbox("Hình thức thanh toán", ["Tiền mặt", "Chuyển khoản", "Công nợ", "Khác"])
    ghi_chu_tt = st.text_input("Ghi chú thanh toán", "")

    st.markdown("**Danh sách sản phẩm trong đơn**")
    n_dong = st.number_input("Số dòng sản phẩm", min_value=1, max_value=20, value=1, step=1)

    line_items = []
    ten_sp_list = san_pham_df["Ten_SP"].tolist()
    for i in range(int(n_dong)):
        c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
        with c1:
            ten_sp = st.selectbox(f"Sản phẩm #{i+1}", ten_sp_list, key=f"sp_{i}")
        with c2:
            sl_dat = st.number_input("SL đặt", min_value=0, value=0, key=f"sl_{i}")
        with c3:
            tang = st.number_input("Tặng", min_value=0, value=0, key=f"tang_{i}")
        with c4:
            chiet_khau = st.number_input("Chiết khấu (đ)", min_value=0, value=0, step=10000, key=f"ck_{i}")
        line_items.append((ten_sp, sl_dat, tang, chiet_khau))

    submitted = st.form_submit_button("✅ Tạo đơn hàng", use_container_width=True)

if submitted:
    valid_items = [li for li in line_items if li[1] > 0]
    if not valid_items:
        st.error("Cần ít nhất 1 dòng sản phẩm có SL đặt > 0.")
        st.stop()

    ma_don = next_code(don_hang_ws, 1, "DH", 5)

    # Ghi Don_hang
    don_hang_ws.append_row([
        ma_don, ngay_len_don.strftime("%Y-%m-%d"), ma_kh, "",  # Sale_phu_trach để trống, đã có công thức tự dò
        "Lên đơn", hinh_thuc_tt, ghi_chu_tt, ngay_len_don.strftime("%Y-%m-%d"), "", ""
    ], value_input_option="USER_ENTERED")

    tong_tien = 0
    for ten_sp, sl_dat, tang, chiet_khau in valid_items:
        ma_sp = san_pham_df[san_pham_df["Ten_SP"] == ten_sp].iloc[0]["Ma_SP"]
        don_gia = lookup_gia(ma_sp, ngay_len_don, lich_su_gia_df)
        thanh_tien = sl_dat * don_gia - chiet_khau
        tong_tien += thanh_tien
        ma_ctdh = next_code(ctdh_ws, 1, "CT", 5)
        ctdh_ws.append_row([
            ma_ctdh, ma_don, ma_sp, sl_dat, tang, don_gia, chiet_khau,
            thanh_tien, sl_dat + tang, "", "", ""
        ], value_input_option="USER_ENTERED")

    st.success(f"Đã tạo đơn **{ma_don}** — tổng giá trị: {tong_tien:,.0f}đ")
    st.balloons()
