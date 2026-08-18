import io
import os
from datetime import date, datetime

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from PIL import Image, ImageDraw, ImageFont

st.set_page_config(page_title="VIFEX - Quản lý đơn hàng", page_icon="📦", layout="centered")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

GREEN = "#15503F"
GREEN_BG = "#E2EDE8"
RED = "#D92B2B"
RED_BG = "#FBE6E6"
AMBER = "#8A5A10"
AMBER_BG = "#FBEEDA"

VALID_STATUSES = ["Đang giao hàng", "Đã nhận hàng"]
ALL_STATUSES = ["Lên đơn", "Gửi kho", "Đang giao hàng", "Đã nhận hàng"]

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
st.markdown(f"""
<style>
.vifex-banner {{
    background: {GREEN};
    color: #fff;
    padding: 16px 20px;
    border-radius: 12px;
    margin-bottom: 18px;
}}
.vifex-banner .brand {{ font-size: 13px; font-weight: 700; letter-spacing: 1px; opacity: 0.85; }}
.vifex-banner .title {{ font-size: 20px; font-weight: 700; margin-top: 2px; }}

div.stButton > button {{
    background-color: {RED} !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
}}
div.stButton > button:hover {{ background-color: #b62222 !important; }}
button[kind="secondary"] {{
    background-color: {GREEN_BG} !important;
    color: {GREEN} !important;
}}

.badge {{
    display: inline-block; padding: 4px 12px; border-radius: 8px;
    font-size: 13px; font-weight: 600; margin-right: 6px;
}}
.order-card {{
    border: 1px solid #e2e0d5; border-radius: 10px; padding: 12px 14px;
    margin-bottom: 8px; background: #fafaf6;
}}
.order-code {{ font-weight: 700; font-size: 14px; }}
.order-cust {{ font-size: 12px; color: #6c6f68; }}
.order-value {{ font-weight: 700; color: {GREEN}; font-size: 14px; margin-top: 2px; }}
.metric-box {{
    background: #fafaf6; border: 1px solid #e2e0d5; border-radius: 10px;
    padding: 12px 14px; text-align: left;
}}
.metric-label {{ font-size: 12px; color: #6c6f68; }}
.metric-value {{ font-size: 20px; font-weight: 700; margin-top: 2px; }}
</style>
""", unsafe_allow_html=True)


def banner(title):
    st.markdown(f"""
    <div class="vifex-banner">
        <div class="brand">VIFEX</div>
        <div class="title">{title}</div>
    </div>
    """, unsafe_allow_html=True)


def status_badge_html(status):
    colors = {
        "Lên đơn": (AMBER, AMBER_BG),
        "Gửi kho": (AMBER, AMBER_BG),
        "Đang giao hàng": (RED, RED_BG),
        "Đã nhận hàng": (GREEN, GREEN_BG),
    }
    fg, bg = colors.get(status, ("#555", "#eee"))
    return f'<span class="badge" style="background:{bg};color:{fg}">{status}</span>'


def money(v):
    try:
        return f"{float(v):,.0f}đ".replace(",", ".")
    except (TypeError, ValueError):
        return "0đ"


# ---------------------------------------------------------------------------
# Kết nối Google Sheets
# ---------------------------------------------------------------------------
@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(creds)


@st.cache_resource
def get_spreadsheet():
    return get_client().open_by_key(st.secrets["SPREADSHEET_ID"])


def serial_to_date(v):
    """Chuyển serial ngày của Google Sheets sang date. Trả về None nếu không hợp lệ."""
    if v is None or v == "":
        return None
    try:
        if isinstance(v, str):
            return pd.to_datetime(v, dayfirst=False, errors="coerce").date()
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))).date()
    except Exception:
        return None


def read_raw(name):
    """Đọc 1 sheet với UNFORMATTED_VALUE (dùng cho ghi/tra cứu đơn lẻ, không dùng cho load chính)."""
    ws = get_spreadsheet().worksheet(name)
    records = ws.get_all_records(value_render_option="UNFORMATTED_VALUE")
    if records:
        df = pd.DataFrame(records)
    else:
        headers = ws.row_values(1)
        df = pd.DataFrame(columns=headers)
    return df, ws


SHEET_NAMES = ["San_pham", "Khach_hang", "Nhan_vien", "Lich_su_gia", "Don_hang", "Chi_tiet_don_hang"]


def read_all_sheets_batch(names):
    """Đọc nhiều sheet trong 1 lần gọi API duy nhất — nhanh hơn nhiều so với gọi từng sheet."""
    sh = get_spreadsheet()
    ranges = [f"{n}!A:Z" for n in names]
    resp = sh.values_batch_get(ranges, params={"valueRenderOption": "UNFORMATTED_VALUE"})
    result = {}
    for n, vr in zip(names, resp.get("valueRanges", [])):
        values = vr.get("values", [])
        if not values:
            result[n] = pd.DataFrame()
            continue
        header = values[0]
        rows = values[1:]
        rows = [r + [None] * (len(header) - len(r)) for r in rows]
        df = pd.DataFrame(rows, columns=header)
        df = df.replace("", None)
        result[n] = df
    return result


def to_date_col(df, col):
    if col in df.columns:
        df[col] = df[col].apply(serial_to_date)
    return df


@st.cache_data(ttl=60)
def load_data():
    raw = read_all_sheets_batch(SHEET_NAMES)
    san_pham_df = raw["San_pham"]
    khach_hang_df = raw["Khach_hang"]
    nhan_vien_df = raw["Nhan_vien"]
    lich_su_gia_df = raw["Lich_su_gia"]
    don_hang_df = raw["Don_hang"]
    ctdh_df = raw["Chi_tiet_don_hang"]

    lich_su_gia_df = to_date_col(lich_su_gia_df, "Ngay_bat_dau")
    lich_su_gia_df = to_date_col(lich_su_gia_df, "Ngay_ket_thuc")
    don_hang_df = to_date_col(don_hang_df, "Ngay_len_don")

    # ép các cột số về đúng kiểu số (batch_get trả nguyên dạng, không tự ép kiểu như get_all_records)
    for col in ["SL_dat", "Tang", "Don_gia_ap_dung", "Chiet_khau", "Thanh_tien", "San_luong_xuat_kho"]:
        if col in ctdh_df.columns:
            ctdh_df[col] = pd.to_numeric(ctdh_df[col], errors="coerce").fillna(0)

    return {
        "san_pham": san_pham_df,
        "khach_hang": khach_hang_df,
        "nhan_vien": nhan_vien_df,
        "lich_su_gia": lich_su_gia_df,
        "don_hang": don_hang_df,
        "ctdh": ctdh_df,
    }


def get_ws(name):
    return get_spreadsheet().worksheet(name)


def refresh():
    load_data.clear()


def next_code(ws, col_index, prefix, width):
    col = ws.col_values(col_index)[1:]
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
    valid = df[(df["Ngay_bat_dau"].notna()) & (df["Ngay_bat_dau"] <= ngay) &
               (df["Ngay_ket_thuc"].isna() | (df["Ngay_ket_thuc"] >= ngay))]
    if valid.empty:
        return 0
    valid = valid.sort_values("Ngay_bat_dau")
    return float(valid.iloc[-1]["Gia_ap_dung"])


def find_row_by_code(ws, code, col_index=1):
    values = ws.col_values(col_index)
    for i, v in enumerate(values, start=1):
        if v == code:
            return i
    return None


def update_order_status(ma_don, new_status):
    today = date.today().strftime("%Y-%m-%d")
    don_hang_ws = get_ws("Don_hang")
    row = find_row_by_code(don_hang_ws, ma_don, col_index=1)
    if row:
        don_hang_ws.update_cell(row, 5, new_status)          # E: Trang_thai
        don_hang_ws.update_cell(row, 8, today)                # H: Ngay_cap_nhat_trang_thai
    # đồng bộ trạng thái sang Chi_tiet_don_hang để dữ liệu Sheets nhất quán
    ctdh_ws = get_ws("Chi_tiet_don_hang")
    ma_don_col = ctdh_ws.col_values(2)
    for i, v in enumerate(ma_don_col, start=1):
        if v == ma_don:
            ctdh_ws.update_cell(i, 10, new_status)            # J: Trang_thai_don
    refresh()


# ---------------------------------------------------------------------------
# Ảnh phiếu xuất đơn hàng
# ---------------------------------------------------------------------------
@st.cache_resource
def get_font(size):
    base = os.path.dirname(__file__)
    candidates = [
        os.path.join(base, "fonts", "NotoSans.ttf"),
        os.path.join(base, "NotoSans.ttf"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    raise FileNotFoundError(
        "Không tìm thấy font NotoSans.ttf. Cần upload file này vào repo (thư mục fonts/ hoặc gốc repo)."
    )


def draw_bold(draw, pos, text, font, fill):
    x, y = pos
    draw.text((x, y), text, font=font, fill=fill)
    draw.text((x + 0.6, y), text, font=font, fill=fill)


def generate_order_slip(ma_don, order_row, items_df, khach_hang_row):
    W = 800
    row_h = 34
    header_h = 260
    footer_h = 90
    H = header_h + row_h * (len(items_df) + 2) + footer_h

    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    f_title = get_font(26)
    f_sub = get_font(14)
    f_h = get_font(16)
    f_n = get_font(15)

    d.rectangle([0, 0, W, 80], fill=GREEN)
    draw_bold(d, (24, 16), "VIFEX", f_title, "white")
    d.text((24, 50), "PHIẾU XUẤT ĐƠN HÀNG (không phải hóa đơn VAT)", font=f_sub, fill="white")

    y = 100
    ten_npp = khach_hang_row.get("Ten_NPP", "")
    dia_chi = khach_hang_row.get("Dia_chi_giao_phu") or khach_hang_row.get("Dia_chi_GPKD", "")
    sdt = khach_hang_row.get("SDT_phu", "")

    draw_bold(d, (24, y), f"Mã đơn: {ma_don}", f_h, "black"); y += 26
    d.text((24, y), f"Ngày lên đơn: {order_row.get('Ngay_len_don')}", font=f_n, fill="black"); y += 24
    d.text((24, y), f"Khách hàng: {ten_npp}", font=f_n, fill="black"); y += 24
    if dia_chi:
        d.text((24, y), f"Địa chỉ giao: {dia_chi}", font=f_n, fill="black"); y += 24
    if sdt:
        d.text((24, y), f"SĐT: {sdt}", font=f_n, fill="black"); y += 24
    y += 10

    d.line([24, y, W - 24, y], fill="#ddd", width=1); y += 10
    cols_x = [24, 380, 470, 560, 670]
    headers = ["Sản phẩm", "SL đặt", "Tặng", "Đơn giá", "Thành tiền"]
    for x, h in zip(cols_x, headers):
        draw_bold(d, (x, y), h, f_n, GREEN)
    y += row_h
    d.line([24, y - 6, W - 24, y - 6], fill="#ddd", width=1)

    total = 0
    for _, r in items_df.iterrows():
        total += r["Thanh_tien"]
        d.text((cols_x[0], y), str(r["Ten_SP"])[:38], font=f_n, fill="black")
        d.text((cols_x[1], y), str(int(r["SL_dat"])), font=f_n, fill="black")
        d.text((cols_x[2], y), str(int(r["Tang"])), font=f_n, fill="black")
        d.text((cols_x[3], y), money(r["Don_gia_ap_dung"]), font=f_n, fill="black")
        d.text((cols_x[4], y), money(r["Thanh_tien"]), font=f_n, fill="black")
        y += row_h

    d.line([24, y, W - 24, y], fill="#ddd", width=1); y += 14
    draw_bold(d, (cols_x[3], y), "TỔNG CỘNG:", f_h, RED)
    draw_bold(d, (cols_x[4], y), money(total), f_h, RED)
    y += 34

    d.text((24, y), f"Hình thức thanh toán: {order_row.get('Hinh_thuc_thanh_toan', '')}", font=f_n, fill="black")
    y += 22
    ghi_chu = order_row.get("Ghi_chu_thanh_toan", "")
    if ghi_chu:
        d.text((24, y), f"Ghi chú: {ghi_chu}", font=f_n, fill="black")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Dữ liệu chung
# ---------------------------------------------------------------------------
data = load_data()
san_pham_df = data["san_pham"]
khach_hang_df = data["khach_hang"]
lich_su_gia_df = data["lich_su_gia"]
don_hang_df = data["don_hang"]
ctdh_df = data["ctdh"]

if khach_hang_df.empty or san_pham_df.empty:
    st.warning("Chưa có dữ liệu Khách hàng hoặc Sản phẩm. Vào Google Sheets nhập trước đã.")
    st.stop()

# ghép Chi_tiet_don_hang với Don_hang để tính toán (không phụ thuộc công thức trong Sheets)
merged = pd.DataFrame()
if not ctdh_df.empty and not don_hang_df.empty:
    merged = ctdh_df.merge(
        don_hang_df[["Ma_don", "Ngay_len_don", "Ma_KH", "Sale_phu_trach", "Trang_thai"]],
        on="Ma_don", how="left"
    )

if "selected_order" not in st.session_state:
    st.session_state.selected_order = None
if "nav" not in st.session_state:
    st.session_state.nav = "🏠 Trang chủ"

# ---------------------------------------------------------------------------
# Điều hướng — dùng nút bấm (tap target lớn hơn radio, bấm chính xác hơn)
# ---------------------------------------------------------------------------
st.markdown(f"""
<style>
.nav-row div.stButton > button {{
    width: 100%;
    padding: 10px 4px;
    font-size: 13px;
    border-radius: 10px !important;
}}
.nav-row div.stButton > button[kind="secondary"] {{
    background-color: #fafaf6 !important;
    color: #444 !important;
    border: 1px solid #e2e0d5 !important;
}}
.nav-row div.stButton > button[kind="primary"] {{
    background-color: {GREEN} !important;
    color: #fff !important;
    border: none !important;
}}
</style>
""", unsafe_allow_html=True)

NAV_OPTIONS = ["🏠 Trang chủ", "📦 Đơn hàng", "➕ Lên đơn", "📊 Dashboard", "👥 Khách hàng"]
st.markdown('<div class="nav-row">', unsafe_allow_html=True)
nav_cols = st.columns(len(NAV_OPTIONS))
for col, opt in zip(nav_cols, NAV_OPTIONS):
    is_active = st.session_state.nav == opt
    if col.button(opt, key=f"nav_{opt}", type="primary" if is_active else "secondary",
                  use_container_width=True):
        st.session_state.nav = opt
        st.rerun()
st.markdown('</div>', unsafe_allow_html=True)
nav = st.session_state.nav
st.divider()


def order_total(ma_don):
    if merged.empty:
        return 0
    return merged.loc[merged["Ma_don"] == ma_don, "Thanh_tien"].sum()


def render_order_detail(ma_don):
    order_rows = don_hang_df[don_hang_df["Ma_don"] == ma_don]
    if order_rows.empty:
        st.error("Không tìm thấy đơn hàng.")
        return
    order_row = order_rows.iloc[0]
    kh_rows = khach_hang_df[khach_hang_df["Ma_KH"] == order_row["Ma_KH"]]
    kh_row = kh_rows.iloc[0] if not kh_rows.empty else {}

    items = ctdh_df[ctdh_df["Ma_don"] == ma_don].copy()
    items = items.merge(san_pham_df[["Ma_SP", "Ten_SP"]], on="Ma_SP", how="left")

    st.markdown(f"### Chi tiết đơn {ma_don}")
    st.markdown(status_badge_html(order_row["Trang_thai"]), unsafe_allow_html=True)
    st.write(f"**Khách hàng:** {kh_row.get('Ten_NPP', '')}")
    st.write(f"**Ngày lên đơn:** {order_row['Ngay_len_don']}")
    st.write(f"**Hình thức thanh toán:** {order_row['Hinh_thuc_thanh_toan']}")
    if order_row.get("Ghi_chu_thanh_toan"):
        st.write(f"**Ghi chú thanh toán:** {order_row['Ghi_chu_thanh_toan']}")

    show_cols = items[["Ten_SP", "SL_dat", "Tang", "Don_gia_ap_dung", "Chiet_khau", "Thanh_tien"]].copy()
    show_cols["Don_gia_ap_dung"] = show_cols["Don_gia_ap_dung"].apply(money)
    show_cols["Chiet_khau"] = show_cols["Chiet_khau"].apply(money)
    show_cols["Thanh_tien"] = show_cols["Thanh_tien"].apply(money)
    show_cols.columns = ["Sản phẩm", "SL đặt", "Tặng", "Đơn giá", "Chiết khấu", "Thành tiền"]
    st.dataframe(show_cols, hide_index=True, use_container_width=True)

    total = items["Thanh_tien"].sum()
    st.markdown(f"**Tổng cộng: {money(total)}**")

    c1, c2 = st.columns(2)
    with c1:
        new_status = st.selectbox("Cập nhật trạng thái", ALL_STATUSES,
                                   index=ALL_STATUSES.index(order_row["Trang_thai"])
                                   if order_row["Trang_thai"] in ALL_STATUSES else 0,
                                   key=f"status_{ma_don}")
        if st.button("Lưu trạng thái", key=f"save_status_{ma_don}"):
            update_order_status(ma_don, new_status)
            st.success("Đã cập nhật trạng thái.")
            st.rerun()
    with c2:
        png_bytes = generate_order_slip(ma_don, order_row, items, kh_row if isinstance(kh_row, dict) else kh_row.to_dict())
        st.download_button("📥 Tải phiếu xuất đơn hàng", data=png_bytes,
                            file_name=f"{ma_don}_phieu_xuat.png", mime="image/png",
                            key=f"dl_{ma_don}")

    if st.button("← Đóng chi tiết", key=f"close_{ma_don}"):
        st.session_state.selected_order = None
        st.rerun()


# ---------------------------------------------------------------------------
# Trang chủ
# ---------------------------------------------------------------------------
if nav == "🏠 Trang chủ":
    banner("Trang chủ")

    if don_hang_df.empty:
        st.info("Chưa có đơn hàng nào.")
    else:
        counts = don_hang_df["Trang_thai"].value_counts()
        c1, c2 = st.columns(2)
        c1.markdown(f"""<div class="metric-box"><div class="metric-label">Gửi kho</div>
            <div class="metric-value" style="color:{AMBER}">{int(counts.get('Gửi kho', 0))}</div></div>""",
            unsafe_allow_html=True)
        c2.markdown(f"""<div class="metric-box"><div class="metric-label">Đang giao</div>
            <div class="metric-value" style="color:{RED}">{int(counts.get('Đang giao hàng', 0))}</div></div>""",
            unsafe_allow_html=True)

        st.markdown("#### Đơn cần xử lý")
        pending = don_hang_df[don_hang_df["Trang_thai"].isin(["Gửi kho", "Đang giao hàng"])]
        pending = pending.sort_values("Ngay_len_don", ascending=False).head(5)
        if pending.empty:
            st.caption("Không có đơn nào đang chờ xử lý.")
        for _, r in pending.iterrows():
            kh = khach_hang_df[khach_hang_df["Ma_KH"] == r["Ma_KH"]]
            ten_kh = kh.iloc[0]["Ten_NPP"] if not kh.empty else r["Ma_KH"]
            st.markdown(f"""
            <div class="order-card">
                <div class="order-code">{r['Ma_don']}</div>
                <div class="order-cust">{ten_kh}</div>
                <div class="order-value">{money(order_total(r['Ma_don']))}</div>
                {status_badge_html(r['Trang_thai'])}
            </div>""", unsafe_allow_html=True)
            if st.button("Xem chi tiết", key=f"home_view_{r['Ma_don']}"):
                st.session_state.selected_order = r["Ma_don"]
                st.rerun()

    if st.button("+ Tạo đơn mới", key="home_new_order"):
        st.session_state.nav = "➕ Lên đơn"
        st.rerun()

    if st.session_state.selected_order:
        st.divider()
        render_order_detail(st.session_state.selected_order)

# ---------------------------------------------------------------------------
# Đơn hàng
# ---------------------------------------------------------------------------
elif nav == "📦 Đơn hàng":
    banner("Danh sách đơn hàng")

    filter_status = st.selectbox("Lọc theo trạng thái", ["Tất cả"] + ALL_STATUSES)
    view_df = don_hang_df.copy()
    if filter_status != "Tất cả":
        view_df = view_df[view_df["Trang_thai"] == filter_status]
    view_df = view_df.sort_values("Ngay_len_don", ascending=False)

    if view_df.empty:
        st.info("Không có đơn hàng nào phù hợp.")
    for _, r in view_df.iterrows():
        kh = khach_hang_df[khach_hang_df["Ma_KH"] == r["Ma_KH"]]
        ten_kh = kh.iloc[0]["Ten_NPP"] if not kh.empty else r["Ma_KH"]
        st.markdown(f"""
        <div class="order-card">
            <div style="display:flex;justify-content:space-between">
                <div class="order-code">{r['Ma_don']}</div>
                {status_badge_html(r['Trang_thai'])}
            </div>
            <div class="order-cust">{ten_kh}</div>
            <div class="order-value">{money(order_total(r['Ma_don']))}</div>
        </div>""", unsafe_allow_html=True)
        if st.button("Xem chi tiết", key=f"list_view_{r['Ma_don']}"):
            st.session_state.selected_order = r["Ma_don"]
            st.rerun()

    if st.session_state.selected_order:
        st.divider()
        render_order_detail(st.session_state.selected_order)

# ---------------------------------------------------------------------------
# Lên đơn
# ---------------------------------------------------------------------------
elif nav == "➕ Lên đơn":
    banner("Lên đơn hàng")

    with st.form("form_don_hang"):
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

        don_hang_ws = get_ws("Don_hang")
        ctdh_ws = get_ws("Chi_tiet_don_hang")
        ma_don = next_code(don_hang_ws, 1, "DH", 5)

        sale_pt = kh_row["Sale_phu_trach"]
        don_hang_ws.append_row([
            ma_don, ngay_len_don.strftime("%Y-%m-%d"), ma_kh, sale_pt,
            "Lên đơn", hinh_thuc_tt, ghi_chu_tt, ngay_len_don.strftime("%Y-%m-%d"),
            ngay_len_don.month, ngay_len_don.year
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
                thanh_tien, sl_dat + tang, "Lên đơn", ngay_len_don.month, ngay_len_don.year
            ], value_input_option="USER_ENTERED")

        refresh()
        st.success(f"Đã tạo đơn **{ma_don}** — tổng giá trị: {money(tong_tien)}")
        st.balloons()

# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
elif nav == "📊 Dashboard":
    banner("Tổng quan")

    if merged.empty:
        st.info("Chưa có dữ liệu đơn hàng để thống kê.")
    else:
        col1, col2 = st.columns(2)
        thang = col1.selectbox("Tháng", list(range(1, 13)),
                                index=(date.today().month - 1))
        nam = col2.number_input("Năm", min_value=2020, max_value=2100, value=date.today().year, step=1)

        valid = merged[merged["Trang_thai"].isin(VALID_STATUSES)].copy()
        valid["Ngay_len_don"] = pd.to_datetime(valid["Ngay_len_don"], errors="coerce")
        period = valid[(valid["Ngay_len_don"].dt.month == thang) & (valid["Ngay_len_don"].dt.year == nam)]

        doanh_thu = period["Thanh_tien"].sum()
        san_luong = period["San_luong_xuat_kho"].sum()
        so_don = period["Ma_don"].nunique()
        cong_no = don_hang_df[don_hang_df["Trang_thai"] != "Đã nhận hàng"].shape[0]

        c1, c2 = st.columns(2)
        c1.markdown(f"""<div class="metric-box"><div class="metric-label">Doanh thu</div>
            <div class="metric-value" style="color:{GREEN}">{money(doanh_thu)}</div></div>""", unsafe_allow_html=True)
        c2.markdown(f"""<div class="metric-box"><div class="metric-label">Số đơn hợp lệ</div>
            <div class="metric-value">{so_don}</div></div>""", unsafe_allow_html=True)
        c3, c4 = st.columns(2)
        c3.markdown(f"""<div class="metric-box"><div class="metric-label">Sản lượng xuất kho</div>
            <div class="metric-value">{int(san_luong)}</div></div>""", unsafe_allow_html=True)
        c4.markdown(f"""<div class="metric-box"><div class="metric-label">Đơn chưa hoàn tất</div>
            <div class="metric-value" style="color:{RED}">{cong_no}</div></div>""", unsafe_allow_html=True)

        st.markdown("#### Doanh thu 6 tháng gần nhất (đơn hợp lệ)")
        valid["ym"] = valid["Ngay_len_don"].dt.to_period("M")
        trend = valid.groupby("ym")["Thanh_tien"].sum().sort_index().tail(6)
        trend.index = trend.index.astype(str)
        st.bar_chart(trend)

        st.markdown("#### Doanh thu & sản lượng theo mặt hàng (tháng đã chọn)")
        by_sp = period.merge(san_pham_df[["Ma_SP", "Ten_SP"]], on="Ma_SP", how="left")
        by_sp = by_sp.groupby("Ten_SP").agg(
            San_luong=("San_luong_xuat_kho", "sum"),
            Doanh_thu=("Thanh_tien", "sum"),
        ).reset_index().sort_values("Doanh_thu", ascending=False)
        by_sp["Doanh_thu"] = by_sp["Doanh_thu"].apply(money)
        by_sp.columns = ["Sản phẩm", "Sản lượng", "Doanh thu"]
        st.dataframe(by_sp, hide_index=True, use_container_width=True)

# ---------------------------------------------------------------------------
# Khách hàng
# ---------------------------------------------------------------------------
elif nav == "👥 Khách hàng":
    banner("Danh sách khách hàng")
    show = khach_hang_df[["Ma_KH", "Ten_NPP", "Sale_phu_trach", "Khu_vuc", "Trang_thai"]].copy()
    show.columns = ["Mã KH", "Tên NPP", "Sale phụ trách", "Khu vực", "Trạng thái"]
    st.dataframe(show, hide_index=True, use_container_width=True)
