import io
import os
import base64
import json
from datetime import date, datetime, timedelta

import gspread
import pandas as pd
import requests
import streamlit as st
from google.oauth2.service_account import Credentials
from PIL import Image, ImageDraw, ImageFont

st.set_page_config(page_title="VIFEX - Quản lý đơn hàng", page_icon="📦", layout="centered")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

GREEN = "#15503F"
GREEN_BG = "#E2EDE8"
RED = "#D92B2B"
RED_BG = "#FDE8E8"
AMBER = "#D97706"
AMBER_BG = "#FEF3C7"
GRAY_BG = "#F3F4F6"
GRAY_TEXT = "#4B5563"

ALL_STATUSES = ["Lên đơn", "Gửi kho", "Đang giao hàng", "Đã nhận hàng", "Chưa Thanh toán"]
VALID_STATUSES = ["Gửi kho", "Đang giao hàng", "Đã nhận hàng", "Chưa Thanh toán"]

VAT_FOLDER_ID = "1HZiL99pNqV31u6Z8q5EqeoyjFJkPJFji"
DEFAULT_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyLfDcPZIs2QEshLL_uYSQ6-N4CnSbJhmorx3EI_28QZGd1EnNUEF9yrzh3Zx8M3bgqNw/exec"

# ---------------------------------------------------------------------------
# HÀM LOAD ẢNH LOGO DƯỚI DẠNG BASE64
# ---------------------------------------------------------------------------
def get_logo_base64():
    candidates = [
        os.path.join(os.path.dirname(__file__), "logo.png"),
        os.path.join(os.path.dirname(__file__), "2.png"),
        os.path.join(os.path.dirname(__file__), "logo.png.png"),
        "logo.png",
        "2.png",
        "logo.png.png",
    ]
    for p in candidates:
        if os.path.exists(p):
            with open(p, "rb") as img_f:
                return base64.b64encode(img_f.read()).decode()
    return None

# ---------------------------------------------------------------------------
# CSS RESPONSIVE & BỐ CỤC CHUẨN
# ---------------------------------------------------------------------------
st.markdown("""
<style>
header[data-testid="stHeader"] {
    background-color: transparent !important;
    z-index: 1 !important;
}

/* 1. Desktop / Màn hình rộng */
.block-container {
    max-width: 960px !important;
    padding-top: 4.5rem !important;
    padding-bottom: 4rem !important;
    padding-left: 1.5rem !important;
    padding-right: 1.5rem !important;
    margin: 0 auto !important;
}

/* Banner Header */
.vifex-banner {
    background: #15503F;
    color: #ffffff;
    padding: 18px 24px;
    border-radius: 16px;
    margin-bottom: 18px;
    box-shadow: 0 4px 14px rgba(21, 80, 63, 0.15);
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.vifex-banner-left {
    flex: 1;
}
.vifex-banner .brand-tag {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 14px;
    font-weight: 700;
    letter-spacing: 1px;
    opacity: 0.9;
    margin-bottom: 4px;
}
.vifex-banner .brand-tag::before {
    content: "";
    display: inline-block;
    width: 6px;
    height: 14px;
    background: #E53E3E;
    border-radius: 3px;
}
.vifex-banner .sub-title {
    font-size: 14px;
    opacity: 0.85;
}
.vifex-banner .main-title {
    font-size: 24px;
    font-weight: 700;
    line-height: 1.25;
    margin-top: 2px;
}
.vifex-banner-logo {
    background: #ffffff;
    width: 68px;
    height: 68px;
    border-radius: 14px;
    padding: 6px;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    margin-left: 16px;
    flex-shrink: 0;
}
.vifex-banner-logo img {
    max-width: 100%;
    max-height: 100%;
    object-fit: contain;
}

/* Thẻ Đơn hàng */
.order-card {
    background: #ffffff;
    border: 1px solid #edf0ed;
    border-radius: 14px;
    padding: 14px 16px;
    margin-bottom: 10px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.03);
}
.order-card-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 4px;
}
.order-code {
    font-weight: 700;
    font-size: 15px;
    color: #111827;
}
.order-cust {
    font-size: 13px;
    color: #4b5563;
    margin-bottom: 6px;
}
.order-value {
    font-weight: 700;
    color: #15503F;
    font-size: 15px;
}

.badge {
    display: inline-block;
    padding: 4px 10px;
    border-radius: 8px;
    font-size: 12px;
    font-weight: 600;
}

.metric-box {
    background: #ffffff;
    border: 1px solid #edf0ed;
    border-radius: 14px;
    padding: 14px 16px;
    text-align: left;
    box-shadow: 0 2px 6px rgba(0,0,0,0.02);
}
.metric-label {
    font-size: 12px;
    color: #6b7280;
    font-weight: 500;
}
.metric-value {
    font-size: 20px;
    font-weight: 700;
    margin-top: 4px;
}

.product-item-title {
    font-size: 13px;
    font-weight: 700;
    color: #15503F;
    margin-bottom: 8px;
}

/* Thanh điều hướng Desktop */
div[class*="st-key-vifex_nav"] {
    margin-bottom: 16px;
}
div[class*="st-key-vifex_nav"] button {
    padding: 10px 4px !important;
    font-size: 13px !important;
    border-radius: 12px !important;
    border: 1px solid #e5e7eb !important;
    background-color: #f9fafb !important;
    color: #374151 !important;
}
div[class*="st-key-vifex_nav"] button[kind="primary"] {
    background-color: #15503F !important;
    color: #ffffff !important;
    border-color: #15503F !important;
    font-weight: 600 !important;
}

div.stButton > button[kind="primary"] {
    border-radius: 12px !important;
    padding: 12px 20px !important;
    font-weight: 600 !important;
}

/* 2. Mobile */
@media screen and (max-width: 768px) {
    .block-container {
        max-width: 100% !important;
        padding-top: 4.8rem !important;
        padding-left: 0.75rem !important;
        padding-right: 0.75rem !important;
    }

    .vifex-banner {
        padding: 14px 16px !important;
        border-radius: 14px !important;
    }
    .vifex-banner .main-title {
        font-size: 19px !important;
    }

    .vifex-banner-logo {
        width: 62px !important;
        height: 62px !important;
        border-radius: 12px !important;
        padding: 5px !important;
        margin-left: 10px !important;
    }

    div[class*="st-key-vifex_nav"] [data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 6px !important;
    }
    div[class*="st-key-vifex_nav"] [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
        width: 33.33% !important;
        min-width: 33.33% !important;
        flex: 1 1 33.33% !important;
        padding: 0 !important;
    }
    div[class*="st-key-vifex_nav"] button {
        padding: 8px 2px !important;
        font-size: 11.5px !important;
        height: 42px !important;
        white-space: nowrap !important;
    }
}
</style>
""", unsafe_allow_html=True)


def banner(title, subtitle=None, highlight_text=None):
    sub_html = f'<div class="sub-title">{subtitle}</div>' if subtitle else ""
    main_text = highlight_text if highlight_text else title
    logo_b64 = get_logo_base64()
    
    if logo_b64:
        logo_html = f'<div class="vifex-banner-logo"><img src="data:image/png;base64,{logo_b64}" alt="VIFEX Logo" /></div>'
    else:
        logo_html = '<div class="vifex-banner-logo" style="font-weight:800;color:#15503F;font-size:12px;">VIFEX</div>'

    html_code = (
        f'<div class="vifex-banner">'
        f'<div class="vifex-banner-left">'
        f'<div class="brand-tag">VIFEX</div>'
        f'{sub_html}'
        f'<div class="main-title">{main_text}</div>'
        f'</div>'
        f'{logo_html}'
        f'</div>'
    )
    st.markdown(html_code, unsafe_allow_html=True)


def status_badge_html(status):
    colors = {
        "Lên đơn": (AMBER, AMBER_BG),
        "Gửi kho": (AMBER, AMBER_BG),
        "Đang giao hàng": (AMBER, AMBER_BG),
        "Đã nhận hàng": (GREEN, GREEN_BG),
        "Chưa Thanh toán": (RED, RED_BG),
    }
    fg, bg = colors.get(status, (GRAY_TEXT, GRAY_BG))
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
def get_credentials():
    return Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )

@st.cache_resource
def get_client():
    return gspread.authorize(get_credentials())

@st.cache_resource
def get_spreadsheet():
    return get_client().open_by_key(st.secrets["SPREADSHEET_ID"])


def serial_to_date(v):
    if v is None or v == "":
        return None
    try:
        if isinstance(v, str):
            return pd.to_datetime(v, dayfirst=False, errors="coerce").date()
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))).date()
    except Exception:
        return None


def read_raw(name):
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
    today = (datetime.utcnow() + timedelta(hours=7)).strftime("%Y-%m-%d")
    don_hang_ws = get_ws("Don_hang")
    row = find_row_by_code(don_hang_ws, ma_don, col_index=1)
    if row:
        don_hang_ws.update_cell(row, 5, new_status)          # E: Trang_thai
        don_hang_ws.update_cell(row, 8, today)                # H: Ngay_cap_nhat_trang_thai
    ctdh_ws = get_ws("Chi_tiet_don_hang")
    ma_don_col = ctdh_ws.col_values(2)
    for i, v in enumerate(ma_don_col, start=1):
        if v == ma_don:
            ctdh_ws.update_cell(i, 10, new_status)            # J: Trang_thai_don
    refresh()


# ---------------------------------------------------------------------------
# XỬ LÝ TẢI TỰ ĐỘNG LÊN GOOGLE DRIVE (GIỜ VIỆT NAM UTC+7 CHUẨN)
# ---------------------------------------------------------------------------
def get_upload_endpoint():
    try:
        secret_url = st.secrets.get("DRIVE_UPLOAD_URL", "").strip()
        if secret_url and secret_url.startswith("http"):
            return secret_url
    except Exception:
        pass
    return DEFAULT_SCRIPT_URL


def get_vat_sheet():
    sh = get_spreadsheet()
    try:
        return sh.worksheet("Hoa_don_VAT")
    except Exception:
        ws = sh.add_worksheet(title="Hoa_don_VAT", rows=100, cols=6)
        ws.append_row(["Ma_HD_VAT", "Ma_don", "Ma_KH", "Ten_file", "Ngay_tai_len", "Link_Drive"])
        return ws


def get_vat_link_from_sheet(ma_don):
    try:
        ws = get_vat_sheet()
        records = ws.get_all_records()
        for r in records:
            if str(r.get("Ma_don")) == ma_don:
                link = r.get("Link_Drive") or r.get("Link_file_PDF") or ""
                if link and str(link).startswith("http"):
                    return link
    except Exception:
        pass
    return None


def upload_vat_directly_to_drive(ma_don, ma_kh, uploaded_file):
    """Tự động đẩy file PDF từ Streamlit thẳng vào folder Drive thông qua Webhook Apps Script"""
    url = get_upload_endpoint()
    _, ext = os.path.splitext(uploaded_file.name)
    file_name = f"{ma_don}_VAT{ext.lower()}"
    
    b64_data = base64.b64encode(uploaded_file.getvalue()).decode()
    
    payload = {
        "folderId": VAT_FOLDER_ID,
        "fileName": file_name,
        "mimeType": uploaded_file.type if uploaded_file.type else "application/pdf",
        "base64Data": b64_data
    }
    
    response = requests.post(
        url,
        data=json.dumps(payload),
        headers={"Content-Type": "application/json"},
        timeout=60,
        allow_redirects=True
    )
    
    try:
        data = response.json()
    except Exception:
        raise Exception(f"Lỗi kết nối Webhook ({response.status_code}): Link đang gọi là: {url}")
    
    if data.get("status") == "success":
        file_url = data.get("fileUrl")
        try:
            ws = get_vat_sheet()
            rows = ws.col_values(2) # Cột B: Ma_don
            target_row = None
            for i, val in enumerate(rows, start=1):
                if val == ma_don:
                    target_row = i
                    break
            
            today_str = (datetime.utcnow() + timedelta(hours=7)).strftime("%Y-%m-%d %H:%M")
            if target_row:
                ws.update_cell(target_row, 4, uploaded_file.name) # Cột D
                ws.update_cell(target_row, 5, today_str)          # Cột E
                ws.update_cell(target_row, 6, file_url)           # Cột F: Link_Drive
            else:
                ma_vat = next_code(ws, 1, "HDVAT", 4)
                ws.append_row([ma_vat, ma_don, ma_kh, uploaded_file.name, today_str, file_url])
        except Exception:
            pass
        return file_url
    else:
        raise Exception(data.get("message", "Lỗi tải lên Google Drive."))


def delete_vat_record(ma_don):
    try:
        ws = get_vat_sheet()
        rows = ws.col_values(2)
        for i, val in enumerate(rows, start=1):
            if val == ma_don:
                ws.delete_rows(i)
                break
    except Exception:
        pass


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
    W = 860
    row_h = 34
    header_h = 270
    footer_h = 90
    H = header_h + row_h * (len(items_df) + 2) + footer_h

    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    f_title = get_font(26)
    f_sub = get_font(14)
    f_h = get_font(16)
    f_n = get_font(14)

    d.rectangle([0, 0, W, 80], fill=GREEN)
    draw_bold(d, (24, 16), "VIFEX", f_title, "white")
    d.text((24, 50), "PHIẾU XUẤT ĐƠN HÀNG (không phải hóa đơn VAT)", font=f_sub, fill="white")

    y = 100
    ten_cty = khach_hang_row.get("Ten_cong_ty_GPKD") or khach_hang_row.get("Ten_cong_ty") or khach_hang_row.get("Ten_GPKD") or khach_hang_row.get("Ten_NPP", "")
    ten_npp = khach_hang_row.get("Ten_NPP", "")
    dia_chi = khach_hang_row.get("Dia_chi_giao_phu") or khach_hang_row.get("Dia_chi_GPKD", "")
    sdt = khach_hang_row.get("SDT_phu", "")

    draw_bold(d, (24, y), f"Mã đơn: {ma_don}", f_h, "black"); y += 26
    d.text((24, y), f"Ngày lên đơn: {order_row.get('Ngay_len_don')}", font=f_n, fill="black"); y += 24
    d.text((24, y), f"Khách hàng: {ten_cty}", font=f_n, fill="black"); y += 24
    if ten_npp:
        d.text((24, y), f"NHÀ PHÂN PHỐI : {ten_npp}", font=f_n, fill="black"); y += 24
    if dia_chi:
        d.text((24, y), f"Địa chỉ giao: {dia_chi}", font=f_n, fill="black"); y += 24
    if sdt:
        d.text((24, y), f"SĐT: {sdt}", font=f_n, fill="black"); y += 24
    y += 10

    d.line([24, y, W - 24, y], fill="#ddd", width=1); y += 10
    cols_x = [24, 270, 370, 440, 520, 640, 750]
    headers = ["Sản phẩm", "Đơn giá", "SL đặt", "Tặng", "Tổng tiền", "Chiết khấu", "Thành tiền"]
    for x, h in zip(cols_x, headers):
        draw_bold(d, (x, y), h, f_n, GREEN)
    y += row_h
    d.line([24, y - 6, W - 24, y - 6], fill="#ddd", width=1)

    total_thanh_tien = 0
    for _, r in items_df.iterrows():
        tong_tien_hang = float(r["SL_dat"]) * float(r["Don_gia_ap_dung"])
        total_thanh_tien += r["Thanh_tien"]
        
        d.text((cols_x[0], y), str(r["Ten_SP"])[:26], font=f_n, fill="black")
        d.text((cols_x[1], y), money(r["Don_gia_ap_dung"]), font=f_n, fill="black")
        d.text((cols_x[2], y), str(int(r["SL_dat"])), font=f_n, fill="black")
        d.text((cols_x[3], y), str(int(r["Tang"])), font=f_n, fill="black")
        d.text((cols_x[4], y), money(tong_tien_hang), font=f_n, fill="black")
        d.text((cols_x[5], y), money(r["Chiet_khau"]), font=f_n, fill="black")
        d.text((cols_x[6], y), money(r["Thanh_tien"]), font=f_n, fill="black")
        y += row_h

    d.line([24, y, W - 24, y], fill="#ddd", width=1); y += 14
    draw_bold(d, (cols_x[5], y), "TỔNG CỘNG:", f_h, RED)
    draw_bold(d, (cols_x[6], y), money(total_thanh_tien), f_h, RED)
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
nhan_vien_df = data["nhan_vien"]
lich_su_gia_df = data["lich_su_gia"]
don_hang_df = data["don_hang"]
ctdh_df = data["ctdh"]

if khach_hang_df.empty or san_pham_df.empty:
    st.warning("Chưa có dữ liệu Khách hàng hoặc Sản phẩm. Vào Google Sheets nhập trước đã.")
    st.stop()

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
if "order_items_count" not in st.session_state:
    st.session_state.order_items_count = 1
if "order_form_version" not in st.session_state:
    st.session_state.order_form_version = 0

# ---------------------------------------------------------------------------
# THANH ĐIỀU HƯỚNG TAB
# ---------------------------------------------------------------------------
NAV_OPTIONS = ["Trang chủ", "Đơn hàng", "Lên đơn", "Lương Sale", "Dashboard", "Khách hàng"]
NAV_ICONS = {"Trang chủ": "🏠", "Đơn hàng": "📦", "Lên đơn": "➕", "Lương Sale": "💰", "Dashboard": "📊", "Khách hàng": "👥"}

with st.container(key="vifex_nav"):
    row1 = st.columns(3)
    row2 = st.columns(3)
    
    for idx, opt in enumerate(NAV_OPTIONS[:3]):
        is_active = (st.session_state.nav == f"{NAV_ICONS[opt]} {opt}")
        if row1[idx].button(f"{NAV_ICONS[opt]} {opt}", key=f"nav_{opt}", type="primary" if is_active else "secondary", use_container_width=True):
            st.session_state.nav = f"{NAV_ICONS[opt]} {opt}"
            st.rerun()

    for idx, opt in enumerate(NAV_OPTIONS[3:]):
        is_active = (st.session_state.nav == f"{NAV_ICONS[opt]} {opt}")
        if row2[idx].button(f"{NAV_ICONS[opt]} {opt}", key=f"nav_{opt}", type="primary" if is_active else "secondary", use_container_width=True):
            st.session_state.nav = f"{NAV_ICONS[opt]} {opt}"
            st.rerun()

nav = st.session_state.nav
st.write("")


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
    ma_kh = order_row["Ma_KH"]

    items = ctdh_df[ctdh_df["Ma_don"] == ma_don].copy()
    items = items.merge(san_pham_df[["Ma_SP", "Ten_SP"]], on="Ma_SP", how="left")

    ten_cty = kh_row.get("Ten_cong_ty_GPKD") or kh_row.get("Ten_cong_ty") or kh_row.get("Ten_GPKD") or kh_row.get("Ten_NPP", "")
    ten_npp = kh_row.get("Ten_NPP", "")

    with st.container(border=True):
        st.markdown(f"#### Chi tiết đơn: `{ma_don}`")
        st.markdown(status_badge_html(order_row["Trang_thai"]), unsafe_allow_html=True)
        st.write(f"**Khách hàng:** {ten_cty}")
        st.write(f"**NHÀ PHÂN PHỐI :** {ten_npp}")
        st.write(f"**Ngày lên đơn:** {order_row['Ngay_len_don']}")
        st.write(f"**Hình thức thanh toán:** {order_row['Hinh_thuc_thanh_toan']}")
        if order_row.get("Ghi_chu_thanh_toan"):
            st.write(f"**Ghi chú:** {order_row['Ghi_chu_thanh_toan']}")

        st.markdown("---")
        
        # Tính cột Tổng tiền hàng trước chiết khấu
        items["Tong_tien_hang"] = items["SL_dat"] * items["Don_gia_ap_dung"]
        
        # Sắp xếp đúng thứ tự: Sản phẩm | Đơn giá | SL đặt | Tặng | Tổng tiền | Chiết khấu | Thành tiền
        show_cols = pd.DataFrame()
        show_cols["Sản phẩm"] = items["Ten_SP"]
        show_cols["Đơn giá"] = items["Don_gia_ap_dung"].apply(money)
        show_cols["SL đặt"] = items["SL_dat"].astype(int)
        show_cols["Tặng"] = items["Tang"].astype(int)
        show_cols["Tổng tiền"] = items["Tong_tien_hang"].apply(money)
        show_cols["Chiết khấu"] = items["Chiet_khau"].apply(money)
        show_cols["Thành tiền"] = items["Thanh_tien"].apply(money)

        st.dataframe(show_cols, hide_index=True, use_container_width=True)

        total = items["Thanh_tien"].sum()
        st.markdown(f"<div style='font-size:16px;font-weight:700;color:{GREEN};text-align:right;'>Tổng cộng: {money(total)}</div>", unsafe_allow_html=True)
        st.markdown("---")

        new_status = st.selectbox("Cập nhật trạng thái", ALL_STATUSES,
                                   index=ALL_STATUSES.index(order_row["Trang_thai"])
                                   if order_row["Trang_thai"] in ALL_STATUSES else 0,
                                   key=f"status_{ma_don}")
        
        # 3 CỘT NÚT HÀNH ĐỘNG
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("💾 Lưu trạng thái", key=f"save_status_{ma_don}", type="primary", use_container_width=True):
                update_order_status(ma_don, new_status)
                st.success("Đã cập nhật trạng thái.")
                st.rerun()
        with c2:
            png_bytes = generate_order_slip(ma_don, order_row, items, kh_row if isinstance(kh_row, dict) else kh_row.to_dict())
            st.download_button("📥 Phiếu xuất (PNG)", data=png_bytes,
                                file_name=f"{ma_don}_phieu_xuat.png", mime="image/png",
                                key=f"dl_{ma_don}", use_container_width=True)
        with c3:
            drive_vat_url = get_vat_link_from_sheet(ma_don)
            if drive_vat_url:
                st.link_button("📄 Xem VAT (Drive)", url=drive_vat_url, use_container_width=True)
            else:
                st.button("☁️ Hóa đơn (Chưa có)", disabled=True, use_container_width=True)

        # KHU VỰC TẢI THẲNG LÊN GOOGLE DRIVE (TỰ ĐỘNG ĐẨY VÀO FOLDER)
        with st.expander("☁️ **Tải lên Hóa đơn VAT tự động vào Google Drive**", expanded=False):
            drive_vat_url = get_vat_link_from_sheet(ma_don)
            if drive_vat_url:
                st.info(f"✅ Đơn hàng **{ma_don}** đã có file lưu trên Google Drive.")
                st.markdown(f"👉 [Bấm vào đây để mở xem file trên Google Drive]({drive_vat_url})")
                if st.button("🗑️ Xóa liên kết hóa đơn này", key=f"btn_del_vat_{ma_don}"):
                    delete_vat_record(ma_don)
                    st.success(f"Đã xóa thông tin hóa đơn VAT của đơn **{ma_don}**.")
                    st.rerun()
                st.markdown("---")
                st.caption("Chọn tệp mới bên dưới nếu muốn **Tải đè / Thay thế file trên Drive**:")
            else:
                st.caption("Kéo thả tệp PDF vào đây, hệ thống sẽ **tự động đẩy thẳng file vào thư mục Google Drive**.")

            uploaded_vat = st.file_uploader(
                f"Chọn tệp hóa đơn cho đơn {ma_don}", 
                type=["pdf", "png", "jpg", "jpeg"], 
                key=f"vat_uploader_{ma_don}",
                label_visibility="collapsed"
            )
            
            if uploaded_vat is not None:
                if st.button("⬆️ Tải lên Google Drive", key=f"btn_save_vat_{ma_don}", type="primary"):
                    with st.spinner("Đang tự động đẩy file lên Google Drive..."):
                        try:
                            file_drive_link = upload_vat_directly_to_drive(ma_don, ma_kh, uploaded_vat)
                            st.success(f"Đã tải thành công file lên Google Drive cho đơn **{ma_don}**!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Lỗi: {e}")

        st.write("")
        if st.button("← Đóng xem chi tiết", key=f"close_{ma_don}", use_container_width=True):
            st.session_state.selected_order = None
            st.rerun()


# ---------------------------------------------------------------------------
# 1. TRANG CHỦ
# ---------------------------------------------------------------------------
if nav == "🏠 Trang chủ":
    pending = don_hang_df[don_hang_df["Trang_thai"].isin(["Gửi kho", "Đang giao hàng", "Chưa Thanh toán"])] if not don_hang_df.empty else pd.DataFrame()
    so_don_pending = len(pending)
    
    banner("Trang chủ", subtitle="Xin chào, Coco", highlight_text=f"{so_don_pending} đơn cần xử lý")

    if don_hang_df.empty:
        st.info("Chưa có đơn hàng nào.")
    else:
        counts = don_hang_df["Trang_thai"].value_counts()
        c1, c2 = st.columns(2)
        c1.markdown(f"""
        <div class="metric-box" style="background:{AMBER_BG};border-color:#FDE68A;">
            <div class="metric-value" style="color:{AMBER};margin-top:0;">{int(counts.get('Gửi kho', 0))}</div>
            <div class="metric-label" style="color:#92400E;font-weight:600;">Gửi kho</div>
        </div>""", unsafe_allow_html=True)
        c2.markdown(f"""
        <div class="metric-box" style="background:{RED_BG};border-color:#FECACA;">
            <div class="metric-value" style="color:{RED};margin-top:0;">{int(counts.get('Đang giao hàng', 0)) + int(counts.get('Chưa Thanh toán', 0))}</div>
            <div class="metric-label" style="color:#991B1B;font-weight:600;">Đang giao / Chưa TT</div>
        </div>""", unsafe_allow_html=True)

        st.write("")
        st.markdown("<div style='font-size:14px;font-weight:700;color:#374151;margin-bottom:8px;'>ĐƠN CẦN XỬ LÝ</div>", unsafe_allow_html=True)
        
        pending_view = pending.sort_values("Ngay_len_don", ascending=False).head(6)
        if pending_view.empty:
            st.caption("Hiện không có đơn nào cần xử lý.")
        
        for _, r in pending_view.iterrows():
            kh = khach_hang_df[khach_hang_df["Ma_KH"] == r["Ma_KH"]]
            ten_kh = kh.iloc[0]["Ten_NPP"] if not kh.empty else r["Ma_KH"]
            st.markdown(f"""
            <div class="order-card">
                <div class="order-card-header">
                    <span class="order-code">{r['Ma_don']}</span>
                    {status_badge_html(r['Trang_thai'])}
                </div>
                <div class="order-cust">NPP {ten_kh}</div>
                <div class="order-value">{money(order_total(r['Ma_don']))}</div>
            </div>""", unsafe_allow_html=True)
            if st.button("Xem chi tiết", key=f"home_view_{r['Ma_don']}", use_container_width=True):
                st.session_state.selected_order = r["Ma_don"]
                st.rerun()

    st.write("")
    if st.button("+ Tạo đơn mới", key="home_new_order", type="primary", use_container_width=True):
        st.session_state.nav = "➕ Lên đơn"
        st.rerun()

    if st.session_state.selected_order:
        st.divider()
        render_order_detail(st.session_state.selected_order)

# ---------------------------------------------------------------------------
# 2. DANH SÁCH ĐƠN HÀNG
# ---------------------------------------------------------------------------
elif nav == "📦 Đơn hàng":
    banner("Danh sách đơn hàng")

    filter_status = st.selectbox("Lọc theo trạng thái đơn hàng", ["Tất cả"] + ALL_STATUSES)
    
    view_df = don_hang_df.copy()
    if filter_status != "Tất cả":
        view_df = view_df[view_df["Trang_thai"] == filter_status]
    view_df = view_df.sort_values("Ngay_len_don", ascending=False)

    st.write("")
    if view_df.empty:
        st.info("Không có đơn hàng nào phù hợp.")
    for _, r in view_df.iterrows():
        kh = khach_hang_df[khach_hang_df["Ma_KH"] == r["Ma_KH"]]
        ten_kh = kh.iloc[0]["Ten_NPP"] if not kh.empty else r["Ma_KH"]
        st.markdown(f"""
        <div class="order-card">
            <div class="order-card-header">
                <span class="order-code">{r['Ma_don']}</span>
                {status_badge_html(r['Trang_thai'])}
            </div>
            <div class="order-cust">NPP {ten_kh}</div>
            <div class="order-value">{money(order_total(r['Ma_don']))}</div>
        </div>""", unsafe_allow_html=True)
        if st.button("Xem chi tiết", key=f"list_view_{r['Ma_don']}", use_container_width=True):
            st.session_state.selected_order = r["Ma_don"]
            st.rerun()

    if st.session_state.selected_order:
        st.divider()
        render_order_detail(st.session_state.selected_order)

# ---------------------------------------------------------------------------
# 3. LÊN ĐƠN HÀNG (RESET TRẠNG THÁI & HIỂN THỊ ĐỊNH DẠNG TIỀN CK TRỰC QUAN)
# ---------------------------------------------------------------------------
elif nav == "➕ Lên đơn":
    banner("Lên đơn hàng")

    v = st.session_state.order_form_version
    ten_npp = st.selectbox("Khách hàng (NPP)", khach_hang_df["Ten_NPP"].dropna().tolist(), key=f"form_npp_{v}")
    kh_matches = khach_hang_df[khach_hang_df["Ten_NPP"] == ten_npp]
    kh_row = kh_matches.iloc[0] if not kh_matches.empty else {}
    ma_kh = kh_row.get("Ma_KH", "")
    sale_pt = kh_row.get("Sale_phu_trach", "")

    st.caption(f"Sale phụ trách: **{sale_pt}**")

    # Hiển thị thông báo nếu vừa tạo đơn thành công
    if "just_created_order" in st.session_state:
        msg, money_val = st.session_state.just_created_order
        st.success(f"Đã tạo đơn **{msg}** — tổng giá trị: **{money(money_val)}**")
        del st.session_state.just_created_order

    with st.container(border=True):
        c_d1, c_d2 = st.columns(2)
        with c_d1:
            ngay_len_don = st.date_input("Ngày lên đơn", value=date.today(), key=f"form_date_{v}")
        with c_d2:
            hinh_thuc_tt = st.selectbox("Hình thức TT", ["Tiền mặt", "Chuyển khoản", "Công nợ", "Khác"], key=f"form_httt_{v}")
        
        ghi_chu_tt = st.text_input("Ghi chú thanh toán", "", key=f"form_note_{v}")

        st.markdown("<div style='font-size:14px;font-weight:700;color:#15503F;margin-12px 0 8px 0;'>DANH SÁCH SẢN PHẨM</div>", unsafe_allow_html=True)

        line_items = []
        ten_sp_list = san_pham_df["Ten_SP"].dropna().tolist()
        
        for i in range(st.session_state.order_items_count):
            st.markdown(f"""<div class="product-item-title">Sản phẩm #{i+1}</div>""", unsafe_allow_html=True)
            
            ten_sp = st.selectbox(f"Chọn SP #{i+1}", ten_sp_list, key=f"sp_{v}_{i}", label_visibility="collapsed")
            
            c1, c2, c3 = st.columns(3)
            with c1:
                sl_dat = st.number_input("SL đặt", min_value=0, value=0, key=f"sl_{v}_{i}")
            with c2:
                tang = st.number_input("Tặng", min_value=0.0, value=0.0, step=0.1, key=f"tang_{v}_{i}")
            with c3:
                chiet_khau = st.number_input("CK (đ)", min_value=0, value=0, step=10000, key=f"ck_{v}_{i}")
                if chiet_khau > 0:
                    st.caption(f"↳ {money(chiet_khau)}")
            
            st.markdown("<div style='height:1px;background:#edf0ed;margin:10px 0;'></div>", unsafe_allow_html=True)
            line_items.append((ten_sp, sl_dat, tang, chiet_khau))

        col_add, col_remove = st.columns(2)
        with col_add:
            if st.button("➕ Thêm sản phẩm", key="btn_add_product", use_container_width=True):
                st.session_state.order_items_count += 1
                st.rerun()
        with col_remove:
            if st.session_state.order_items_count > 1:
                if st.button("➖ Bớt sản phẩm", key="btn_remove_product", use_container_width=True):
                    st.session_state.order_items_count -= 1
                    st.rerun()

        st.write("")
        submitted = st.button("✅ Tạo đơn hàng", use_container_width=True, type="primary")

    if submitted:
        valid_items = [li for li in line_items if li[1] > 0]
        if not valid_items:
            st.error("Cần ít nhất 1 dòng sản phẩm có SL đặt > 0.")
            st.stop()

        don_hang_ws = get_ws("Don_hang")
        ctdh_ws = get_ws("Chi_tiet_don_hang")
        ma_don = next_code(don_hang_ws, 1, "DH", 5)

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

        # Reset form về trạng thái trống ban đầu
        st.session_state.order_items_count = 1
        st.session_state.order_form_version += 1
        st.session_state.just_created_order = (ma_don, tong_tien)
        refresh()
        st.rerun()

# ---------------------------------------------------------------------------
# 4. LƯƠNG SALE
# ---------------------------------------------------------------------------
elif nav == "💰 Lương Sale":
    banner("Doanh số & Lương Sale")

    if nhan_vien_df.empty:
        st.info("Chưa có dữ liệu nhân viên Sale.")
    elif merged.empty:
        st.info("Chưa có dữ liệu đơn hàng để tính doanh số.")
    else:
        ten_nv = st.selectbox("Chọn nhân viên Sale", nhan_vien_df["Ten_NV"].tolist())
        c_m, c_y = st.columns(2)
        with c_m:
            thang = st.selectbox("Tháng", list(range(1, 13)), index=(date.today().month - 1), key="sale_thang")
        with c_y:
            nam = st.number_input("Năm", min_value=2020, max_value=2100, value=date.today().year, step=1, key="sale_nam")

        nv_row = nhan_vien_df[nhan_vien_df["Ten_NV"] == ten_nv].iloc[0]
        ma_nv = nv_row["Ma_NV"]
        ty_le_hh = float(nv_row.get("Ty_le_hoa_hong") or 0.02)

        valid = merged[merged["Trang_thai"].isin(VALID_STATUSES)].copy()
        valid["Ngay_len_don"] = pd.to_datetime(valid["Ngay_len_don"], errors="coerce")
        of_sale = valid[(valid["Sale_phu_trach"] == ma_nv) &
                         (valid["Ngay_len_don"].dt.month == thang) &
                         (valid["Ngay_len_don"].dt.year == nam)]

        doanh_thu = of_sale["Thanh_tien"].sum()
        doanh_thu_sau_vat = doanh_thu * (1 - 0.08)
        luong = doanh_thu_sau_vat * ty_le_hh

        st.caption("Tính theo tất cả các đơn hàng đã duyệt xuất kho trở đi (ngoại trừ trạng thái 'Lên đơn').")

        c1, c2 = st.columns(2)
        c1.markdown(f"""<div class="metric-box"><div class="metric-label">Doanh thu hợp lệ</div>
            <div class="metric-value" style="color:{GREEN}">{money(doanh_thu)}</div></div>""", unsafe_allow_html=True)
        c2.markdown(f"""<div class="metric-box"><div class="metric-label">Sau trừ VAT 8%</div>
            <div class="metric-value">{money(doanh_thu_sau_vat)}</div></div>""", unsafe_allow_html=True)
        
        st.write("")
        c3, c4 = st.columns(2)
        c3.markdown(f"""<div class="metric-box"><div class="metric-label">Tỷ lệ hoa hồng</div>
            <div class="metric-value">{ty_le_hh*100:.1f}%</div></div>""", unsafe_allow_html=True)
        c4.markdown(f"""<div class="metric-box"><div class="metric-label">Lương thực nhận</div>
            <div class="metric-value" style="color:{RED}">{money(luong)}</div></div>""", unsafe_allow_html=True)

        st.write("")
        st.markdown(f"#### 📋 Doanh số theo Nhà phân phối & Nhóm hàng (Tháng {thang}/{nam})")
        if of_sale.empty:
            st.caption("Chưa có đơn hợp lệ nào trong tháng này.")
        else:
            df_nhom = of_sale.merge(khach_hang_df[["Ma_KH", "Ten_NPP"]], on="Ma_KH", how="left")
            df_nhom = df_nhom.merge(san_pham_df[["Ma_SP", "Nhom_danh_muc"]], on="Ma_SP", how="left")
            df_nhom["Nhom_danh_muc"] = df_nhom["Nhom_danh_muc"].fillna("Khác")

            by_npp = df_nhom.groupby(["Ten_NPP", "Nhom_danh_muc"]).agg(
                San_luong=("San_luong_xuat_kho", "sum"),
                Doanh_thu=("Thanh_tien", "sum"),
            ).reset_index().sort_values(["Ten_NPP", "Doanh_thu"], ascending=[True, False])

            by_npp["Doanh_thu"] = by_npp["Doanh_thu"].apply(money)
            by_npp.columns = ["Nhà phân phối (NPP)", "Nhóm danh mục", "Sản lượng (Thùng)", "Doanh thu"]
            st.dataframe(by_npp, hide_index=True, use_container_width=True)

        st.write("")
        st.markdown(f"#### 📦 Chi tiết từng mặt hàng đã bán")
        if not of_sale.empty:
            by_sp = of_sale.merge(san_pham_df[["Ma_SP", "Ten_SP"]], on="Ma_SP", how="left")
            by_sp = by_sp.groupby("Ten_SP").agg(
                San_luong=("San_luong_xuat_kho", "sum"),
                Doanh_thu=("Thanh_tien", "sum"),
            ).reset_index().sort_values("Doanh_thu", ascending=False)
            by_sp["Doanh_thu"] = by_sp["Doanh_thu"].apply(money)
            by_sp.columns = ["Sản phẩm", "Sản lượng", "Doanh thu"]
            st.dataframe(by_sp, hide_index=True, use_container_width=True)

# ---------------------------------------------------------------------------
# 5. DASHBOARD TỔNG QUAN
# ---------------------------------------------------------------------------
elif nav == "📊 Dashboard":
    banner("Tổng quan")

    if merged.empty:
        st.info("Chưa có dữ liệu đơn hàng để thống kê.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            thang = st.selectbox("Tháng", list(range(1, 13)), index=(date.today().month - 1))
        with col2:
            nam = st.number_input("Năm", min_value=2020, max_value=2100, value=date.today().year, step=1)

        valid = merged[merged["Trang_thai"].isin(VALID_STATUSES)].copy()
        valid["Ngay_len_don"] = pd.to_datetime(valid["Ngay_len_don"], errors="coerce")
        period = valid[(valid["Ngay_len_don"].dt.month == thang) & (valid["Ngay_len_don"].dt.year == nam)]

        doanh_thu = period["Thanh_tien"].sum()
        san_luong = period["San_luong_xuat_kho"].sum()
        so_don = period["Ma_don"].nunique()
        cong_no_count = don_hang_df[don_hang_df["Trang_thai"] != "Đã nhận hàng"].shape[0]

        c1, c2 = st.columns(2)
        c1.markdown(f"""
        <div class="metric-box">
            <div class="metric-label">Doanh thu</div>
            <div class="metric-value" style="color:{GREEN}">{money(doanh_thu)}</div>
        </div>""", unsafe_allow_html=True)
        c2.markdown(f"""
        <div class="metric-box">
            <div class="metric-label">Đơn hàng hợp lệ</div>
            <div class="metric-value">{so_don}</div>
        </div>""", unsafe_allow_html=True)

        st.write("")
        c3, c4 = st.columns(2)
        c3.markdown(f"""
        <div class="metric-box">
            <div class="metric-label">Đang giao</div>
            <div class="metric-value">{int(don_hang_df[don_hang_df['Trang_thai'] == 'Đang giao hàng'].shape[0])}</div>
        </div>""", unsafe_allow_html=True)
        c4.markdown(f"""
        <div class="metric-box">
            <div class="metric-label">Đơn chưa nhận</div>
            <div class="metric-value" style="color:{RED}">{cong_no_count}</div>
        </div>""", unsafe_allow_html=True)

        st.write("")
        st.markdown("<div style='font-size:14px;font-weight:700;color:#374151;margin-bottom:8px;'>DOANH THU 6 THÁNG GẦN NHẤT</div>", unsafe_allow_html=True)
        valid["ym"] = valid["Ngay_len_don"].dt.to_period("M")
        trend = valid.groupby("ym")["Thanh_tien"].sum().sort_index().tail(6)
        trend.index = trend.index.astype(str)
        st.bar_chart(trend)

        st.markdown("<div style='font-size:14px;font-weight:700;color:#374151;margin-bottom:8px;'>SẢN LƯỢNG & DOANH THU THEO MẶT HÀNG</div>", unsafe_allow_html=True)
        by_sp = period.merge(san_pham_df[["Ma_SP", "Ten_SP"]], on="Ma_SP", how="left")
        by_sp = by_sp.groupby("Ten_SP").agg(
            San_luong=("San_luong_xuat_kho", "sum"),
            Doanh_thu=("Thanh_tien", "sum"),
        ).reset_index().sort_values("Doanh_thu", ascending=False)
        by_sp["Doanh_thu"] = by_sp["Doanh_thu"].apply(money)
        by_sp.columns = ["Sản phẩm", "Sản lượng", "Doanh thu"]
        st.dataframe(by_sp, hide_index=True, use_container_width=True)

# ---------------------------------------------------------------------------
# 6. KHÁCH HÀNG
# ---------------------------------------------------------------------------
elif nav == "👥 Khách hàng":
    banner("Danh sách khách hàng")
    show = khach_hang_df[["Ma_KH", "Ten_NPP", "Sale_phu_trach", "Khu_vuc", "Trang_thai"]].dropna(subset=["Ten_NPP"]).copy()
    show.columns = ["Mã KH", "Tên NPP", "Sale phụ trách", "Khu vực", "Trạng thái"]
    st.dataframe(show, hide_index=True, use_container_width=True)
