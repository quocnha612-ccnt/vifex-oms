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
BLUE = "#2563EB"
BLUE_BG = "#EFF6FF"
PURPLE = "#7C3AED"
PURPLE_BG = "#F5F3FF"
TEAL = "#0D9488"
TEAL_BG = "#CCFBF1"
ORANGE = "#EA580C"
ORANGE_BG = "#FFEDD5"
GRAY_BG = "#F3F4F6"
GRAY_TEXT = "#4B5563"

# 1. Trạng thái vận hành đơn hàng
ORDER_STATUSES = ["Lên đơn", "Gửi kho", "Đang giao", "Đã giao"]
VALID_ORDER_STATUSES = ["Gửi kho", "Đang giao", "Đang giao hàng", "Đã giao", "Đã giao hàng", "Đã nhận hàng"]

# 2. Trạng thái thanh toán & VAT
PAYMENT_STATUSES = [
    "Chưa thanh toán",
    "Chưa TT - Nháp",
    "Chưa TT - Đã VAT",
    "TT - Chưa VAT",
    "TT - Nháp VAT",
    "TT - không VAT",
    "TT - Đã VAT"
]

VAT_FOLDER_ID = "1HZiL99pNqV31u6Z8q5EqeoyjFJkPJFji"
DEFAULT_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyLfDcPZIs2QEshLL_uYSQ6-N4CnSbJhmorx3EI_28QZGd1EnNUEF9yrzh3Zx8M3bgqNw/exec"

def safe_str(v):
    if v is None or pd.isna(v):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s

# ---------------------------------------------------------------------------
# HÀM LOAD ẢNH LOGO & QR DƯỚI DẠNG BASE64 / PIL
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

def get_qr_bank_image():
    candidates = [
        os.path.join(os.path.dirname(__file__), "qr_bank.png"),
        os.path.join(os.path.dirname(__file__), "qr.png"),
        "qr_bank.png",
        "qr.png",
        "input_file_27.png"
    ]
    for p in candidates:
        if os.path.exists(p):
            try:
                return Image.open(p).convert("RGBA")
            except Exception:
                pass
    return None

# ---------------------------------------------------------------------------
# CSS RESPONSIVE, THU GỌN KHOẢNG CÁCH & BỐ CỤC GIAO DIỆN
# ---------------------------------------------------------------------------
st.markdown("""
<style>
header[data-testid="stHeader"] {
    background-color: transparent !important;
    z-index: 1 !important;
}

/* Desktop */
.block-container {
    max-width: 960px !important;
    padding-top: 3.5rem !important;
    padding-bottom: 3rem !important;
    padding-left: 1.25rem !important;
    padding-right: 1.25rem !important;
    margin: 0 auto !important;
}

/* Banner Header */
.vifex-banner {
    background: #15503F;
    color: #ffffff;
    padding: 16px 22px;
    border-radius: 14px;
    margin-bottom: 14px;
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
    font-size: 13.5px;
    font-weight: 700;
    letter-spacing: 1px;
    opacity: 0.9;
    margin-bottom: 3px;
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
    font-size: 13.5px;
    opacity: 0.85;
}
.vifex-banner .main-title {
    font-size: 23px;
    font-weight: 700;
    line-height: 1.25;
    margin-top: 2px;
}
.vifex-banner-logo {
    background: #ffffff;
    width: 64px;
    height: 64px;
    border-radius: 12px;
    padding: 5px;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    margin-left: 14px;
    flex-shrink: 0;
}
.vifex-banner-logo img {
    max-width: 100%;
    max-height: 100%;
    object-fit: contain;
}

/* Thẻ Đơn hàng Compact */
div[data-testid="stVerticalBlockBorderWrapper"] {
    margin-bottom: 8px !important;
    border-radius: 12px !important;
    border-color: #edf0ed !important;
    background: #ffffff !important;
    box-shadow: 0 2px 5px rgba(0,0,0,0.02) !important;
}

.order-code-compact {
    font-weight: 700;
    font-size: 15.5px;
    color: #111827;
    margin-bottom: 1px;
}
.order-cust-compact {
    font-size: 13px;
    color: #4b5563;
    margin-bottom: 3px;
}
.order-value-compact {
    font-weight: 700;
    color: #15503F;
    font-size: 15px;
}

.badge {
    display: inline-block;
    padding: 3px 8px;
    border-radius: 6px;
    font-size: 11.5px;
    font-weight: 600;
    text-align: center;
    white-space: nowrap;
}

.metric-box {
    background: #ffffff;
    border: 1px solid #edf0ed;
    border-radius: 12px;
    padding: 12px 14px;
    text-align: left;
    box-shadow: 0 2px 5px rgba(0,0,0,0.02);
}
.metric-label {
    font-size: 12px;
    color: #6b7280;
    font-weight: 500;
}
.metric-value {
    font-size: 19px;
    font-weight: 700;
    margin-top: 3px;
}

.product-item-title {
    font-size: 13px;
    font-weight: 700;
    color: #15503F;
    margin-bottom: 6px;
}

/* Khung hiển thị thông tin gửi hàng nhanh */
.quick-order-card {
    background-color: #F8FAF9;
    border: 1px solid #D1E3DC;
    border-left: 4px solid #15503F;
    border-radius: 10px;
    padding: 14px 18px;
    font-family: inherit;
    font-size: 13.8px;
    line-height: 1.7;
    color: #1F2937;
    margin-bottom: 10px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
}
.quick-order-card .label {
    font-weight: 700;
    color: #111827;
    display: inline-block;
    min-width: 105px;
}
.quick-order-card .value {
    color: #374151;
    font-weight: 400;
}

/* TỐI ƯU KHOẢNG CÁCH FORM LÊN ĐƠN GỌN GÀNG */
.order-item-box {
    background-color: #ffffff;
    border: 1px solid #E5E7EB;
    border-radius: 10px;
    padding: 10px 14px;
    margin-bottom: 8px;
}
.order-item-title {
    font-size: 13px;
    font-weight: 700;
    color: #15503F;
    margin-bottom: 4px;
}
.order-live-total-card {
    background-color: #E2EDE8;
    border: 1.5px solid #15503F;
    border-radius: 12px;
    padding: 12px 16px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin: 12px 0;
}
.order-live-total-title {
    font-size: 13.5px;
    font-weight: 700;
    color: #15503F;
}
.order-live-total-value {
    font-size: 21px;
    font-weight: 800;
    color: #D92B2B;
}

/* Thanh điều hướng */
div[class*="st-key-vifex_nav"] {
    margin-bottom: 12px;
}
div[class*="st-key-vifex_nav"] button {
    padding: 8px 4px !important;
    font-size: 13px !important;
    border-radius: 10px !important;
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
    border-radius: 10px !important;
    padding: 10px 18px !important;
    font-weight: 600 !important;
}

/* STYLING RIÊNG CHO TỪNG THẺ STATS NỀN MÀU / SỐ TO / CHỮ RÕ */
.card-wrapper {
    position: relative;
    border-radius: 14px;
    padding: 8px 10px;
    text-align: left;
    transition: all 0.15s ease;
    cursor: pointer;
    box-shadow: 0 2px 6px rgba(0,0,0,0.02);
    min-height: 66px;
    display: flex;
    flex-direction: column;
    justify-content: center;
}
.card-wrapper:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 10px rgba(0,0,0,0.06);
}
.card-num {
    font-size: 22px;
    font-weight: 800;
    line-height: 1.1;
}
.card-title {
    font-size: 12px;
    font-weight: 600;
    margin-top: 3px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

div[class*="st-key-btn_filter_"] {
    margin-top: -65px;
    opacity: 0;
}
div[class*="st-key-btn_filter_"] button {
    height: 65px !important;
    width: 100% !important;
    cursor: pointer !important;
}

/* Mobile */
@media screen and (max-width: 768px) {
    .block-container {
        max-width: 100% !important;
        padding-top: 4.5rem !important;
        padding-left: 0.6rem !important;
        padding-right: 0.6rem !important;
    }

    .vifex-banner {
        padding: 12px 14px !important;
        border-radius: 12px !important;
    }
    .vifex-banner .main-title {
        font-size: 18px !important;
    }

    .vifex-banner-logo {
        width: 56px !important;
        height: 56px !important;
        border-radius: 10px !important;
        padding: 4px !important;
        margin-left: 8px !important;
    }

    div[class*="st-key-vifex_nav"] [data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 4px !important;
    }
    div[class*="st-key-vifex_nav"] [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
        width: 33.33% !important;
        min-width: 33.33% !important;
        flex: 1 1 33.33% !important;
        padding: 0 !important;
    }
    div[class*="st-key-vifex_nav"] button {
        padding: 6px 2px !important;
        font-size: 11px !important;
        height: 38px !important;
        white-space: nowrap !important;
    }
    
    .card-num {
        font-size: 18px !important;
    }
    .card-title {
        font-size: 10.5px !important;
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


def order_status_badge_html(status):
    st_clean = safe_str(status)
    if "lên đơn" in st_clean.lower():
        return f'<span class="badge" style="background:#EEF6FF;color:#1D4ED8;border:1px solid #B9D7FE;">Lên đơn</span>'
    elif "gửi kho" in st_clean.lower():
        return f'<span class="badge" style="background:#F5EEFF;color:#7C3AED;border:1px solid #E0CFFE;">Gửi kho</span>'
    elif any(k in st_clean.lower() for k in ["đang giao", "giao hàng"]):
        return f'<span class="badge" style="background:#FFF6DA;color:#D97706;border:1px solid #FDE08B;">Đang giao</span>'
    elif any(k in st_clean.lower() for k in ["đã giao", "đã nhận"]):
        return f'<span class="badge" style="background:#E6F4EA;color:#065F46;border:1px solid #BBE1C7;">Đã giao</span>'
    return f'<span class="badge" style="background:{GRAY_BG};color:{GRAY_TEXT}">{status}</span>'


def payment_status_badge_html(status):
    st_clean = safe_str(status).strip()
    s_low = st_clean.lower()
    
    if "chưa tt - nháp" in s_low:
        return f'<span class="badge" style="background:#FFF1F2;color:#BE123C;border:1px solid #FECDD3;">Chưa TT - Nháp</span>'
    elif "chưa tt - đã vat" in s_low or "chưa tt - vat" in s_low:
        return f'<span class="badge" style="background:#FDF2F8;color:#DB2777;border:1px solid #FBCFE8;">Chưa TT - Đã VAT</span>'
    elif "tt - chưa vat" in s_low:
        return f'<span class="badge" style="background:#FEF3C7;color:#D97706;border:1px solid #FDE68A;">TT - Chưa VAT</span>'
    elif "tt - nháp vat" in s_low:
        return f'<span class="badge" style="background:#FFF0E5;color:#EA580C;border:1px solid #FED7AA;">TT - Nháp VAT</span>'
    elif "không vat" in s_low:
        return f'<span class="badge" style="background:#DCFAF4;color:#0D9488;border:1px solid #99F6E4;">TT - không VAT</span>'
    elif "tt - đã vat" in s_low:
        return f'<span class="badge" style="background:#E5F0EC;color:#2D6A4F;border:1px solid #C4E3D7;">TT - Đã VAT</span>'
    elif any(k in s_low for k in ["chưa thanh toán", "chưa tt"]):
        return f'<span class="badge" style="background:#FFEBEF;color:#E11D48;border:1px solid #FECDD3;">Chưa TT</span>'
    return f'<span class="badge" style="background:{GRAY_BG};color:{GRAY_TEXT}">{status}</span>'


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

    if "Trang_thai_TT" not in don_hang_df.columns:
        def map_payment_status(r):
            st_val = safe_str(r.get("Trang_thai"))
            for p in PAYMENT_STATUSES:
                if p.lower() == st_val.lower():
                    return p
            for p in PAYMENT_STATUSES:
                if p.lower() in st_val.lower():
                    return p
            return "Chưa thanh toán"
        don_hang_df["Trang_thai_TT"] = don_hang_df.apply(map_payment_status, axis=1)
    else:
        don_hang_df["Trang_thai_TT"] = don_hang_df["Trang_thai_TT"].fillna("Chưa thanh toán")

    def normalize_order_status(st_val):
        s = safe_str(st_val).lower()
        if "lên đơn" in s:
            return "Lên đơn"
        elif "gửi kho" in s:
            return "Gửi kho"
        elif any(k in s for k in ["đang giao", "giao hàng"]):
            return "Đang giao"
        elif any(k in s for k in ["đã giao", "đã nhận"]):
            return "Đã giao"
        return "Lên đơn"
    
    don_hang_df["Trang_thai_Don"] = don_hang_df["Trang_thai"].apply(normalize_order_status)

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
        if str(v).strip() == str(code).strip():
            return i
    return None


def update_order_both_statuses(ma_don, new_order_status, new_payment_status):
    today = (datetime.utcnow() + timedelta(hours=7)).strftime("%Y-%m-%d")
    don_hang_ws = get_ws("Don_hang")
    row = find_row_by_code(don_hang_ws, ma_don, col_index=1)
    if row:
        don_hang_ws.update_cell(row, 5, new_order_status)
        don_hang_ws.update_cell(row, 8, today)
        try:
            headers = don_hang_ws.row_values(1)
            if "Trang_thai_TT" in headers:
                col_k = headers.index("Trang_thai_TT") + 1
                don_hang_ws.update_cell(row, col_k, new_payment_status)
            else:
                don_hang_ws.update_cell(1, 11, "Trang_thai_TT")
                don_hang_ws.update_cell(row, 11, new_payment_status)
        except Exception:
            pass

    ctdh_ws = get_ws("Chi_tiet_don_hang")
    ma_don_col = ctdh_ws.col_values(2)
    for i, v in enumerate(ma_don_col, start=1):
        if str(v).strip() == str(ma_don).strip():
            ctdh_ws.update_cell(i, 10, new_order_status)
    refresh()


def delete_order_completely(ma_don):
    try:
        ctdh_ws = get_ws("Chi_tiet_don_hang")
        col_vals = ctdh_ws.col_values(2)
        for i in range(len(col_vals), 0, -1):
            if str(col_vals[i - 1]).strip() == str(ma_don).strip():
                ctdh_ws.delete_rows(i)
    except Exception:
        pass

    try:
        don_hang_ws = get_ws("Don_hang")
        row = find_row_by_code(don_hang_ws, ma_don, col_index=1)
        if row:
            don_hang_ws.delete_rows(row)
    except Exception:
        pass

    delete_vat_record(ma_don)
    refresh()


# ---------------------------------------------------------------------------
# XỬ LÝ TẢI TỰ ĐỘNG LÊN GOOGLE DRIVE
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
            if str(r.get("Ma_don")).strip() == str(ma_don).strip():
                link = r.get("Link_Drive") or r.get("Link_file_PDF") or ""
                if link and str(link).startswith("http"):
                    return link
    except Exception:
        pass
    return None


def upload_vat_directly_to_drive(ma_don, ma_kh, uploaded_file):
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
            rows = ws.col_values(2)
            target_row = None
            for i, val in enumerate(rows, start=1):
                if str(val).strip() == str(ma_don).strip():
                    target_row = i
                    break
            
            today_str = (datetime.utcnow() + timedelta(hours=7)).strftime("%Y-%m-%d %H:%M")
            if target_row:
                ws.update_cell(target_row, 4, uploaded_file.name)
                ws.update_cell(target_row, 5, today_str)
                ws.update_cell(target_row, 6, file_url)
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
            if str(val).strip() == str(ma_don).strip():
                ws.delete_rows(i)
                break
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Ảnh phiếu xuất đơn hàng (MST góc phải + QR Vietinbank)
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
    raise FileNotFoundError("Không tìm thấy font NotoSans.ttf.")


def draw_bold(draw, pos, text, font, fill):
    x, y = pos
    draw.text((x, y), text, font=font, fill=fill)
    draw.text((x + 0.6, y), text, font=font, fill=fill)


def draw_wrapped_text(d, pos, text, font, fill, max_width, line_spacing=6):
    x, y = pos
    words = text.split(" ")
    lines = []
    curr = ""
    for w in words:
        test = (curr + " " + w).strip()
        bbox = d.textbbox((0, 0), test, font=font)
        w_px = bbox[2] - bbox[0]
        if w_px <= max_width:
            curr = test
        else:
            if curr:
                lines.append(curr)
            curr = w
    if curr:
        lines.append(curr)
    
    for line in lines:
        d.text((x, y), line, font=font, fill=fill)
        bbox = d.textbbox((0, 0), line, font=font)
        h_px = bbox[3] - bbox[1]
        y += h_px + line_spacing
    return y


def generate_order_slip(ma_don, order_row, items_df, khach_hang_row):
    W = 860
    row_h = 34
    
    f_title = get_font(26)
    f_sub = get_font(14)
    f_h = get_font(16)
    f_n = get_font(14)
    f_small = get_font(13)

    H = 540 + row_h * (len(items_df) + 2) + 140

    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    d.rectangle([0, 0, W, 80], fill=GREEN)
    draw_bold(d, (24, 16), "VIFEX", f_title, "white")
    d.text((24, 50), "PHIẾU XUẤT ĐƠN HÀNG (không phải hóa đơn VAT)", font=f_sub, fill="white")

    y = 100
    ten_cty = safe_str(khach_hang_row.get("Ten_cong_ty_GPKD")) or safe_str(khach_hang_row.get("Ten_NPP"))
    ten_npp = safe_str(khach_hang_row.get("Ten_NPP"))
    dia_chi_giao = safe_str(khach_hang_row.get("Dia_chi_giao_phu")) or safe_str(khach_hang_row.get("Dia_chi_GPKD"))
    sdt = safe_str(khach_hang_row.get("SDT_phu"))
    mst = safe_str(khach_hang_row.get("MST"))

    draw_bold(d, (24, y), f"Mã đơn: {ma_don}", f_h, "black")
    
    if mst:
        mst_text = f"MST: {mst}"
        bbox_mst = d.textbbox((0, 0), mst_text, font=f_n)
        mst_w = bbox_mst[2] - bbox_mst[0]
        draw_bold(d, (W - 24 - mst_w, y + 1), mst_text, f_n, "black")
        
    y += 26
    d.text((24, y), f"Ngày lên đơn: {order_row.get('Ngay_len_don')}", font=f_n, fill="black"); y += 24
    
    if ten_cty:
        y = draw_wrapped_text(d, (24, y), f"Khách hàng: {ten_cty}", f_n, "black", max_width=810)
    if ten_npp:
        y = draw_wrapped_text(d, (24, y), f"NHÀ PHÂN PHỐI : {ten_npp}", f_n, "black", max_width=810)
    if dia_chi_giao:
        y = draw_wrapped_text(d, (24, y), f"Địa chỉ giao: {dia_chi_giao}", f_n, "black", max_width=810)
    if sdt:
        d.text((24, y), f"SĐT người nhận: {sdt}", font=f_n, fill="black"); y += 24
    
    y += 8
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
    y += 32

    d.text((24, y), f"Hình thức thanh toán: {order_row.get('Hinh_thuc_thanh_toan', '')}", font=f_n, fill="black")
    y += 22
    
    ghi_chu = safe_str(order_row.get("Ghi_chu_thanh_toan"))
    if ghi_chu:
        y = draw_wrapped_text(d, (24, y), f"Ghi chú: {ghi_chu}", f_n, "black", max_width=810)

    # KHU VỰC THÔNG TIN THANH TOÁN & MÃ QR
    y += 10
    d.line([24, y, W - 24, y], fill="#eee", width=1); y += 14
    
    qr_img = get_qr_bank_image()
    if qr_img:
        qr_size = 120
        qr_resized = qr_img.resize((qr_size, qr_size), Image.Resampling.LANCZOS)
        img.paste(qr_resized, (24, y), qr_resized if qr_resized.mode == "RGBA" else None)
        
        tx = 24 + qr_size + 18
        ty = y + 10
        draw_bold(d, (tx, ty), "THÔNG TIN THANH TOÁN CHUYỂN KHOẢN:", f_h, GREEN); ty += 26
        draw_bold(d, (tx, ty), "STK Vietinbank: 116003017106", f_n, "black"); ty += 22
        d.text((tx, ty), "Chủ TK: Công ty TNHH VIFEX", font=f_n, fill="black"); ty += 22
        d.text((tx, ty), "(Quét mã QR bên cạnh để chuyển khoản chính xác và nhanh chóng)", font=f_small, fill="#4B5563")
        y += qr_size + 15
    else:
        draw_bold(d, (24, y), "THÔNG TIN THANH TOÁN: STK Vietinbank 116003017106 - Công ty TNHH VIFEX", f_n, GREEN)
        y += 28

    final_h = max(y + 20, 320)
    cropped_img = img.crop((0, 0, W, final_h))

    buf = io.BytesIO()
    cropped_img.save(buf, format="PNG")
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
        don_hang_df[["Ma_don", "Ngay_len_don", "Ma_KH", "Sale_phu_trach", "Trang_thai", "Trang_thai_Don", "Trang_thai_TT"]],
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
if "home_filter_category" not in st.session_state:
    st.session_state.home_filter_category = None
if "home_filter_value" not in st.session_state:
    st.session_state.home_filter_value = None

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


def render_order_detail_inline(ma_don):
    order_rows = don_hang_df[don_hang_df["Ma_don"] == ma_don]
    if order_rows.empty:
        st.error("Không tìm thấy đơn hàng.")
        return
    order_row = order_rows.iloc[0]
    kh_rows = khach_hang_df[khach_hang_df["Ma_KH"] == order_row["Ma_KH"]]
    kh_row = kh_rows.iloc[0] if not kh_rows.empty else {}
    ma_kh = order_row["Ma_KH"]

    items = ctdh_df[ctdh_df["Ma_don"] == ma_don].copy()
    items = items.merge(san_pham_df[["Ma_SP", "Ten_SP", "Nhom_danh_muc"]], on="Ma_SP", how="left")

    ten_cty = safe_str(kh_row.get("Ten_cong_ty_GPKD")) or safe_str(kh_row.get("Ten_NPP"))
    ten_npp = safe_str(kh_row.get("Ten_NPP"))
    dia_chi_giao = safe_str(kh_row.get("Dia_chi_giao_phu")) or safe_str(kh_row.get("Dia_chi_GPKD"))
    sdt_nhan = safe_str(kh_row.get("SDT_phu"))
    ten_nguoi_nhan = safe_str(kh_row.get("Ten_nguoi_nhan_phu"))

    current_order_st = order_row.get("Trang_thai_Don", "Lên đơn")
    current_pay_st = order_row.get("Trang_thai_TT", "Chưa thanh toán")

    with st.container(border=True):
        st.markdown(f"#### Chi tiết đơn: `{ma_don}`")
        
        col_b1, col_b2 = st.columns([1, 1])
        with col_b1:
            st.markdown(f"**Giao hàng:** {order_status_badge_html(current_order_st)}", unsafe_allow_html=True)
        with col_b2:
            st.markdown(f"**Thanh toán:** {payment_status_badge_html(current_pay_st)}", unsafe_allow_html=True)

        st.write(f"**Khách hàng:** {ten_cty}")
        st.write(f"**NHÀ PHÂN PHỐI :** {ten_npp}")
        st.write(f"**Ngày lên đơn:** {order_row['Ngay_len_don']}")
        st.write(f"**Hình thức thanh toán:** {order_row['Hinh_thuc_thanh_toan']}")
        if safe_str(order_row.get("Ghi_chu_thanh_toan")):
            st.write(f"**Ghi chú:** {order_row['Ghi_chu_thanh_toan']}")

        # -------------------------------------------------------------------
        # KHU VỰC VĂN BẢN ĐƠN HÀNG NHANH
        # -------------------------------------------------------------------
        nhom_dm_list = items["Nhom_danh_muc"].dropna().unique().tolist()
        nhom_dm_str = ", ".join(nhom_dm_list) if nhom_dm_list else "Hàng hóa"
        
        nguoi_nhan_parts = []
        if ten_nguoi_nhan:
            nguoi_nhan_parts.append(ten_nguoi_nhan)
        if sdt_nhan:
            nguoi_nhan_parts.append(sdt_nhan)
        if dia_chi_giao:
            nguoi_nhan_parts.append(dia_chi_giao)
        nguoi_nhan_str = " - ".join(nguoi_nhan_parts) if nguoi_nhan_parts else "Chưa có thông tin nhận"

        tong_sl_giao = int(items["SL_dat"].sum() + items["Tang"].sum())
        ghi_chu_don = safe_str(order_row.get("Ghi_chu_thanh_toan")) or "Không có"

        with st.expander("📋 **Xem nhanh thông tin gửi hàng (dạng chữ):**", expanded=True):
            html_text_card = f"""
            <div class="quick-order-card">
                <div><span class="label">Đặt hàng:</span> <span class="value">{nhom_dm_str}</span></div>
                <div><span class="label">NPP:</span> <span class="value">{ten_npp}</span></div>
                <div><span class="label">Ngày lên đơn:</span> <span class="value">{order_row.get('Ngay_len_don')}</span></div>
                <div><span class="label">Người nhận:</span> <span class="value">{nguoi_nhan_str}</span></div>
                <div><span class="label">Số lượng:</span> <span class="value">{tong_sl_giao}</span></div>
                <div><span class="label">Ghi chú:</span> <span class="value">{ghi_chu_don}</span></div>
            </div>
            """
            st.markdown(html_text_card, unsafe_allow_html=True)

        items["Tong_tien_hang"] = items["SL_dat"] * items["Don_gia_ap_dung"]
        
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
        
        # CẬP NHẬT 2 TRẠNG THÁI
        st.markdown("<div style='font-size:13px;font-weight:700;color:#15503F;margin:8px 0 4px 0;'>CẬP NHẬT TIẾN ĐỘ ĐƠN HÀNG & THANH TOÁN</div>", unsafe_allow_html=True)
        c_st1, c_st2, c_st3 = st.columns([1.5, 1.5, 1])
        
        with c_st1:
            idx_order = ORDER_STATUSES.index(current_order_st) if current_order_st in ORDER_STATUSES else 0
            new_order_status = st.selectbox("📦 Tiến độ giao hàng", ORDER_STATUSES, index=idx_order, key=f"sel_order_{ma_don}")
        
        with c_st2:
            idx_pay = PAYMENT_STATUSES.index(current_pay_st) if current_pay_st in PAYMENT_STATUSES else 0
            new_payment_status = st.selectbox("💳 Thanh toán & VAT", PAYMENT_STATUSES, index=idx_pay, key=f"sel_pay_{ma_don}")
            
        with c_st3:
            st.write("")
            st.write("")
            if st.button("💾 Lưu thay đổi", key=f"save_status_{ma_don}", type="primary", use_container_width=True):
                update_order_both_statuses(ma_don, new_order_status, new_payment_status)
                st.success("Đã cập nhật!")
                st.rerun()

        # NÚT PHIẾU XUẤT & VAT
        c_action1, c_action2 = st.columns(2)
        with c_action1:
            preview_key = f"show_preview_{ma_don}"
            if preview_key not in st.session_state:
                st.session_state[preview_key] = False
            btn_label = "🙈 Đóng phiếu xuất" if st.session_state[preview_key] else "🖼️ Xem Phiếu xuất đơn"
            if st.button(btn_label, key=f"btn_toggle_preview_{ma_don}", use_container_width=True):
                st.session_state[preview_key] = not st.session_state[preview_key]
                st.rerun()
                
        with c_action2:
            drive_vat_url = get_vat_link_from_sheet(ma_don)
            if drive_vat_url:
                st.link_button("📄 Mở File Hóa đơn VAT (Drive)", url=drive_vat_url, use_container_width=True)
            else:
                st.button("☁️ Chưa có File VAT", disabled=True, use_container_width=True)

        if st.session_state.get(f"show_preview_{ma_don}", False):
            png_bytes = generate_order_slip(ma_don, order_row, items, kh_row if isinstance(kh_row, dict) else kh_row.to_dict())
            with st.container(border=True):
                st.image(png_bytes, use_container_width=True)
                st.download_button("📥 Tải ảnh phiếu xuất (PNG)", data=png_bytes,
                                    file_name=f"{ma_don}_phieu_xuat.png", mime="image/png",
                                    key=f"dl_inside_{ma_don}", use_container_width=True)

        c_exp1, c_exp2 = st.columns(2)
        with c_exp1:
            with st.expander("☁️ **Quản lý Hóa đơn VAT**", expanded=False):
                drive_vat_url = get_vat_link_from_sheet(ma_don)
                if drive_vat_url:
                    st.info(f"✅ Đơn hàng đã có file lưu trên Google Drive.")
                    st.markdown(f"👉 [Mở xem file trên Google Drive]({drive_vat_url})")
                    if st.button("🗑️ Xóa hóa đơn này", key=f"btn_del_vat_{ma_don}"):
                        delete_vat_record(ma_don)
                        st.success(f"Đã xóa hóa đơn VAT của đơn **{ma_don}**.")
                        st.rerun()
                else:
                    st.caption("Kéo thả file PDF để tự động lưu vào Google Drive:")

                uploaded_vat = st.file_uploader(
                    f"Chọn tệp hóa đơn cho đơn {ma_don}", 
                    type=["pdf", "png", "jpg", "jpeg"], 
                    key=f"vat_uploader_{ma_don}",
                    label_visibility="collapsed"
                )
                
                if uploaded_vat is not None:
                    if st.button("⬆️ Tải lên Google Drive", key=f"btn_save_vat_{ma_don}", type="primary"):
                        with st.spinner("Đang tải file..."):
                            try:
                                file_drive_link = upload_vat_directly_to_drive(ma_don, ma_kh, uploaded_vat)
                                st.success(f"Đã tải thành công file lên Google Drive!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Lỗi: {e}")
                                
        with c_exp2:
            with st.expander("⚠️ **Xóa đơn hàng này**", expanded=False):
                st.caption(f"Xóa vĩnh viễn đơn **{ma_don}** và toàn bộ chi tiết trên Google Sheets.")
                confirm_del = st.checkbox(f"Xác nhận xóa {ma_don}", key=f"chk_del_{ma_don}")
                if confirm_del:
                    if st.button("🗑️ Xóa vĩnh viễn", key=f"btn_confirm_del_order_{ma_don}", type="primary"):
                        with st.spinner("Đang xóa đơn hàng..."):
                            delete_order_completely(ma_don)
                            st.session_state.selected_order = None
                            st.success(f"Đã xóa đơn {ma_don}!")
                            st.rerun()

        st.write("")
        if st.button("← Đóng xem chi tiết", key=f"close_{ma_don}", use_container_width=True):
            st.session_state.selected_order = None
            st.rerun()


# ---------------------------------------------------------------------------
# 1. TRANG CHỦ (ĐẦY ĐỦ 7 TRẠNG THÁI THANH TOÁN + MÀU SẮC ĐẸP BẢN CŨ)
# ---------------------------------------------------------------------------
if nav == "🏠 Trang chủ":
    pending_delivery = don_hang_df[don_hang_df["Trang_thai_Don"].isin(["Lên đơn", "Gửi kho", "Đang giao"])] if not don_hang_df.empty else pd.DataFrame()
    so_don_pending = len(pending_delivery)
    banner("Trang chủ", subtitle="Xin chào, Coco", highlight_text=f"{so_don_pending} đơn đang vận hành")

    if don_hang_df.empty:
        st.info("Chưa có đơn hàng nào.")
    else:
        counts_order = don_hang_df["Trang_thai_Don"].value_counts()
        counts_pay = don_hang_df["Trang_thai_TT"].value_counts()

        # HÀNG 1: TIẾN ĐỘ ĐƠN HÀNG (4 Ô)
        st.markdown("<div style='font-size:13px;font-weight:700;color:#15503F;margin-bottom:6px;'>📦 TIẾN ĐỘ VẬN HÀNH ĐƠN HÀNG (Nhấn ô để lọc danh sách)</div>", unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)

        order_metric_defs = [
            ("Lên đơn", int(counts_order.get('Lên đơn', 0)), "#EEF6FF", "#B9D7FE", "#1D4ED8", "#1E3A8A", c1),
            ("Gửi kho", int(counts_order.get('Gửi kho', 0)), "#F5EEFF", "#E0CFFE", "#7C3AED", "#374151", c2),
            ("Đang giao", int(counts_order.get('Đang giao', 0)), "#FFF6DA", "#FDE08B", "#D97706", "#374151", c3),
            ("Đã giao", int(counts_order.get('Đã giao', 0)), "#E6F4EA", "#BBE1C7", "#065F46", "#1F2937", c4)
        ]

        for st_name, count_val, bg_c, bdr_c, num_c, label_c, col_obj in order_metric_defs:
            is_active = (st.session_state.home_filter_category == "order" and st.session_state.home_filter_value == st_name)
            border_style = f"3px solid {num_c}" if is_active else f"1.5px solid {bdr_c}"
            active_badge = " ✔️" if is_active else ""
            
            with col_obj:
                html_card = f"""
                <div class="card-wrapper" style="background-color: {bg_c}; border: {border_style};">
                    <div class="card-num" style="color: {num_c};">{count_val}</div>
                    <div class="card-title" style="color: {label_c};">{st_name}{active_badge}</div>
                </div>
                """
                st.markdown(html_card, unsafe_allow_html=True)
                
                with st.container(key=f"btn_filter_order_{st_name}"):
                    if st.button(st_name, key=f"btn_act_order_{st_name}", use_container_width=True):
                        if is_active:
                            st.session_state.home_filter_category = None
                            st.session_state.home_filter_value = None
                        else:
                            st.session_state.home_filter_category = "order"
                            st.session_state.home_filter_value = st_name
                        st.session_state.selected_order = None
                        st.rerun()

        st.write("")
        # HÀNG 2: TIẾN ĐỘ THANH TOÁN & VAT (CHIA ĐỀU GỌN GÀNG)
        st.markdown("<div style='font-size:13px;font-weight:700;color:#15503F;margin-bottom:6px;'>💳 TIẾN ĐỘ THANH TOÁN & HÓA ĐƠN VAT (Nhấn ô để lọc danh sách)</div>", unsafe_allow_html=True)
        
        # Hàng 2.1: Nhóm Chưa thanh toán (3 ô)
        p_row1_cols = st.columns(3)
        pay_row1_defs = [
            ("Chưa thanh toán", "Chưa TT", int(counts_pay.get('Chưa thanh toán', 0)), "#FFEBEF", "#FECDD3", "#E11D48", "#881337", p_row1_cols[0]),
            ("Chưa TT - Nháp", "Chưa TT - Nháp", int(counts_pay.get('Chưa TT - Nháp', 0)), "#FFF1F2", "#FECDD3", "#BE123C", "#881337", p_row1_cols[1]),
            ("Chưa TT - Đã VAT", "Chưa TT - Đã VAT", int(counts_pay.get('Chưa TT - Đã VAT', 0)), "#FDF2F8", "#FBCFE8", "#DB2777", "#831843", p_row1_cols[2]),
        ]

        for full_pay_st, short_name, count_val, bg_c, bdr_c, num_c, label_c, col_obj in pay_row1_defs:
            is_active = (st.session_state.home_filter_category == "pay" and st.session_state.home_filter_value == full_pay_st)
            border_style = f"3px solid {num_c}" if is_active else f"1.5px solid {bdr_c}"
            active_badge = " ✔️" if is_active else ""
            
            with col_obj:
                html_card = f"""
                <div class="card-wrapper" style="background-color: {bg_c}; border: {border_style};">
                    <div class="card-num" style="color: {num_c};">{count_val}</div>
                    <div class="card-title" style="color: {label_c};">{short_name}{active_badge}</div>
                </div>
                """
                st.markdown(html_card, unsafe_allow_html=True)
                
                with st.container(key=f"btn_filter_pay_{short_name}"):
                    if st.button(short_name, key=f"btn_act_pay_{short_name}", use_container_width=True):
                        if is_active:
                            st.session_state.home_filter_category = None
                            st.session_state.home_filter_value = None
                        else:
                            st.session_state.home_filter_category = "pay"
                            st.session_state.home_filter_value = full_pay_st
                        st.session_state.selected_order = None
                        st.rerun()

        st.write("")
        # Hàng 2.2: Nhóm Đã thanh toán (4 ô)
        p_row2_cols = st.columns(4)
        pay_row2_defs = [
            ("TT - Chưa VAT", "TT - Chưa VAT", int(counts_pay.get('TT - Chưa VAT', 0)), "#FEF3C7", "#FDE68A", "#D97706", "#92400E", p_row2_cols[0]),
            ("TT - Nháp VAT", "TT - Nháp VAT", int(counts_pay.get('TT - Nháp VAT', 0)), "#FFF0E5", "#FED7AA", "#EA580C", "#431407", p_row2_cols[1]),
            ("TT - không VAT", "TT - không VAT", int(counts_pay.get('TT - không VAT', 0)), "#DCFAF4", "#99F6E4", "#0D9488", "#0F766E", p_row2_cols[2]),
            ("TT - Đã VAT", "TT - Đã VAT", int(counts_pay.get('TT - Đã VAT', 0)), "#E5F0EC", "#C4E3D7", "#2D6A4F", "#1F2937", p_row2_cols[3])
        ]

        for full_pay_st, short_name, count_val, bg_c, bdr_c, num_c, label_c, col_obj in pay_row2_defs:
            is_active = (st.session_state.home_filter_category == "pay" and st.session_state.home_filter_value == full_pay_st)
            border_style = f"3px solid {num_c}" if is_active else f"1.5px solid {bdr_c}"
            active_badge = " ✔️" if is_active else ""
            
            with col_obj:
                html_card = f"""
                <div class="card-wrapper" style="background-color: {bg_c}; border: {border_style};">
                    <div class="card-num" style="color: {num_c};">{count_val}</div>
                    <div class="card-title" style="color: {label_c};">{short_name}{active_badge}</div>
                </div>
                """
                st.markdown(html_card, unsafe_allow_html=True)
                
                with st.container(key=f"btn_filter_pay_{short_name}"):
                    if st.button(short_name, key=f"btn_act_pay_{short_name}", use_container_width=True):
                        if is_active:
                            st.session_state.home_filter_category = None
                            st.session_state.home_filter_value = None
                        else:
                            st.session_state.home_filter_category = "pay"
                            st.session_state.home_filter_value = full_pay_st
                        st.session_state.selected_order = None
                        st.rerun()

        st.write("")

        # -------------------------------------------------------------------
        # DANH SÁCH ĐƠN HÀNG THEO BỘ LỌC KHI NHẤN VÀO Ô
        # -------------------------------------------------------------------
        display_orders_df = don_hang_df.copy()
        list_title = "ĐƠN HÀNG CẦN XỬ LÝ GẦN ĐÂY"

        if st.session_state.home_filter_category == "order":
            st_val = st.session_state.home_filter_value
            display_orders_df = display_orders_df[display_orders_df["Trang_thai_Don"] == st_val]
            list_title = f"DANH SÁCH ĐƠN: {st_val.upper()} ({len(display_orders_df)} đơn)"
        elif st.session_state.home_filter_category == "pay":
            st_val = st.session_state.home_filter_value
            display_orders_df = display_orders_df[display_orders_df["Trang_thai_TT"] == st_val]
            list_title = f"DANH SÁCH ĐƠN: {st_val.upper()} ({len(display_orders_df)} đơn)"
        else:
            display_orders_df = display_orders_df.head(6)

        col_head1, col_head2 = st.columns([3, 1])
        with col_head1:
            st.markdown(f"<div style='font-size:14px;font-weight:700;color:#15503F;margin-bottom:8px;'>📋 {list_title}</div>", unsafe_allow_html=True)
        with col_head2:
            if st.session_state.home_filter_category is not None:
                if st.button("✖️ Bỏ lọc", key="btn_clear_home_filter", use_container_width=True):
                    st.session_state.home_filter_category = None
                    st.session_state.home_filter_value = None
                    st.session_state.selected_order = None
                    st.rerun()

        display_orders_df = display_orders_df.sort_values("Ma_don", ascending=False)
        
        if display_orders_df.empty:
            st.info("Không có đơn hàng nào thuộc trạng thái này.")
        
        for _, r in display_orders_df.iterrows():
            kh = khach_hang_df[khach_hang_df["Ma_KH"] == r["Ma_KH"]]
            ten_kh = safe_str(kh.iloc[0]["Ten_NPP"]) if not kh.empty else safe_str(r["Ma_KH"])
            if not ten_kh:
                ten_kh = "Chưa có tên NPP"
                
            is_expanded = (st.session_state.selected_order == r["Ma_don"])
            
            with st.container(border=True):
                col_left, col_right = st.columns([2.6, 1.4])
                with col_left:
                    st.markdown(f"""
                    <div class="order-code-compact">{r['Ma_don']}</div>
                    <div class="order-cust-compact">NPP {ten_kh}</div>
                    <div class="order-value-compact">{money(order_total(r['Ma_don']))}</div>
                    """, unsafe_allow_html=True)
                with col_right:
                    st.markdown(f"""
                    <div style='text-align:right;margin-bottom:4px;'>
                        {order_status_badge_html(r.get('Trang_thai_Don'))} {payment_status_badge_html(r.get('Trang_thai_TT'))}
                    </div>
                    """, unsafe_allow_html=True)
                    btn_text = "Đóng chi tiết" if is_expanded else "Xem chi tiết"
                    if st.button(btn_text, key=f"home_view_{r['Ma_don']}", use_container_width=True):
                        st.session_state.selected_order = None if is_expanded else r["Ma_don"]
                        st.rerun()

            if is_expanded:
                render_order_detail_inline(r["Ma_don"])

    st.write("")
    if st.button("+ Tạo đơn mới", key="home_new_order", type="primary", use_container_width=True):
        st.session_state.nav = "➕ Lên đơn"
        st.rerun()

# ---------------------------------------------------------------------------
# 2. DANH SÁCH ĐƠN HÀNG (3 Ô LỌC TRÊN CÙNG 1 HÀNG: TIẾN ĐỘ, THANH TOÁN, NPP)
# ---------------------------------------------------------------------------
elif nav == "📦 Đơn hàng":
    banner("Danh sách đơn hàng")

    raw_npp_list = sorted([str(x).strip() for x in khach_hang_df["Ten_NPP"].dropna().unique() if str(x).strip()])
    list_npp_filter = ["Tất cả"] + raw_npp_list

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        filter_order_st = st.selectbox("Lọc tiến độ giao hàng", ["Tất cả"] + ORDER_STATUSES, key="filter_order_st")
    with col_f2:
        filter_pay_st = st.selectbox("Lọc thanh toán & VAT", ["Tất cả"] + PAYMENT_STATUSES, key="filter_pay_st")
    with col_f3:
        filter_npp = st.selectbox("Lọc theo Nhà phân phối", list_npp_filter, key="filter_npp")
    
    view_df = don_hang_df.copy()
    if filter_order_st != "Tất cả":
        view_df = view_df[view_df["Trang_thai_Don"] == filter_order_st]
    if filter_pay_st != "Tất cả":
        view_df = view_df[view_df["Trang_thai_TT"] == filter_pay_st]
    if filter_npp != "Tất cả":
        matching_kh_ids = khach_hang_df[khach_hang_df["Ten_NPP"] == filter_npp]["Ma_KH"].dropna().tolist()
        view_df = view_df[view_df["Ma_KH"].isin(matching_kh_ids)]
    
    view_df = view_df.sort_values("Ma_don", ascending=False)

    st.write("")
    if view_df.empty:
        st.info("Không có đơn hàng nào phù hợp với bộ lọc.")
    
    for _, r in view_df.iterrows():
        kh = khach_hang_df[khach_hang_df["Ma_KH"] == r["Ma_KH"]]
        ten_kh = safe_str(kh.iloc[0]["Ten_NPP"]) if not kh.empty else safe_str(r["Ma_KH"])
        if not ten_kh:
            ten_kh = "Chưa có tên NPP"
            
        is_expanded = (st.session_state.selected_order == r["Ma_don"])
        
        with st.container(border=True):
            col_info, col_action = st.columns([2.6, 1.4])
            with col_info:
                st.markdown(f"""
                <div class="order-code-compact">{r['Ma_don']}</div>
                <div class="order-cust-compact">NPP {ten_kh}</div>
                <div class="order-value-compact">{money(order_total(r['Ma_don']))}</div>
                """, unsafe_allow_html=True)
            with col_action:
                st.markdown(f"""
                <div style='text-align:right;margin-bottom:4px;'>
                    {order_status_badge_html(r.get('Trang_thai_Don'))} {payment_status_badge_html(r.get('Trang_thai_TT'))}
                </div>
                """, unsafe_allow_html=True)
                btn_text = "Đóng chi tiết" if is_expanded else "Xem chi tiết"
                if st.button(btn_text, key=f"list_view_{r['Ma_don']}", use_container_width=True):
                    st.session_state.selected_order = None if is_expanded else r["Ma_don"]
                    st.rerun()

        if is_expanded:
            render_order_detail_inline(r["Ma_don"])

# ---------------------------------------------------------------------------
# 3. LÊN ĐƠN HÀNG (KHOẢNG CÁCH THU GỌN + LIVE TỔNG GIÁ TRỊ ĐƠN HÀNG + HIỂN THỊ CHẤM HÀNG ĐƠN VỊ TIỀN)
# ---------------------------------------------------------------------------
elif nav == "➕ Lên đơn":
    banner("Lên đơn hàng")

    v = st.session_state.order_form_version
    ten_npp = st.selectbox("Khách hàng (NPP)", khach_hang_df["Ten_NPP"].dropna().tolist(), key=f"form_npp_{v}")
    kh_matches = khach_hang_df[khach_hang_df["Ten_NPP"] == ten_npp]
    kh_row = kh_matches.iloc[0] if not kh_matches.empty else {}
    ma_kh = safe_str(kh_row.get("Ma_KH"))
    sale_pt = safe_str(kh_row.get("Sale_phu_trach"))

    st.caption(f"Sale phụ trách: **{sale_pt}**")

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

        st.markdown("<div style='font-size:14px;font-weight:700;color:#15503F;margin:10px 0 6px 0;'>DANH SÁCH SẢN PHẨM</div>", unsafe_allow_html=True)

        line_items = []
        ten_sp_list = san_pham_df["Ten_SP"].dropna().tolist()
        estimated_total_order = 0.0
        
        for i in range(st.session_state.order_items_count):
            with st.container():
                st.markdown(f"""
                <div class="order-item-box">
                    <div class="order-item-title">Sản phẩm #{i+1}</div>
                </div>
                """, unsafe_allow_html=True)
                
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

                # Tính nhẩm tổng tiền tạm tính trực tiếp theo giá ngày chọn
                ma_sp_lookup = safe_str(san_pham_df[san_pham_df["Ten_SP"] == ten_sp].iloc[0]["Ma_SP"]) if not san_pham_df[san_pham_df["Ten_SP"] == ten_sp].empty else ""
                current_price = float(lookup_gia(ma_sp_lookup, ngay_len_don, lich_su_gia_df)) if ma_sp_lookup else 0.0
                row_thanh_tien = max(0.0, (float(sl_dat) * current_price) - float(chiet_khau))
                estimated_total_order += row_thanh_tien

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

        # Ô HIỂN THỊ TỔNG GIÁ TRỊ ĐƠN HÀNG DỰ KIẾN (TRỰC QUAN, RÕ RÀNG)
        st.markdown(f"""
        <div class="order-live-total-card">
            <span class="order-live-total-title">TỔNG GIÁ TRỊ ĐƠN HÀNG (DỰ KIẾN):</span>
            <span class="order-live-total-value">{money(estimated_total_order)}</span>
        </div>
        """, unsafe_allow_html=True)

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

        don_hang_row_data = [
            str(ma_don),
            str(ngay_len_don.strftime("%Y-%m-%d")),
            str(ma_kh),
            str(sale_pt),
            "Lên đơn",
            str(hinh_thuc_tt),
            str(ghi_chu_tt),
            str(ngay_len_don.strftime("%Y-%m-%d")),
            int(ngay_len_don.month),
            int(ngay_len_don.year),
            "Chưa thanh toán"
        ]
        
        don_hang_ws.append_row(don_hang_row_data, value_input_option="USER_ENTERED")

        tong_tien = 0
        for ten_sp, sl_dat, tang, chiet_khau in valid_items:
            ma_sp = safe_str(san_pham_df[san_pham_df["Ten_SP"] == ten_sp].iloc[0]["Ma_SP"])
            don_gia = float(lookup_gia(ma_sp, ngay_len_don, lich_su_gia_df))
            sl_dat_num = int(sl_dat)
            tang_num = float(tang)
            ck_num = float(chiet_khau)
            thanh_tien = float(sl_dat_num * don_gia - ck_num)
            tong_tien += thanh_tien
            ma_ctdh = next_code(ctdh_ws, 1, "CT", 5)
            
            ctdh_row_data = [
                str(ma_ctdh),
                str(ma_don),
                str(ma_sp),
                int(sl_dat_num),
                float(tang_num),
                float(don_gia),
                float(ck_num),
                float(thanh_tien),
                float(sl_dat_num + tang_num),
                "Lên đơn",
                int(ngay_len_don.month),
                int(ngay_len_don.year)
            ]
            ctdh_ws.append_row(ctdh_row_data, value_input_option="USER_ENTERED")

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

        valid = merged[merged["Trang_thai_Don"].isin(["Gửi kho", "Đang giao", "Đã giao"])].copy()
        valid["Ngay_len_don"] = pd.to_datetime(valid["Ngay_len_don"], errors="coerce")
        of_sale = valid[(valid["Sale_phu_trach"] == ma_nv) &
                         (valid["Ngay_len_don"].dt.month == thang) &
                         (valid["Ngay_len_don"].dt.year == nam)]

        doanh_thu = of_sale["Thanh_tien"].sum()
        doanh_thu_sau_vat = doanh_thu * (1 - 0.08)
        luong_co_ban = doanh_thu_sau_vat * ty_le_hh

        is_dat = any(kw in str(ten_nv).lower() for kw in ["đạt", "dat"])
        thuong_tan = 0
        thuong_duc = 0
        ten_nv_tan = ""
        ten_nv_duc = ""

        if is_dat:
            tan_rows = nhan_vien_df[nhan_vien_df["Ten_NV"].str.lower().str.contains("tân|tan", na=False)]
            if not tan_rows.empty:
                ma_tan = tan_rows.iloc[0]["Ma_NV"]
                ten_nv_tan = tan_rows.iloc[0]["Ten_NV"]
                of_tan = valid[(valid["Sale_phu_trach"] == ma_tan) &
                               (valid["Ngay_len_don"].dt.month == thang) &
                               (valid["Ngay_len_don"].dt.year == nam)]
                doanh_thu_tan = of_tan["Thanh_tien"].sum()
                doanh_thu_tan_sau_vat = doanh_thu_tan * (1 - 0.08)
                thuong_tan = doanh_thu_tan_sau_vat * 0.01

            duc_rows = nhan_vien_df[nhan_vien_df["Ten_NV"].str.lower().str.contains("đức|duc", na=False)]
            if not duc_rows.empty:
                ma_duc = duc_rows.iloc[0]["Ma_NV"]
                ten_nv_duc = duc_rows.iloc[0]["Ten_NV"]
                of_duc = valid[(valid["Sale_phu_trach"] == ma_duc) &
                               (valid["Ngay_len_don"].dt.month == thang) &
                               (valid["Ngay_len_don"].dt.year == nam)]
                doanh_thu_duc = of_duc["Thanh_tien"].sum()
                doanh_thu_duc_sau_vat = doanh_thu_duc * (1 - 0.08)
                thuong_duc = doanh_thu_duc_sau_vat * 0.01

        luong_tong = luong_co_ban + thuong_tan + thuong_duc

        st.caption("Căn cứ tính: Tất cả đơn hàng hợp lệ đã xuất kho / đang giao trở đi (loại trừ trạng thái 'Lên đơn').")

        if is_dat and (thuong_tan > 0 or thuong_duc > 0):
            c1, c2 = st.columns(2)
            c1.markdown(f"""<div class="metric-box"><div class="metric-label">Doanh thu cá nhân (sau VAT)</div>
                <div class="metric-value" style="color:{GREEN}">{money(doanh_thu_sau_vat)}</div></div>""", unsafe_allow_html=True)
            c2.markdown(f"""<div class="metric-box"><div class="metric-label">Hoa hồng cá nhân ({ty_le_hh*100:.1f}%)</div>
                <div class="metric-value">{money(luong_co_ban)}</div></div>""", unsafe_allow_html=True)
            
            st.write("")
            c3, c4, c5 = st.columns(3)
            c3.markdown(f"""<div class="metric-box"><div class="metric-label">+ 1% {ten_nv_tan or 'Tân'}</div>
                <div class="metric-value" style="color:{AMBER}">+{money(thuong_tan)}</div></div>""", unsafe_allow_html=True)
            c4.markdown(f"""<div class="metric-box"><div class="metric-label">+ 1% {ten_nv_duc or 'Đức'}</div>
                <div class="metric-value" style="color:{AMBER}">+{money(thuong_duc)}</div></div>""", unsafe_allow_html=True)
            c5.markdown(f"""<div class="metric-box"><div class="metric-label">Tổng thực nhận</div>
                <div class="metric-value" style="color:{RED}">{money(luong_tong)}</div></div>""", unsafe_allow_html=True)
        else:
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
                <div class="metric-value" style="color:{RED}">{money(luong_tong)}</div></div>""", unsafe_allow_html=True)

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
# 5. DASHBOARD TỔNG QUAN (TÍNH THEO TIẾN ĐỘ ĐƠN HỢP LỆ VÀ CÔNG NỢ)
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

        valid = merged[merged["Trang_thai_Don"].isin(["Gửi kho", "Đang giao", "Đã giao"])].copy()
        valid["Ngay_len_don"] = pd.to_datetime(valid["Ngay_len_don"], errors="coerce")
        period = valid[(valid["Ngay_len_don"].dt.month == thang) & (valid["Ngay_len_don"].dt.year == nam)]

        doanh_thu = period["Thanh_tien"].sum()
        san_luong = period["San_luong_xuat_kho"].sum()
        so_don = period["Ma_don"].nunique()
        cong_no_count = don_hang_df[don_hang_df["Trang_thai_TT"].str.startswith("Chưa", na=False)].shape[0]

        c1, c2 = st.columns(2)
        c1.markdown(f"""
        <div class="metric-box">
            <div class="metric-label">Doanh thu hợp lệ</div>
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
            <div class="metric-label">Đơn đang giao hàng</div>
            <div class="metric-value">{int(don_hang_df[don_hang_df['Trang_thai_Don'] == 'Đang giao'].shape[0])}</div>
        </div>""", unsafe_allow_html=True)
        c4.markdown(f"""
        <div class="metric-box">
            <div class="metric-label">Đơn chưa thanh toán</div>
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
