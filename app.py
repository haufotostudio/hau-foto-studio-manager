import streamlit as st
import datetime
import os.path
import pandas as pd
import io
import unicodedata
import json
import urllib.parse
import re
from pandas.api.types import is_numeric_dtype
import streamlit.components.v1 as components 
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont, TTFError # Import TTFError

# --- CẤU HÌNH HỆ THỐNG ---
SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
TOKEN_FILE = 'token.json'
CREDENTIALS_FILE = 'credentials.json'
SHEET_NAME = "Data Studio"
USERS_FILE = 'users.json'
STAFF_DB_FILE = 'staff_full_db.json'
SETTINGS_FILE = 'settings.json' 
SESSION_FILE = 'session_state.json' 
# --- FILE MỚI: THEO DÕI TRẠNG THÁI THANH TOÁN CAST ---
CAST_PAYMENT_FILE = 'cast_payment_status.json'

# --- URL Google Forms ---
FORM_URL_PERSONAL = "https://forms.gle/avMDvrKpS6E3gXji8" 
FORM_URL_WEDDING = "https://forms.gle/AmcFfNhvjcrkQDPP9" 

# --- CỘT SHEET HỢP ĐỒNG MỚI ---
CONTRACT_COLS = [
    'Mã hợp đồng', 'Khách hàng', 'Số điện thoại', 'Gói chụp', 'NgàyGiờ', 
    'Địa điểm', 'Nhân viên', 'Cast Nhân Viên', 'Giá trị hợp đồng', 
    'Nội dung phát sinh', 'Chi phí phát sinh', 'Tiền cọc', 'Giảm giá', 
    'Chi phí cost', 'Số tiền còn lại', 'Trạng thái', 'Người liên hệ', 
    'Ghi chú', 'Sản phẩm bao gồm', 'Link sản phẩm'
]

# --- GOOGLE AUTH IMPORTS ---
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import gspread

# --- HELPER FUNCTIONS (GLOBAL SCOPE) ---

# --- Ánh xạ Role App sang tiếng Việt ---
ROLE_MAP_VN_TO_KEY = {
    'Quản Lý': 'admin', 
    'Nhân Viên': 'staff', 
    'Đối tác': 'partner'
}
# ĐÃ SỬA LỖI: Sử dụng k (key) và v (value) cho đúng
ROLE_MAP_KEY_TO_VN = {v: k for k, v in ROLE_MAP_VN_TO_KEY.items()}

# --- Cấu hình Tab Order (9 tabs mới) ---
DEFAULT_STAFF_TABS_CONFIG = {
    "CRM": "👥 CRM", 
    "BOOKING_NEW": "📅 BOOKING", 
    "SALARY": "💰 LƯƠNG",
    "CLIENT": "👤 KHÁCH HÀNG", 
    "CONTRACT": "📄 HỢP ĐỒNG", 
    "CALENDAR": "🗓️ LỊCH",
    "PNL": "📈 PHÂN TÍCH P&L", # NEW TAB FOR P&L ANALYSIS
    "ADMIN": "🛠️ ADMIN", 
    "STAFF": "👥 QUẢN LÝ NHÂN SỰ", 
    "SETTINGS": "⚙️ CÀI ĐẶT"
}

# --- KEYWORD MAPPING DỮ LIỆU KHẢO SÁT (TÙY CHỈNH THEO CỘT SHEET) ---
# Sử dụng Tiếng Việt KHÔNG DẤU, viết thường.
# Bạn có thể thêm các tên cột khác nhau từ sheet của mình vào đây.
KEYWORD_MAPPING = {
    # Wedding Keywords
    'wedding': {
        'ten_co_dau': ["co dau", "ten co dau"], # Ưu tiên "Cô dâu"
        'ten_chu_re': ["chu re", "ten chu re"], # Ưu tiên "Chú Rể"
        'sdt_co_dau': ["so dien thoai co dau", "sdt co dau"], # Ưu tiên "Số điện thoại cô dâu"
        'sdt_chu_re': ["so dien thoai chu re", "sdt chu re"], # Ưu tiên "Số điện thoại chú rể"
        'dia_chi_co_dau': ["dia chi co dau", "dia chi nha gai"],
        'dia_chi_chu_re': ["dia chi chu re", "dia chi nha trai"],
        'dinh_vi_co_dau': ["dinh vi nha gai", "dinh vi nha co dau"], # Thêm định vị
        'dinh_vi_chu_re': ["dinh vi nha trai", "dinh vi nha chu re"], # Thêm định vị
        'goi_chup': ["goi chup", "dich vu"],
        'ghi_chu': ["yeu cau dac biet", "ghi chu"],
        'ngay_chup': ["ngay chup"],
        'gio_chup': ["gio chup", "thoi gian", "gio checkin"],
        'facebook': ["facebook", "lien he facebook"], 
    },
    # Personal Keywords
    'personal': {
        'ten_khach_hang': ["ho va ten", "ten khach hang", "khach hang"], # Ưu tiên "Họ và tên"
        'sdt_khach': ["so dien thoai", "sdt", "phone"], # Ưu tiên "Số điện thoại"
        'goi_chup': ["dich vu", "goi chup"],
        'dia_diem': ["boi canh", "dia diem", "dia chi"],
        'dinh_vi_gen': ["dinh vi"], # Thêm định vị chung
        'ghi_chu': ["yeu cau dac biet", "ghi chu"],
        'ngay_chup': ["ngay chup"],
        'gio_chup': ["gio chup", "thoi gian", "gio checkin"],
        'facebook': ["facebook"],
    }
}


def Load_Tab_Order():
    """Tải thứ tự tabs từ settings.json"""
    default_order = list(DEFAULT_STAFF_TABS_CONFIG.keys())
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                 settings = json.load(f)
                 if isinstance(settings, dict): 
                    loaded_order = settings.get('tab_order', [])
                    final_order = [key for key in loaded_order if key in default_order]
                    if not set(default_order).issubset(set(final_order)):
                        new_base_order = []
                        existing_keys = set(final_order)
                        for key in default_order:
                            if key in existing_keys:
                                new_base_order.append(key)
                                existing_keys.remove(key)
                            else:
                                new_base_order.append(key)
                        return new_base_order
                    return final_order if final_order else default_order
        except: pass
    return default_order

def Save_Tab_Order(order_list):
    """Lưu thứ tự tabs vào settings.json"""
    settings = {}
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                 loaded_settings = json.load(f)
                 if isinstance(loaded_settings, dict):
                     settings = loaded_settings
        except: pass
        
    settings['tab_order'] = order_list 
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=4)

def Remove_Accents(s):
    """Loại bỏ dấu tiếng Việt"""
    if not s: return ""
    return "".join([c for c in unicodedata.normalize('NFKD', str(s)) if not unicodedata.combining(c)])

# --- HÀM MỚI: QUẢN LÝ TRẠNG THÁI THANH TOÁN CAST TẠM THỜI (DÙNG JSON) ---

def Load_Cast_Payment_Status():
    """Tải trạng thái thanh toán Cast từ file JSON (Cast ID: True/False)"""
    if not os.path.exists(CAST_PAYMENT_FILE):
        return {}
    try:
        with open(CAST_PAYMENT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def Save_Cast_Payment_Status(status_data):
    """Lưu trạng thái thanh toán Cast vào file JSON"""
    with open(CAST_PAYMENT_FILE, "w", encoding="utf-8") as f:
        json.dump(status_data, f, ensure_ascii=False, indent=4)

def Update_Cast_Payment_Status(cast_id, is_paid):
    """Cập nhật trạng thái thanh toán cho một Cast ID cụ thể"""
    status_data = Load_Cast_Payment_Status()
    status_data[cast_id] = is_paid
    Save_Cast_Payment_Status(status_data)

@st.cache_resource(show_spinner="Đang kết nối dữ liệu...")
import streamlit as st # Đảm bảo đã import st
import os.path
import json # Cần thiết để xử lý chuỗi JSON từ secrets
# ... (các imports khác)

@st.cache_resource(show_spinner="Đang kết nối dữ liệu...")
def Init_Connection():
    """Thiết lập kết nối với Google API (Sheets, Drive, Calendar) và thêm sheets mới"""
    
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    # Biến theo dõi file secrets đang dùng
    client_secrets_file_path = CREDENTIALS_FILE
    is_temp_file = False
    
    if not creds or not creds.valid:
        
        # --- BƯỚC MỚI: KIỂM TRA STREAMLIT SECRETS ---
        if 'google_auth' in st.secrets and 'credentials_json' in st.secrets.google_auth:
            
            # Đang chạy trên Cloud, sử dụng secrets và tạo file tạm
            client_secrets_file_path = "temp_credentials_oauth.json"
            is_temp_file = True
            
            try:
                # Ghi nội dung JSON từ secret vào file tạm
                credentials_data = st.secrets.google_auth['credentials_json']
                with open(client_secrets_file_path, "w", encoding="utf-8") as f:
                    f.write(credentials_data.strip())
            except Exception as e:
                st.error(f"Lỗi tạo file credentials tạm thời từ Secrets: {e}")
                return None, None, None, None, None, None
        
        # --- LOGIC XÁC THỰC CŨ (SỬ DỤNG FILE PATH ĐÃ XÁC ĐỊNH) ---
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:
                # Sử dụng client_secrets_file_path (có thể là CREDENTIALS_FILE hoặc file tạm)
                flow = InstalledAppFlow.from_client_secrets_file(client_secrets_file_path, SCOPES)
                creds = flow.run_local_server(port=0)
            except Exception as e:
                # Dọn dẹp file tạm nếu quá trình xác thực thất bại
                if is_temp_file and os.path.exists(client_secrets_file_path):
                    os.remove(client_secrets_file_path)
                
                st.error(f"Lỗi khởi tạo OAuth: Vui lòng kiểm tra file {CREDENTIALS_FILE} hoặc Secrets. Lỗi: {e}")
                return None, None, None, None, None, None
        
        # Dọn dẹp file tạm sau khi đã xác thực thành công
        if is_temp_file and os.path.exists(client_secrets_file_path):
             os.remove(client_secrets_file_path)

        with open(TOKEN_FILE, 'w') as token: token.write(creds.to_json())
    
    # --- PHẦN KẾT NỐI API GOOGLE (GIỮ NGUYÊN) ---
    try:
        cal_service = build('calendar', 'v3', credentials=creds) 
        gc = gspread.authorize(creds)
        sh = gc.open(SHEET_NAME)
        
        # SỬA LỖI: Dùng sh.worksheets() thay vì sh.worksheet_title()
        all_sheets = [s.title for s in sh.worksheets()]
        
        # --- 1. Sheets Booking & Contract
        ws_personal = sh.worksheet("Booking Cá Nhân") if "Booking Cá Nhân" in all_sheets else None
        ws_wedding = sh.worksheet("Booking Weddings") if "Booking Weddings" in all_sheets else None
        ws_main = sh.worksheet("Quan Ly Hop Dong") if "Quan Ly Hop Dong" in all_sheets else sh.sheet1 
        
        # --- 2. Sheet Hop Dong (Mới)
        try: ws_contract = sh.worksheet("Hop Dong")
        except: ws_contract = sh.add_worksheet(title="Hop Dong", rows="100", cols=len(CONTRACT_COLS))
        
        # --- 3. Sheet Danh Sach Khach Hang (Mới)
        try: ws_client_db = sh.worksheet("Danh Sach Khach Hang")
        except: ws_client_db = sh.add_worksheet(title="Danh Sach Khach Hang", rows="100", cols="6")

        return cal_service, ws_main, ws_personal, ws_wedding, ws_contract, ws_client_db
    except Exception as e: 
        st.error(f"Lỗi kết nối GSheet/GCal: {e}")
        return None, None, None, None, None, None

def Load_Users():
    """Tải CSDL người dùng"""
    default = [{"username": "admin", "password": "123", "role": "admin", "name": "Administrator"}]
    if not os.path.exists(USERS_FILE):
        with open(USERS_FILE, "w", encoding="utf-8") as f: json.dump(default, f, ensure_ascii=False)
    try: return json.load(open(USERS_FILE, "r", encoding="utf-8"))
    except: return default

def Load_Staff_Db():
    """Tải CSDL nhân sự với cấu trúc mới: Thêm Email"""
    # ... (Hàm load_staff_db giữ nguyên)
    default = [{
        "Tên": "Hậu", "Khả năng": "Photo", "Địa điểm": "Sóc Trăng",
        "Account": "hau", "Password": "123", "Role App": "admin",
        "Số điện thoại": "0901xxxxxx", "Email": "hau@studio.com"
    }]
    if not os.path.exists(STAFF_DB_FILE):
        with open(STAFF_DB_FILE, "w", encoding="utf-8") as f: json.dump(default, f, ensure_ascii=False)
    
    try:
        data = json.load(open(STAFF_DB_FILE, "r", encoding="utf-8"))
        for item in data:
            if 'phone' in item: item['Số điện thoại'] = item.pop('phone')
            if 'name' in item and 'Tên' not in item: item['Tên'] = item.pop('name')
            if 'Email' not in item: item['Email'] = '' 
            role_key = item.get("Role App", "staff")
            if role_key in ROLE_MAP_VN_TO_KEY.keys():
                item["Role App"] = ROLE_MAP_VN_TO_KEY[role_key]
        return data
    except Exception as e: return default

def Save_Staff_Db(data):
    """Lưu CSDL nhân sự"""
    with open(STAFF_DB_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False)

def Update_User_From_Staff(staff_db):
    """Đồng bộ users.json từ staff_db.json"""
    users_data = []
    for staff in staff_db:
        role_app_key = staff.get("Role App", "staff")
        
        if role_app_key in ["admin", "staff"]:
            if staff.get("Account") and staff.get("Password"):
                users_data.append({
                    "username": staff["Account"],
                    "password": staff["Password"],
                    "role": role_app_key, 
                    "name": staff.get("Tên", staff["Account"])
                })
    with open(USERS_FILE, "w", encoding="utf-8") as f: json.dump(users_data, f, ensure_ascii=False)

def Load_Packages():
    """Tải CSDL gói chụp (Chỉ giữ cấu trúc file cũ)"""
    default = ["Pre-Wedding Studio", "Pre-Wedding Ngoại Cảnh"]
    return default 

def Save_Packages(pkg_list):
    """Lưu CSDL gói chụp"""
    pass

def Authenticate():
    """Xử lý đăng nhập và duy trì session"""
    st.sidebar.subheader("🔒 Đăng Nhập")
    users = Load_Users()
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False; st.session_state.role = None
        if os.path.exists(SESSION_FILE):
            try:
                session = json.load(open(SESSION_FILE, 'r'))
                user = next((u for u in users if u['username'] == session.get('username')), None)
                if user:
                    st.session_state.authenticated = True; st.session_state.role = user['role']
                    st.session_state.username = user['username']; st.session_state.staff_name = user['name']
            except: 
                if os.path.exists(SESSION_FILE): os.remove(SESSION_FILE)
    
    if st.session_state.authenticated:
        st.sidebar.success(f"Xin chào, {st.session_state.staff_name}")
        if st.sidebar.button("Đăng Xuất"):
            st.session_state.authenticated = False; os.remove(SESSION_FILE) if os.path.exists(SESSION_FILE) else None; st.rerun()
        return True
    
    with st.sidebar.form("login"):
        # Dịch các trường trong form đăng nhập
        u = st.text_input("Tên đăng nhập"); p = st.text_input("Mật khẩu", type="password"); r = st.checkbox("Ghi nhớ")
        if st.form_submit_button("Đăng nhập"):
            user = next((x for x in users if x['username'] == u and x['password'] == p), None)
            if user:
                st.session_state.authenticated = True; st.session_state.role = user['role']
                st.session_state.username = user['username']; st.session_state.staff_name = user['name']
                if r: json.dump({'username': u}, open(SESSION_FILE, 'w'))
                st.rerun()
            else: st.error("Sai thông tin")
    return st.session_state.authenticated

def Clear_Data_Cache():
    """Xóa tất cả cache dữ liệu và session state liên quan đến data"""
    st.cache_resource.clear()
    for k in ['df_survey_personal', 'df_survey_wedding', 'df_prog', 'df_fin_main', 'df_fin_expense', 'cal_events', 'df_admin', 'current_crm_data', 'df_contract', 'df_client_db']:
        if k in st.session_state: del st.session_state[k]

def Load_Booking_Data_From_Sheet(worksheet):
    """
    Tải dữ liệu từ sheet Booking, trả về DataFrame và header.
    Đã chỉnh sửa để xử lý tên cột trùng lặp đơn giản hơn, chỉ giữ tên gốc (Streamlit sẽ xử lý hiển thị).
    """
    if worksheet is None: return pd.DataFrame(), [] 
    try:
        rd = worksheet.get_all_values()
        if not rd: return pd.DataFrame(), [] 

        # Lấy tên cột gốc từ hàng đầu tiên
        cols = [str(h).strip() for h in rd[0]]
        
        # Xử lý tên cột trùng lặp để DataFrame.loc hoạt động ổn định
        final_cols = []
        counts = {}
        for h in rd[0]:
            h = str(h).strip()
            clean_h = re.sub(r'[\(\)\/]', '', h).strip()
            # Nếu tên đã làm sạch KHÔNG phải là tên mặc định (CONTRACT_COLS) thì mới thêm số
            counts[clean_h] = counts.get(clean_h, 0) + 1
            
            # Cột trùng lặp (ví dụ: 'Thời gian_1', 'Thời gian_2')
            if counts[clean_h] > 1:
                 final_cols.append(f"{clean_h}_{counts[clean_h]}")
            else:
                 final_cols.append(clean_h)

        if len(rd) > 1:
            df = pd.DataFrame(rd[1:], columns=final_cols)
            return df, final_cols
        else:
            return pd.DataFrame(columns=final_cols), final_cols
            
    except Exception as e:
        st.error(f"Lỗi khi tải dữ liệu từ Google Sheet {worksheet.title}: {e}")
        return pd.DataFrame(), []

def Get_Survey_Value(df_row, search_keywords):
    """Tìm kiếm giá trị trong df_row dựa trên list keyword (có Remove_Accents và khớp một phần)"""
    for kw in search_keywords:
        # Tìm các cột có chứa keyword hoặc khớp chính xác (sau khi loại bỏ dấu)
        col_names = [col for col in df_row.index if kw.lower() in Remove_Accents(col).lower() or kw.lower() == Remove_Accents(str(col)).lower()]
        if col_names:
            val = str(df_row[col_names[0]]).strip()
            if val and val != "None" and val != "nan" and not pd.isna(val) and val != '': return val
    return ""

def Search_Survey_Data_V2(df_survey, phone_lookup, booking_type):
    """
    Tìm kiếm dữ liệu từ sheet Wedding hoặc Cá Nhân dựa trên SĐT đã được làm sạch.
    Sử dụng KEYWORD_MAPPING để tìm kiếm tên cột và trích xuất dữ liệu.
    """
    if df_survey.empty or not phone_lookup: return None
    
    # 1. Làm sạch SĐT từ input lookup
    clean_number_lookup = re.sub(r'\D', '', phone_lookup)
    if not clean_number_lookup: return None

    # 2. Xác định các cột tiềm năng chứa SĐT dựa trên loại booking (Sử dụng keyword không dấu)
    keywords = KEYWORD_MAPPING[booking_type]
    
    if booking_type == 'wedding':
        # Ghép từ khóa SĐT cô dâu và chú rể
        phone_keywords = keywords.get('sdt_co_dau', []) + keywords.get('sdt_chu_re', [])
    else: # personal
        # Ghép từ khóa SĐT khách
        phone_keywords = keywords.get('sdt_khach', [])

    # Tìm các cột thực tế có SĐT trong DataFrame
    phone_cols = []
    for col in df_survey.columns:
        clean_col = Remove_Accents(col).lower()
        if any(kw in clean_col for kw in phone_keywords):
            phone_cols.append(col)

    if not phone_cols: return None

    df_temp = df_survey.copy()
    
    # 3. Làm sạch dữ liệu SĐT trong các cột tiềm năng của DataFrame
    for col in phone_cols: 
         # Bước 1: Chuyển sang string và loại bỏ dấu nháy đơn (') ở đầu (Nếu có)
         clean_col = df_temp[col].astype(str).str.lstrip("'")
         # Bước 2: Loại bỏ tất cả các ký tự không phải chữ số
         df_temp[col] = clean_col.apply(lambda x: re.sub(r'\D', '', x))
         
    # 4. Tạo mask tìm kiếm: So sánh SĐT đã làm sạch với SĐT lookup đã làm sạch
    # Sử dụng `==` (so sánh bằng) để tìm kiếm SĐT đã được làm sạch hoàn toàn
    mask = df_temp[phone_cols].apply(lambda x: x == clean_number_lookup).any(axis=1)

    search_results = df_survey[mask]
    
    if search_results.empty: return None
    latest_row = search_results.iloc[-1]
    data = {'is_found': True, 'type': booking_type}
    
    # --- TRÍCH XUẤT DỮ LIỆU SỬ DỤNG MAPPING KEYWORD ---
    
    if booking_type == 'wedding':
        # Tên Khách Hàng: Cô dâu & Chú rể
        cd_name = Get_Survey_Value(latest_row, keywords.get('ten_co_dau', []))
        cr_name = Get_Survey_Value(latest_row, keywords.get('ten_chu_re', []))
        data['khach_hang'] = f"{cd_name} & {cr_name}" if cd_name and cr_name else (cd_name or cr_name or "")
        
        # SĐT: SĐT cô dâu & SĐT chú rể
        sdt_cd = Get_Survey_Value(latest_row, keywords.get('sdt_co_dau', []))
        sdt_cr = Get_Survey_Value(latest_row, keywords.get('sdt_chu_re', []))
        data['sdt'] = f"{sdt_cd} & {sdt_cr}" if sdt_cd and sdt_cr else (sdt_cd or sdt_cr or "")
        
        # Địa điểm: Địa chỉ cô dâu + Địa chỉ chú rể
        addr_cd = Get_Survey_Value(latest_row, keywords.get('dia_chi_co_dau', []))
        addr_cr = Get_Survey_Value(latest_row, keywords.get('dia_chi_chu_re', []))
        data['dia_diem'] = f"Nhà gái: {addr_cd}\nNhà trai: {addr_cr}" if addr_cd or addr_cr else ""
        
        # Gói Chụp & Ghi Chú
        data['goi_chup'] = Get_Survey_Value(latest_row, keywords.get('goi_chup', []))
        data['ghi_chu'] = Get_Survey_Value(latest_row, keywords.get('ghi_chu', []))
        
        # Lấy SĐT và Tên riêng để lưu CSDL Khách Hàng
        data['ten_cd'] = cd_name
        data['ten_cr'] = cr_name
        data['sdt_cd'] = sdt_cd
        data['sdt_cr'] = sdt_cr
        data['facebook'] = Get_Survey_Value(latest_row, keywords.get('facebook', []))
        data['dia_chi_cd'] = addr_cd
        data['dia_chi_cr'] = addr_cr

        # NEW: Extract Map Links
        data['map_cd'] = Get_Survey_Value(latest_row, keywords.get('dinh_vi_co_dau', []))
        data['map_cr'] = Get_Survey_Value(latest_row, keywords.get('dinh_vi_chu_re', []))


    else: # personal
        data['khach_hang'] = Get_Survey_Value(latest_row, keywords.get('ten_khach_hang', []))
        data['sdt'] = Get_Survey_Value(latest_row, keywords.get('sdt_khach', []))
        data['goi_chup'] = Get_Survey_Value(latest_row, keywords.get('goi_chup', []))
        data['dia_diem'] = Get_Survey_Value(latest_row, keywords.get('dia_diem', []))
        data['ghi_chu'] = Get_Survey_Value(latest_row, keywords.get('ghi_chu', []))
        
        # Lấy SĐT và Tên riêng để lưu CSDL Khách Hàng
        data['facebook'] = Get_Survey_Value(latest_row, keywords.get('facebook', []))
        data['dia_chi_gen'] = data['dia_diem']

        # NEW: Extract Map Link
        data['map_gen'] = Get_Survey_Value(latest_row, keywords.get('dinh_vi_gen', []))

    # Lịch trình (Giả định sẽ tìm Ngày chụp đầu tiên)
    date_str = Get_Survey_Value(latest_row, keywords.get('ngay_chup', []))
    time_str = Get_Survey_Value(latest_row, keywords.get('gio_chup', []))
    
    # Cột 'Thời gian' (Timestamp)
    timestamp_str = Get_Survey_Value(latest_row, ["thoi gian"]) # Thêm tìm kiếm cột 'Thời gian' (Timestamp) mặc định của Google Forms
    
    # Ưu tiên cột ngày/giờ nếu có, nếu không dùng timestamp
    if date_str and time_str:
        try: data['ngay_gio'] = datetime.datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M:%S")
        except: 
            try: data['ngay_gio'] = datetime.datetime.strptime(f"{date_str}", "%Y-%m-%d").replace(hour=8, minute=0)
            except: data['ngay_gio'] = datetime.datetime.now()
    elif timestamp_str:
         try: 
            # Giả định format Timestamp Google Form mặc định: MM/DD/YYYY HH:MM:SS
            # Hoặc DD/MM/YYYY HH:MM:SS (Tùy theo cài đặt Sheet)
            data['ngay_gio'] = datetime.datetime.strptime(timestamp_str.split()[0], '%d/%m/%Y')
         except:
             try: data['ngay_gio'] = datetime.datetime.strptime(timestamp_str.split()[0], '%m/%d/%Y')
             except: data['ngay_gio'] = datetime.datetime.now()
    else:
        # Nếu không tìm thấy ngày nào, sẽ trả về ngày hiện tại, nhưng chúng ta sẽ
        # bỏ qua nó trong logic gán mặc định của form để ngày luôn là today() khi form được reset
        data['ngay_gio'] = datetime.datetime.now() 

    return data

def Generate_Smart_Contract_Id(date_obj, ws_contract):
    """Tạo Mã hợp đồng HF01ddmmyyyy"""
    try:
        date_str = date_obj.strftime('%d%m%Y')
        records = ws_contract.get_all_values() # Lấy tất cả giá trị để tìm kiếm Mã hợp đồng
        df = pd.DataFrame(records[1:], columns=records[0]) if records and len(records) > 1 else pd.DataFrame(columns=CONTRACT_COLS)
        
        # Đếm số lượng job trong ngày
        count = 0
        if 'Mã hợp đồng' in df.columns:
            # Tìm kiếm mã hợp đồng có chứa chuỗi ddmmyyyy (ví dụ: HF0128112025)
            mask = df['Mã hợp đồng'].astype(str).str.contains(date_str, na=False)
            count = mask.sum()
            
        sequence = count + 1
        return f"HF{sequence:02d}{date_str}"
    except Exception as e: 
        print(f"Lỗi tạo Mã HĐ: {e}")
        return datetime.datetime.now().strftime("HF%H%M%y%m%d")

def Update_Client_Db(ws_client_db, contract_id, client_data):
    """Cập nhật CSDL Khách Hàng (Danh Sach Khach Hang)"""
    # ... (Hàm này cần được gọi sau khi tạo hợp đồng)
    try:
        data = ws_client_db.get_all_values()
        # Header mặc định cho sheet Danh Sach Khach Hang
        client_db_cols = ['STT', 'Họ tên khách hàng', 'Số điện thoại', 'Facebook', 'Địa chỉ', 'Dịch vụ đã chọn']
        df = pd.DataFrame(data[1:], columns=data[0]) if data and len(data) > 1 else pd.DataFrame(columns=client_db_cols)
        
        # Trường hợp Wedding: Cập nhật CĐ và CR riêng biệt
        if client_data.get('type') == 'wedding':
            # Cô dâu
            Update_Client_Row(df, ws_client_db, contract_id, client_data['ten_cd'], client_data['sdt_cd'], client_data['facebook'], client_data['dia_chi_cd'], client_db_cols)
            # Chú Rể (nếu có)
            if client_data['sdt_cr']:
                Update_Client_Row(df, ws_client_db, contract_id, client_data['ten_cr'], client_data['sdt_cr'], client_data['facebook'], client_data['dia_chi_cr'], client_db_cols)
        
        # Trường hợp Cá nhân: Cập nhật 1 dòng
        else: # personal
            Update_Client_Row(df, ws_client_db, contract_id, client_data['khach_hang'], client_data['sdt'], client_data['facebook'], client_data['dia_chi_gen'], client_db_cols)

        # Cập nhật lại cache DB
        df_new, _ = Load_Booking_Data_From_Sheet(ws_client_db)
        st.session_state['df_client_db'] = df_new
        st.toast("✅ Đã cập nhật CSDL Khách Hàng.", icon="👤")
    except Exception as e:
        st.warning(f"⚠️ Lỗi cập nhật CSDL Khách Hàng: {e}")

def Update_Client_Row(df, ws_client_db, contract_id, name, phone, facebook, address, client_db_cols):
    """Logic cập nhật một khách hàng cụ thể"""
    if not phone: return # Bỏ qua nếu không có SĐT
    
    # Kiểm tra cột có tồn tại trong df không
    if 'Họ tên khách hàng' not in df.columns or 'Số điện thoại' not in df.columns or 'Dịch vụ đã chọn' not in df.columns:
         print("Lỗi: Cấu trúc sheet 'Danh Sach Khach Hang' không đúng.")
         return
         
    # Tìm kiếm chính xác (Nếu có thể: tên và SĐT trùng khớp, hoặc chỉ SĐT khớp)
    match_row = df[(df['Họ tên khách hàng'] == name) & (df['Số điện thoại'] == phone)]
    
    if not match_row.empty:
        # Khách hàng cũ: Thêm Mã hợp đồng vào cột 'Dịch vụ đã chọn'
        idx = match_row.index[0]
        current_services = df.loc[idx, 'Dịch vụ đã chọn']
        
        col_index = client_db_cols.index('Dịch vụ đã chọn') + 1 # +1 vì gspread index từ 1
        
        if current_services and contract_id not in current_services:
            new_services = f"{current_services}, {contract_id}"
            # idx + 2: +1 cho header, +1 cho index (vì df là 0-index)
            ws_client_db.update_cell(idx + 2, col_index, new_services) 
        elif not current_services:
             ws_client_db.update_cell(idx + 2, col_index, contract_id)
        
    else:
        # Khách hàng mới: Thêm dòng mới
        new_stt = len(df) + 1
        new_row = [
            new_stt, 
            name, 
            phone, 
            facebook, 
            address, 
            contract_id
        ]
        ws_client_db.append_row(new_row)

def Fmt_Money_Str(value):
    """Format tiền tệ"""
    try: return "{:,.0f}".format(float(value))
    except: return "0"

def Create_Pro_Pdf(data):
    """Tạo file PDF hợp đồng chuyên nghiệp"""
    
    # Khởi tạo mặc định: Sử dụng font chuẩn ReportLab (KHÔNG hỗ trợ Tiếng Việt Unicode)
    TITLE_FONT = 'Helvetica-Bold'
    TEXT_FONT = 'Helvetica'
    
    # --- CỐ GẮNG ĐĂNG KÝ FONT TIẾNG VIỆT ---
    # NOTE: Để PDF hiển thị Tiếng Việt, người dùng CẦN phải có file DejaVuSans.ttf 
    # và DejaVuSans-Bold.ttf trong cùng thư mục với app.py.
    try:
        if "DejaVuSans" not in pdfmetrics.getRegisteredFontNames():
            # Đăng ký font hỗ trợ Tiếng Việt (cần file font thực tế)
            pdfmetrics.registerFont(TTFont('DejaVuSans', 'DejaVuSans.ttf')) 
            pdfmetrics.registerFont(TTFont('DejaVuSansBd', 'DejaVuSans-Bold.ttf'))
            
        # Nếu đăng ký thành công, sử dụng font đã đăng ký
        TITLE_FONT = 'DejaVuSansBd'
        TEXT_FONT = 'DejaVuSans'
        
    except TTFError as e:
        # Nếu lỗi mở file font (ví dụ: file không tồn tại), báo lỗi và dùng font cơ bản
        print(f"Lỗi: Không tìm thấy file font DejaVuSans.ttf/DejaVuSans-Bold.ttf ({e}). PDF sẽ không hiển thị Tiếng Việt có dấu.")
        st.warning("⚠️ Lỗi font PDF: Không tìm thấy file file font DejaVuSans.ttf. PDF có thể không hiển thị Tiếng Việt có dấu.")
    except Exception as e:
        print(f"Lỗi đăng ký font không xác định: {e}")

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    # --- Header ---
    c.setFillColorRGB(0.1, 0.1, 0.1)
    c.setFont(TITLE_FONT, 26)
    c.drawCentredString(width/2, height - 2*cm, "HAU FOTO STUDIO")
    c.setFont(TEXT_FONT, 10)
    c.drawCentredString(width/2, height - 2.8*cm, "Sóc Trăng, Việt Nam | Hotline: 09xx.xxx.xxx")
    c.setLineWidth(1.5)
    c.line(2*cm, height - 3.2*cm, width - 2*cm, height - 3.2*cm)
    
    # --- Title ---
    c.setFont(TITLE_FONT, 18)
    c.drawCentredString(width/2, height - 5*cm, "HỢP ĐỒNG DỊCH VỤ")
    c.setFont(TEXT_FONT, 11)
    c.drawCentredString(width/2, height - 5.6*cm, f"Mã Hợp Đồng: {data['ma_hop_dong']}")
    
    y = height - 7.5*cm
    x = 2.5*cm
    
    # --- A. THÔNG TIN CHUNG ---
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(2*cm, y + 0.5*cm, width-4*cm, -0.6*cm, fill=1, stroke=0)
    c.setFillColorRGB(0, 0, 0)
    c.setFont(TITLE_FONT, 12)
    c.drawString(x, y, "A. THÔNG TIN KHÁCH HÀNG & DỊCH VỤ")
    
    c.setFont(TEXT_FONT, 10)
    y -= 0.8*cm
    c.drawString(x, y, f"Khách Hàng: {data['khach_hang']}")
    y -= 0.6*cm
    c.drawString(x, y, f"Số Điện Thoại: {data['sdt']}")
    y -= 0.6*cm
    c.drawString(x, y, f"Gói Chụp: {data['goi_chup']}")
    y -= 0.6*cm
    c.drawString(x, y, f"Địa Điểm: {data['dia_diem']}")
    y -= 0.6*cm
    c.drawString(x, y, f"Người Liên Hệ: {data['nguoi_lien_he']}")
    y -= 0.6*cm
    c.drawString(x, y, f"Trạng Thái: {data['trang_thai']}")
    y -= 0.6*cm
    c.drawString(x, y, f"Ghi Chú: {data['ghi_chu']}")
    
    y -= 1.5*cm
    
    # --- B. LỊCH CHỤP ---
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(2*cm, y + 0.5*cm, width-4*cm, -0.6*cm, fill=1, stroke=0)
    c.setFillColorRGB(0, 0, 0)
    c.setFont(TITLE_FONT, 12)
    c.drawString(x, y, "B. LỊCH CHỤP")
    
    c.setFont(TEXT_FONT, 10)
    y -= 0.8*cm
    
    # Thêm lịch trình chi tiết
    y_sch = y
    
    if isinstance(data['lich_trinh'], list):
         lt = "\n".join([f"- Ngày: {s['date']} | Giờ Checkin: {s['time_in']} | Giờ Checkout: {s['time_out']} | Nội dung: {s['content']}" for s in data['lich_trinh']])
    else:
         lt = data['lich_trinh'] # Nếu là string
         
    for line in lt.split('\n'):
        if y_sch < 4*cm: # Kiểm tra tràn trang
            c.showPage()
            y_sch = height - 2*cm
            c.setFont(TITLE_FONT, 12)
            c.drawString(x, y_sch, "B. LỊCH CHỤP (Tiếp theo)")
            c.setFont(TEXT_FONT, 10)
            y_sch -= 0.8*cm
            
        c.drawString(x, y_sch, line)
        y_sch -= 0.6*cm
        
    y = y_sch - 1.5*cm
    
    # --- C. TỔNG KẾT VÀ CHỮ KÝ ---
    
    # Tiền tệ
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(2*cm, y + 0.5*cm, width-4*cm, -0.6*cm, fill=1, stroke=0)
    c.setFillColorRGB(0, 0, 0)
    c.setFont(TITLE_FONT, 12)
    c.drawString(x, y, "C. TỔNG KẾT GIÁ TRỊ HỢP ĐỒNG")
    
    c.setFont(TEXT_FONT, 10)
    y -= 0.8*cm
    c.drawString(x, y, f"Trị giá Hợp đồng (chưa giảm giá): {data['tri_gia_hd']} VNĐ")
    y -= 0.6*cm
    c.drawString(x, y, f"Giảm giá: {data['giam_gia']} VNĐ")
    y -= 0.6*cm
    c.drawString(x, y, f"Tiền cọc: {data['tien_coc']} VNĐ")
    y -= 0.6*cm
    c.drawString(x, y, f"Chi phí phát sinh (nếu có): {data['chi_phi_phat_sinh']} VNĐ")
    y -= 0.6*cm
    c.drawString(x, y, f"Tổng Chi phí Ekip (Cast): {data['tong_chi_cast']} VNĐ")
    
    # Thêm Lợi nhuận
    y -= 0.8*cm
    
    # Cần tính lại Lợi nhuận trong PDF creator (Tổng thu - Tổng chi phí)
    # Total Revenue (Sau giảm giá và phát sinh) = Trị giá HD + Chi phí phát sinh - Giảm giá
    try:
        tri_gia_hd_val = float(data['tri_gia_hd'].replace(',', '').replace(' VNĐ', ''))
        giam_gia_val = float(data['giam_gia'].replace(',', '').replace(' VNĐ', ''))
        chi_phi_phat_sinh_val = float(data['chi_phi_phat_sinh'].replace(',', '').replace(' VNĐ', ''))
        tong_chi_cast_val = float(data['tong_chi_cast'].replace(',', '').replace(' VNĐ', ''))
        
        tong_thu = tri_gia_hd_val + chi_phi_phat_sinh_val - giam_gia_val
        loi_nhuan = tong_thu - tong_chi_cast_val
        
        c.setFont(TITLE_FONT, 10)
        c.drawString(x, y, f"Lợi nhuận (Tổng Thu - Tổng Chi): {Fmt_Money_Str(loi_nhuan)} VNĐ")
    except:
        c.setFont(TITLE_FONT, 10)
        c.drawString(x, y, "Lợi nhuận (Tổng Thu - Tổng Chi): Không xác định")
        
    y -= 0.8*cm
    c.setFont(TITLE_FONT, 12)
    c.drawString(x, y, f"SỐ TIỀN CÒN LẠI: {data['so_tien_con_lai']} VNĐ")

    # Chữ ký
    y_sign = 3*cm
    c.setFont(TEXT_FONT, 10)
    c.drawString(x, y_sign+1.5*cm, f"Ngày lập: {datetime.date.today().strftime('%d/%m/%Y')}")
    c.setFont(TITLE_FONT, 11)
    c.drawString(x + 2*cm, y_sign, "ĐẠI DIỆN STUDIO")
    c.drawString(width - 6*cm, y_sign, "KHÁCH HÀNG")

    c.save(); buffer.seek(0)
    return buffer

def Update_Staff_Db(staff_db, current_username, new_data, is_acc_change=False):
    """Hàm chung để cập nhật CSDL Nhân sự và đồng bộ Users"""
    idx = next(i for i, item in enumerate(staff_db) if item.get('Account') == current_username)
    
    for key, value in new_data.items():
        if key in staff_db[idx]:
            staff_db[idx][key] = value
            
    if is_acc_change:
        st.session_state.authenticated = False
        st.error(f"✅ Đã đổi Tên đăng nhập thành **{new_data.get('Account')}**. Vui lòng **Đăng nhập lại** bằng Account mới.")
    
    # 3. Save and Rerun
    with open(STAFF_DB_FILE, "w", encoding="utf-8") as f: json.dump(staff_db, f, ensure_ascii=False)
    Update_User_From_Staff(staff_db) 
    st.rerun()

# --- FORM RENDERER VÀ LOGIC TẠO HỢP ĐỒNG ---

def Init_Form_State(booking_type):
    """Khởi tạo cấu trúc Session State cho Form"""
    # Khởi tạo form_state dictionary nếu chưa có
    if 'form_state' not in st.session_state:
        st.session_state.form_state = {}

    # Nếu booking_type chưa có state, khởi tạo state mặc định cho nó
    if booking_type not in st.session_state.form_state:
        st.session_state.form_state[booking_type] = {
            # Metadata
            'is_new': True,
            'form_reset_key': 0,
            'form_data': {},
            'phone_lookup': '',
            
            # Input data
            'inp_khach_hang': '',
            'inp_sdt': '',
            'inp_goi_chup': '',
            'inp_dia_diem': '',
            'inp_ghi_chu': '',
            'inp_trang_thai': 'Chưa cọc',
            'inp_nguoi_lien_he': st.session_state.get('staff_name', ''),
            
            # Financial
            'inp_tri_gia_hd': 5000000,
            'inp_da_coc': 1000000,
            'inp_giam_gia': 0,
            
            # Chi phí phát sinh (Outside form)
            'inp_noi_dung_phat_sinh_checkbox': False,
            'inp_chi_phi_phat_sinh': 0,
            'inp_noi_dung_phat_sinh_text': '',
            
            # Dynamic lists
            'products': [
                 {'so_luong': 1, 'san_pham': '', 'size': '', 'chi_phi_cost': 0},
                 {'so_luong': 1, 'san_pham': '', 'size': '', 'chi_phi_cost': 0}
            ],
            'ekip': [],
        }
        
    return st.session_state.form_state[booking_type]

# --- HÀM MỚI: RENDER KHỐI SẢN PHẨM ---
def Render_Product_Block(current_state, booking_type):
    """Render khối Sản phẩm và quản lý state"""
    st.markdown("---")
    c_prod_title, c_prod_add_btn = st.columns([0.8, 0.2])
    c_prod_title.markdown("##### 🖼️ Sản phẩm")

    products = current_state['products']
    
    # Nút THÊM SẢN PHẨM (Đã chuyển ra ngoài st.form)
    with c_prod_add_btn.container(): 
        # Cập nhật state trước khi reruns
        if st.button("➕", help="Thêm Sản phẩm", key=f"add_prod_btn_out_{booking_type}"):
            products.append({'so_luong': 1, 'san_pham': '', 'size': '', 'chi_phi_cost': 0})
            current_state['products'] = products # Cập nhật lại list vào state
            st.rerun()
            
    st.caption("Nhập thông tin sản phẩm và chi phí Cost (sản xuất):")
    
    # Header
    c_h1, c_h2, c_h3, c_h4, c_h5 = st.columns([1, 2, 1, 2, 0.5])
    c_h1.write("**SL**")
    c_h2.write("**Sản phẩm**")
    c_h3.write("**Size**")
    c_h4.write("**Cost/SP**")
    
    products_to_keep = []
    has_delete_action = False

    for i, prod in enumerate(products):
        # SỬ DỤNG KEY ĐỘNG TRONG VÒNG LẶP CHO SẢN PHẨM
        dynamic_prod_key = f"{booking_type}_prod_{i}_{current_state['form_reset_key']}"

        c_p1, c_p2, c_p3, c_p4, c_p5 = st.columns([1, 2, 1, 2, 0.5])
        
        # Input/Widget cho Sản phẩm (Dùng key độc lập + dynamic_prod_key)
        # Bắt buộc cập nhật giá trị vào dictionary để tránh mất data khi reruns
        products[i]['so_luong'] = c_p1.number_input("", value=prod.get('so_luong', 1), min_value=1, key=f"prod_sl_{dynamic_prod_key}", label_visibility="collapsed")
        products[i]['san_pham'] = c_p2.text_input("", value=prod.get('san_pham', ''), key=f"prod_sp_{dynamic_prod_key}", label_visibility="collapsed")
        products[i]['size'] = c_p3.text_input("", value=prod.get('size', ''), key=f"prod_sz_{dynamic_prod_key}", label_visibility="collapsed")
        products[i]['chi_phi_cost'] = c_p4.number_input("", value=prod.get('chi_phi_cost', 0), min_value=0, step=10000, key=f"prod_cost_{dynamic_prod_key}", label_visibility="collapsed")
        
        # Nút xóa sản phẩm (Đã chuyển ra ngoài st.form)
        if c_p5.button("❌", key=f"del_prod_out_{booking_type}_{i}", help="Xóa sản phẩm này"):
            has_delete_action = True
            # Không cần pass, logic xóa sẽ được xử lý bên ngoài vòng lặp
        else:
            products_to_keep.append(products[i])
        
    # Cập nhật lại list sản phẩm nếu có nút xóa nào được nhấn
    if has_delete_action:
        current_state['products'] = products_to_keep
        st.rerun()
    
    # Tính tổng cost sản phẩm (Sử dụng list hiện tại sau khi đã xử lý xóa)
    final_products = current_state['products'] # Luôn lấy từ state sau khi reruns
    tong_chi_cost_sp = sum(p['chi_phi_cost'] * p['so_luong'] for p in final_products)
    st.markdown(f"**Tổng Chi Phí Sản Xuất (Cost Sản phẩm):** **{Fmt_Money_Str(tong_chi_cost_sp)} VNĐ**")
    return tong_chi_cost_sp

# --- HÀM MỚI: RENDER KHỐI EKIP ---
def Render_Ekip_Block(current_state, booking_type):
    """Render khối Phân công nhân sự và quản lý state"""
    st.markdown("---")
    c_ekip_title, c_ekip_add_btn = st.columns([0.8, 0.2])
    c_ekip_title.markdown("##### 🧑‍🤝‍🧑 Phân công nhân sự (Ekip)")

    staff_db = Load_Staff_Db()
    staff_list = current_state['ekip']
    
    # SỬ DỤNG KEY ĐỘNG CHO INPUT CỦA EKIP
    dynamic_ekip_key = f"{booking_type}_ekip_input_{current_state['form_reset_key']}"

    # Input fields cho Ekip 
    c_e1, c_e2, c_e3, c_e4 = st.columns([2, 2, 2, 1])
    # Lưu giá trị input vào session state tạm
    sel_staff = c_e1.selectbox("Nhân Viên", ["--"] + [s['Tên'] for s in staff_db], key=f"sel_staff_book_out_{dynamic_ekip_key}")
    role_staff = c_e2.text_input("Nhiệm vụ", "Chụp chính", key=f"role_staff_book_out_{dynamic_ekip_key}")
    cast_staff = c_e3.number_input("Giá Cast", value=500000, min_value=0, step=100000, key=f"m_cast_book_val_out_{dynamic_ekip_key}")
    ghi_chu_ekip = c_e4.text_input("Ghi chú NV", "", key=f"note_ekip_book_out_{dynamic_ekip_key}")

    col_add, col_clear = st.columns([0.5, 0.5])
    
    # Nút THÊM NHÂN SỰ (Đã chuyển ra ngoài st.form)
    if col_add.button("➕ Thêm Ekip", key=f"add_ekip_{booking_type}_btn_out", type="primary"):
        # Lấy giá trị từ session state với key động
        sel_staff_val = st.session_state.get(f"sel_staff_book_out_{dynamic_ekip_key}", "--")
        role_staff_val = st.session_state.get(f"role_staff_book_out_{dynamic_ekip_key}", "Chụp chính")
        cast_staff_val = st.session_state.get(f"m_cast_book_val_out_{dynamic_ekip_key}", 500000)
        ghi_chu_ekip_val = st.session_state.get(f"note_ekip_book_out_{dynamic_ekip_key}", "")
        
        if sel_staff_val != "--":
            # Kiểm tra trùng lặp
            if any(e['name'] == sel_staff_val and e['role'] == role_staff_val for e in staff_list):
                 st.warning(f"Nhân viên {sel_staff_val} đã được thêm với vai trò {role_staff_val}.")
            else:
                 staff_list.append({"name": sel_staff_val, "role": role_staff_val, "cast": cast_staff_val, "note": ghi_chu_ekip_val}); st.rerun()

    # Nút xóa hết Ekip (Đã chuyển ra ngoài st.form)
    if col_clear.button("🗑️ Xóa hết Ekip", key=f"clear_ekip_{booking_type}_btn_out"):
        staff_list.clear(); st.rerun()

    # Hiển thị danh sách Ekip và nút xóa từng dòng
    staff_list_to_keep = []
    has_delete_action = False
    
    if staff_list:
        st.write("###### Danh sách Ekip:")
        
        # Sử dụng container để định dạng đẹp hơn
        with st.container(border=True):
            cols_head = st.columns([2, 2, 2, 1, 0.5])
            cols_head[0].markdown("**Tên NV**")
            cols_head[1].markdown("**Nhiệm vụ**")
            cols_head[2].markdown("**Giá Cast (VNĐ)**")
            cols_head[3].markdown("**Ghi chú**")

            for i, member in enumerate(staff_list):
                cols = st.columns([2, 2, 2, 1, 0.5])
                cols[0].write(member['name'])
                cols[1].write(member['role'])
                cols[2].write(Fmt_Money_Str(member['cast']))
                cols[3].write(member['note'])
                
                # Nút xóa từng dòng Ekip
                if cols[4].button("❌", key=f"del_ekip_row_{booking_type}_{i}", help="Xóa thành viên này"):
                    has_delete_action = True
                else:
                    staff_list_to_keep.append(member)

    # Cập nhật lại list Ekip nếu có nút xóa nào được nhấn
    if has_delete_action:
        current_state['ekip'] = staff_list_to_keep
        st.rerun()


    # Tính tổng cost Ekip
    tong_chi_cast = sum(e['cast'] for e in staff_list_to_keep) if has_delete_action else sum(e['cast'] for e in staff_list)
    st.markdown(f"**Tổng Chi Phí Ekip (Cast):** **{Fmt_Money_Str(tong_chi_cast)} VNĐ**")
    return tong_chi_cast

# --- HÀM MỚI: RENDER KHỐI PHÁT SINH (ĐÃ BỎ EXPANDER) ---
def Render_Expense_Block(current_state, booking_type):
    """Render khối Chi phí Phát sinh tinh gọn (Không dùng Expander)"""
    st.markdown("---")
    st.markdown("##### 📝 Chi phí phát sinh")

    # SỬ DỤNG KEY ĐỘNG CHO INPUT PHÁT SINH
    dynamic_expense_key = f"{booking_type}_expense_{current_state['form_reset_key']}"

    # Checkbox (Dùng để kiểm soát trạng thái)
    noi_dung_phat_sinh_checkbox = st.checkbox(
        "Có Nội dung phát sinh?", 
        value=current_state['inp_noi_dung_phat_sinh_checkbox'], 
        key=f"inp_noi_dung_phat_sinh_checkbox_{dynamic_expense_key}", # Key độc lập
    )
    
    # Cập nhật state ngay lập tức (Chỉ cần cho checkbox)
    current_state['inp_noi_dung_phat_sinh_checkbox'] = noi_dung_phat_sinh_checkbox

    r_g4, r_g5 = st.columns(2)
    
    # Number input
    chi_phi_phat_sinh = r_g4.number_input(
        "Chi phí phát sinh (VNĐ)", 
        value=current_state['inp_chi_phi_phat_sinh'], 
        min_value=0, 
        step=100000, 
        key=f"inp_chi_phi_phat_sinh_out_{dynamic_expense_key}", # Key độc lập
        disabled=not noi_dung_phat_sinh_checkbox
    )
    
    # Text Area cho Nội dung phát sinh
    noi_dung_phat_sinh_text = st.text_area(
        "Chi tiết nội dung phát sinh:",
        value=current_state['inp_noi_dung_phat_sinh_text'],
        key=f"inp_noi_dung_phat_sinh_text_out_{dynamic_expense_key}", # Key độc lập
        disabled=not noi_dung_phat_sinh_checkbox
    )
    
    # Cập nhật các giá trị input vào state
    current_state['inp_chi_phi_phat_sinh'] = chi_phi_phat_sinh
    current_state['inp_noi_dung_phat_sinh_text'] = noi_dung_phat_sinh_text
    
    # Logic reset (Nếu bỏ chọn checkbox)
    is_reset_needed = False
    if not noi_dung_phat_sinh_checkbox:
        if chi_phi_phat_sinh != 0 or noi_dung_phat_sinh_text != "":
             current_state['inp_chi_phi_phat_sinh'] = 0
             current_state['inp_noi_dung_phat_sinh_text'] = ""
             is_reset_needed = True
        
    if is_reset_needed:
         st.rerun()

    return current_state['inp_chi_phi_phat_sinh']

# --- HÀM GỐC ĐÃ CHỈNH SỬA VỀ BỐ CỤC VÀ LOGIC CHECK TRẠNG THÁI (FIX LỖI TỰ ĐIỀN KEY TĨNH) ---
def Render_Booking_Form(worksheet_name, booking_type, ws_contract, ws_client_db, cal_service):
    """Render form cho cả Wedding và Cá Nhân (sử dụng Session State chung)"""
    
    # --- 1. LẤY STATE HIỆN TẠI ---
    current_state = Init_Form_State(booking_type)

    # Load data sheet tương ứng
    if booking_type == 'wedding':
         ws_data = st.session_state.get('ws_wedding')
         search_key = "SĐT CĐ/CR"
    else:
         ws_data = st.session_state.get('ws_personal')
         search_key = "SĐT Khách"

    if ws_data is None:
        st.warning(f"Sheet **'{worksheet_name}'** không tồn tại. Vui lòng tạo sheet này để sử dụng tính năng tự điền.")

    # --- TỐI ƯU: Chỉ load data nếu cần thiết (Khi form đang trống hoặc có yêu cầu lookup) ---
    df_data = pd.DataFrame()
    if ws_data and (current_state['is_new'] or current_state['phone_lookup']):
         # Chỉ load 1000 dòng gần nhất để tăng tốc độ nếu data quá lớn
         # Tuy nhiên, do gspread không có get_all_records_limit, giữ nguyên get_all_values()
         df_data, _ = Load_Booking_Data_From_Sheet(ws_data) 


    # --- 2. TỰ ĐIỀN THÔNG MINH ---
    st.write("##### 🔎 Tự điền thông tin từ Form Khảo Sát")
    c_s1, c_s2 = st.columns([3, 1])
    
    # Widget lookup_phone được lưu vào current_state['phone_lookup']
    phone_lookup = c_s1.text_input(
        f"Nhập {search_key} để tự điền:", 
        value=current_state['phone_lookup'], 
        key=f"lookup_phone_{booking_type}"
    )
    # Cập nhật state ngay sau khi widget thay đổi giá trị
    current_state['phone_lookup'] = st.session_state[f"lookup_phone_{booking_type}"]
    
    # Nút "Tự điền"
    if c_s2.button("🔎 Tự điền", type="primary", key=f"lookup_btn_{booking_type}"):
        if df_data.empty: 
            st.warning(f"Dữ liệu Booking trong sheet **'{worksheet_name}'** rỗng. Không thể tra cứu.")
        elif not phone_lookup: 
            st.warning("Vui lòng nhập Số điện thoại.")
        else:
            found_data = Search_Survey_Data_V2(df_data, phone_lookup, booking_type)
            if found_data and found_data.get('is_found'):
                # CẬP NHẬT TẤT CẢ CÁC INPUT STATE TỪ DỮ LIỆU TÌM THẤY TRƯỚC KHI RERUN (Sửa lỗi mất tự điền)
                current_state['form_data'] = found_data
                current_state['inp_khach_hang'] = found_data.get('khach_hang', '')
                current_state['inp_sdt'] = found_data.get('sdt', '')
                current_state['inp_goi_chup'] = found_data.get('goi_chup', '')
                current_state['inp_dia_diem'] = found_data.get('dia_diem', '') 
                current_state['inp_ghi_chu'] = found_data.get('ghi_chu', '')
                current_state['is_new'] = False
                
                # CẢI TIẾN QUAN TRỌNG: Tăng key động để buộc toàn bộ form reset
                current_state['form_reset_key'] += 1
                
                st.toast(f"✅ Đã tìm thấy dữ liệu loại {booking_type.upper()}! Đang tự động điền...", icon="👤")
                st.rerun() 
            else: 
                st.toast("❌ Không tìm thấy SĐT này!", icon="❌")
                
    # Lấy hoặc khởi tạo dữ liệu form
    form_data = current_state['form_data']

    # --- 3. KHỞI TẠO GIÁ TRỊ MẶC ĐỊNH CHO CÁC TRƯỜNG FORM ---
    
    default_datetime = datetime.datetime.now()
    if isinstance(form_data.get('ngay_gio'), datetime.datetime):
         default_time_in = form_data['ngay_gio'].time().replace(second=0, microsecond=0)
    else:
         default_time_in = default_datetime.time().replace(minute=0, second=0, microsecond=0)

    default_date_end = default_datetime.date()
    default_time_out = (default_datetime + datetime.timedelta(hours=4)).time().replace(second=0, microsecond=0)
    
    current_date = default_datetime.date() 
    ma_hop_dong = Generate_Smart_Contract_Id(current_date, ws_contract)

    # Lấy giá trị TEXT từ State
    default_khach_hang = current_state['inp_khach_hang']
    default_sdt = current_state['inp_sdt']
    default_goi_chup = current_state['inp_goi_chup']
    default_ghi_chu = current_state['inp_ghi_chu']
    
    # --- CONSTRUCT DIA DIEM WITH MAP LINKS ---
    default_dia_diem_base = current_state['inp_dia_diem']
    map_links = []
    if booking_type == 'wedding':
        map_cd = form_data.get('map_cd', '')
        map_cr = form_data.get('map_cr', '')
        if map_cd: map_links.append(f"Định vị Nhà gái: {map_cd}")
        if map_cr: map_links.append(f"Định vị Nhà trai: {map_cr}")
    else: # personal
        map_gen = form_data.get('map_gen', '')
        if map_gen: map_links.append(f"Định vị: {map_gen}")

    formatted_location = default_dia_diem_base.strip()
    map_suffix = " | ".join(map_links)
    if formatted_location and map_suffix:
        formatted_location += " - " + map_suffix
    elif map_suffix:
        formatted_location = map_suffix
        
    default_dia_diem_final = formatted_location
    
    # --- 4. TÍNH TOÁN COST VÀ PHÁT SINH (SỬ DỤNG HÀM RENDER NGOÀI FORM) ---
    
    # === KHỐI CHI PHÍ PHÁT SINH (RENDER KHÔNG BOX) ===
    chi_phi_phat_sinh_val = Render_Expense_Block(current_state, booking_type) 
    
    # === KHỐI SẢN PHẨM ===
    tong_chi_cost_sp = Render_Product_Block(current_state, booking_type)
    
    # === KHỐI PHÂN CÔNG NHÂN SỰ (EKIP) ===
    tong_chi_cast = Render_Ekip_Block(current_state, booking_type)

    # --- 5. FORM CHÍNH (Sử dụng key reset động) ---
    form_key = f"contract_form_{booking_type}_{current_state['form_reset_key']}"
    dynamic_key_suffix = current_state['form_reset_key'] # Suffix mới cho các widget

    with st.form(form_key, clear_on_submit=False):
        
        # --- TITLE 1: THÔNG TIN KHÁCH HÀNG ---
        st.markdown("##### 👤 Thông tin khách hàng")
        
        r1, r2 = st.columns(2)
        # Sử dụng KEY ĐỘNG
        r1.text_input("Mã hợp đồng", value=ma_hop_dong, disabled=True, key=f"inp_ma_hop_dong_{booking_type}_{dynamic_key_suffix}")
        khach_hang = r2.text_input("Khách hàng", value=default_khach_hang, key=f"inp_khach_hang_{booking_type}_{dynamic_key_suffix}")
        
        r3, r4 = st.columns(2)
        sdt = r3.text_input("Số điện thoại", value=default_sdt, key=f"inp_sdt_{booking_type}_{dynamic_key_suffix}")
        goi_chup = r4.text_input("Gói chụp", value=default_goi_chup, key=f"inp_goi_chup_{booking_type}_{dynamic_key_suffix}")
        
        dia_diem = st.text_area("Địa điểm", value=default_dia_diem_final, key=f"inp_dia_diem_{booking_type}_{dynamic_key_suffix}")
        
        r5, r6 = st.columns(2)
        trang_thai_options = ['Chưa cọc', 'Đã lên lịch', 'Hậu kỳ', 'Chọn ảnh in', 'Hoàn thành']
        default_status = current_state.get('inp_trang_thai', 'Chưa cọc') # Dùng state hiện tại
        try: 
            default_index = trang_thai_options.index(default_status)
        except ValueError:
            default_index = 0
            
        trang_thai = r5.selectbox("Trạng thái", options=trang_thai_options, index=default_index, key=f"inp_trang_thai_{booking_type}_{dynamic_key_suffix}")
        nguoi_lien_he = r6.text_input("Người liên hệ (Tên NV)", value=current_state.get('inp_nguoi_lien_he', st.session_state.get('staff_name', '')), key=f"inp_nguoi_lien_he_{booking_type}_{dynamic_key_suffix}")
        ghi_chu = st.text_area("Ghi chú", value=default_ghi_chu, key=f"inp_ghi_chu_{booking_type}_{dynamic_key_suffix}")

        # Cập nhật input data state
        # NOTE: Giữ nguyên logic cập nhật này
        current_state['inp_khach_hang'] = khach_hang
        current_state['inp_sdt'] = sdt
        current_state['inp_goi_chup'] = goi_chup
        current_state['inp_dia_diem'] = dia_diem
        current_state['inp_trang_thai'] = trang_thai
        current_state['inp_nguoi_lien_he'] = nguoi_lien_he
        current_state['inp_ghi_chu'] = ghi_chu


        # --- TITLE 2: LÊN LỊCH ---
        st.markdown("---")
        st.markdown("##### 📅 Lên Lịch (Start - End)")
        
        r7, r8 = st.columns(2)
        # Sử dụng KEY ĐỘNG
        ngay_bat_dau = r7.date_input("Ngày Bắt đầu", default_datetime.date(), key=f"inp_ngay_bat_dau_{booking_type}_{dynamic_key_suffix}")
        ngay_ket_thuc = r8.date_input("Ngày Kết thúc", default_date_end, key=f"inp_ngay_ket_thuc_{booking_type}_{dynamic_key_suffix}")

        r9, r10 = st.columns(2)
        # Sử dụng KEY ĐỘNG
        gio_checkin = r9.time_input("Giờ Checkin", default_time_in, key=f"inp_gio_in_{booking_type}_{dynamic_key_suffix}")
        gio_checkout = r10.time_input("Giờ Checkout", default_time_out, key=f"inp_gio_out_{booking_type}_{dynamic_key_suffix}")
        

        # === KHỐI GIÁ TRỊ HỢP ĐỒNG (VẪN Ở TRONG FORM) ---
        st.markdown("---")
        st.markdown("##### 💵 Giá trị hợp đồng")
        
        r_g1, r_g2, r_g3 = st.columns(3)
        # Sử dụng KEY ĐỘNG
        tri_gia_hd = r_g1.number_input("Trị giá hợp đồng (Chưa giảm)", value=current_state.get('inp_tri_gia_hd', 5000000), min_value=0, step=100000, key=f"inp_tri_gia_hd_{booking_type}_{dynamic_key_suffix}")
        da_coc = r_g2.number_input("Đã cọc", value=current_state.get('inp_da_coc', 1000000), min_value=0, step=100000, key=f"inp_da_coc_{booking_type}_{dynamic_key_suffix}")
        giam_gia = r_g3.number_input("Giảm giá", value=current_state.get('inp_giam_gia', 0), min_value=0, step=100000, key=f"inp_giam_gia_{booking_type}_{dynamic_key_suffix}")
        
        # Cập nhật financial state
        current_state['inp_tri_gia_hd'] = tri_gia_hd
        current_state['inp_da_coc'] = da_coc
        current_state['inp_giam_gia'] = giam_gia
        
        
        # --- TÍNH TOÁN CUỐI CÙNG (Đã được sắp xếp logic) ---
        tong_chi_thuc_te = tong_chi_cost_sp + tong_chi_cast
        tong_thu = tri_gia_hd + chi_phi_phat_sinh_val - giam_gia
        con_lai = tong_thu - da_coc
        loi_nhuan = tong_thu - tong_chi_thuc_te
        
        st.markdown("---")
        st.markdown("##### 📈 Tổng kết & Chi tiết Chi phí")
        
        # --- GIAO DIỆN TỔNG KẾT CHI TIẾT VÀ DỄ NHÌN HƠN ---
        
        st.write("**Tổng Thu Nhận:**")
        col_t1, col_t2, col_t3 = st.columns(3)
        col_t1.metric("Trị giá HĐ (Gốc)", f"{Fmt_Money_Str(tri_gia_hd)} VNĐ")
        col_t2.metric("Chi phí Phát sinh", f"{Fmt_Money_Str(chi_phi_phat_sinh_val)} VNĐ")
        col_t3.metric("Giảm giá", f"- {Fmt_Money_Str(giam_gia)} VNĐ")
        
        st.markdown("---")
        
        st.write("**Tổng Chi Phí (Cost):**")
        col_c1, col_c2, col_c3 = st.columns(3)
        col_c1.metric("Cost Sản phẩm", f"{Fmt_Money_Str(tong_chi_cost_sp)} VNĐ")
        col_c2.metric("Cast Ekip", f"{Fmt_Money_Str(tong_chi_cast)} VNĐ")
        col_c3.metric("Tổng Chi Thực tế", f"{Fmt_Money_Str(tong_chi_thuc_te)} VNĐ")

        st.markdown("---")

        st.write("##### **KẾT QUẢ KINH DOANH & THU HỒI VỐN**")
        col_f1, col_f2, col_f3 = st.columns(3)
        col_f1.metric("Tổng Giá Trị Thu (Net)", f"{Fmt_Money_Str(tong_thu)} VNĐ")
        col_f2.metric("LỢI NHUẬN", f"{Fmt_Money_Str(loi_nhuan)} VNĐ", delta_color="inverse", delta=f"{'Tăng' if loi_nhuan >= 0 else 'Giảm'}")
        col_f3.metric("CÒN LẠI PHẢI THU", f"{Fmt_Money_Str(con_lai)} VNĐ", delta=f"Đã cọc: {Fmt_Money_Str(da_coc)} VNĐ")
        # --- KẾT THÚC GIAO DIỆN TỔNG KẾT ---

        st.markdown("---")
        
        submit_contract = st.form_submit_button("🚀 LƯU SHOW, CẬP NHẬT LỊCH & XUẤT HỢP ĐỒNG", type="primary")

        if submit_contract:
            # Kiểm tra logic ngày/giờ
            start_dt = datetime.datetime.combine(ngay_bat_dau, gio_checkin)
            end_dt = datetime.datetime.combine(ngay_ket_thuc, gio_checkout)
            
            if start_dt >= end_dt:
                 st.error("Ngày/Giờ Bắt đầu phải trước Ngày/Giờ Kết thúc.")
                 st.stop()

            if not khach_hang or not sdt or not goi_chup:
                st.error("Vui lòng điền đủ Tên Khách hàng, SĐT và Gói chụp.")
            
            # Đảm bảo trạng thái không phải là "Chưa cọc" nếu muốn tạo lịch Google
            # Tuy nhiên, theo yêu cầu mới là "Trạng thái chưa cọc vẫn có thể lên lịch được",
            # nên chỉ cần đảm bảo có trạng thái Hợp đồng được chọn, ví dụ:
            if not trang_thai:
                 st.error("Vui lòng chọn Trạng thái hợp đồng.")
                 st.stop()
            
            # --- A. CHUẨN BỊ DỮ LIỆU ---
            full_sdt_clean = re.sub(r'[^\d&]', '', sdt)
            
            # Format Ngày/Giờ theo yêu cầu mới: "YYYY-MM-DD HH:MM - YYYY-MM-DD HH:MM"
            ngay_gio_hop_dong = f"{ngay_bat_dau.strftime('%Y-%m-%d')} {gio_checkin.strftime('%H:%M')} - {ngay_ket_thuc.strftime('%Y-%m-%d')} {gio_checkout.strftime('%H:%M')}"
            
            # Chi tiết Ekip và Sản phẩm cho Sheet
            ekip_txt = "\n".join([f"{e['name']} ({e['role']}) - Cast: {Fmt_Money_Str(e['cast'])} VNĐ" for e in current_state['ekip']])
            san_pham_txt = "\n".join([f"{p['san_pham']} ({p['size']}) - SL: {p['so_luong']}" for p in current_state['products']])
            
            # Nội dung phát sinh cho sheet
            noi_dung_phat_sinh_checkbox_final = current_state['inp_noi_dung_phat_sinh_checkbox']
            noi_dung_phat_sinh_text_final = current_state['inp_noi_dung_phat_sinh_text']

            noi_dung_phat_sinh_sheet = noi_dung_phat_sinh_text_final if noi_dung_phat_sinh_checkbox_final else 'Không'

            # --- B. LƯU VÀO SHEET HOP DONG ---
            row_data = [
                ma_hop_dong, khach_hang, full_sdt_clean, goi_chup, ngay_gio_hop_dong,
                dia_diem, ekip_txt, tong_chi_cast, tri_gia_hd, 
                noi_dung_phat_sinh_sheet, chi_phi_phat_sinh_val, da_coc, # Sử dụng chi_phi_phat_sinh_val đã tính
                giam_gia, tong_chi_cost_sp, con_lai, trang_thai, 
                nguoi_lien_he, ghi_chu, san_pham_txt, '' # Link sản phẩm (empty)
            ]

            ws_contract.append_row(row_data)
            st.toast("✅ Đã Lưu Hợp Đồng vào Google Sheet 'Hop Dong'.")
            
            # --- C. CẬP NHẬT CSDL KHÁCH HÀNG (Danh Sach Khach Hang) ---
            client_data_for_db = {
                'type': booking_type,
                'khach_hang': khach_hang,
                'sdt': full_sdt_clean,
                'facebook': form_data.get('facebook', ''),
                'dia_chi_gen': form_data.get('dia_diem', ''),
                'ten_cd': form_data.get('ten_cd', khach_hang),
                'ten_cr': form_data.get('ten_cr', ''),
                'sdt_cd': form_data.get('sdt_cd', full_sdt_clean),
                'sdt_cr': form_data.get('sdt_cr', ''),
                'dia_chi_cd': form_data.get('dia_diem', ''),
                'dia_chi_cr': form_data.get('dia_diem', ''),
            }
            Update_Client_Db(st.session_state.ws_client_db, ma_hop_dong, client_data_for_db)

            # --- D. GOOGLE CALENDAR ---
            try:
                # Dùng start_dt và end_dt đã tính toán
                evt = {
                    'summary': f"📸 HĐ {ma_hop_dong}: {goi_chup} - {khach_hang}",
                    'description': f"SĐT: {sdt}\nTrị giá: {Fmt_Money_Str(tong_thu)} VNĐ\nCòn lại: {Fmt_Money_Str(con_lai)} VNĐ\nEkip: {ekip_txt}",
                    'location': dia_diem.split('\n')[0].strip()[:50], 
                    'start': {'dateTime': start_dt.isoformat(), 'timeZone': 'Asia/Ho_Chi_Minh'},
                    'end': {'dateTime': end_dt.isoformat(), 'timeZone': 'Asia/Ho_Chi_Minh'}, 'colorId': '10'
                }
                cal_service.events().insert(calendarId='primary', body=evt).execute()
                st.toast("✅ Đã thêm lịch vào Google Calendar.")
            except Exception as e: 
                st.warning(f"⚠️ Lỗi Calendar: Không thể thêm lịch. Lỗi: {e}")

            # --- E. XUẤT PDF ---
            pdf_d = {
                'ma_hop_dong': ma_hop_dong, 'khach_hang': khach_hang, 'sdt': sdt, 
                'goi_chup': goi_chup, 'dia_diem': dia_diem, 'nguoi_lien_he': nguoi_lien_he,
                'trang_thai': trang_thai, 'ghi_chu': ghi_chu,
                # Lịch trình cho PDF cũng dùng định dạng mới
                'lich_trinh': [
                    {'date': ngay_bat_dau.strftime('%d/%m/%Y'), 'time_in': gio_checkin.strftime('%H:%M'), 'time_out': gio_checkout.strftime('%H:%M'), 'content': goi_chup}
                ],
                'tri_gia_hd': Fmt_Money_Str(tri_gia_hd), 'giam_gia': Fmt_Money_Str(giam_gia), 
                'tien_coc': Fmt_Money_Str(da_coc), 'chi_phi_phat_sinh': Fmt_Money_Str(chi_phi_phat_sinh_val), # Sử dụng chi phí phát sinh chính xác
                'tong_chi_cast': Fmt_Money_Str(tong_chi_thuc_te), 'so_tien_con_lai': Fmt_Money_Str(con_lai)
            }
            
            st.session_state.last_pdf = Create_Pro_Pdf(pdf_d)
            st.session_state.last_pdf_name = f"HD_{ma_hop_dong}.pdf"
            st.success(f"✅ Hợp đồng {ma_hop_dong} đã được tạo thành công!")

            # --- F. RESET STATE CỦA LOẠI BOOKING HIỆN TẠI ---
            current_state['form_reset_key'] += 1
            
            # Reset data fields về giá trị default ban đầu
            default_new_state = Init_Form_State(booking_type).copy()
            for key in default_new_state.keys():
                if key not in ['form_reset_key', 'inp_nguoi_lien_he']: # Giữ lại key reset và người liên hệ
                    current_state[key] = default_new_state[key]
            
            Clear_Data_Cache() # Clear sheet data cache
            st.rerun()

    # --- Nút Download PDF (Giữ nguyên vị trí cuối cùng)
    if 'last_pdf' in st.session_state: 
        st.markdown("---")
        # THÊM KEY DUY NHẤT
        st.download_button("📥 Tải Hợp Đồng PDF", st.session_state.last_pdf, st.session_state.last_pdf_name, "application/pdf", key=f"download_pdf_{booking_type}")
        
# --- HÀM MỚI: TÁCH NGÀY KÝ HĐ ---
def Extract_Date_From_Contract_Id(contract_id):
    """Tách ngày ký Hợp đồng (ddmmyyyy) từ Mã hợp đồng (HFXXddmmyyyy)"""
    if isinstance(contract_id, str) and len(contract_id) >= 10 and contract_id.startswith('HF'):
        date_part = contract_id[-8:] # last 8 chars: ddmmyyyy
        try:
            return pd.to_datetime(date_part, format='%d%m%Y').date()
        except:
            return None
    return None

# --- HÀM MỚI: PHÂN TÍCH CHUỖI EKIP ĐỂ TÍNH TOÁN CAST CHI TIẾT ---
def Parse_Ekip_String(ekip_string):
    """
    Phân tích chuỗi chi tiết Ekip (dạng multiline string từ cột 'Nhân viên') 
    thành danh sách các thành viên và Cast của họ.
    """
    if not isinstance(ekip_string, str) or not ekip_string.strip():
        return []
    
    # Ví dụ: "Tên NV (Vai trò) - Cast: 500,000 VNĐ\n Tên NV2 (Vai trò2) - Cast: 200,000 VNĐ"
    members = []
    lines = ekip_string.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Regex để tìm kiếm tên, vai trò và giá trị Cast (hỗ trợ cả dấu phẩy)
        # 1: Tên (Vai trò), 2: Giá trị Cast
        match = re.search(r'(.+?)\s+-\s+Cast:\s+([\d,.]+)\s+VNĐ', line)
        
        if match:
            name_role_part = match.group(1).strip()
            cast_value_str = match.group(2).replace(',', '').replace('.', '') # Loại bỏ dấu phân cách
            
            try:
                cast_value = int(float(cast_value_str))
            except ValueError:
                cast_value = 0
            
            # Tên: Phần trước dấu ngoặc đơn hoặc toàn bộ nếu không có ngoặc đơn
            name_match = re.match(r'(.+?)\s+\(', name_role_part)
            name = name_match.group(1).strip() if name_match else name_role_part.split('(')[0].strip()
            
            # Vai trò: Phần trong dấu ngoặc đơn
            role_match = re.search(r'\((.*?)\)', name_role_part)
            role = role_match.group(1).strip() if role_match else "N/A"
            
            members.append({
                'Tên NV': name,
                'Vai trò': role,
                'Cast': cast_value
            })
    return members


# --- HÀM MỚI: HIỂN THỊ DANH SÁCH HỢP ĐỒNG ---
def Render_Contract_List(ws_contract):
    """Tải và hiển thị danh sách Hợp đồng từ sheet "Hop Dong" với tùy chọn lọc ngày."""
    
    st.write("##### 📋 Danh sách Hợp đồng đã tạo")
    
    if ws_contract is None:
        st.warning("Không thể kết nối đến sheet 'Hop Dong'. Vui lòng kiểm tra lại Google Sheet.")
        return

    # Tải dữ liệu (sử dụng cache nếu có)
    df_contract = st.session_state.get('df_contract')
    
    if df_contract is None:
        try:
            df, cols = Load_Booking_Data_From_Sheet(ws_contract)
            df_contract = df
            st.session_state['df_contract'] = df_contract
        except Exception as e:
            st.error(f"Lỗi khi tải dữ liệu Hợp đồng: {e}")
            return
            
    # Hiển thị nút refresh
    col_refresh, col_info = st.columns([1, 4])
    # FIX: Thêm key duy nhất cho nút Tải lại dữ liệu
    if col_refresh.button("🔄 Tải lại dữ liệu", key="refresh_contract_list_btn"):
        Clear_Data_Cache()
        st.rerun()

    if df_contract.empty:
        st.info("Hiện chưa có Hợp đồng nào được lưu. Hãy tạo một Hợp đồng mới từ tab '📅 BOOKING'.")
        return
        
    # --- 1. LỌC NGÀY KÝ HỢP ĐỒNG ---
    st.markdown("---")
    st.write("##### 📈 Hiệu suất Kinh doanh theo Ngày Ký Hợp đồng")
    
    filter_options = {
        "Tất cả": 0,
        "3 ngày gần nhất": 3,
        "7 ngày gần nhất": 7,
        "15 ngày gần nhất": 15,
        "30 ngày gần nhất": 30,
    }
    
    selected_filter = st.selectbox(
        "Chọn Khung thời gian:", 
        options=list(filter_options.keys()), 
        key="contract_date_filter"
    )
    
    days_to_filter = filter_options[selected_filter]
    
    # 2. Chuẩn bị DataFrame cho hiển thị và lọc
    df_display = df_contract.copy()
    
    # 3. Thêm cột Ngày Ký HĐ
    if 'Mã hợp đồng' in df_display.columns:
        df_display['Ngày Ký HĐ'] = df_display['Mã hợp đồng'].apply(Extract_Date_From_Contract_Id)
    else:
        st.warning("Không tìm thấy cột 'Mã hợp đồng' trong Sheet. Không thể lọc theo ngày ký.")
        df_display['Ngày Ký HĐ'] = None # Ensure column exists

    # 4. Áp dụng Lọc
    df_filtered = df_display.copy()
    if days_to_filter > 0 and 'Ngày Ký HĐ' in df_filtered.columns:
        today = datetime.date.today()
        # Tính ngày bắt đầu lọc
        cutoff_date = today - datetime.timedelta(days=days_to_filter)
        
        # Lọc các dòng có Ngày Ký HĐ (không None) và >= cutoff_date
        mask = df_filtered['Ngày Ký HĐ'].apply(lambda x: x is not None and x >= cutoff_date)
        df_filtered = df_filtered[mask]

    # Cột tiền tệ
    money_cols = ['Cast Nhân Viên', 'Giá trị hợp đồng', 'Chi phí phát sinh', 'Tiền cọc', 'Giảm giá', 'Chi phí cost', 'Số tiền còn lại']
    
    # Chuyển đổi các cột tiền tệ sang số để dễ sắp xếp và tính toán
    for col in money_cols:
        if col in df_filtered.columns:
            # Thay thế chuỗi rỗng bằng 0 trước khi chuyển đổi
            df_filtered[col] = pd.to_numeric(df_filtered[col].replace('', '0'), errors='coerce').fillna(0).astype(int)

    # 5. Tính toán tổng hợp (Tổng doanh thu, Tổng lợi nhuận) trên dữ liệu đã lọc
    total_revenue = 0
    total_profit = 0
    
    if 'Giá trị hợp đồng' in df_filtered.columns and 'Chi phí phát sinh' in df_filtered.columns and 'Giảm giá' in df_filtered.columns and 'Chi phí cost' in df_filtered.columns and 'Cast Nhân Viên' in df_filtered.columns:
        tong_thu = df_filtered['Giá trị hợp đồng'] + df_filtered['Chi phí phát sinh'] - df_filtered['Giảm giá']
        tong_chi = df_filtered['Chi phí cost'] + df_filtered['Cast Nhân Viên']
        df_filtered.insert(len(df_filtered.columns), 'Lợi nhuận', tong_thu - tong_chi)
        
        total_revenue = tong_thu.sum() # Tổng giá trị hợp đồng (Net Revenue)
        total_profit = df_filtered['Lợi nhuận'].sum() # Tổng lợi nhuận

    # 6. Format các cột tiền tệ thành chuỗi VÀ LỌC CỘT HIỂN THỊ
    df_display_final = df_filtered.copy()
    
    for col in money_cols:
         if col in df_display_final.columns:
              df_display_final[col] = df_display_final[col].apply(Fmt_Money_Str) # Format tiền tệ

    if 'Lợi nhuận' in df_display_final.columns:
         df_display_final['Lợi nhuận'] = df_display_final['Lợi nhuận'].apply(Fmt_Money_Str) # Format Lợi nhuận

    # Lọc các cột cần thiết cho hiển thị tổng quan
    display_cols = ['Mã hợp đồng', 'Ngày Ký HĐ', 'Khách hàng', 'Số điện thoại', 'Gói chụp', 'NgàyGiờ', 'Trạng thái', 'Giá trị hợp đồng', 'Số tiền còn lại']
    if 'Lợi nhuận' in df_display_final.columns:
        display_cols.append('Lợi nhuận')
        
    # Sắp xếp lại cột hiển thị
    final_display_cols = [col for col in display_cols if col in df_display_final.columns]

    if df_filtered.empty:
        st.info("Không có hợp đồng nào được ký trong khung thời gian đã chọn.")
    else:
        # Hiển thị DataFrame
        st.dataframe(
            df_display_final[final_display_cols], # Sử dụng DataFrame đã được Format chuỗi
            use_container_width=True,
            hide_index=True
        )

        # Hiển thị Tổng hợp
        st.markdown("---")
        st.write("##### Tổng hợp Hiệu suất Kinh doanh:")
        
        col_rev, col_prof = st.columns(2)
        col_rev.metric("Tổng Thu (Net Revenue)", f"{Fmt_Money_Str(total_revenue)} VNĐ") 
        col_prof.metric("Tổng Lợi nhuận Thu về", f"{Fmt_Money_Str(total_profit)} VNĐ", delta_color="inverse", delta=f"{'Tăng' if total_profit >= 0 else 'Giảm'}")

# --- HÀM MỚI: RENDER CALENDAR ---
@st.cache_data(ttl=600)
def Get_Calendar_Events(_cal_service):
    """Tải sự kiện Calendar với phạm vi 90 ngày."""
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    time_max = (datetime.datetime.utcnow() + datetime.timedelta(days=90)).isoformat() + 'Z'
    
    events_result = _cal_service.events().list(calendarId='primary', timeMin=now,
                                            timeMax=time_max, maxResults=20, singleEvents=True,
                                            orderBy='startTime').execute()
    return events_result.get('items', [])

def Render_Calendar(cal_service):
    """Hiển thị lịch làm việc từ Google Calendar với thông tin chi tiết."""
    
    st.write("##### 🗓️ Lịch Chụp & Lịch Làm Việc Sắp Tới")
    
    try:
        events = Get_Calendar_Events(cal_service)

        if not events:
            st.info('Không có lịch chụp nào sắp tới trong 90 ngày. Hãy tạo Hợp đồng mới để cập nhật lịch.')
            return

        data_rows = []
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            
            date_display = "Cả ngày"
            time_display = "Không rõ"
            
            try:
                # Chuyển đổi start/end time sang datetime object và format lại (chuyển Z về +00:00)
                start_dt_utc = datetime.datetime.fromisoformat(start.replace('Z', '+00:00'))
                
                # Lấy ngày (Format: DD/MM/YYYY)
                date_display = start_dt_utc.strftime('%d/%m/%Y')
                
                # Lấy giờ check-in/out
                end = event['end'].get('dateTime', event['end'].get('date'))
                end_dt_utc = datetime.datetime.fromisoformat(end.replace('Z', '+00:00'))
                time_display = f"{start_dt_utc.strftime('%H:%M')} - {end_dt_utc.strftime('%H:%M')}"
                
            except:
                date_display = f"Ngày {start.split('T')[0]}"
                time_display = "Cả ngày"

            summary = event.get('summary', 'Sự kiện không tên')
            description = event.get('description', '')
            location = event.get('location', 'Chưa xác định')
            
            # Extract key details from description (assuming format is "Ekip: <details>\n..." and "SĐT: <details>\n...")
            # SĐT
            sdt_match = re.search(r'SĐT: ([^\n]*)', description)
            sdt = sdt_match.group(1).strip() if sdt_match else 'N/A'
            
            # Ekip
            ekip_match = re.search(r'Ekip: ([^\n]*)', description)
            ekip = ekip_match.group(1).split('\n')[0].strip() if ekip_match else 'Chưa phân công'

            data_rows.append({
                'Ngày': date_display,
                'Giờ': time_display,
                'Sự kiện (HĐ)': summary,
                'Địa điểm': location,
                'Ekip': ekip,
                'SĐT Khách': sdt,
                'Link GCal': event.get('htmlLink', '#')
            })

        df_events = pd.DataFrame(data_rows)
        # Sắp xếp lại cột để dễ xem cho thợ ảnh
        df_events_display = df_events[['Ngày', 'Giờ', 'Sự kiện (HĐ)', 'Địa điểm', 'Ekip', 'SĐT Khách']]
        
        st.dataframe(df_events_display, use_container_width=True, hide_index=True)

        st.caption("Hiển thị 20 sự kiện sắp tới trong 90 ngày. (Giờ hiển thị là giờ UTC/Địa phương, tùy thuộc vào cách Calendar API trả về).")
        
        # Thêm link trực tiếp
        st.markdown("---")
        st.write("##### Xem chi tiết trên Google Calendar")
        for i, row in df_events.iterrows():
            st.markdown(f"- **{row['Ngày']} - {row['Giờ']}**: [{row['Sự kiện (HĐ)']}]({row['Link GCal']})")
            
    except Exception as e:
        st.error(f"⚠️ Lỗi khi tải lịch Google Calendar: {e}. Vui lòng kiểm tra kết nối API.")

# --- HÀM MỚI: RENDER DANH SÁCH KHÁCH HÀNG (CLIENT) ---
def Render_Client_Db(ws_client_db):
    """Hiển thị và lọc CSDL Khách Hàng"""
    st.write("##### 👤 Danh sách Khách Hàng đã giao dịch")
    
    if ws_client_db is None:
        st.warning("Không thể kết nối đến sheet 'Danh Sach Khach Hang'. Vui lòng kiểm tra lại Google Sheet.")
        return

    # Tải dữ liệu (sử dụng cache nếu có)
    df_client_db = st.session_state.get('df_client_db')
    
    if df_client_db is None or df_client_db.empty:
        try:
            df, cols = Load_Booking_Data_From_Sheet(ws_client_db)
            df_client_db = df
            st.session_state['df_client_db'] = df_client_db
        except Exception as e:
            st.error(f"Lỗi khi tải dữ liệu Khách Hàng: {e}")
            return
    
    col_refresh, col_filter = st.columns([1, 4])
    if col_refresh.button("🔄 Tải lại dữ liệu", key="refresh_client_db_btn"):
        Clear_Data_Cache()
        st.rerun()
        
    if df_client_db.empty:
        st.info("Hiện chưa có dữ liệu Khách Hàng nào được lưu.")
        return
        
    st.markdown("---")
    
    # --- Lọc và tìm kiếm ---
    search_term = col_filter.text_input("Tìm kiếm (Tên, SĐT, Facebook...)", key="client_search_term")
    
    df_filtered = df_client_db.copy()
    
    if search_term:
        search_term_lower = Remove_Accents(search_term).lower()
        
        # Tìm kiếm trên tất cả các cột
        mask = df_filtered.apply(lambda row: any(search_term_lower in Remove_Accents(str(x)).lower() for x in row), axis=1)
        df_filtered = df_filtered[mask]
        
    st.write(f"Đã tìm thấy **{len(df_filtered)}** khách hàng trong tổng số **{len(df_client_db)}**.")

    display_cols = ['STT', 'Họ tên khách hàng', 'Số điện thoại', 'Facebook', 'Địa chỉ', 'Dịch vụ đã chọn']
    final_display_cols = [col for col in display_cols if col in df_filtered.columns]

    st.dataframe(
        df_filtered[final_display_cols],
        use_container_width=True,
        hide_index=True
    )

# --- HÀM MỚI: RENDER SALARY (LƯƠNG) ---
def Render_Salary(ws_contract):
    """Hiển thị tổng quan Lương/Cast Ekip từ các Hợp đồng và quản lý thanh toán Cast"""
    st.write("##### 💰 Tổng hợp Cast Ekip theo tháng (Chi tiết Nhân sự)")
    
    if ws_contract is None:
         st.warning("Không thể truy cập sheet 'Hop Dong' để tính toán Cast.")
         return

    # Tải dữ liệu hợp đồng (sử dụng cache)
    df_contract = st.session_state.get('df_contract')
    if df_contract is None:
        try:
            df_contract, _ = Load_Booking_Data_From_Sheet(ws_contract)
            st.session_state['df_contract'] = df_contract
        except: return
        
    if df_contract.empty:
        st.info("Chưa có Hợp đồng nào được lưu để tính toán Cast.")
        return

    # Chuẩn bị DataFrame
    df_salary = df_contract.copy()
    
    # 1. Chuyển đổi NgàyGiờ thành datetime (Sử dụng cột NgàyGiờ format: YYYY-MM-DD HH:MM - YYYY-MM-DD HH:MM)
    if 'NgàyGiờ' in df_salary.columns:
        df_salary['Ngày Bắt đầu'] = df_salary['NgàyGiờ'].str.split(' - ').str[0].str.split(' ').str[0]
        df_salary['Ngày Bắt đầu'] = pd.to_datetime(df_salary['Ngày Bắt đầu'], errors='coerce')
        df_salary.dropna(subset=['Ngày Bắt đầu'], inplace=True)
        df_salary['Tháng'] = df_salary['Ngày Bắt đầu'].dt.to_period('M')
    else:
        st.warning("Không tìm thấy cột 'NgàyGiờ' để tính toán tháng.")
        return

    # 2. Đảm bảo cột Cast Nhân Viên là số
    if 'Cast Nhân Viên' in df_salary.columns:
        df_salary['Cast Nhân Viên'] = pd.to_numeric(df_salary['Cast Nhân Viên'].replace('', '0'), errors='coerce').fillna(0).astype(int)
    else:
        st.warning("Không tìm thấy cột 'Cast Nhân Viên'.")
        return

    # 3. Tính tổng Cast theo tháng (Tổng Cast Ekip)
    monthly_cast = df_salary.groupby('Tháng')['Cast Nhân Viên'].sum().reset_index()
    monthly_cast['Tháng'] = monthly_cast['Tháng'].astype(str)
    
    # 4. Hiển thị tổng hợp theo tháng
    monthly_cast['Tổng Cast'] = monthly_cast['Cast Nhân Viên'].apply(Fmt_Money_Str)
    
    st.write("###### Tổng Cast Ekip (Tổng chi phí lương/cast trên tất cả hợp đồng):")
    st.dataframe(
        monthly_cast[['Tháng', 'Tổng Cast']].sort_values(by='Tháng', ascending=False),
        use_container_width=True,
        hide_index=True
    )
    
    # --- 5. TÍNH TOÁN CAST CHI TIẾT THEO TỪNG NHÂN VIÊN ---
    st.markdown("---")
    st.write("##### 🧑‍🤝‍🧑 Chi tiết Cast theo từng Nhân viên & Quản lý Thanh toán")

    staff_cast_data = []

    if 'Nhân viên' in df_salary.columns and 'Mã hợp đồng' in df_salary.columns and 'Ngày Bắt đầu' in df_salary.columns:
        for index, row in df_salary.iterrows():
            ekip_members = Parse_Ekip_String(row['Nhân viên'])
            
            for i, member in enumerate(ekip_members):
                # Tạo Cast ID duy nhất: HĐ ID + Tên NV + Vai trò (để phân biệt nếu 1 NV có 2 vai trò)
                cast_id = f"{row['Mã hợp đồng']}_{member['Tên NV']}_{member['Vai trò']}".replace(' ', '_')
                staff_cast_data.append({
                    'Cast ID': cast_id,
                    'Tên NV': member['Tên NV'],
                    'Vai trò': member['Vai trò'],
                    'Cast': member['Cast'],
                    'Mã HĐ': row['Mã hợp đồng'],
                    'Ngày Chụp': row['Ngày Bắt đầu'] # Already a datetime/date
                })

        if staff_cast_data:
            df_staff_cast = pd.DataFrame(staff_cast_data)

            # --- Tích hợp Trạng thái Thanh toán ---
            cast_payment_status = Load_Cast_Payment_Status()
            
            # Thêm cột trạng thái thanh toán
            df_staff_cast['Đã Thanh Toán'] = df_staff_cast['Cast ID'].apply(lambda x: cast_payment_status.get(x, False))
            df_staff_cast['Trạng thái Thanh toán'] = df_staff_cast['Đã Thanh Toán'].apply(lambda x: "✅ Đã TT" if x else "❌ Chưa TT")


            # 5a. Tính tổng Cast theo Nhân viên (Aggregate)
            # Tính Tổng Cast cần trả (chưa thanh toán)
            df_staff_summary_agg = df_staff_cast.groupby('Tên NV').agg(
                {'Cast': 'sum', 'Đã Thanh Toán': 'sum'}
            ).reset_index()
            
            # Lọc chỉ Cast CHƯA THANH TOÁN để tính Tổng Cast còn nợ
            df_unpaid = df_staff_cast[df_staff_cast['Đã Thanh Toán'] == False]
            df_unpaid_agg = df_unpaid.groupby('Tên NV')['Cast'].sum().reset_index()
            df_unpaid_agg.columns = ['Tên NV', 'Tổng Cast Còn Nợ']
            
            df_staff_summary = df_staff_summary_agg.merge(df_unpaid_agg, on='Tên NV', how='left').fillna(0)
            df_staff_summary.rename(columns={'Cast': 'Tổng Cast', 'Tên NV': 'Tên Nhân Viên'}, inplace=True)
            
            # Thêm cột hiển thị
            df_staff_summary['Tổng Cast (VNĐ)'] = df_staff_summary['Tổng Cast'].apply(Fmt_Money_Str)
            df_staff_summary['Còn Nợ (VNĐ)'] = df_staff_summary['Tổng Cast Còn Nợ'].apply(Fmt_Money_Str)
            
            # Sắp xếp theo Tổng Cast còn nợ giảm dần (ưu tiên thanh toán)
            df_staff_summary = df_staff_summary.sort_values(by='Tổng Cast Còn Nợ', ascending=False)
            
            st.write("###### Tổng hợp Cast cần thanh toán cho từng Nhân viên:")
            st.dataframe(
                df_staff_summary[['Tên Nhân Viên', 'Tổng Cast (VNĐ)', 'Còn Nợ (VNĐ)']], 
                use_container_width=True, 
                hide_index=True
            )
            
            # 5b. Hiển thị chi tiết từng job của nhân viên
            st.markdown("---")
            st.write("###### Chi tiết từng Show của Nhân viên (Quản lý Thanh toán):")
            
            # Cho phép lọc theo trạng thái thanh toán
            pay_filter = st.radio(
                "Lọc theo Trạng thái Thanh toán:",
                options=["Tất cả", "❌ Chưa thanh toán", "✅ Đã thanh toán"],
                index=1,
                horizontal=True,
                key="cast_pay_filter"
            )
            
            df_cast_detail = df_staff_cast[['Ngày Chụp', 'Mã HĐ', 'Tên NV', 'Vai trò', 'Cast', 'Cast ID', 'Trạng thái Thanh toán']].copy()
            df_cast_detail['Cast (VNĐ)'] = df_cast_detail['Cast'].apply(Fmt_Money_Str)

            # Áp dụng lọc
            if pay_filter == "❌ Chưa thanh toán":
                df_cast_detail = df_cast_detail[df_cast_detail['Trạng thái Thanh toán'] == "❌ Chưa TT"]
            elif pay_filter == "✅ Đã thanh toán":
                df_cast_detail = df_cast_detail[df_cast_detail['Trạng thái Thanh toán'] == "✅ Đã TT"]
                
            # Tạo form cho chức năng thanh toán
            with st.form("cast_payment_form"):
                st.markdown("###### ✅ Đánh dấu Thanh toán Cast")
                
                # Lọc danh sách Cast ID chưa thanh toán để chọn
                unpaid_cast_ids = df_cast_detail[df_cast_detail['Trạng thái Thanh toán'] == "❌ Chưa TT"]['Cast ID'].tolist()
                
                if unpaid_cast_ids:
                    # Lấy thông tin hiển thị chi tiết (ví dụ: Tên NV - Mã HĐ - Cast)
                    unpaid_details = []
                    for cast_id in unpaid_cast_ids:
                        row = df_cast_detail[df_cast_detail['Cast ID'] == cast_id].iloc[0]
                        unpaid_details.append(f"{row['Tên NV']} ({row['Mã HĐ']}) - {row['Cast (VNĐ)']} - {row['Ngày Chụp'].strftime('%d/%m/%Y')}")
                    
                    selected_unpaid = st.selectbox(
                        "Chọn Cast Job để đánh dấu Đã Thanh toán:",
                        options=["-- Chọn Cast Job --"] + unpaid_details,
                        key="selected_cast_to_pay"
                    )
                    
                    submit_pay = st.form_submit_button("✅ Xác nhận Đã Thanh toán Cast", type="primary")

                    if submit_pay and selected_unpaid != "-- Chọn Cast Job --":
                        # Trích xuất Cast ID từ chuỗi hiển thị
                        selected_cast_id_raw = selected_unpaid.split('(')[1].split(')')[0]
                        selected_nv_name = selected_unpaid.split('(')[0].strip()
                        # Tìm Cast ID gốc trong list unpaid_cast_ids (có thể hơi phức tạp do chuỗi hiển thị)
                        
                        # Cách đơn giản: Lặp qua df_cast_detail để tìm Cast ID khớp
                        found_row = df_cast_detail[df_cast_detail.apply(lambda row: f"{row['Tên NV']} ({row['Mã HĐ']})" in selected_unpaid, axis=1)]
                        
                        if not found_row.empty:
                            cast_id_to_update = found_row.iloc[0]['Cast ID']
                            Update_Cast_Payment_Status(cast_id_to_update, True)
                            st.toast(f"✅ Đã đánh dấu Cast Job của {found_row.iloc[0]['Tên NV']} ({found_row.iloc[0]['Mã HĐ']}) là Đã Thanh toán!")
                            Clear_Data_Cache() # Clear cache để tải lại trạng thái
                            st.rerun()
                        else:
                             st.error("Lỗi: Không tìm thấy Cast ID hợp lệ để cập nhật.")
                             
                else:
                    st.info("Không có Cast Job nào chưa thanh toán để đánh dấu.")
                    st.form_submit_button("✅ Xác nhận Đã Thanh toán Cast", disabled=True)
            
            # Hiển thị bảng chi tiết (đã lọc)
            st.dataframe(
                df_cast_detail[['Ngày Chụp', 'Mã HĐ', 'Tên NV', 'Vai trò', 'Cast (VNĐ)', 'Trạng thái Thanh toán']].sort_values(by='Ngày Chụp', ascending=False),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("Không tìm thấy thông tin chi tiết Ekip trong các Hợp đồng (kiểm tra cột 'Nhân viên').")
    
    # 6. Chi tiết Cast theo Hợp đồng (dữ liệu cũ đã được hiển thị bên trên trong mục 4)
    st.markdown("---")
    st.write("##### Chi tiết Cast theo từng Hợp đồng (Toàn bộ Ekip/Show):")
    
    df_detail = df_salary[['Mã hợp đồng', 'Ngày Bắt đầu', 'Khách hàng', 'Nhân viên', 'Cast Nhân Viên', 'Trạng thái']].copy()
    
    # Format cột tiền tệ trong chi tiết
    if 'Cast Nhân Viên' in df_detail.columns:
         df_detail['Cast Nhân Viên'] = df_detail['Cast Nhân Viên'].apply(Fmt_Money_Str)
         
    # Đổi tên cột hiển thị
    df_detail.columns = ['Mã HĐ', 'Ngày Chụp', 'Khách Hàng', 'Ekip Ghi Chú (Chi tiết NV)', 'Tổng Cast Ekip', 'Trạng Thái']
    
    # FIX LỖI: Cột dùng để sort đã được đổi tên từ 'Ngày Bắt đầu' thành 'Ngày Chụp'
    st.dataframe(
        df_detail.sort_values(by='Ngày Chụp', ascending=False),
        use_container_width=True,
        hide_index=True
    )

# --- HÀM MỚI: PHÂN TÍCH LỢI NHUẬN (P&L) ---
def Render_PNL_Analysis(ws_contract):
    """Phân tích Lợi nhuận (P&L) cho các Gói Chụp (Packages)"""
    st.write("##### 📊 Phân tích hiệu suất tài chính chi tiết theo từng Gói Dịch vụ.")

    if ws_contract is None:
         st.warning("Không thể truy cập sheet 'Hop Dong' để tính toán P&L.")
         return

    # Tải dữ liệu hợp đồng (sử dụng cache)
    df_contract = st.session_state.get('df_contract')
    if df_contract is None:
        try:
            df_contract, _ = Load_Booking_Data_From_Sheet(ws_contract)
            st.session_state['df_contract'] = df_contract
        except: return
        
    if df_contract.empty:
        st.info("Chưa có Hợp đồng nào được lưu để thực hiện phân tích.")
        return

    df_pnl = df_contract.copy()
    
    # 1. Cleaning & Conversion
    money_cols = ['Giá trị hợp đồng', 'Chi phí phát sinh', 'Tiền cọc', 'Giảm giá', 'Chi phí cost', 'Cast Nhân Viên', 'Số tiền còn lại']
    
    # Chuyển đổi các cột tiền tệ sang số
    for col in money_cols:
        if col in df_pnl.columns:
            df_pnl[col] = pd.to_numeric(df_pnl[col].replace('', '0'), errors='coerce').fillna(0).astype(int)
    
    # Kiểm tra cột bắt buộc
    required_pnl_cols = ['Giá trị hợp đồng', 'Chi phí phát sinh', 'Giảm giá', 'Chi phí cost', 'Cast Nhân Viên', 'Gói chụp']
    if not all(col in df_pnl.columns for col in required_pnl_cols):
        st.error("Dữ liệu trong sheet 'Hop Dong' bị thiếu các cột bắt buộc cho Phân tích P&L (ví dụ: 'Giá trị hợp đồng', 'Chi phí cost', 'Cast Nhân Viên', 'Gói chụp').")
        return

    # 2. Tính toán Lợi nhuận (P&L)
    df_pnl['Doanh Thu Net'] = df_pnl['Giá trị hợp đồng'] + df_pnl['Chi phí phát sinh'] - df_pnl['Giảm giá']
    df_pnl['Tổng Chi Phí'] = df_pnl['Chi phí cost'] + df_pnl['Cast Nhân Viên']
    df_pnl['Lợi Nhuận'] = df_pnl['Doanh Thu Net'] - df_pnl['Tổng Chi Phí']
    
    # 3. Widget chọn Gói chụp
    packages = ['Tất cả'] + df_pnl['Gói chụp'].unique().tolist()
    selected_package = st.selectbox("Chọn Gói Dịch vụ để phân tích:", options=packages, key="pnl_package_select")
    
    df_filtered = df_pnl
    if selected_package != 'Tất cả':
        df_filtered = df_pnl[df_pnl['Gói chụp'] == selected_package].copy()
        
    if df_filtered.empty:
        st.info(f"Không có hợp đồng nào được tìm thấy cho Gói chụp: **{selected_package}**.")
        return
        
    # 4. Tính toán Metrics Tổng hợp
    num_contracts = len(df_filtered)
    total_revenue_net = df_filtered['Doanh Thu Net'].sum()
    total_cost = df_filtered['Tổng Chi Phí'].sum()
    total_profit = df_filtered['Lợi Nhuận'].sum()
    
    # Tính lợi nhuận ròng (%)
    profit_margin_percent = (total_profit / total_revenue_net) * 100 if total_revenue_net > 0 else 0
    
    # Tính trung bình trên mỗi hợp đồng
    avg_revenue = total_revenue_net / num_contracts
    avg_cost = total_cost / num_contracts
    avg_profit = total_profit / num_contracts

    st.markdown("---")
    st.write(f"##### 🎯 Kết quả Tổng hợp cho Gói: **{selected_package}** ({num_contracts} Hợp đồng)")

    col1, col2, col3, col4 = st.columns(4)
    
    col1.metric("Tổng Doanh Thu Net", f"{Fmt_Money_Str(total_revenue_net)} VNĐ")
    col2.metric("Tổng Chi Phí (Cost + Cast)", f"{Fmt_Money_Str(total_cost)} VNĐ")
    col3.metric("Tổng Lợi Nhuận", f"{Fmt_Money_Str(total_profit)} VNĐ", delta_color="inverse", delta=f"{'Tăng' if total_profit >= 0 else 'Giảm'}")
    col4.metric("Tỷ suất Lợi nhuận (P%)", f"{profit_margin_percent:,.2f} %")
    
    st.markdown("---")
    st.write("##### 📈 Lợi nhuận Trung bình trên mỗi Hợp đồng")

    col5, col6, col7 = st.columns(3)
    col5.metric("Doanh thu TB/HĐ", f"{Fmt_Money_Str(avg_revenue)} VNĐ")
    col6.metric("Chi phí TB/HĐ", f"{Fmt_Money_Str(avg_cost)} VNĐ")
    col7.metric("Lợi nhuận TB/HĐ", f"{Fmt_Money_Str(avg_profit)} VNĐ", delta_color="inverse", delta=f"{'Tăng' if avg_profit >= 0 else 'Giảm'}")

    st.markdown("---")
    st.write("##### Chi tiết các Hợp đồng (Đã tính Lợi nhuận)")
    
    # Chuẩn bị DataFrame chi tiết
    df_detail_display = df_filtered.copy()
    
    # Chọn và format các cột cần thiết
    detail_cols = ['Mã hợp đồng', 'Khách hàng', 'NgàyGiờ', 'Trạng thái', 
                   'Doanh Thu Net', 'Tổng Chi Phí', 'Lợi Nhuận', 'Giá trị hợp đồng', 'Chi phí cost', 'Cast Nhân Viên']
    
    # Đảm bảo chỉ hiển thị các cột tồn tại sau tính toán
    final_detail_cols = [col for col in detail_cols if col in df_detail_display.columns]

    for col in ['Doanh Thu Net', 'Tổng Chi Phí', 'Lợi Nhuận']:
        if col in df_detail_display.columns:
            df_detail_display[col] = df_detail_display[col].apply(Fmt_Money_Str)

    # Hiển thị
    st.dataframe(
        df_detail_display[final_detail_cols].sort_values(by='NgàyGiờ', ascending=False),
        use_container_width=True,
        hide_index=True,
        column_order=['Mã hợp đồng', 'Khách hàng', 'NgàyGiờ', 'Doanh Thu Net', 'Tổng Chi Phí', 'Lợi Nhuận', 'Trạng thái']
    )


# --- HÀM MỚI: RENDER ADMIN DATA (QUẢN LÝ HỢP ĐỒNG TRỰC TIẾP) ---
def Render_Admin_Data(ws_contract):
    """
    Hiển thị dữ liệu Hợp đồng trong một Data Editor để Quản lý (Admin) có thể chỉnh sửa trực tiếp.
    Lưu ý: Chức năng này sẽ chỉ hoạt động nếu user có quyền ghi vào sheet.
    """
    st.write("##### 🛠️ Quản Lý Hợp Đồng Trực Tiếp (Admin Edit)")
    st.warning("⚠️ **CHỨC NĂNG NÀY CHỈ DÀNH CHO QUẢN LÝ.** Mọi thay đổi sẽ được ghi trực tiếp vào Google Sheet 'Hop Dong'.")

    if st.session_state.get('role') != 'admin':
         st.error("Bạn không có quyền truy cập chức năng này.")
         return
         
    if ws_contract is None:
        st.error("Không thể kết nối đến sheet 'Hop Dong' để chỉnh sửa.")
        return

    # Lấy data và columns
    df_admin = st.session_state.get('df_admin')
    if df_admin is None:
        try:
            df, cols = Load_Booking_Data_From_Sheet(ws_contract)
            df_admin = df
            df_admin.columns = cols # Đảm bảo tên cột khớp
            st.session_state['df_admin'] = df_admin
        except Exception as e:
            st.error(f"Lỗi khi tải dữ liệu Admin: {e}")
            return
            
    col_refresh, col_save = st.columns([1, 4])
    if col_refresh.button("🔄 Tải lại dữ liệu Sheet", key="refresh_admin_data_btn"):
        del st.session_state['df_admin'] # Xóa cache admin
        Clear_Data_Cache()
        st.rerun()

    # --- Data Editor ---
    st.markdown("---")
    
    # Chuyển các cột tiền tệ sang kiểu int để data editor hiển thị rõ ràng hơn (nếu là số)
    df_for_edit = df_admin.copy()
    money_cols = ['Cast Nhân Viên', 'Giá trị hợp đồng', 'Chi phí phát sinh', 'Tiền cọc', 'Giảm giá', 'Chi phí cost', 'Số tiền còn lại']
    
    # Loại bỏ các cột không tồn tại trước khi xử lý
    valid_money_cols = [col for col in money_cols if col in df_for_edit.columns]
    
    for col in valid_money_cols:
         df_for_edit[col] = pd.to_numeric(df_for_edit[col].replace('', '0'), errors='coerce').fillna(0).astype('Int64') # Int64 để cho phép NaN

    # Cấu hình widget edit
    column_config = {
        "Mã hợp đồng": st.column_config.TextColumn(disabled=True),
        "Trạng thái": st.column_config.SelectboxColumn(options=['Chưa cọc', 'Đã lên lịch', 'Hậu kỳ', 'Chọn ảnh in', 'Hoàn thành']),
        "Giá trị hợp đồng": st.column_config.NumberColumn(format="%.0f VNĐ", step=100000),
        "Tiền cọc": st.column_config.NumberColumn(format="%.0f VNĐ", step=100000),
        "Giảm giá": st.column_config.NumberColumn(format="%.0f VNĐ", step=10000),
        "Chi phí phát sinh": st.column_config.NumberColumn(format="%.0f VNĐ", step=10000),
        "Chi phí cost": st.column_config.NumberColumn(format="%.0f VNĐ", step=10000),
        "Cast Nhân Viên": st.column_config.NumberColumn(format="%.0f VNĐ", step=10000),
        "NgàyGiờ": st.column_config.TextColumn("Ngày Giờ (Start - End)"), # Giữ nguyên text để tránh lỗi format
        "Link sản phẩm": st.column_config.TextColumn(width="small"),
    }
    
    edited_df = st.data_editor(
        df_for_edit,
        column_config=column_config,
        num_rows="dynamic",
        use_container_width=True,
        key="admin_contract_editor"
    )

    if col_save.button("💾 LƯU CÁC THAY ĐỔI VÀO GOOGLE SHEET", type="primary"):
        try:
            # So sánh edited_df với df_admin gốc để tìm thay đổi
            df_edited_raw = edited_df.fillna('').astype(str)
            df_admin_raw = df_admin.fillna('').astype(str)

            # --- Logic Cập nhật GSheet ---
            # 1. Chuẩn bị dữ liệu để gửi lên sheet (bao gồm header)
            data_to_save = [list(df_admin.columns)] + edited_df.fillna('').values.tolist()
            
            # 2. Xóa và ghi lại toàn bộ sheet (Cách an toàn nhất cho gspread data editor)
            ws_contract.clear()
            ws_contract.update(data_to_save, value_input_option='USER_ENTERED')
            
            # 3. Cập nhật cache và thông báo
            st.session_state['df_admin'] = edited_df # Cập nhật cache admin
            Clear_Data_Cache()
            st.success("✅ Đã lưu thành công các thay đổi vào Google Sheet 'Hop Dong'.")
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ Lỗi khi lưu vào Google Sheet: {e}. Vui lòng thử lại.")
            
    # Luôn hiển thị dữ liệu gốc nếu chưa lưu để tránh mất dữ liệu
    st.caption("Dữ liệu hiện tại trong bảng là phiên bản đang chỉnh sửa. Nhấn nút **LƯU** để lưu vào Google Sheet.")

# --- MAIN APPLICATION LOGIC ---

def main():
    st.set_page_config(page_title="HauFoto ERP", layout="wide", page_icon="📸")
    # --- Tối ưu hóa: Giảm hiệu ứng CSS cho button ---
    st.markdown("""<style>.stButton>button {height: 2.5em; font-weight: bold; border-radius: 6px; padding: 0.5em 1em; margin: 0; transition: none;} div[data-testid="stExpander"] {border: 1px solid #ddd; border-radius: 8px;}</style>""", unsafe_allow_html=True)
    st.title("📸 HAU FOTO STUDIO MANAGER")

    is_logged_in = Authenticate()

    # Khởi tạo/duy trì Session State
    if 'booking_type' not in st.session_state: st.session_state.booking_type = None
    if 'current_crm_data' not in st.session_state: st.session_state.current_crm_data = None
    if 'current_data_title' not in st.session_state: st.session_state.current_data_title = "Vui lòng chọn loại Booking để xem"
    if 'current_data_cols' not in st.session_state: st.session_state.current_data_cols = []
    
    # Khởi tạo state cho việc chọn loại booking trong tab BOOKING_NEW
    if 'selected_booking_type' not in st.session_state: st.session_state.selected_booking_type = "wedding"
    if 'df_client_db' not in st.session_state: st.session_state.df_client_db = None
    if 'df_admin' not in st.session_state: st.session_state.df_admin = None


    tab_order_keys = Load_Tab_Order()

    # --- KHỐI XỬ LÝ KẾT NỐI VÀ KHỞI TẠO TABS ---
    if is_logged_in:
        # Tải 6 giá trị trả về
        result = Init_Connection()
        # Đảm bảo kết nối trả về đủ 6 giá trị
        if result is None or len(result) < 6 or result[0] is None: 
            st.error("Lỗi kết nối Google!"); 
            if st.button("🔧 Xóa Cache & Reconnect"): Clear_Data_Cache(); os.remove(TOKEN_FILE) if os.path.exists(TOKEN_FILE) else None; st.rerun()
            return
            
        cal_service, ws_main, ws_personal, ws_wedding, ws_contract, ws_client_db = result
        
        # Lưu các worksheet vào session state để dễ dàng truy cập trong các hàm con
        st.session_state.ws_personal = ws_personal
        st.session_state.ws_wedding = ws_wedding
        st.session_state.ws_contract = ws_contract
        st.session_state.ws_client_db = ws_client_db
        
        ordered_tab_names = [DEFAULT_STAFF_TABS_CONFIG[key] for key in tab_order_keys]
        tabs = st.tabs(ordered_tab_names)
        ordered_tabs_by_key = {key: tabs[i] for i, key in enumerate(tab_order_keys)}

    else:
        tabs = st.tabs(["📝 BOOKING KHÁCH"])
        
    # --- KHỐI LOGIC HIỂN THỊ NỘI DUNG ---

    if not is_logged_in:
        # --- 0. BOOKING KHÁCH (Chưa đăng nhập) ---
        with tabs[0]: 
            st.subheader("📝 Đặt Lịch Chụp Hình - Hau Foto Studio")
            st.info("Vui lòng chọn loại dịch vụ bạn quan tâm để điền thông tin khảo sát booking.")

            c1, c2 = st.columns(2)
            if c1.button("👤 Booking Cá Nhân / Gia Đình", type="primary", use_container_width=True): st.session_state.booking_type = 'personal'
            if c2.button("👰🤵 Booking Weddings / Cặp Đôi", type="primary", use_container_width=True): st.session_state.booking_type = 'wedding'

            st.markdown("---")
            
            if st.session_state.booking_type == 'personal':
                st.write(f"##### Form Khảo Sát Booking Cá Nhân")
                FORM_URL_PERSONAL = "https://forms.gle/avMDvrKpS6E3gXji8" 
                components.html(f'<iframe src="{FORM_URL_PERSONAL}" width="100%" height="800" frameborder="0"></iframe>', height=850)
            elif st.session_state.booking_type == 'wedding':
                st.write(f"##### Form Khảo Sát Booking Weddings")
                FORM_URL_WEDDING = "https://forms.gle/AmcFfNhvjcrkQDPP9" # Đã sửa lại URL FORM_URL_WEDDING
                components.html(f'<iframe src="{FORM_URL_WEDDING}" width="100%" height="800" frameborder="0"></iframe>', height=850)
            elif st.session_state.booking_type is None:
                st.caption("👈 Nhấn vào một trong hai nút trên để hiển thị form khảo sát.")
    
    else: # is_logged_in == True
        # --- 1. CRM ---
        if "CRM" in ordered_tabs_by_key:
            with ordered_tabs_by_key["CRM"]: 
                st.subheader("📋 Quản Lý & Phân loại Booking Khảo Sát")
                st.write("##### Tải Dữ liệu Booking từ Google Sheets")
                
                col_b1, col_b2 = st.columns(2)
                
                def handle_load_data(worksheet, title):
                    df, cols = Load_Booking_Data_From_Sheet(worksheet)
                    st.session_state.current_crm_data = df
                    st.session_state.current_data_cols = cols
                    st.session_state.current_data_title = title
                    if worksheet is None: st.warning(f"Sheet **'{title}'** không tồn tại. Vui lòng tạo sheet này.")

                if col_b1.button("👤 Tải & Xem: BOOKING CÁ NHÂN", use_container_width=True, type="primary", key="load_crm_personal_btn"):
                    handle_load_data(ws_personal, "Booking Cá Nhân")

                if col_b2.button("👰🤵 Tải & Xem: BOOKING WEDDINGS", use_container_width=True, type="primary", key="load_crm_wedding_btn"):
                    handle_load_data(ws_wedding, "Booking Weddings")

                df = st.session_state.get('current_crm_data')
                cols = st.session_state.get('current_data_cols')
                st.markdown("---")
                st.write(f"##### {st.session_state.current_data_title}")
                
                if df is not None and not df.empty:
                    st.dataframe(df.dropna(axis=1, how='all'), use_container_width=True)
                elif cols:
                    empty_df_with_cols = pd.DataFrame(columns=cols)
                    st.info(f"Sheet **'{st.session_state.current_data_title}'** hiện **chưa có dữ liệu**, nhưng đã tìm thấy cấu trúc cột (header).")
                    st.dataframe(empty_df_with_cols, use_container_width=True)
                else:
                    st.info("Vui lòng chọn một loại Booking để tải dữ liệu.")
            
        # --- 2. BOOKING (Pos 2) ---
        if "BOOKING_NEW" in ordered_tabs_by_key:
            with ordered_tabs_by_key["BOOKING_NEW"]:
                st.subheader("📅 Quản Lý Booking")
                
                # --- LOGIC GỘP TAB BOOKING ---
                st.write("##### 📝 Tạo Hợp Đồng Mới")
                
                # Tạo selectbox để chọn loại booking
                booking_options = {
                    "wedding": "👰🤵 Booking Weddings (Hợp đồng CĐ/CR)",
                    "personal": "👤 Booking Cá Nhân (Hợp đồng Cá nhân/Gia đình)"
                }
                
                # Lưu key của loại booking đang chọn vào session state để dùng cho Render_Booking_Form
                selected_booking_key = st.selectbox(
                    "Chọn Loại Booking:",
                    options=list(booking_options.keys()),
                    format_func=lambda x: booking_options[x],
                    key='selected_booking_type' # key này được dùng để render form
                )
                
                # Render Form dựa trên lựa chọn
                if selected_booking_key:
                    worksheet_name = "Booking Weddings" if selected_booking_key == 'wedding' else "Booking Cá Nhân"
                    Render_Booking_Form(worksheet_name, selected_booking_key, ws_contract, ws_client_db, cal_service)

        # --- 3. SALARY (Pos 3) ---
        if "SALARY" in ordered_tabs_by_key:
            with ordered_tabs_by_key["SALARY"]:
                st.subheader("💰 Quản Lý Lương & Thu Chi Ekip")
                Render_Salary(ws_contract)

        # --- 4. CLIENT (Pos 4) ---
        if "CLIENT" in ordered_tabs_by_key:
            with ordered_tabs_by_key["CLIENT"]:
                st.subheader("👤 Danh sách Khách Hàng")
                Render_Client_Db(ws_client_db)
        
        # --- 5. CONTRACT (Pos 5) ---
        if "CONTRACT" in ordered_tabs_by_key:
            with ordered_tabs_by_key["CONTRACT"]:
                st.subheader("📄 Quản Lý Hợp Đồng")
                # GỌI HÀM HIỂN THỊ HỢP ĐỒNG ĐÃ ĐƯỢC PHÁT TRIỂN
                Render_Contract_List(ws_contract)
                
        # --- 6. CALENDAR (Pos 6) ---
        if "CALENDAR" in ordered_tabs_by_key:
            with ordered_tabs_by_key["CALENDAR"]:
                st.subheader("🗓️ Lịch Chụp & Lịch Làm Việc")
                # GỌI HÀM HIỂN THỊ LỊCH
                Render_Calendar(cal_service)

        # --- 7. PNL (Pos 7 Mới) ---
        if "PNL" in ordered_tabs_by_key:
            with ordered_tabs_by_key["PNL"]:
                st.subheader("📈 Phân Tích Lợi Nhuận (P&L) theo Gói Chụp")
                Render_PNL_Analysis(ws_contract)
                
        # --- 8. ADMIN --- (Pos 8 Mặc định mới)
        if "ADMIN" in ordered_tabs_by_key:
            with ordered_tabs_by_key["ADMIN"]:
                st.subheader("🛠️ Admin Data: Quản Lý Hợp Đồng Trực Tiếp")
                Render_Admin_Data(ws_contract)
            
        # --- 9. NHÂN SỰ --- (Pos 9 Mặc định mới)
        if "STAFF" in ordered_tabs_by_key:
            with ordered_tabs_by_key["STAFF"]:
                st.subheader("👥 Quản Lý Nhân sự")
                
                staff_db = Load_Staff_Db()
                df_staff = pd.DataFrame(staff_db)
                
                # --- KHỐI THÊM/CHỈNH SỬA NHÂN SỰ ---
                st.write("##### 📝 Thêm/Chỉnh sửa Nhân sự")
                
                staff_names = ["(Thêm mới)"] + df_staff['Tên'].tolist()
                selected_name = st.selectbox("Chọn nhân sự để chỉnh sửa/xóa:", options=staff_names)
                
                if selected_name != "(Thêm mới)":
                    current_staff = next(item for item in staff_db if item.get('Tên') == selected_name)
                    is_editing = True
                else:
                    current_staff = {}
                    is_editing = False

                default_role_key = current_staff.get("Role App", "staff")
                
                if 'form_role_app' not in st.session_state or is_editing: 
                    st.session_state.form_role_app = default_role_key

                # --- SỬ DỤNG FORM ĐỂ THÊM/SỬA (LỖI data_editor đã được FIX bằng cách không dùng nó)
                with st.form("staff_form", clear_on_submit=False):
                    
                    # --- Hàng Dọc (Tên, Khả năng, SĐT, Email, Địa điểm) ---
                    new_name = st.text_input("Tên", value=current_staff.get("Tên", ""), disabled=is_editing and selected_name != "(Thêm mới)")
                    new_kh = st.text_input("Khả năng", value=current_staff.get("Khả năng", ""))
                    phone_input = st.text_input("Số điện thoại", value=current_staff.get("Số điện thoại", ""), key="staff_phone_input")
                    new_email = st.text_input("Email", value=current_staff.get("Email", ""))
                    new_loc = st.text_input("Địa điểm", value=current_staff.get("Địa điểm", ""))

                    st.markdown("---")

                    # --- Hàng Ngang (Account/Password vs Role/Xóa) ---
                    c1, c2 = st.columns(2)
                    
                    # Role App (Quyền) - Cột 2
                    role_options_vn = list(ROLE_MAP_VN_TO_KEY.keys())
                    default_role_vn = ROLE_MAP_KEY_TO_VN.get(default_role_key, "Nhân Viên")
                    
                    new_role_vn = c2.selectbox(
                        "Role App (Quyền)", 
                        options=role_options_vn, 
                        index=role_options_vn.index(default_role_vn),
                        key='role_app_select'
                    )
                    
                    st.session_state.form_role_app = ROLE_MAP_VN_TO_KEY[new_role_vn]
                    is_partner = (st.session_state.form_role_app == 'partner')
                    
                    # Account - Cột 1
                    new_acc = c1.text_input(
                        "Tên đăng nhập (Account)", 
                        value=current_staff.get("Account", ""), 
                        disabled=is_partner,
                        help="Account bị vô hiệu hóa khi Role là Đối tác."
                    )
                    # Password - Cột 1
                    new_pass = c1.text_input(
                        "Mật khẩu (Password)", 
                        value=current_staff.get("Password", ""), 
                        type="password", 
                        disabled=is_partner,
                        help="Password bị vô hiệu hóa khi Role là Đối tác."
                    )

                    col_save, col_delete = st.columns([1, 1])
                    save_button = col_save.form_submit_button("💾 Lưu/Cập nhật", type="primary")
                    delete_button = col_delete.form_submit_button("🗑️ Xóa Nhân sự", disabled=not is_editing)
                    
                    if delete_button and is_editing:
                        staff_db[:] = [item for item in staff_db if item['Tên'] != selected_name]
                        Save_Staff_Db(staff_db)
                        Update_User_From_Staff(staff_db) 
                        st.toast(f"🗑️ Đã xóa nhân sự: {selected_name}.")
                        st.rerun()

                    if save_button:
                        if not new_name or not new_kh: st.error("Tên và Khả năng là bắt buộc."); st.stop()
                        if not is_partner and (not new_acc or not new_pass): st.error("Account và Password là bắt buộc cho Quản Lý và Nhân Viên."); st.stop()
                        
                        new_staff_data = {
                            "Tên": new_name, "Khả năng": new_kh, "Địa điểm": new_loc,
                            "Account": new_acc if not is_partner else "", "Password": new_pass if not is_partner else "",
                            "Role App": st.session_state.form_role_app, "Số điện thoại": phone_input, "Email": new_email
                        }
                        
                        if is_editing:
                            idx = next(i for i, item in enumerate(staff_db) if item.get('Tên') == selected_name)
                            staff_db[idx] = new_staff_data
                            st.toast(f"✅ Đã cập nhật thông tin cho {new_name}.")
                        else:
                            if new_acc != "" and any(item.get('Account') == new_acc for item in staff_db):
                                st.error(f"Tên đăng nhập '{new_acc}' đã tồn tại. Vui lòng chọn Account khác."); st.stop()
                            staff_db.append(new_staff_data); st.toast(f"✅ Đã thêm nhân sự mới: {new_name}.")

                        Save_Staff_Db(staff_db); Update_User_From_Staff(staff_db); st.rerun()

                # --- KHỐI DANH SÁCH NHÂN SỰ ---
                st.markdown("---")
                st.write("##### 📋 Danh sách Nhân sự hiện tại")
                
                if not df_staff.empty:
                    df_staff['Password (Ẩn)'] = df_staff['Password'].apply(lambda x: x[:3] + '***' if isinstance(x, str) and len(x) > 3 else '***')
                    df_staff['Role App (Quyền)'] = df_staff['Role App'].map(lambda x: ROLE_MAP_KEY_TO_VN.get(x, x))
                    
                    cols_order = ['Tên', 'Khả năng', 'Số điện thoại', 'Email', 'Địa điểm', 'Account', 'Password (Ẩn)', 'Role App (Quyền)']
                    cols_to_display = [col for col in cols_order if col in df_staff.columns]
                    # FIX: Sử dụng st.dataframe thay vì st.data_editor
                    st.dataframe(df_staff[cols_to_display], use_container_width=True, hide_index=True)
                else: st.info("Chưa có nhân sự nào được thêm.")


        # --- 10. CÀI ĐẶT --- (Pos 10 Mặc định mới)
        if "SETTINGS" in ordered_tabs_by_key:
            with ordered_tabs_by_key["SETTINGS"]:
                st.subheader("⚙️ Cài Đặt Hệ Thống")
                st.info("Khu vực này sẽ chứa các thiết lập chính cho ứng dụng (ví dụ: Cấu hình Sheet, Lịch...).")

                st.markdown("---")
                
                current_username = st.session_state.username
                current_staff_name = st.session_state.staff_name
                staff_db = Load_Staff_Db()
                current_staff_entry = next((s for s in staff_db if s.get('Account') == current_username), None)
                
                is_valid_role = current_staff_entry and current_staff_entry.get("Role App") in ["admin", "staff"]

                if not is_valid_role:
                    st.error("Bạn không có quyền chỉnh sửa thông tin cá nhân và mật khẩu.")
                else:
                    # --- Expander 1: THÔNG TIN CÁ NHÂN ---
                    with st.expander("👤 Thông tin cá nhân"):
                        st.write("##### Cập nhật Thông tin cá nhân (Tên, Email, SĐT)")
                        
                        stored_name = current_staff_entry.get("Tên", current_staff_name) 
                        stored_phone = current_staff_entry.get("Số điện thoại", "") 
                        stored_email = current_staff_entry.get("Email", "") 

                        with st.form("personal_info_form", clear_on_submit=False):
                            new_name = st.text_input("Tên hiển thị:", value=stored_name, key="info_new_name_input")
                            new_phone = st.text_input("Số điện thoại:", value=stored_phone, key="info_new_phone_input")
                            new_email = st.text_input("Email:", value=stored_email, key="info_new_email_input")
                            
                            submit_info_button = st.form_submit_button("✅ Cập nhật Thông tin cá nhân", type="primary")

                            if submit_info_button:
                                if not new_name: st.error("Tên hiển thị không được để trống."); st.stop()
                                Update_Staff_Db(staff_db, current_username, {"Tên": new_name, "Số điện thoại": new_phone, "Email": new_email})
                    
                    st.markdown("---")

                    # --- Expander 2: ĐỔI MẬT KHẨU & ACCOUNT ---
                    with st.expander("🔑 Đổi Mật khẩu & Account"):
                        st.write("##### Đổi Mật khẩu và Tên đăng nhập (Username)")
                        
                        stored_password = current_staff_entry.get("Password")
                        
                        with st.form("change_credentials_form", clear_on_submit=False):
                            st.caption(f"Đang chỉnh sửa cho tài khoản: **{current_username}**")
                            
                            new_username = st.text_input("Tên đăng nhập (Account) mới:", value=current_username, key="change_new_username")
                            
                            st.markdown("###### Xác nhận Thông tin hiện tại")
                            confirm_username = st.text_input("Nhập lại Account (Username) hiện tại:", key="change_confirm_username_input")
                            current_password_input = st.text_input("Nhập lại Mật khẩu hiện tại:", type="password", key="change_current_password_input")
                            
                            st.markdown("###### Mật khẩu mới")
                            new_password = st.text_input("Mật khẩu mới:", type="password", key="change_new_password_input")
                            confirm_password = st.text_input("Xác nhận Mật khẩu mới:", type="password", key="change_confirm_password_input")
                            
                            submit_button = st.form_submit_button("✅ Cập nhật Mật khẩu & Account", type="primary")
                            
                            if submit_button:
                                if confirm_username != current_username: st.error("Xác nhận Account hiện tại không khớp."); st.stop()
                                if current_password_input != stored_password: st.error("Mật khẩu hiện tại không đúng. Vui lòng kiểm tra lại."); st.stop()
                                if new_username != current_username and any(item.get('Account') == new_username for item in staff_db if item.get('Account') != current_username): st.error(f"Tên đăng nhập '{new_username}' đã tồn tại. Vui lòng chọn Account khác."); st.stop()
                                if not new_password: st.error("Mật khẩu mới không được để trống."); st.stop()
                                if new_password != confirm_password: st.error("Mật khẩu mới và xác nhận mật khẩu không khớp."); st.stop()
                                
                                Update_Staff_Db(staff_db, current_username, {"Password": new_password, "Account": new_username}, is_acc_change=(new_username != current_username))

                    st.markdown("---")

                    # --- Expander 3: Cấu hình Thứ tự Tabs ---
                    with st.expander("🔧 Cấu hình Thứ tự Tabs"):
                        st.write("##### Thay đổi vị trí các tab")
                        
                        current_order_keys = Load_Tab_Order()
                        new_order_keys = current_order_keys[:]
                        col_sel, col_btn = st.columns([3, 1])

                        is_changed = False
                        for i, tab_key in enumerate(current_order_keys):
                            tab_name = DEFAULT_STAFF_TABS_CONFIG[tab_key]
                            
                            selected_pos = col_sel.selectbox(
                                f"Vị trí Tab **{tab_name}**:",
                                options=range(1, len(DEFAULT_STAFF_TABS_CONFIG) + 1),
                                index=new_order_keys.index(tab_key), 
                                key=f"tab_pos_{tab_key}"
                            )
                            
                            target_index = selected_pos - 1
                            
                            if new_order_keys.index(tab_key) != target_index:
                                new_order_keys.remove(tab_key)
                                new_order_keys.insert(target_index, tab_key)
                                is_changed = True
                                break 

                        if is_changed:
                            Save_Tab_Order(new_order_keys)
                            st.toast(f"Đã thay đổi thứ tự Tab. Đang tải lại...", icon="🔄")
                            st.rerun()
                        
                        st.markdown("---")
                        st.write("###### Thứ tự Tabs hiện tại:")
                        st.code(" -> ".join([DEFAULT_STAFF_TABS_CONFIG[key] for key in current_order_keys]))


if __name__ == '__main__':
    # Khối try/except ở cuối main đã được dọn dẹp vì logic font đã được chuyển vào hàm Create_Pro_Pdf
    main()
