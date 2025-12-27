import os
import re
import time
import platform
import logging
import json
import pickle
import socket
from threading import Thread, Event

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import openpyxl 
from openpyxl.utils import get_column_letter

# Kivy Imports
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.scrollview import ScrollView
from kivy.uix.progressbar import ProgressBar
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.utils import get_color_from_hex
from kivy.properties import StringProperty, NumericProperty, BooleanProperty
from kivy.metrics import dp
from kivy.core.text import LabelBase

# --- ARABIC SUPPORT FIX ---
try:
    import arabic_reshaper
    ARABIC_RESHAPER_AVAILABLE = True
except ImportError:
    ARABIC_RESHAPER_AVAILABLE = False

try:
    from bidi.algorithm import get_display
    BIDI_LIBS_AVAILABLE = True
except ImportError:
    BIDI_LIBS_AVAILABLE = False

def fix_text(text):
    """إصلاح السطر بالكامل ليدعم العربية المختلطة مع الإنجليزية"""
    if not text: return ""
    text = str(text)
    processed_text = text
    if ARABIC_RESHAPER_AVAILABLE:
        try:
            # تشكيل الحروف
            processed_text = arabic_reshaper.reshape(text)
        except Exception: pass 
    if BIDI_LIBS_AVAILABLE:
        try:
            # ضبط الاتجاه (Bidi Algorithm) للسطر كاملاً
            processed_text = get_display(processed_text)
        except Exception: pass
    return processed_text

# --- FONT CONFIGURATION ---
FONT_NAME = 'ArabicFont'
FONT_FILENAME = 'font.ttf'
APP_FONT = 'Roboto' 
if os.path.exists(FONT_FILENAME):
    try:
        LabelBase.register(name=FONT_NAME, fn_regular=FONT_FILENAME)
        APP_FONT = FONT_NAME
    except Exception: pass

# --- Configuration ---
BASE_URL = "https://arkan-int.joodbooking.com"
LOGIN_PAGE = "/Account/Login?ReturnUrl=%2FBookingWorkflow%2FIndex"
LOGIN_POST = "/Account/Login"
FINANCIAL_STATUS_PAGE = "/FinancialStatus/CustomerFinancialStatus"
ACCOUNT_STATEMENT_PAGE = "/Finance/AccountStatement"
# تم تحديث الرابط بناءً على تحليل الملفات المرفقة
GET_ACCOUNT_STATEMENT_API = "/Finance/GetAccountStatement" 
REPORT_GEN_POST = "/Finance/ReportAccountStatement"
REPORT_VIEWER_BASE = "/Reports/Viewer.aspx"
PDF_AXD_ENDPOINT = "/Reserved.ReportViewerWebControl.axd"
CUSTOMER_FINANCIAL_STATUS_GET = "/FinancialStatus/GetCustomerFinancialStatus"

USERNAME = os.getenv('BOOKING_USERNAME', '')
PASSWORD = os.getenv('BOOKING_PASSWORD', '')
SESSION_FILE = "session_cookies.pkl"
CUSTOMERS_CACHE_FILE = "customers_cache.json"
CACHE_EXPIRY_HOURS = 24

BASE_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
    'Origin': BASE_URL
}

socket.setdefaulttimeout(30)
Window.clearcolor = get_color_from_hex('#0f1419')

# --- Backend Functions ---

def save_session_cookies(session, filename=SESSION_FILE):
    try:
        with open(filename, 'wb') as f:
            pickle.dump(session.cookies.get_dict(), f)
    except Exception: pass

def load_session_cookies(session, filename=SESSION_FILE):
    if not os.path.exists(filename): return False
    try:
        with open(filename, 'rb') as f:
            cookies = pickle.load(f)
            session.cookies.update(cookies)
        return True
    except Exception: return False

def perform_full_login(session, username, password):
    session.headers['Referer'] = BASE_URL + LOGIN_PAGE
    session.headers.pop('X-Requested-With', None)
    get_response = session.get(BASE_URL + LOGIN_PAGE)
    session.headers['X-Requested-With'] = 'XMLHttpRequest'
    if get_response.status_code != 200: return False
    soup = BeautifulSoup(get_response.text, 'html.parser')
    token_tag = soup.find('input', {'name': '__RequestVerificationToken'})
    if not token_tag: return False
    dynamic_token = token_tag.get('value')
    session.headers['Content-Type'] = 'application/x-www-form-urlencoded'
    payload = {'__RequestVerificationToken': dynamic_token, 'UserName': username, 'Password': password, 'RememberMe': 'true'}
    post_response = session.post(BASE_URL + LOGIN_POST, data=payload, allow_redirects=False)
    if post_response.status_code == 200:
        try: return post_response.json().get('success') is True
        except: return False
    elif post_response.status_code == 302: return True
    return False

def download_and_cache_customers(session):
    all_customers = {}
    page, page_size = 1, 500
    session.headers['Referer'] = BASE_URL + FINANCIAL_STATUS_PAGE
    session.headers['X-Requested-With'] = 'XMLHttpRequest'
    while True:
        search_params = {'AgencyType': '4', 'page': str(page), 'pageSize': str(page_size)}
        response = session.get(BASE_URL + CUSTOMER_FINANCIAL_STATUS_GET, params=search_params)
        if response.status_code == 200:
            try:
                data = response.json()
                customers = data.get('data', [])
                if not customers: break
                for c in customers:
                    cid, cname = str(c.get('CustomerId', '')), c.get('CustomerName', '')
                    if cid and cname: all_customers[cid] = cname
                if len(customers) < page_size: break
                page += 1
            except: break
        else: break
    cache_data = {'timestamp': time.time(), 'customers': all_customers}
    try:
        with open(CUSTOMERS_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache_data, f, ensure_ascii=False, indent=2)
    except: pass
    return all_customers

def load_customers_cache():
    if not os.path.exists(CUSTOMERS_CACHE_FILE): return None
    try:
        with open(CUSTOMERS_CACHE_FILE, 'r', encoding='utf-8') as f:
            cache_data = json.load(f)
        if (time.time() - cache_data.get('timestamp', 0)) / 3600 > CACHE_EXPIRY_HOURS: return None
        return cache_data.get('customers', {})
    except: return None

def get_all_client_names(session):
    cached = load_customers_cache()
    return cached if cached else download_and_cache_customers(session)

def get_client_name_from_dict(client_id, customers_dict):
    name = customers_dict.get(str(client_id))
    return name if name else f"Client_{client_id}"

def access_account_statement_page(session, agency_id):
    session.headers.pop('X-Requested-With', None)
    session.headers['Referer'] = BASE_URL + FINANCIAL_STATUS_PAGE
    account_url = f"{BASE_URL}{ACCOUNT_STATEMENT_PAGE}/{agency_id}"
    response = session.get(account_url)
    session.headers['X-Requested-With'] = 'XMLHttpRequest'
    return response.text if response.status_code == 200 else None

def extract_transactions_from_page(html_content):
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        checkboxes = soup.find_all('input', {'name': 'Transactions', 'type': 'checkbox'})
        if checkboxes:
            ids = [cb.get('value') for cb in checkboxes if cb.get('value')]
            if ids: return ','.join(ids)
        return ','.join(map(str, range(30000)))
    except: return ','.join(map(str, range(30000)))

def get_report_id(session, agency_id, report_token, transactions_list, from_date='01/01/2025', to_date=''):
    session.headers['Content-Type'] = 'application/x-www-form-urlencoded; charset=UTF-8'
    gen_payload = {
        '__RequestVerificationToken': report_token, 
        'Transactions': transactions_list, 
        'fromDate': from_date, 
        'toDate': to_date, 
        'AgencyId': agency_id, 
        'CurrencyId': 'SAR', 
        'BookingStatus': '3', 
        'RoomStatus': '2'
    }
    response = session.post(BASE_URL + REPORT_GEN_POST, data=gen_payload, timeout=300)
    if response.status_code == 200:
        match = re.search(r'id=([0-9a-f\-]{36})', response.text, re.IGNORECASE)
        if match: return match.group(1)
    return None

def get_control_id_and_download_pdf(session, report_id, client_name, client_id, final_filename, progress_callback=None, stop_event=None):
    report_viewer_url = BASE_URL + REPORT_VIEWER_BASE + f"?id={report_id}"
    session.headers.pop('X-Requested-With', None)
    try:
        report_view_response = session.get(report_viewer_url, timeout=180)
        match = re.search(r'ControlID=([0-9a-fA-F]{32})', report_view_response.text)
        if not match: return False
        control_id = match.group(1)
        
        import urllib.parse
        encoded_name = urllib.parse.quote(f"Account statement: {client_name}")
        download_params = {
            'Culture': '1033', 'CultureOverrides': 'True', 'UICulture': '1033', 'UICultureOverrides': 'True',
            'ReportStack': '1', 'ControlID': control_id, 'Mode': 'true', 'OpType': 'Export',
            'FileName': encoded_name, 'ContentDisposition': 'OnlyHtmlInline', 'Format': 'PDF'
        }
        
        pdf_response = session.get(BASE_URL + PDF_AXD_ENDPOINT, params=download_params, stream=True, timeout=600)
        if pdf_response.status_code == 200:
            os.makedirs(os.path.dirname(final_filename) or '.', exist_ok=True)
            total_size = int(pdf_response.headers.get('content-length', 0))
            downloaded = 0
            with open(final_filename, 'wb') as f:
                for chunk in pdf_response.iter_content(8192):
                    if stop_event and stop_event.is_set(): return False
                    if chunk:
                        f.write(chunk); downloaded += len(chunk)
                        if progress_callback: progress_callback(downloaded, total_size)
            return True
        return False
    except: return False

# -[span_0](start_span)[span_1](start_span)-- FIXED BALANCE FUNCTION (Using GetAccountStatement)[span_0](end_span)[span_1](end_span) ---
def get_customer_balance(session, agency_id, from_date='01/01/2025'):
    """
    جلب الرصيد من نقطة النهاية الصحيحة التي تحتوي على TotalBalance
    """
    # Parameters matched from [40] request
    params = {
        'AgencyId': agency_id,
        'HotelId': 'null',
        'OperationType': '',
        'BookingStatus': '3',
        'RoomStatus': '2',
        'PostingStatus': '',
        'PaymentStatus': '',
        'findBy': '0',
        'fromDate': from_date,
        'toDate': '',
        'CurrencyId': '',
        'AmountTypeSelected': '1',
        'Amount': '',
        'HidePreviousBalance': 'false',
        'DisplayBookingDateSelected': '0',
        'GroupByDocNumber': 'false'
    }
    
    # [span_2](start_span)Headers crucial for ASP.NET MVC Ajax[span_2](end_span)
    req_headers = session.headers.copy()
    req_headers.update({
        'X-Requested-With': 'XMLHttpRequest',
        'Referer': BASE_URL + "/Finance/AccountStatement"
    })

    try:
        # Using the correct endpoint that returns TotalBalance JSON
        response = session.get(BASE_URL + GET_ACCOUNT_STATEMENT_API, params=params, headers=req_headers, timeout=20)
        
        if response.status_code == 200:
            data = response.json()
            # [span_3](start_span)Extract TotalBalance directly from the JSON root[span_3](end_span)
            # Example: "TotalBalance": "22,835.03"
            raw_balance = data.get('TotalBalance', '0.00')
            
            # Clean string for float conversion
            clean_balance = str(raw_balance).replace(',', '')
            try:
                f_val = float(clean_balance)
            except:
                f_val = 0.0
            
            # Format nicely with SAR
            formatted_balance = f"SAR {raw_balance}"
            return formatted_balance, f_val
            
    except Exception as e:
        print(f"Balance Fetch Error for {agency_id}: {e}")
    
    return "N/A", 0.0

def resolve_client_balance(session, client_id, from_date):
    balance_raw, balance_float = get_customer_balance(session, client_id, from_date)
    return balance_raw, balance_float

def download_single_pdf(session, client_id, client_name, output_dir, from_date, to_date, report_token, progress_callback=None, stop_event=None):
    try:
        if stop_event and stop_event.is_set(): return False, "Cancelled"
        html = access_account_statement_page(session, client_id)
        txs = extract_transactions_from_page(html)
        rid = get_report_id(session, client_id, report_token, txs, from_date, to_date)
        if not rid: return False, "Report failed"
        safe_name = re.sub(r'[<>:"/\\|?*]', '_', client_name)[:50]
        final_filename = os.path.join(output_dir, f"{safe_name}_Statement_{client_id}.pdf")
        if get_control_id_and_download_pdf(session, rid, client_name, client_id, final_filename, progress_callback, stop_event):
            return True, f" "
        return False, "Download failed"
    except Exception as e: return False, str(e)

def parse_client_ids(input_str):
    return [id.strip() for id in input_str.replace('\n', ' ').replace(',', ' ').split() if id.strip()]

def create_session_with_retry():
    session = requests.Session()
    session.mount("https://", HTTPAdapter(max_retries=Retry(total=3, backoff_factor=1)))
    session.headers.update(BASE_HEADERS)
    return session

def export_to_excel(data, filename):
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["ID", "Name", "Balance"])
        for row in data: ws.append([row['id'], row['name'], row['balance']])
        wb.save(filename)
        return True
    except: return False

# --- Kivy UI ---

class LogEntry(BoxLayout):
    text = StringProperty('')
    status = StringProperty('')
    
    def __init__(self, text, status='info', **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'horizontal'
        self.size_hint_y = None
        self.height = dp(35)
        self.padding = [dp(10), dp(5)]
        
        # ألوان الحالات
        self.status_colors = {'success': '#00d26a', 'error': '#ff4757', 'info': '#1e90ff', 'warning': '#ffa502'}
        
        # النص يمين
        self.label = Label(
            text=fix_text(text), 
            size_hint_x=1, 
            halign='right', 
            valign='middle', 
            font_size=dp(12), 
            font_name=APP_FONT
        )
        self.label.bind(width=lambda *x: setattr(self.label, 'text_size', (self.label.width - dp(5), None)))
        
        # الأيقونة يسار (نص فقط بدون رموز خاصة لحل مشكلة المربع)
        self.status_label = Label(
            text='-', 
            size_hint_x=None, 
            width=dp(40), 
            bold=True, 
            font_name=APP_FONT,
            font_size=dp(11),
            color=get_color_from_hex(self.status_colors.get(status, '#1e90ff'))
        )
        
        self.add_widget(self.status_label)
        self.add_widget(self.label)
        
        self.bind(text=self._update_text, status=self._update_status)

    def _update_text(self, instance, value):
        self.label.text = fix_text(value)

    def _update_status(self, instance, value):
        self.status_label.color = get_color_from_hex(self.status_colors.get(value, '#1e90ff'))
        # استخدام كلمات بدلاً من الرموز لتجنب مشاكل الخطوط
        chars = {'success': 'DONE', 'error': 'ERR', 'warning': '!', 'info': '...'} 
        self.status_label.text = chars.get(value, '-')

class FinancialStatementApp(App):
    status_text = StringProperty("Ready")
    progress_value = NumericProperty(0)
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.stop_event = Event()
        self.log_entries = {}
    
    def build(self):
        self.title = "Booking Statement Tool"
        default_output = os.path.expanduser('~/Downloads')
        if platform.system() == 'Linux' and os.path.exists('/storage/emulated/0/Download'):
            default_output = '/storage/emulated/0/Download'
        
        # الجذر الرئيسي مع هوامش مريحة
        root = BoxLayout(orientation='vertical', padding=dp(15), spacing=dp(15))
        
        # Header
        header = Label(
            text=fix_text("أداة تنزيل الكشوفات المالية"), 
            font_size=dp(22), 
            bold=True, 
            size_hint_y=None, 
            height=dp(50), 
            font_name=APP_FONT
        )
        root.add_widget(header)

        # Inputs Area
        input_box = BoxLayout(orientation='vertical', spacing=dp(15), size_hint_y=None, height=dp(330))
        input_box.padding = [0, dp(10), 0, dp(10)]
        
        def add_input(label, default_val, is_pass=False):
            # سطر الإدخال
            row = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(45), spacing=dp(15))
            
            # مربع النص
            ti = TextInput(
                text=default_val, 
                password=is_pass, 
                multiline=False, 
                font_name=APP_FONT,
                padding=[dp(10), dp(10)],
                write_tab=False,
                background_color=(0.15, 0.2, 0.25, 1),
                foreground_color=(1, 1, 1, 1)
            )
            
            # التسمية (Label)
            lbl = Label(
                text=fix_text(label), 
                size_hint_x=0.35, 
                font_name=APP_FONT, 
                halign='right',
                valign='middle'
            )
            lbl.bind(size=lbl.setter('text_size'))
            
            row.add_widget(ti)
            row.add_widget(lbl)
            input_box.add_widget(row)
            return ti

        self.username_input = add_input("المستخدم:", USERNAME)
        self.password_input = add_input("كلمة المرور:", PASSWORD, True)
        self.from_date_input = add_input("من تاريخ:", "01/01/2025")
        self.to_date_input = add_input("إلى تاريخ:", "")
        self.output_input = add_input("مسار الحفظ:", default_output)

        root.add_widget(input_box)

        # IDs Area
        id_label = Label(text=fix_text("أرقام هويات العملاء:"), size_hint_y=None, height=dp(30), font_name=APP_FONT, halign='right')
        id_label.bind(size=id_label.setter('text_size'))
        root.add_widget(id_label)
        
        self.client_input = TextInput(multiline=True, font_name=APP_FONT, padding=[dp(10), dp(10)], background_color=(0.15, 0.2, 0.25, 1), foreground_color=(1, 1, 1, 1))
        root.add_widget(self.client_input)

        # Controls
        btn_box = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(55), spacing=dp(15))
        self.stop_btn = Button(text=fix_text("إيقاف"), background_color=get_color_from_hex('#ff4757'), disabled=True, font_name=APP_FONT, bold=True)
        self.stop_btn.bind(on_press=self.stop_download)
        self.download_btn = Button(text=fix_text("بدء التنزيل"), background_color=get_color_from_hex('#00d26a'), font_name=APP_FONT, bold=True)
        self.download_btn.bind(on_press=self.start_download)
        btn_box.add_widget(self.stop_btn); btn_box.add_widget(self.download_btn)
        root.add_widget(btn_box)

        # Status & Progress
        self.status_label = Label(text=fix_text(self.status_text), size_hint_y=None, height=dp(35), font_name=APP_FONT)
        self.bind(status_text=lambda i, v: setattr(self.status_label, 'text', fix_text(v)))
        root.add_widget(self.status_label)
        
        self.progress_bar = ProgressBar(max=100, size_hint_y=None, height=dp(15))
        self.bind(progress_value=self.progress_bar.setter('value'))
        root.add_widget(self.progress_bar)

        # Logs
        log_scroll = ScrollView(size_hint_y=1)
        self.log_layout = BoxLayout(orientation='vertical', size_hint_y=None, spacing=dp(5))
        self.log_layout.bind(minimum_height=self.log_layout.setter('height'))
        log_scroll.add_widget(self.log_layout)
        root.add_widget(log_scroll)

        return root

    def add_log(self, text, status='info', client_id=None):
        def _add():
            if client_id and client_id in self.log_entries:
                entry = self.log_entries[client_id]
                entry.text, entry.status = text, status
            else:
                entry = LogEntry(text, status)
                if client_id: self.log_entries[client_id] = entry
                self.log_layout.add_widget(entry, index=0)
        Clock.schedule_once(lambda dt: _add(), 0)

    def start_download(self, instance):
        client_ids = parse_client_ids(self.client_input.text)
        if not client_ids: return
        self.stop_event.clear()
        self.download_btn.disabled, self.stop_btn.disabled = True, False
        self.log_layout.clear_widgets(); self.log_entries = {}
        Thread(target=self.download_thread, args=(client_ids,)).start()

    def stop_download(self, instance):
        self.stop_event.set()
        self.status_text = "جاري الإيقاف..."

    def download_thread(self, client_ids):
        try:
            s = create_session_with_retry()
            user, pw = self.username_input.text, self.password_input.text
            from_d, to_d = self.from_date_input.text, self.to_date_input.text
            out_dir = self.output_input.text

            if not perform_full_login(s, user, pw) and not load_session_cookies(s):
                self.add_log("فشل الدخول: تأكد من البيانات", 'error'); return
            
            save_session_cookies(s)
            customers = get_all_client_names(s)
            
            try:
                token_resp = s.get(BASE_URL + FINANCIAL_STATUS_PAGE)
                report_token_match = re.search(r'value="([^"]+)"', token_resp.text)
                report_token = report_token_match.group(1) if report_token_match else ""
            except:
                report_token = ""
            
            total_sum = 0.0
            summary = []

            for idx, cid in enumerate(client_ids, 1):
                if self.stop_event.is_set(): break
                name = get_client_name_from_dict(cid, customers)
                self.progress_value = (idx / len(client_ids)) * 100
                self.status_text = f"جاري معالجة: {name}"
                
                # جلب الرصيد من API كشف الحساب الصحيح (GetAccountStatement)
                bal_raw, bal_float = resolve_client_balance(s, cid, from_d)
                
                def prog(cur, tot):
                    pct = (cur/tot*100) if tot else 0
                    msg = f"[{idx}/{len(client_ids)}] {name} - المستحق: {bal_raw} - جاري التحميل: {pct:.0f}%"
                    self.add_log(msg, 'info', cid)

                ok, msg = download_single_pdf(s, cid, name, out_dir, from_d, to_d, report_token, prog, self.stop_event)
                
                # السطر النهائي للعميل
                final_msg = f"[{idx}/{len(client_ids)}] {name} - المستحق: {bal_raw} - {msg}"
                self.add_log(final_msg, 'success' if ok else 'error', cid)
                
                if ok: total_sum += bal_float
                summary.append({'id': cid, 'name': name, 'balance': bal_raw})
                time.sleep(0.5)

            export_to_excel(summary, os.path.join(out_dir, "Summary.xlsx"))
            final_status = f"الإجمالي: SAR {total_sum:,.2f}"
            self.status_text = final_status
            self.add_log(final_status, 'success')

        except Exception as e:
            self.add_log(f"خطأ: {str(e)}", 'error')
        finally:
            Clock.schedule_once(lambda dt: self.finish(), 0)

    def finish(self):
        self.download_btn.disabled, self.stop_btn.disabled = False, True

if __name__ == '__main__':
    FinancialStatementApp().run()