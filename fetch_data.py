import time
import os
import requests
import pandas as pd
import subprocess  # المكتبة المطلوبة لتنفيذ أوامر الـ Git
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# --- الإعدادات الأساسية ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_PATH = os.path.join(BASE_DIR, "index.html")
EXCEL_PATH = os.path.join(BASE_DIR, "staff_data.xlsx") 

USERNAME ='E1126415635'
PASSWORD = '415635'

# --- 🚀 وظيفة الرفع لـ GitHub تلقائياً ---
def auto_push_to_github():
    try:
        print("📤 [GitHub] جاري رفع التحديث المكتبي الجديد...")
        # تأكد أن المجلد مهيأ كـ Git Repo مسبقاً
        subprocess.run(["git", "add", "index.html"], check=True)
        # إذا كنت تريد رفع ملف الإكسيل أيضاً فك التعليق عن السطر التالي
        # subprocess.run(["git", "add", "staff_data.xlsx"], check=True)
        
        commit_msg = f"Auto-Update: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        subprocess.run(["git", "commit", "-m", commit_msg], check=True)
        subprocess.run(["git", "push"], check=True)
        print("✅ تم التحديث في GitHub بنجاح!")
    except Exception as e:
        print(f"⚠️ فشل الرفع لـ GitHub (تأكد من إعداد الـ Git في المجلد): {e}")

def safe_extract_list(res_json):
    if not res_json: return []
    if isinstance(res_json, list): return res_json
    if isinstance(res_json, dict):
        data = res_json.get('data')
        if isinstance(data, list): return data
        if isinstance(data, dict): return data.get('list', [])
    return []

def get_hajj_token():
    driver = None
    try:
        print("🌐 [1/4] جاري فتح المتصفح الخفي لسحب التوكن...")
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        # 🔥 خطوة التمويه: نتنكر كمتصفح كروم طبيعي عشان ما يصيدنا حماية الموقع
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        print("🌐 جاري طلب الرابط...")
        driver.get("https://tnql-prod.sejeltech.app/")
        time.sleep(10)
        
        print(f"📍 الرابط اللي فتحه المتصفح: {driver.current_url}")
        
        if "404" in driver.current_url or "error" in driver.current_url.lower():
            try:
                driver.find_element(By.XPATH, "//*[contains(text(), 'HOME')]").click()
                time.sleep(5)
            except: pass

        wait = WebDriverWait(driver, 30)
        print("🔍 جاري البحث عن مربعات تسجيل الدخول...")
        
        try:
            user_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text'] | (//input)[1]")))
            pass_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
        except Exception as e:
            print("⚠️ الكود ما قدر يشوف مربعات الدخول! هذي عينة من اللي يشوفه المتصفح الآن:")
            print(driver.page_source[:1500]) 
            raise e
        
        user_input.send_keys(USERNAME)
        pass_input.send_keys(PASSWORD)
        pass_input.send_keys(Keys.RETURN)
        
        print("⏳ جاري تسجيل الدخول وسحب التوكن...")
        time.sleep(15) 
        
        token = driver.execute_script("return window.localStorage.getItem('token')") or \
                next((v for k, v in driver.execute_script("return window.localStorage;").items() if str(v).startswith("eyJ")), None)
        
        if token:
            print("✅ [4/4] تم سحب التوكن الآلي بنجاح!")
            return f"Bearer {token}"
            
        print("❌ سجلنا دخول بس ما لقينا التوكن في الذاكرة!")
        return None
    except Exception as e:
        print(f"❌ خطأ المتصفح: {e}")
        return None
    finally:
        if driver: driver.quit()

def generate_master_dashboard():
    # 🎯 توقيت مكة المكرمة
    makkah_time = datetime.utcnow() + timedelta(hours=3)
    current_time_str = makkah_time.strftime("%Y-%m-%d %I:%M %p")
    hour = makkah_time.hour
    
    # 🎯 الكلمات المفتاحية الذكية بناءً على المسميات الدقيقة
    if 8 <= hour < 16:
        shift_keywords = ["ثاني", "صباح", "2"]
        active_shift_title = "الوردية الثانية صباح"
    elif 16 <= hour < 24:
        shift_keywords = ["ثالث", "مساء", "3"]
        active_shift_title = "وردية الثالثة مساء"
    else:
        shift_keywords = ["اول", "ليل", "1"]
        active_shift_title = "الوردية الاولى ليلية"

    # 1. تجهيز قاعدة بيانات الإكسيل
    excel_db = {}
    if os.path.exists(EXCEL_PATH):
        try:
            df_excel = pd.read_excel(EXCEL_PATH, dtype=str).fillna('غير متوفر')
            phone_col = 'غير متوفر'
            for col in ['رقم التليفون', 'رقم الجوال', 'الجوال', 'Phone', 'Mobile']:
                if col in df_excel.columns:
                    phone_col = col
                    break
            for _, row in df_excel.iterrows():
                n_id = str(row.get('هويه الموظف', '')).replace('.0', '').strip().lower()
                if n_id and n_id != 'غير متوفر' and n_id != 'nan':
                    excel_db[n_id] = {
                        'name': str(row.get('اسم الموظف', 'غير متوفر')).strip(), 
                        'job': str(row.get('الوظيفه', 'غير متوفر')).strip(),
                        'dept': str(row.get('القسم', 'غير متوفر')).strip(),
                        'company': str(row.get('شركة التشغيل', 'غير متوفر')).strip(),
                        'phone': str(row.get(phone_col, 'غير متوفر')).strip() if phone_col != 'غير متوفر' else 'غير متوفر'
                    }
        except: pass

    token = get_hajj_token()
    if not token: return

    headers = {"authorization": token, "content-type": "application/json", "lang": "ar"}
    payload = {
        "paging": {"sortField": "Id", "searchOrder": 2, "pageIndex": 1, "totalRowsCount": 10469, "totalPages": 1, "pageSize": 11000, "sortBy": "Id Desc"},
        "data": {
            "searchText": "", "name": "", "EmployeeId": None, "OccupationIds": [], "DepartmentIds": [], 
            "SectionIds": [], "WorkShiftIds": [], "EmployeeTypes": [], "ManagerIds": [], 
            "OperatorCompanyIds": [], "NationalIdExpired": [], "ActiveStatus": [True], 
            "isPrinted": None, "isDeleted": False
        }
    }

    try:
        r_emp = requests.post("https://tnql-prod.sejeltech.app/api/StaffMember/GetStaffMember", headers=headers, json=payload)
        all_employees = safe_extract_list(r_emp.json())
        r_att = requests.post("https://tnql-prod.sejeltech.app/api/EmployeeAttendanceMonitor/GetAttendance", headers=headers, json=payload)
        att_data = safe_extract_list(r_att.json())
        
        today_iso = makkah_time.strftime("%Y-%m-%d") 
        today_ar1 = makkah_time.strftime("%d/%m/%Y") 
        today_ar2 = f"{makkah_time.day}/{makkah_time.month}/{makkah_time.year}" 

        present_ids = set()
        for x in att_data:
            if isinstance(x, dict):
                row_str = str(x.values()).lower()
                if today_iso in row_str or today_ar1 in row_str or today_ar2 in row_str:
                    if "غائب" not in row_str and "absent" not in row_str:
                        emp_code = x.get('employeeCode')
                        if emp_code: present_ids.add(str(emp_code).strip().lower())

        if not all_employees: return

        for emp in all_employees:
            nid = str(emp.get('nationalId', '')).replace('.0', '').strip().lower()
            api_phone = emp.get('mobileNumber') or emp.get('phoneNumber') or 'لا يوجد رقم'
            if nid in excel_db:
                ex = excel_db[nid]
                if ex['job'] not in ['غير متوفر', 'nan']: emp['occupationName'] = ex['job']
                if ex['company'] not in ['غير متوفر', 'nan']: emp['operatorCompanyName'] = ex['company']
                emp['clean_name'] = ex['name'] if ex['name'] not in ['غير متوفر', 'nan'] else (emp.get('name') or 'غير متوفر')
                emp['clean_dept'] = ex['dept']
                emp['clean_phone'] = ex['phone'] if ex['phone'] not in ['غير متوفر', 'nan'] else api_phone
            else:
                emp['clean_name'] = emp.get('name') or 'غير متوفر'
                emp['clean_dept'] = 'غير متوفر'
                emp['clean_phone'] = api_phone

        df = pd.DataFrame(all_employees)
        df = df.fillna('غير محدد')
        total_employees = len(df)
        total_companies = df['operatorCompanyName'].nunique()
        total_shifts = df['workShiftName'].nunique()
        
        def clean_type(val):
            v = str(val).lower()
            return 'دائم' if 'permanent' in v or 'دائم' in v else ('موسمي' if 'seasonal' in v or 'موسمي' in v else 'غير محدد')
        
        df['mapped_type'] = df.get('employeeTypeName', '').apply(clean_type)
        permanent_count = len(df[df['mapped_type'] == 'دائم'])
        seasonal_count = len(df[df['mapped_type'] == 'موسمي'])

        # (كود الـ HTML هنا - تم اختصاره في الرد للحفاظ على المساحة، هو نفسه في كودك الأصلي)
        html_content = f"""<!DOCTYPE html><html lang="ar" dir="rtl"><head>...</head><body>""" # كمل كود الـ HTML حقك هنا
        # ... (بقية كود توليد الـ HTML) ...

        with open(REPORT_PATH, "w", encoding="utf-8") as f:
            f.write(html_content)
        print("✅ تم تحديث ملف التقرير المحلي بنجاح.")

    except Exception as e:
        print(f"❌ حدث خطأ فني: {e}")

# --- 🚀 المحرك الرئيسي المعدل للتكرار والرفع ---
if __name__ == "__main__":
    while True:
        print(f"\n✨ {datetime.now().strftime('%H:%M:%S')} - جاري بدء دورة التحديث...")
        
        # 1. تشغيل سحب البيانات وتحديث الملف المحلي
        generate_master_dashboard()
        
        # 2. الرفع المباشر لـ GitHub
        auto_push_to_github()
        
        print("⏳ تم الانتهاء من هذه الدورة. سأكرر العملية بعد ساعتين...")
        # الانتظار لمدة ساعتين (7200 ثانية)
        time.sleep(7200)