import time
import json
import os
import logging
import traceback
import tkinter as tk
import customtkinter as ctk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

# =========================================================
# 1. KHAI BÁO XPATH CHÍNH XÁC NHƯ BẠN CUNG CẤP
# =========================================================
XPATH_DROPDOWN_INPUT = "//input[@value='Mầm non' and @type='text']"
XPATH_OPTION_THCS = "//li[text()='Trung học cơ sở']"

# Phường/Xã
XPATH_DROPDOWN_PHUONGXA_INPUT = "//input[@id='ctl00_ContentPlaceHolder1_rcbPhongGD_Input']"
XPATH_OPTION_HANHTHONG = "//li[contains(text(), 'Phường Hạnh Thông')]"

# Chọn trường
XPATH_INPUT_TRUONGHOC = "//input[@id='ctl00_ContentPlaceHolder1_cbTruongInput']"
XPATH_ARROW_TRUONGHOC = "//a[@id='ctl00_ContentPlaceHolder1_cbTruong_Arrow']"

# =========================================================
# 2. CẤU HÌNH LOGGING FORCE FLUSH
# =========================================================
log_file = "debug_log.txt"
if os.path.exists(log_file): os.remove(log_file)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler(log_file, encoding='utf-8'), logging.StreamHandler()]
)

def log_info(msg):
    logging.info(msg)
    for handler in logging.root.handlers: handler.flush()

# =========================================================
# 3. MODULE UI CUSTOMTKINTER (GHI NHẬN & LƯU TRỮ)
# =========================================================
CONFIG_FILE = "login_config.json"

def show_ctk_ui():
    log_info("--- Mở UI cấu hình ---")
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    config = {"username": "", "password": "", "remember": False}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f: config = json.load(f)
        except: pass

    root = ctk.CTk()
    root.title("RPA Đăng Nhập")
    root.geometry("350x300")
    root.attributes("-topmost", True)

    ctk.CTkLabel(root, text="Tài khoản (CCCD):", font=ctk.CTkFont(size=12, weight="bold")).pack(pady=(20,5))
    ent_user = ctk.CTkEntry(root, width=250)
    ent_user.insert(0, config.get("username", ""))
    ent_user.pack()

    ctk.CTkLabel(root, text="Mật khẩu:", font=ctk.CTkFont(size=12, weight="bold")).pack(pady=5)
    ent_pass = ctk.CTkEntry(root, width=250, show="*")
    ent_pass.insert(0, config.get("password", ""))
    ent_pass.pack()

    var_rem = tk.BooleanVar(value=config.get("remember", False))
    ctk.CTkCheckBox(root, text="Ghi nhớ tài khoản", variable=var_rem).pack(pady=10)

    def save_and_run():
        data = {
            "username": ent_user.get() if var_rem.get() else "",
            "password": ent_pass.get() if var_rem.get() else "",
            "remember": var_rem.get()
        }
        with open(CONFIG_FILE, "w") as f: json.dump(data, f)
        log_info("--- Đã lưu tài khoản. Đóng UI để chạy Bot ---")
        root.destroy()

    ctk.CTkButton(root, text="BẮT ĐẦU CHẠY", command=save_and_run).pack(pady=15)
    root.mainloop()

# =========================================================
# 4. MODULE RPA (DỰA TRÊN 100% CHI TIẾT GỐC)
# =========================================================
def run_bot():
    log_info("--- KHỞI ĐỘNG CHROME ---")
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)
    
    try:
        driver.get("https://hcm.quanlytruonghoc.edu.vn")
        
        # Bước 2: Chọn Cấp trường
        log_info("2. Chọn Cấp trường (Sử dụng XPATH giá trị Mầm non)...")
        dropdown_cap = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_DROPDOWN_INPUT)))
        dropdown_cap.click()
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_OPTION_THCS))).click()
        time.sleep(3)

        # Bước 3: Thao tác Phường/Xã (Sử dụng ID bạn cung cấp)
        log_info("3. Nhập Phường/Xã...")
        input_px = wait.until(EC.visibility_of_element_located((By.XPATH, XPATH_DROPDOWN_PHUONGXA_INPUT)))
        input_px.click()
        input_px.send_keys("Thông")
        time.sleep(2)
        
        input_px.send_keys(Keys.ARROW_DOWN)
        time.sleep(1)
        input_px.send_keys(Keys.TAB)
        time.sleep(2)
        
        # TAB lần 2 sang Trường học
        ActionChains(driver).send_keys(Keys.TAB).perform()
        time.sleep(1)

        # Bước 4: Thao tác Trường học (Nhấn xuống 4 lần)
        log_info("4. Nhấn xuống 4 lần chọn Trường...")
        for i in range(1, 5):
            ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
            time.sleep(1)
        
        ActionChains(driver).send_keys(Keys.ENTER).perform()
        log_info("--- ĐÃ CHUYỂN SANG TRANG SSO ---")
        time.sleep(5)
        # BƯỚC 5: TỰ ĐỘNG ĐIỀN TÀI KHOẢN TỪ FILE JSON
        log_info("5. Đang đọc cấu hình và điền tài khoản vào SSO...")
        
        # Đọc dữ liệu từ file JSON đã lưu từ UI
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                config_data = json.load(f)
                u_name = config_data.get("username", "")
                p_word = config_data.get("password", "")
        else:
            log_info("!!! Không tìm thấy file cấu hình login_config.json")
            return

        # Đợi các ô input xuất hiện (Sử dụng ID từ hình Inspect bạn gửi)
        user_field = wait.until(EC.visibility_of_element_located((By.ID, "UserName")))
        pass_field = driver.find_element(By.ID, "Password") # Thường Password sẽ đi cùng UserName
        
        # Điền thông tin
        user_field.send_keys(u_name)
        log_info(f"   -> Đã điền tài khoản: {u_name}")
        time.sleep(0.5)
        pass_field.send_keys(p_word)
        log_info("   -> Đã điền mật khẩu.")
        
        # Click nút Đăng nhập (Dùng XPATH dựa trên cấu trúc container-login100)
        btn_submit = driver.find_element(By.XPATH, "//button[contains(@class, 'login100-form-btn')]")
        btn_submit.click()
        
        log_info("--- ĐÃ NHẤN ĐĂNG NHẬP THÀNH CÔNG ---")
        time.sleep(5)

    except Exception:
        log_info(f"!!! LỖI TẠI BƯỚC 5:\n{traceback.format_exc()}")
    finally:
        input("Nhấn ENTER để đóng trình duyệt...")
    
    def thuc_hien_nhap_diem(driver, wait):
    log_info("--- BẮT ĐẦU QUY TRÌNH GÕ ĐIỂM TỪ EXCEL ---")
    try:
        # Đọc file Excel
        df = pd.read_excel("danh_sach_nhap_diem.xlsx")
        
        # 1. Click vào ô input của học sinh đầu tiên (Dựa trên ID Inspect bạn gửi)
        # Ô đầu tiên thường là dòng ctl04
        first_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@id, 'txtDIEM_HK')]")))
        first_input.click()
        time.sleep(1)
        
        # 2. Vòng lặp gõ điểm
        for index, row in df.iterrows():
            diem_so = str(row['Diem']).replace('.', ',') # Đổi sang dấu phẩy nếu web yêu cầu
            ten_hs = row['HoTen']
            
            # Gõ điểm
            ActionChains(driver).send_keys(diem_so).perform()
            log_info(f"   [OK] Đã gõ {diem_so} cho {ten_hs}")
            time.sleep(1) # Nghỉ 1s sau khi gõ
            
            # Nhấn phím mũi tên xuống
            if index < len(df) - 1:
                ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
                time.sleep(1) # Nghỉ 1s sau khi xuống dòng
                
        log_info("--- HOÀN TẤT NHẬP ĐIỂM ---")
    except Exception:
        log_info(f"!!! LỖI NHẬP ĐIỂM:\n{traceback.format_exc()}")
if __name__ == "__main__":
    show_ctk_ui()
    run_bot()
