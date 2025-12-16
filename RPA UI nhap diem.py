# Tool_nhap_diem.py
import time
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging #

# --- CẤU HÌNH GHI LOG ---
LOG_FILE = 'rpa_log.txt'
logging.basicConfig(
    filename=LOG_FILE, 
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s'
)
# --- HẾT CẤU HÌNH LOG ---

# --- CẤU HÌNH CỐ ĐỊNH ---
DRIVER_PATH = 'C:/RPA nhap diem/chromedriver-win64/chromedriver.exe'
BROWSER_PATH = 'C:/RPA nhap diem/chrome-win64/chrome.exe'
TARGET_URL = 'https://hcm.quanlytruonghoc.edu.vn'

# XPath của phần tử dropdown đầu tiên (Mầm non/Tiểu học/Trung học...)
XPATH_DROPDOWN_INPUT = "//input[@name='ctl00$ContentPlaceHolder1$cboCapTruong_Input']"
# XPath của option "Trung học cơ sở" khi danh sách xổ xuống
XPATH_OPTION_THCS = "//div[@class='rcbList']//li[text()='Trung học cơ sở']"


# ----------------------------------------------------------------------
# HÀM CHÍNH: THỰC THI RPA 
# ----------------------------------------------------------------------
def run_rpa_process():
    """Thực hiện quy trình mở trình duyệt và chọn cấp trường."""
    
    driver = None
    try:
        # 1. KHỞI TẠO TRÌNH DUYỆT
        options = Options()
        options.binary_location = BROWSER_PATH 
        service = Service(executable_path=DRIVER_PATH)
        options.add_experimental_option("detach", True) 
        
        driver = webdriver.Chrome(service=service, options=options) 
        # THÊM DÒNG NÀY: Maximize cửa sổ trình duyệt
        driver.maximize_window()
        driver.get(TARGET_URL)
        
        wait = WebDriverWait(driver, 10)
        
        # 2. CHỌN "TRUNG HỌC CƠ SỞ"
        
        # 2.1. Chờ input dropdown xuất hiện và click để mở danh sách
        dropdown_input = wait.until(
            EC.element_to_be_clickable((By.XPATH, XPATH_DROPDOWN_INPUT))
        )
        dropdown_input.click()
        
        # 2.2. Chờ option "Trung học cơ sở" xuất hiện và click
        option_thcs = wait.until(
            EC.presence_of_element_located((By.XPATH, XPATH_OPTION_THCS))
        )
        option_thcs.click()
        
        # 3. THÔNG BÁO THÀNH CÔNG VÀ CHỜ HƯỚNG DẪN TIẾP THEO
        messagebox.showinfo("THÀNH CÔNG", "✅ Đã chọn 'Trung học cơ sở' thành công. Bot đang chờ hướng dẫn tiếp theo.")
        
    except Exception as e:
        messagebox.showerror("LỖI TỰ ĐỘNG HÓA", f"Không thể hoàn tất thao tác. Kiểm tra lại cấu hình hoặc XPath. Chi tiết lỗi: {e}")
        
    finally:
        # Giữ trình duyệt mở
        pass


# ----------------------------------------------------------------------
# XÂY DỰNG GIAO DIỆN NGƯỜI DÙNG (UI)
# ----------------------------------------------------------------------
def create_ui():
    root = tk.Tk()
    root.title("Công Cụ Tự Động Đăng Nhập (Tool nhập điểm)")
    
    def on_start_click():
        root.withdraw() # Ẩn cửa sổ UI khi đang chạy
        run_rpa_process()
        root.deiconify() # Hiển thị lại cửa sổ UI sau khi xong

    # 1. Tiêu đề
    tk.Label(root, text="Tự động truy cập và chọn 'Trung học cơ sở':", font=('Arial', 10, 'bold')).pack(pady=5, padx=10, anchor='w')

    # 2. Ô Hiển thị URL
    url_entry = tk.Entry(root, width=70, bd=2, relief="groove")
    url_entry.pack(pady=5, padx=10)
    url_entry.insert(0, TARGET_URL)
    url_entry.config(state='readonly') # Không cho người dùng sửa

    # 3. Nút Bắt đầu
    start_button = tk.Button(root, text="🚀 BẮT ĐẦU TỰ ĐỘNG HÓA", command=on_start_click, 
                             bg='#4CAF50', fg='white', font=('Arial', 12, 'bold'))
    start_button.pack(pady=20, padx=10)
    
    # 4. Ghi chú
    tk.Label(root, text="Bot sẽ tự động mở Chrome và chọn mục đầu tiên.", 
             fg='gray', font=('Arial', 8)).pack(pady=5, padx=10)

    root.mainloop()

# Chạy giao diện
if __name__ == '__main__':
    create_ui()
