import time
import traceback
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
# PHẢI CÓ DÒNG NÀY (ĐÃ SỬA TỪ SERVICE THÀNH Service)
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains # Nếu bạn đã thêm
import logging
from webdriver_manager.chrome import ChromeDriverManager # Thư viện mới

# --- CẤU HÌNH GHI LOG ---
LOG_FILE = 'rpa_log.txt'
logging.basicConfig(
    filename=LOG_FILE, 
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s'
)
# --- HẾT CẤU HÌNH LOG ---

# --- CẤU HÌNH CỐ ĐỊNH ---
BROWSER_PATH = 'C:/RPA NHAP DIEM/chrome-win64/chrome.exe'
TARGET_URL = 'https://hcm.quanlytruonghoc.edu.vn'

# XPath của phần tử dropdown đầu tiên (Mầm non/Tiểu học/Trung học...)
XPATH_DROPDOWN_INPUT = "//input[@value='Mầm non' and @type='text']"
# XPATH OPTION (Đã được xác nhận là đúng)
XPATH_OPTION_THCS = "//li[text()='Trung học cơ sở']"
# Cho Phường/Xã
XPATH_DROPDOWN_PHUONGXA_INPUT = "//input[@id='ctl00_ContentPlaceHolder1_rcbPhongGD_Input']"
XPATH_OPTION_HANHTHONG = "//li[text()='Phường Hạnh Thông']"
# SỬA LẠI XPATH TÌM BOX CUỘN BẰNG ID ĐẦY ĐỦ
dropdown_list_box = wait.until(
    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_rcbPhongGD_DropDown"))
)
driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight/2", dropdown_list_box)
# ----------------------------------------------------------------------
# HÀM CHÍNH: THỰC THI RPA 
# ----------------------------------------------------------------------
def run_rpa_process():
    """Thực hiện quy trình mở trình duyệt và chọn cấp trường (có ghi log)."""
    
    driver = None
    logging.info("--- BẮT ĐẦU QUÁ TRÌNH TỰ ĐỘNG HÓA ---")
    try:
        # 1. KHỞI TẠO TRÌNH DUYỆT (TỰ ĐỘNG TÌM DRIVER)
        logging.info("1. Đang khởi tạo trình duyệt Chrome (Sử dụng WebDriverManager).")
        
        # --- THIẾT LẬP OPTIONS ---
        options = Options()
        # Chỉ định đường dẫn của Chrome for Testing (Không thay đổi)
        options.binary_location = BROWSER_PATH 
        options.add_experimental_option("detach", True) 
        
        # --- KHỞI TẠO DRIVER BẰNG ChromeDriverManager ---
        # ChromeDriverManager().install() sẽ tự động tải và cache Driver tương thích
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()), 
            options=options
        ) 
        
        logging.info("   -> DRIVER VÀ BROWSER ĐÃ KHỞI TẠO THÀNH CÔNG.")
        
        # --- MAXIMIZE VÀ TRUY CẬP URL ---
        driver.maximize_window() 
        logging.info("   -> Đã maximize cửa sổ trình duyệt.")
        
        driver.get(TARGET_URL)
        logging.info(f"   -> Đã truy cập URL: {TARGET_URL}")
        
        # *** GIẢI PHÁP CHỜ ĐỢI TẢI JS BẮT BUỘC ***
        logging.info("1a. Đang chờ tải JavaScript và DOM hoàn tất (Tối đa 30s)...")

        # 1. Đợi trạng thái tài liệu chuyển sang 'complete'
        WebDriverWait(driver, 30).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        logging.info("   -> Trạng thái tải trang đã hoàn tất ('complete').")

        # 2. Thêm một chút chờ cứng để UI ổn định (rất cần thiết cho các form phức tạp)
        time.sleep(2) 
        # *****************************************

        
        # *** KHẮC PHỤC LỖI IFRAME ***
        # Dùng WebDriverWait mặc định 7 giây cho lần chờ này
        wait_iframe = WebDriverWait(driver, 7) 
        
        try:
            # 1a. Thử chuyển đổi sang iframe đầu tiên (thường là iframe duy nhất)
            logging.info("1b. Đang thử chuyển đổi sang iframe (Nếu có)...")
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe")))
            logging.info("   -> Đã chuyển đổi thành công sang iframe.")
        except:
            # Bỏ qua nếu không tìm thấy iframe (Không phải mọi trang đều dùng iframe)
            logging.info("   -> Không tìm thấy iframe hoặc lỗi chuyển đổi. Tiếp tục ở khung chính.")
            pass
        # *****************************
        
        # ĐỊNH NGHĨA LẠI ĐỐI TƯỢNG CHỜ CHUNG (ĐÃ GIẢM TỪ 20S CÒN 7S)
        wait = WebDriverWait(driver, 7)
        
        # 2. CHỌN "TRUNG HỌC CƠ SỞ"
        logging.info("2. Đang thực hiện tương tác UI.")
        
        # 2.1. Chờ input dropdown xuất hiện (CHỈ CẦN TỒN TẠI)
        logging.info("2.1. Đang tìm kiếm dropdown bằng Value và click (JavaScript)...")

        try:
            # Chỉ cần chờ phần tử có mặt (presence)
            dropdown_input = wait.until(
                EC.presence_of_element_located((By.XPATH, XPATH_DROPDOWN_INPUT))
            )

            # 1. Thử click mô phỏng chuột (ActionChains)
            ActionChains(driver).move_to_element(dropdown_input).click().perform()
            logging.info("   -> Đã thử click thành công bằng ActionChains.")
            
        except Exception as e:
            # 2. Nếu ActionChains thất bại, thử click bằng JavaScript (ép buộc)
            logging.warning(f"   -> ActionChains thất bại ({e}). Thử click bằng JavaScript...")
            driver.execute_script("arguments[0].click();", dropdown_input)
            logging.info("   -> Đã click thành công bằng JavaScript (ép buộc).")


        time.sleep(1) 
        
        # 2.2. Chờ option "Trung học cơ sở" xuất hiện và click
        logging.info("2.2. Đang tìm kiếm và click vào option 'Trung học cơ sở'...")
        
        # Chờ option THCS xuất hiện (presence_of_element_located)
        option_thcs = wait.until(
            EC.presence_of_element_located((By.XPATH, XPATH_OPTION_THCS))
        )
        
        # Click vào option THCS
        option_thcs.click()
        
        logging.info("   -> THÀNH CÔNG: Đã chọn 'Trung học cơ sở'.")
        
        # THÊM BƯỚC CHỜ ĐỒNG BỘ NỘI DUNG MỚI (Khắc phục lỗi TimeoutException mới)
        logging.info("   -> Chờ 2 giây để danh sách Phường/Xã tải lại...")
        time.sleep(2)
    
        # 3. CHỌN "PHƯỜNG HẠNH THÔNG"
        logging.info("3. Đang thực hiện chọn Phường/Xã.")
        
        # 3.1. Click vào input dropdown Phường/Xã
        logging.info("3.1. Đang tìm kiếm và click vào dropdown chọn Phường/Xã...")
        
        # Chiến lược click mạnh mẽ
        try:
            # Chờ Input Field xuất hiện
            dropdown_input_px = wait.until(
                EC.presence_of_element_located((By.XPATH, XPATH_DROPDOWN_PHUONGXA_INPUT))
            )
            # Thử click mô phỏng chuột (ActionChains)
            ActionChains(driver).move_to_element(dropdown_input_px).click().perform()
            logging.info("   -> Đã thử click thành công bằng ActionChains.")
            
        except Exception as e:
            # Nếu ActionChains thất bại, thử click bằng JavaScript (ép buộc)
            logging.warning(f"   -> ActionChains thất bại. Thử click bằng JavaScript...")
            driver.execute_script("arguments[0].click();", dropdown_input_px)
            logging.info("   -> Đã click thành công bằng JavaScript (ép buộc).")
            
        time.sleep(1) # Chờ 1 giây để danh sách tùy chọn tải
        # *** BƯỚC MỚI: CUỘN CHUỘT TRONG DANH SÁCH DÀI ***
        logging.info("3.1.5. Đang cuộn chuột tìm 'Phường Hạnh Thông'...")
        
        # 1. Tìm phần tử chứa danh sách (Box chứa các li)
        dropdown_list_box = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'rcbList')]"))
        )
        
        # 2. Thực hiện cuộn chuột 5 lần (hoặc 100 pixels)
        # Sử dụng ActionChains để cuộn xuống (move_by_offset) hoặc JS
        # Tôi sẽ dùng JavaScript vì nó ổn định hơn cho scroll
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight/2", dropdown_list_box)
        logging.info("   -> Đã cuộn giữa danh sách (Bằng JavaScript).")
        
        time.sleep(1) # Chờ 1 giây sau khi cuộn để phần tử xuất hiện trong DOM
        
        # 3.2. Chờ option "Phường Hạnh Thông" xuất hiện và click
        logging.info("3.2. Đang tìm kiếm và click vào option 'Phường Hạnh Thông'...")
        
        # Chờ option Phường Hạnh Thông xuất hiện và click (Sử dụng text để tìm)
        option_hanhthong = wait.until(
            EC.presence_of_element_located((By.XPATH, XPATH_OPTION_HANHTHONG))
        )
        option_hanhthong.click()
        
        logging.info("   -> Đã click thành công vào 'Phường Hạnh Thông'.")
        time.sleep(1)

        # 4. CHUẨN BỊ ĐĂNG NHẬP (Phần tiếp theo)
        logging.info("4. Đã chọn Phường/Xã. Chuẩn bị bước Đăng nhập...")
        
    except Exception as e:
        logging.error(f"!!! LỖI QUAN TRỌNG TẠI BƯỚC TỰ ĐỘNG HÓA: {e}", exc_info=True)
        messagebox.showerror("LỖI TỰ ĐỘNG HÓA", f"Không thể hoàn tất thao tác. Vui lòng kiểm tra file LOG ({LOG_FILE}). Chi tiết lỗi: {e}")
        
    finally:
        logging.info("--- KẾT THÚC QUÁ TRÌNH TỰ ĐỘNG HÓA ---")
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
