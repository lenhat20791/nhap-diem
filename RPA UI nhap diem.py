# RPA_UI_Nhap_Diem.py
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import tkinter as tk
from tkinter import messagebox

# --- CẤU HÌNH CỐ ĐỊNH (Không cần chỉnh sửa) ---
DRIVER_PATH = 'C:/RPA nhap diem/chromedriver/chromedriver.exe'
BROWSER_PATH = 'C:/RPA nhap diem/chrome-win64/chrome.exe' 
DATA_FILE = 'C:/RPA nhap diem/diem_hoc_sinh.xlsx'

# Chỉ số vị trí cột điểm trong bảng HTML (đã phân tích trước đó)
COT_DIEM = {
    'Cot_3': 9,  # Cột 3 (Dưới DDGtx) tương ứng với td[9]
 
}
# Lưu ý: Cần đảm bảo file Excel có cột tên là 'Ho_Ten', 'Cot_1', 'Cot_3', 'Cot_4'

# ----------------------------------------------------------------------
# HÀM CHÍNH: THỰC THI RPA 
# ----------------------------------------------------------------------
def run_rpa_process(web_url):
    """Thực hiện quy trình đọc Excel và nhập liệu tự động."""
    
    # Kiểm tra URL trống
    if not web_url.startswith("http"):
        messagebox.showerror("LỖI", "Vui lòng nhập đường link (URL) hợp lệ.")
        return

    # 1. ĐỌC DỮ LIỆU TỪ EXCEL
    try:
        df = pd.read_excel(DATA_FILE)
        df['Cot_1'] = df['Cot_1'].astype(str)
        df['Cot_3'] = df['Cot_3'].astype(str)
        df['Cot_4'] = df['Cot_4'].astype(str)
    except FileNotFoundError:
        messagebox.showerror("LỖI", f"Không tìm thấy file dữ liệu tại {DATA_FILE}")
        return

    # 2. KHỞI TẠO TRÌNH DUYỆT
    driver = None
    try:
        options = Options()
        options.binary_location = BROWSER_PATH 
        service = Service(executable_path=DRIVER_PATH)
        
        # Thêm tùy chọn để trình duyệt không đóng ngay cả khi có lỗi
        options.add_experimental_option("detach", True) 
        
        driver = webdriver.Chrome(service=service, options=options) 
        driver.get(web_url)
        # THÊM BƯỚC DỪNG VÀ CHỜ (Sửa 2 dòng này)
        print("\n>>> CHỜ BẠN THAO TÁC: Vui lòng ĐĂNG NHẬP, chọn Khối/Lớp/Môn và điều hướng đến TRANG NHẬP ĐIỂM.")
        input(">>> Sau khi trang nhập điểm tải xong, bấm ENTER trong cửa sổ CMD để bắt đầu nhập điểm...")
        wait = WebDriverWait(driver, 20) # Tăng thời gian chờ tải trang
        
        # 3. TỰ ĐỘNG NHẬP ĐIỂM
        success_count = 0
        failure_list = []
        
        for index, row in df.iterrows():
            ten_hoc_sinh = row['Ho_Ten'].strip() 
            ROW_XPATH = f"//tr[contains(., '{ten_hoc_sinh}')]"
            
            try:
                # Chờ dòng học sinh xuất hiện 
                row_element = wait.until(EC.presence_of_element_located((By.XPATH, ROW_XPATH)))
                
                for cot_ten, cot_index in COT_DIEM.items():
                    diem_can_nhap = row[cot_ten]
                    INPUT_XPATH = f"{ROW_XPATH}/td[{cot_index}]/input"
                    
                    input_field = row_element.find_element(By.XPATH, INPUT_XPATH)
                    input_field.clear() 
                    input_field.send_keys(diem_can_nhap) 
                
                success_count += 1
                
            except Exception as e: # Bắt lỗi và đặt tên là 'e'
                failure_list.append(ten_hoc_sinh)
                # THÊM DÒNG NÀY ĐỂ BÁO CÁO LỖI NGAY LẬP TỨC
                print(f"LỖI TÌM KIẾM HỌC SINH '{ten_hoc_sinh}': {e}")
                # Dừng lại nếu lỗi xảy ra với học sinh đầu tiên để kiểm tra XPath
                if index == 0:
                    break

        # 4. HIỂN THỊ KẾT QUẢ VÀ CHỜ KIỂM TRA
        result_message = f"✅ Hoàn tất nhập điểm tự động!\n\n"
        result_message += f" - Số học sinh đã nhập thành công: {success_count}/{len(df)}\n"
        
        if failure_list:
            result_message += f" - ❌ LỖI KHÔNG TÌM THẤY: {len(failure_list)} học sinh. Vui lòng nhập thủ công:\n"
            result_message += ", ".join(failure_list[:5]) + "..."
        else:
             result_message += f" - Tất cả điểm đã được nhập thành công."

        messagebox.showinfo("KẾT QUẢ RPA", result_message)
        
    except Exception as e:
        messagebox.showerror("LỖI KHỞI TẠO CHUNG", f"Không thể chạy Selenium. Kiểm tra lại: {e}")
        
    finally:
        # Giữ trình duyệt mở, người dùng tự đóng sau khi kiểm tra và lưu
        if driver:
             # Tuy nhiên, chúng ta vẫn cần phải loại bỏ các kết nối service
             pass


# ----------------------------------------------------------------------
# XÂY DỰNG GIAO DIỆN NGƯỜI DÙNG (UI)
# ----------------------------------------------------------------------
def create_ui():
    root = tk.Tk()
    root.title("Công Cụ Nhập Điểm Tự Động (RPA)")
    
    # Hàm xử lý khi nút bấm được nhấn
    def on_start_click():
        web_url = url_entry.get()
        # Chạy quy trình RPA trong một luồng riêng (tính năng nâng cao)
        # Tạm thời chạy trong luồng chính để đơn giản
        root.withdraw() # Ẩn cửa sổ UI khi đang chạy
        run_rpa_process(web_url)
        root.deiconify() # Hiển thị lại cửa sổ UI sau khi xong

    # 1. Tiêu đề
    tk.Label(root, text="Dán đường link (URL) trang nhập điểm vào đây:", font=('Arial', 10, 'bold')).pack(pady=10, padx=10, anchor='w')

    # 2. Ô Nhập liệu URL
    url_entry = tk.Entry(root, width=70, bd=2, relief="groove")
    url_entry.pack(pady=5, padx=10)
    url_entry.insert(0, "https://") # Gợi ý ban đầu

    # 3. Nút Bắt đầu
    start_button = tk.Button(root, text="🚀 BẮT ĐẦU NHẬP ĐIỂM TỰ ĐỘNG", command=on_start_click, 
                             bg='#4CAF50', fg='white', font=('Arial', 12, 'bold'))
    start_button.pack(pady=20, padx=10)
    
    # 4. Ghi chú
    tk.Label(root, text="Lưu ý: File diem_hoc_sinh.xlsx phải nằm trong cùng thư mục.\nSau khi nhập xong, trình duyệt sẽ mở để bạn kiểm tra và tự Cập nhật/Lưu.", 
             fg='gray', font=('Arial', 8)).pack(pady=5, padx=10)

    root.mainloop()

# Chạy giao diện
if __name__ == '__main__':
    create_ui()
