import cv2
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
def open_image():
    # Mở hộp thoại chọn tệp
    file_path = filedialog.askopenfilename()
    if file_path:
        # Kiểm tra kiểu tệp
        allowed_extensions = ['.jpg', '.jpeg', '.png']
        file_extension = file_path[file_path.rfind('.'):]
        if file_extension.lower() not in allowed_extensions:
            messagebox.showerror("Lỗi", "Chọn sai tệp ảnh! Vui lòng chọn lại")
            return
        # Xóa ảnh cũ nếu có
        global image
        try:
            del gauss_label.image
            del image
        except:
            pass
        # Đọc ảnh
        image = cv2.imread(file_path,cv2.IMREAD_COLOR)
        img = Image.fromarray(cv2.cvtColor(image, cv2.COLOR_BGR2GRAY))
        img = img.resize((400, 400))
        img_tk = ImageTk.PhotoImage(img)
        img_label.configure(image=img_tk)
        img_label.image = img_tk

        gray = img = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        gauss = cv2.GaussianBlur(gray, (0, 0), sigmaX=1.0)
        gauss_img = Image.fromarray(gauss)
        gauss_img = gauss_img.resize((400, 400))
        gauss_tk = ImageTk.PhotoImage(gauss_img)
        gauss_label.configure(image=gauss_tk)
        gauss_label.image = gauss_tk


# Tạo cửa sổ giao diện
window = tk.Tk()
window.title("Làm mịn ảnh")
window.geometry("850x450")
scale_percent =1
# Tạo nút "Open Image"
open_button = tk.Button(window, text="Mở ảnh", command=open_image)
open_button.pack(pady=10)


button_frame = tk.Frame(window)
button_frame.pack(pady=10)

# Tạo hai khung hình để hiển thị ảnh gốc và ảnh làm mịn
img_frame = tk.Frame(window)
img_frame.pack(side=tk.LEFT, padx=10)
gauss_frame = tk.Frame(window)
gauss_frame.pack(side=tk.LEFT, padx=10)

img_label = tk.Label(img_frame, bg="white")
img_label.pack(pady=10)
gauss_label = tk.Label(gauss_frame, bg="white")
gauss_label.pack(pady=10)

# Chạy ứng dụng
window.mainloop()