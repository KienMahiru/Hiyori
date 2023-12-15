import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import cv2
import matplotlib.pyplot as plt
import numpy as np
# Hàm để cập nhật và hiển thị ảnh sau khi zoom và xoay

# Hàm xử lý sự kiện khi nút "Chọn ảnh" được nhấn
def select_image():
    global original_img, current_scale
    file_path = filedialog.askopenfilename()
    if file_path:
        original_img = Image.open(file_path)
        original_img = cv2.cvtColor(np.array(original_img), cv2.COLOR_RGB2BGR)
        # Chuyển đổi BGR -> RGB trước khi hiển thị
        original_img = cv2.cvtColor(original_img, cv2.COLOR_BGR2RGB)

        update_image()
        current_scale = 1.0
        update_image()
def update_image():
    if original_img is not None:
        # Áp dụng zoom và xoay vào ảnh gốc
        pil_img = Image.fromarray(original_img)  # Convert NumPy array to PIL Image
        zoomed_img = pil_img.resize((int(pil_img.width * current_scale), int(pil_img.height * current_scale)), Image.LANCZOS)

        img = ImageTk.PhotoImage(zoomed_img)
        img_label.config(image=img)
        img_label.image = img

        # Cập nhật tỷ lệ zoom
        zoom_label.config(text=f"Zoom: {current_scale:.1f}")


# Hàm xử lý sự kiện khi thanh trượt thay đổi giá trị
def on_scale_change(event):
    global current_scale
    current_scale = scale.get()
    update_image()



# Hàm xử lý sự kiện khi nút "Xoay trái" được nhấn
# Hàm xoay ảnh 90 độ sang trái
# Hàm xoay 90 độ sang trái
def rotate_left():
    global original_img

    h, w = original_img.shape[:2]
    center = (w / 2, h / 2)
    M = cv2.getRotationMatrix2D(center, 90, 1.0)
    original_img = cv2.warpAffine(original_img, M, (w, h))

    update_image()

# Hàm xoay ảnh 90 độ sang phải
# Hàm xoay 90 độ sang phải
def rotate_right():
  global original_img

  h, w = original_img.shape[:2]
  center = (w/2, h/2)

  # Xoay -90 độ => tương đương xoay 90 độ sang phải
  M = cv2.getRotationMatrix2D(center, -90, 1.0)

  original_img = cv2.warpAffine(original_img, M, (w,h))

  update_image()
# Hàm xử lý sự kiện khi nhấp chuột trên ảnh để zoom in vào vị trí tùy chọn
def zoom_in(event):
    global current_scale
    x = event.x
    y = event.y
    # Lấy kích thước từ shape
    h, w = original_img.shape[:2]

    new_width = int(w * current_scale * zoom_factor)
    new_height = int(h * current_scale * zoom_factor)
    # Tính toán lại vị trí x và y tương đối trên ảnh sau khi zoom
    x_ratio = x / img_label.winfo_width()
    y_ratio = y / img_label.winfo_height()
    new_x = int(x_ratio * new_width)
    new_y = int(y_ratio * new_height)

    # Tạo ảnh mới đã được zoom in
      # Chuyển NumPy array sang PIL Image
    pil_img = Image.fromarray(original_img)

    # Resize PIL Image
    zoomed_img = pil_img.resize((new_width, new_height), Image.LANCZOS)


    # Cắt ảnh theo vị trí mới
    zoomed_img = zoomed_img.crop((new_x - img_label.winfo_width() // 2, new_y - img_label.winfo_height() // 2,
                                  new_x + img_label.winfo_width() // 2, new_y + img_label.winfo_height() // 2))

    img = ImageTk.PhotoImage(zoomed_img)
    img_label.config(image=img)
    img_label.image = img

    # Cập nhật tỷ lệ zoom
    current_scale *= zoom_factor
    zoom_label.config(text=f"Zoom: {current_scale:.1f}")

def detect_edges():
    if original_img is not None:
        gray_img = cv2.cvtColor(original_img, cv2.COLOR_BGR2GRAY)
        edges = cv2.Canny(gray_img, 90, 100)  # Thay đổi các giá trị ngưỡng tùy theo ảnh

        # Hiển thị ảnh được tách biên bằng Matplotlib
        plt.figure()
        plt.imshow(edges, cmap='gray')
        plt.title("Ảnh Tách Biên")
        plt.show()

def loc_thong_thap_5x5():
    global original_img

    if original_img is not None:
        img = original_img.copy()  # Tạo một bản sao của ảnh gốc để thực hiện bộ lọc

        rows, cols = img.shape[:2]
        kernel_identity = np.array([[0, 0, 0], [0, 1, 0], [0, 0, 0]])

        kernel_5x5 = np.ones((5, 5), np.float32) / 25.0

        output = cv2.filter2D(img, -1, kernel_identity)


        output = cv2.filter2D(img, -1, kernel_5x5)
        plt.figure()
        plt.title("Ảnh Làm Mịn 5x5")
        plt.imshow(output, cmap='gray')
        plt.show()

def loc_thong_thap_3x3():
    global original_img
    if original_img is not None:
        img = original_img.copy()  # Tạo một bản sao của ảnh gốc để thực hiện bộ lọc
        rows, cols = img.shape[:2]
        kernel_identity = np.array([[0, 0, 0], [0, 1, 0], [0, 0, 0]])
        kernel_3x3 = np.ones((3, 3), np.float32) / 9.0
        output = cv2.filter2D(img, -1, kernel_identity)
        output = cv2.filter2D(img, -1, kernel_3x3)
        plt.figure()
        plt.title("Ảnh Làm Mịn 3x3")
        plt.imshow(output, cmap='gray')
        plt.show()
# Tạo cửa sổ
window = tk.Tk()
window.title("Sử dụng bộ lọc")
window.geometry('950x450')

# Khai báo biến toàn cục
current_scale= 1.0
original_img = None
rotation_angle = 0
zoom_factor = 1  # Hệ số zoom tùy chỉnh

# Tạo frame chứa thanh trượt và nút chọn ảnh
control_frame = tk.Frame(window)
control_frame.pack(side="left")

# Tạo thanh trượt
scale = tk.Scale(control_frame, from_=0.1, to=3.0, resolution=0.1, orient=tk.HORIZONTAL, command=on_scale_change)
scale.set(current_scale)
scale.pack()

# Tạo nút "Chọn ảnh" và đặt sự kiện khi nút này được nhấn
select_button = tk.Button(control_frame, text="Chọn ảnh", command=select_image)
select_button.pack()

# Tạo nút "Xoay trái" và đặt sự kiện khi nút này được nhấn
rotate_left_button = tk.Button(control_frame, text="Xoay trái", command=rotate_left)
rotate_left_button.pack()

# Tạo nút "Xoay phải" và đặt sự kiện khi nút này được nhấn
rotate_right_button = tk.Button(control_frame, text="Xoay phải", command=rotate_right)
rotate_right_button.pack()

# Label hiển thị ảnh
img_label = tk.Label(window)
img_label.pack()

edge_detection_button = tk.Button(control_frame, text="Tách biên", command=detect_edges)
edge_detection_button.pack()
# Bộ lọc  thông thấp 3x3
img_label = tk.Label(window)
img_label.pack()
edge_detection_button = tk.Button(control_frame, text="Bộ lọc 3x3", command=loc_thong_thap_3x3)
edge_detection_button.pack()
# Bộ lọc  thông thấp 5x5
img_label = tk.Label(window)
img_label.pack()
edge_detection_button = tk.Button(control_frame, text="Bộ lọc 5x5", command=loc_thong_thap_5x5)
edge_detection_button.pack()
# Label hiển thị tỷ lệ zoom
zoom_label = tk.Label(control_frame, text=f"Zoom: {current_scale:.1f}")
zoom_label.pack()

# Đặt sự kiện nhấp chuột trái để zoom in
img_label.bind("<Button-1>", zoom_in)

window.mainloop()