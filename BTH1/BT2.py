import tkinter as tk
from sympy import *
from tkinter import ttk

# Tạo một style mới

def open_limit_window():
    limit_window = tk.Toplevel(app)
    limit_window.title("Tính Giới Hạn")
    limit_label = tk.Label(limit_window, text="Nhập biểu thức cần tính giới hạn:")
    limit_label.pack()
    limit_entry = tk.Entry(limit_window)
    limit_entry.pack()
    limit_value_label = tk.Label(limit_window, text="Nhập giá trị tiến đến:")
    limit_value_label.pack()
    limit_value_entry = tk.Entry(limit_window)
    limit_value_entry.pack()
    calculate_limit_button = tk.Button(limit_window, text="Tính Giới Hạn",
                                       command=lambda: calculate_limit(limit_entry.get(), limit_value_entry.get()))
    clear_button = tk.Button(limit_window, text="Xóa", command=lambda: result_label.config(text=""))
    clear_button.pack()

    close_button = tk.Button(limit_window, text="Thoát", command=limit_window.destroy)
    close_button.pack()
    calculate_limit_button.pack()
def open_derivative_window():
    derivative_window = tk.Toplevel(app)
    derivative_window.title("Tính Đạo Hàm")

    derivative_label = tk.Label(derivative_window, text="Nhập biểu thức cần tính đạo hàm:")
    derivative_label.pack()

    derivative_entry = tk.Entry(derivative_window)
    derivative_entry.pack()
    calculate_derivative_button = tk.Button(derivative_window, text="Tính Đạo Hàm",
                                            command=lambda: calculate_derivative(derivative_entry.get()))
    clear_button = tk.Button(derivative_window, text="Xóa", command=lambda: result_label.config(text=""))
    clear_button.pack()
    close_button = tk.Button(derivative_window, text="Thoát", command=derivative_window.destroy)  # Sửa đoạn này
    close_button.pack()
    calculate_derivative_button.pack()
def open_integral_window():
    integral_window = tk.Toplevel(app)
    integral_window.title("Tính nguyên hàm")
    integral_label = tk.Label(integral_window, text="Nhập biểu thức cần tính nguyên hàm:")
    integral_label.pack()
    integral_entry = tk.Entry(integral_window)
    integral_entry.pack()
    calculate_integral_button = tk.Button(integral_window, text="Tính Nguyên hàm",
                                          command=lambda: calculate_integral(integral_entry.get()))
    clear_button = tk.Button(integral_window, text="Xóa", command=lambda: result_label.config(text=""))
    clear_button.pack()
    close_button = tk.Button(integral_window, text="Thoát", command=integral_window.destroy)  # Sửa đoạn này
    close_button.pack()
    calculate_integral_button.pack()
def open_polynomial_division_window():
    division_window = tk.Toplevel(app)
    division_window.title("Chia Đa Thức")
    dividend_label = tk.Label(division_window, text="Nhập biểu thức cho số tử của đa thức:")
    dividend_label.pack()
    dividend_entry = tk.Entry(division_window)
    dividend_entry.pack()
    divisor_label = tk.Label(division_window, text="Nhập biểu thức cho số mẫu của đa thức:")
    divisor_label.pack()
    divisor_entry = tk.Entry(division_window)
    divisor_entry.pack()
    calculate_polynomial_division_button = tk.Button(division_window, text="Chia Đa Thức",
                                                     command=lambda: calculate_polynomial_division(dividend_entry.get(),
                                                                                                   divisor_entry.get()))
    clear_button = tk.Button(division_window, text="Xóa", command=lambda: result_label.config(text=""))
    clear_button.pack()

    close_button = tk.Button(division_window, text="Thoát", command=division_window.destroy)
    close_button.pack()
    calculate_polynomial_division_button.pack()
def calculate_limit(expression, limit_value):
    if not expression or not limit_value:
        result_label.config(text="Vui lòng nhập biểu thức và giá trị tiến đến trước khi tính toán.")
        return
    x = Symbol('x')
    result = limit(sympify(expression), x, float(limit_value))
    result_label.config(text=f"Giới hạn của {expression} khi x tiến đến {limit_value} là: {result}")

def calculate_derivative(expression,k):
    if not expression or not k:
        result_label.config(text="Vui lòng nhập biểu thức và giá trị tiến đến trước khi tính toán.")
        return
    x = Symbol('x')
    result = diff(sympify(expression), x)
    result_label.config(text=f"Đạo hàm của {expression} là: {result}")
def calculate_integral(expression,l):
    if not expression or not l:
        result_label.config(text="Vui lòng nhập biểu thức và giá trị tiến đến trước khi tính toán.")
        return
    x = Symbol('x')
    result = integrate(sympify(expression), x)
    result_label.config(text=f"Nguyên hàm của {expression} là: {result}")
def calculate_polynomial_division(dividend, divisor):
    if not dividend or not divisor:
        result_label.config(text="Vui lòng nhập cả số tử và số mẫu đa thức trước khi tính toán.")
        return
    x = Symbol('x')
    quotient, remainder = div(sympify(dividend), sympify(divisor))
    result_label.config(text=f"Kết quả chia đa thức {dividend} cho {divisor} là:\nKết quả: {quotient}\nDư: {remainder}")
app = tk.Tk()
app.title("Ứng dụng Giải Tích")
app.geometry("500x200")

result_label = ttk.Label(app, text="", wraplength=400)
result_label.pack()

# Frame chứa các nút
btn_frame = tk.Frame()
btn_frame.pack()

# Tạo các nút
limit_button = tk.Button(btn_frame, text="Tính Giới Hạn", width=15, height=2, command=open_limit_window)
derivative_button = tk.Button(btn_frame, text="Tính Đạo Hàm", width=15, height=2, command=open_derivative_window)
integral_button = tk.Button(btn_frame, text="Tính Nguyên Hàm", width=15, height=2, command=open_integral_window)
division_button = tk.Button(btn_frame, text="Chia Đa Thức", width=15, height=2, command=open_polynomial_division_window)

# Bố trí theo grid
limit_button.grid(row=0, column=0, padx=10, pady=10)
derivative_button.grid(row=0, column=1, padx=10, pady=10)
integral_button.grid(row=1, column=0, padx=10, pady=10)
division_button.grid(row=1, column=1, padx=10, pady=10)

app.mainloop()