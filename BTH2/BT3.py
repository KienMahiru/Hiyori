import math
from tkinter import *

class TamGiac:
    def __init__(self, a, b, c):
        self.a = a
        self.b = b
        self.c = c

    def chuvi(self):
        return self.a + self.b + self.c

    def dientich(self):
        s = (self.a + self.b + self.c) / 2
        return math.sqrt(s * (s - self.a) * (s - self.b) * (s - self.c))

    def tinh_goc(self):
        radian_a = math.acos((self.b**2 + self.c**2 - self.a**2) / (2 * self.b * self.c))
        radian_b = math.acos((self.a**2 + self.c**2 - self.b**2) / (2 * self.a * self.c))
        radian_c = math.acos((self.a**2 + self.b**2 - self.c**2) / (2 * self.a * self.b))

        degree_a = math.degrees(radian_a)
        degree_b = math.degrees(radian_b)
        degree_c = math.degrees(radian_c)

        return degree_a, degree_b, degree_c


class TuGiac:
    def __init__(self, a, b, c, d):
        self.a = a
        self.b = b
        self.c = c
        self.d = d

    def chuvi(self):
        return self.a + self.b + self.c + self.d

    def dientich(self):
        s = (self.a + self.b + self.c + self.d) / 2
        return math.sqrt((s - self.a) * (s - self.b) * (s - self.c) * (s - self.d))

    def tinh_goc(self):
        cos_a = (self.b**2 + self.d**2 - self.c**2 - self.a**2) / (2 * self.b * self.d)
        cos_b = (self.a**2 + self.c**2 - self.d**2 - self.b**2) / (2 * self.a * self.c)
        cos_c = (self.a**2 + self.d**2 - self.b**2 - self.c**2) / (2 * self.a * self.d)
        cos_d = (self.b**2 + self.c**2 - self.a**2 - self.d**2) / (2 * self.b * self.c)

        radian_a = math.acos(cos_a)
        radian_b = math.acos(cos_b)
        radian_c = math.acos(cos_c)
        radian_d = math.acos(cos_d)

        degree_a = math.degrees(radian_a)
        degree_b = math.degrees(radian_b)
        degree_c = math.degrees(radian_c)
        degree_d = math.degrees(radian_d)

        return degree_a, degree_b, degree_c, degree_d


def open_tamgiac_window():
    tamgiac_window = Toplevel(root)
    tamgiac_window.title("Tam giác")

    label3 = Label(tamgiac_window, text='Hãy nhập 3 cạnh: ')
    label3.pack()

    canh1 = DoubleVar()
    canh2 = DoubleVar()
    canh3 = DoubleVar()

    E1 = Entry(tamgiac_window, textvariable=canh1)
    E1.pack()
    E2 = Entry(tamgiac_window, textvariable=canh2)
    E2.pack()
    E3 = Entry(tamgiac_window, textvariable=canh3)
    E3.pack()

    b1 = Button(tamgiac_window, text='Thêm dữ liệu', command=lambda: nhap_du_lieu_tamgiac(canh1.get(), canh2.get(), canh3.get()))
    b1.pack()
    b2 = Button(tamgiac_window, text='Tính toán', command=tinh_toan_tamgiac)
    b2.pack()

    global label4
    label4 = Label(tamgiac_window)
    label4.pack()


def tinh_toan_tugiac():
    if len(ds) >= 1 and isinstance(ds[-1], TuGiac):
        tugiac = ds[-1]
        chu_vi = tugiac.chuvi()
        dien_tich = tugiac.dientich()
        goc_a, goc_b, goc_c, goc_d = tugiac.tinh_goc()

        s = f"Chu vi: {chu_vi}\nDiện tích: {dien_tich}\nGóc A: {goc_a} độ\nGóc B: {goc_b} độ\nGóc C: {goc_c} độ\nGóc D: {goc_d} độ"
        label2.config(text=s)
    else:
        label2.config(text="Không đủ dữ liệu hoặc dữ liệu không phù hợp")


def open_tugiac_window():
    tugiac_window = Toplevel(root)
    tugiac_window.title("Tứ giác")

    label1 = Label(tugiac_window, text='Hãy nhập 4 cạnh: ')
    label1.pack()

    canh1 = DoubleVar()
    canh2 = DoubleVar()
    canh3 = DoubleVar()
    canh4 = DoubleVar()

    E1 = Entry(tugiac_window, textvariable=canh1)
    E1.pack()
    E2 = Entry(tugiac_window, textvariable=canh2)
    E2.pack()
    E3 = Entry(tugiac_window, textvariable=canh3)
    E3.pack()
    E4 = Entry(tugiac_window, textvariable=canh4)
    E4.pack()

    b1 = Button(tugiac_window, text='Thêm dữ liệu', command=lambda: nhap_du_lieu_tugiac(canh1.get(), canh2.get(), canh3.get(), canh4.get()))
    b1.pack()
    b2 = Button(tugiac_window, text='Tính toán', command=tinh_toan_tugiac)
    b2.pack()

    global label2  # Khai báo label2 là biến toàn cục
    label2 = Label(tugiac_window)
    label2.pack()

def nhap_du_lieu_tamgiac(a, b, c):
    if not a or not b or not c:
        # Hiển thị thông báo lỗi
        label4.config(text="Các cạnh không được để trống")
    elif a <= 0 or b <= 0 or c <= 0:
        # Hiển thị thông báo lỗi
        label4.config(text="Các cạnh phải là số dương")
    else:
        tamgiac = TamGiac(a, b, c)
        ds.append(tamgiac)

def nhap_du_lieu_tugiac(a, b, c, d):
    if not a or not b or not c or not d:
        # Hiển thị thông báo lỗi
        label2.config(text="Các cạnh không được để trống")
    elif a <= 0 or b <= 0 or c <= 0 or d <= 0:
        # Hiển thị thông báo lỗi
        label2.config(text="Các cạnh phải là số dương")
    else:
        tugiac = TuGiac(a, b, c, d)
        ds.append(tugiac)
def tinh_toan_tamgiac():
    if len(ds) >= 1 and isinstance(ds[-1], TamGiac):
        tamgiac = ds[-1]
        chu_vi = tamgiac.chuvi()
        dien_tich = tamgiac.dientich()
        goc_a, goc_b, goc_c = tamgiac.tinh_goc()

        s = f"Chu vi: {chu_vi}\nDiện tích: {dien_tich}\nGóc A: {goc_a} độ\nGóc B: {goc_b} độ\nGóc C: {goc_c} độ"
        label4.config(text=s)
    else:
        label4.config(text="Không đủ dữ liệu hoặc dữ liệu không phù hợp")




ds = []

root = Tk()
root.geometry('300x200')
root.title('Lựa chọn hình')

tamgiac_button = Button(root, text="Tam giác", command=open_tamgiac_window)
tamgiac_button.pack()

tugiac_button = Button(root, text="Tứ giác", command=open_tugiac_window)
tugiac_button.pack()

root.mainloop()