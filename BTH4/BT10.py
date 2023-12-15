import pandas as pd
import matplotlib.pyplot as plt

# Đọc dữ liệu từ file CSV
data = pd.read_csv('Student_Performance.csv')

# Hiển thị dữ liệu
print("Dữ liệu:")
print(data.head())

# Tìm giá trị lớn nhất (max) của từng cột
max_values = data.max()
print("Giá trị lớn nhất của từng cột:")
print(max_values)

# Tìm giá trị nhỏ nhất (min) của từng cột
min_values = data.min()
print("Giá trị nhỏ nhất của từng cột:")
print(min_values)

# Tính giá trị trung bình (mean) của từng cột
mean_values = data.mean()
print("Giá trị trung bình của từng cột:")
print(mean_values)

# Vẽ đồ thị phân bố cho từng cột
data.hist(bins=10, figsize=(10, 8))
plt.tight_layout()
plt.show()