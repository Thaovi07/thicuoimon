import pandas as pd
import numpy as np
import os

os.makedirs('sample_data', exist_ok=True)

# Tạo bộ dữ liệu mẫu kinh doanh thực tế
data = [
    {"Ma_Don_Hang": "DH-1001", "Ngay": "2026-01-05", "Khach_Hang": "Nguyễn Văn An", "Khu_Vuc": "Miền Bắc", "Nganh_Hang": "Điện tử", "San_Pham": "Laptop Dell XPS", "So_Luong": 2, "Don_Gia": 28000000, "Doanh_Thu": 56000000, "Chi_Phi": 46000000, "Loi_Nhuan": 10000000, "Danh_Gia": 5},
    {"Ma_Don_Hang": "DH-1002", "Ngay": "2026-01-07", "Khach_Hang": "Trần Thị Mai", "Khu_Vuc": "Miền Nam", "Nganh_Hang": "Gia dụng", "San_Pham": "Nồi chiên không dầu", "So_Luong": 5, "Don_Gia": 2500000, "Doanh_Thu": 12500000, "Chi_Phi": 8500000, "Loi_Nhuan": 4000000, "Danh_Gia": 4},
    {"Ma_Don_Hang": "DH-1003", "Ngay": "2026-01-12", "Khach_Hang": "Lê Hoàng Long", "Khu_Vuc": "Miền Trung", "Nganh_Hang": "Thời trang", "San_Pham": "Áo khoác gió cao cấp", "So_Luong": 10, "Don_Gia": 850000, "Doanh_Thu": 8500000, "Chi_Phi": 5000000, "Loi_Nhuan": 3500000, "Danh_Gia": 5},
    {"Ma_Don_Hang": "DH-1004", "Ngay": "2026-01-15", "Khach_Hang": "Phạm Minh Tuấn", "Khu_Vuc": "Miền Bắc", "Nganh_Hang": "Điện tử", "San_Pham": "Màn hình 4K LG", "So_Luong": 3, "Don_Gia": 9500000, "Doanh_Thu": None, "Chi_Phi": 22000000, "Loi_Nhuan": 6500000, "Danh_Gia": 4}, # Missing Doanh_Thu
    {"Ma_Don_Hang": "DH-1005", "Ngay": "2026-01-18", "Khach_Hang": "Vũ Thu Trang", "Khu_Vuc": "Miền Nam", "Nganh_Hang": "Mỹ phẩm", "San_Pham": "Bộ dưỡng da Skincare", "So_Luong": 8, "Don_Gia": 1200000, "Doanh_Thu": 9600000, "Chi_Phi": 6000000, "Loi_Nhuan": 3600000, "Danh_Gia": 5},
    {"Ma_Don_Hang": "DH-1006", "Ngay": "2026-01-20", "Khach_Hang": "Đỗ Quang Hải", "Khu_Vuc": "Tây Nguyên", "Nganh_Hang": "Gia dụng", "San_Pham": "Máy hút bụi Robot", "So_Luong": 2, "Don_Gia": 7000000, "Doanh_Thu": 14000000, "Chi_Phi": 10500000, "Loi_Nhuan": 3500000, "Danh_Gia": 3},
    {"Ma_Don_Hang": "DH-1007", "Ngay": "2026-02-02", "Khach_Hang": "Hoàng Yến Chi", "Khu_Vuc": "Miền Bắc", "Nganh_Hang": "Thiết bị số", "San_Pham": "Tai nghe chống ồn", "So_Luong": 15, "Don_Gia": 1500000, "Doanh_Thu": 22500000, "Chi_Phi": 15000000, "Loi_Nhuan": 7500000, "Danh_Gia": 5},
    {"Ma_Don_Hang": "DH-1008", "Ngay": "2026-02-05", "Khach_Hang": "Đặng Quốc Bảo", "Khu_Vuc": "Miền Nam", "Nganh_Hang": "Điện tử", "San_Pham": "Bàn phím cơ không dây", "So_Luong": 6, "Don_Gia": 2200000, "Doanh_Thu": 13200000, "Chi_Phi": 9000000, "Loi_Nhuan": 4200000, "Danh_Gia": 4},
    {"Ma_Don_Hang": "DH-1008", "Ngay": "2026-02-05", "Khach_Hang": "Đặng Quốc Bảo", "Khu_Vuc": "Miền Nam", "Nganh_Hang": "Điện tử", "San_Pham": "Bàn phím cơ không dây", "So_Luong": 6, "Don_Gia": 2200000, "Doanh_Thu": 13200000, "Chi_Phi": 9000000, "Loi_Nhuan": 4200000, "Danh_Gia": 4}, # Hàng trùng lặp (duplicate)
    {"Ma_Don_Hang": "DH-1009", "Ngay": "2026-02-10", "Khach_Hang": "Bùi Thảo My", "Khu_Vuc": "Miền Trung", "Nganh_Hang": "Thời trang", "San_Pham": "Giày thể thao Sneaker", "So_Luong": 4, "Don_Gia": 1800000, "Doanh_Thu": 7200000, "Chi_Phi": None, "Loi_Nhuan": 2400000, "Danh_Gia": 4}, # Missing Chi_Phi
    {"Ma_Don_Hang": "DH-1010", "Ngay": "2026-02-14", "Khach_Hang": "Ngô Gia Huy", "Khu_Vuc": "Miền Bắc", "Nganh_Hang": "Mỹ phẩm", "San_Pham": "Nước hoa cao cấp", "So_Luong": 3, "Don_Gia": 3500000, "Doanh_Thu": 10500000, "Chi_Phi": 6500000, "Loi_Nhuan": 4000000, "Danh_Gia": 5},
    {"Ma_Don_Hang": "DH-1011", "Ngay": "2026-02-18", "Khach_Hang": "Trịnh Thu Hà", "Khu_Vuc": "Miền Nam", "Nganh_Hang": "Gia dụng", "San_Pham": "Nồi cơm điện tử", "So_Luong": 7, "Don_Gia": 1800000, "Doanh_Thu": 12600000, "Chi_Phi": 8400000, "Loi_Nhuan": 4200000, "Danh_Gia": 4},
    {"Ma_Don_Hang": "DH-1012", "Ngay": "2026-02-22", "Khach_Hang": "Dương Văn Đức", "Khu_Vuc": "Tây Nguyên", "Nganh_Hang": "Điện tử", "San_Pham": "Chuột Gaming Logitech", "So_Luong": 12, "Don_Gia": 1100000, "Doanh_Thu": 13200000, "Chi_Phi": 8500000, "Loi_Nhuan": 4700000, "Danh_Gia": 5},
    {"Ma_Don_Hang": "DH-1013", "Ngay": "2026-03-01", "Khach_Hang": "Lý Thanh Vân", "Khu_Vuc": "Miền Bắc", "Nganh_Hang": "Thời trang", "San_Pham": "Túi xách da nữ", "So_Luong": 5, "Don_Gia": 2600000, "Doanh_Thu": 13000000, "Chi_Phi": 8000000, "Loi_Nhuan": 5000000, "Danh_Gia": 4},
    {"Ma_Don_Hang": "DH-1014", "Ngay": "2026-03-05", "Khach_Hang": "Đinh Xuân Bách", "Khu_Vuc": "Miền Nam", "Nganh_Hang": "Thiết bị số", "San_Pham": "Ổ cứng SSD 1TB", "So_Luong": 8, "Don_Gia": 2100000, "Doanh_Thu": 16800000, "Chi_Phi": 12000000, "Loi_Nhuan": 4800000, "Danh_Gia": 5},
    {"Ma_Don_Hang": "DH-1015", "Ngay": "2026-03-10", "Khach_Hang": "Lê Tuyết Mai", "Khu_Vuc": "Miền Trung", "Nganh_Hang": "Mỹ phẩm", "San_Pham": "Son môi dưỡng ẩm", "So_Luong": 20, "Don_Gia": 450000, "Doanh_Thu": 9000000, "Chi_Phi": 5500000, "Loi_Nhuan": 3500000, "Danh_Gia": 4},
    {"Ma_Don_Hang": "DH-1016", "Ngay": "2026-03-15", "Khach_Hang": "Phan Hữu Nghĩa", "Khu_Vuc": "Miền Bắc", "Nganh_Hang": "Điện tử", "San_Pham": "Máy tính bảng iPad Air", "So_Luong": 3, "Don_Gia": 16500000, "Doanh_Thu": 49500000, "Chi_Phi": 41000000, "Loi_Nhuan": 8500000, "Danh_Gia": 5},
    {"Ma_Don_Hang": "DH-1017", "Ngay": "2026-03-20", "Khach_Hang": "Vương Kiều Trang", "Khu_Vuc": "Miền Nam", "Nganh_Hang": "Gia dụng", "San_Pham": "Quạt không cánh thông minh", "So_Luong": 4, "Don_Gia": 3200000, "Doanh_Thu": 12800000, "Chi_Phi": 8800000, "Loi_Nhuan": 4000000, "Danh_Gia": 4},
    {"Ma_Don_Hang": "DH-1018", "Ngay": "2026-03-25", "Khach_Hang": "Tạ Quang Dũng", "Khu_Vuc": "Tây Nguyên", "Nganh_Hang": "Thời trang", "San_Pham": "Thắt lưng da bò", "So_Luong": 10, "Don_Gia": 600000, "Doanh_Thu": 6000000, "Chi_Phi": 3500000, "Loi_Nhuan": 2500000, "Danh_Gia": 4},
    {"Ma_Don_Hang": "DH-1019", "Ngay": "2026-03-28", "Khach_Hang": "Chu Ngọc Hân", "Khu_Vuc": "Miền Bắc", "Nganh_Hang": "Thiết bị số", "San_Pham": "Loa Bluetooth JBL", "So_Luong": 6, "Don_Gia": 1900000, "Doanh_Thu": 11400000, "Chi_Phi": 7500000, "Loi_Nhuan": 3900000, "Danh_Gia": 5},
    {"Ma_Don_Hang": "DH-1020", "Ngay": "2026-03-30", "Khach_Hang": "Nguyễn Hoàng Nam", "Khu_Vuc": "Miền Nam", "Nganh_Hang": "Điện tử", "San_Pham": "Webcam Full HD Logitech", "So_Luong": 100, "Don_Gia": 1500000, "Doanh_Thu": 150000000, "Chi_Phi": 110000000, "Loi_Nhuan": 40000000, "Danh_Gia": 5} # Đơn hàng ngoại lai đột biến (outlier B2B)
]

df = pd.DataFrame(data)
csv_path = os.path.join('sample_data', 'du_lieu_kinh_doanh_mau.csv')
xlsx_path = os.path.join('sample_data', 'du_lieu_kinh_doanh_mau.xlsx')

df.to_csv(csv_path, index=False, encoding='utf-8-sig')
df.to_excel(xlsx_path, index=False)

print(f"Created {csv_path} and {xlsx_path} with {len(df)} rows.")
