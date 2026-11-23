 

# KTCM AI Analytics Studio

**Nền Tảng Phân Tích Dữ Liệu & Báo Cáo Điều Hành Doanh Nghiệp Tự Động**  
*Sáng lập & Trưởng nhóm Phân tích (Founder & Lead Data Analyst):* **Mai Thảo Vi**

Công cụ AI phân tích dữ liệu tự động 6 giai đoạn khép kín, hoạt động độc lập trên môi trường `localhost`, hỗ trợ đa dạng định dạng dữ liệu (CSV, Excel), tích hợp đối soát dữ liệu (Human-in-the-Loop Audit), giao diện tối giản hiện đại chuẩn phong cách FinTech/SaaS và xuất báo cáo đa định dạng (HTML, PDF, DOCX) với hệ thống 6 biểu đồ phân tích chuyên sâu.

---

## I. Quy Trình 6 Giai Đoạn Bắt Buộc

1. **Giai đoạn 1: Nhập Dữ Liệu & Mục Tiêu Người Dùng**

   - Hỗ trợ tải file Excel (`.xlsx`, `.xls`) và `.csv` bằng thao tác kéo thả (Drag & Drop).
   - Tự động nhận diện cấu trúc (Auto-Schema Profiling): Tên cột, kiểu dữ liệu (Numeric, Datetime, Categorical, Text), số dòng khuyết, số giá trị duy nhất.
   - **Tính năng mở rộng theo yêu cầu**: Ô nhập mục tiêu nghiệp vụ (*User Goal*) kèm các thẻ prompt gợi ý để AI định hướng phân tích đúng trọng tâm.
   - Xem trước 10 dòng đầu (Sample preview).
2. **Giai đoạn 2: Làm Sạch Dữ Liệu (Automated Data Cleaning)**

   - Quét và phát hiện 100% dòng trùng lặp (*Duplicates*), ô dữ liệu khuyết (*Missing values*), điểm dị biệt (*Outliers* theo IQR).
   - Đề xuất phương án tự động: Tự động tính toán theo công thức quan hệ (`Doanh_Thu = So_Luong * Don_Gia`), điền Trung vị (Median), Trung bình (Mean), hoặc Mode.
   - Chuẩn hóa chuỗi (Strip whitespace), chuẩn hóa định dạng ngày tháng.
   - Giữ nguyên bộ dữ liệu gốc bất biến (*Raw data*) để đối chiếu.
3. **Giai đoạn 3: Duyệt Dữ Liệu & Đối Soát (Human-in-the-Loop Audit)**

   - Bảng Diff đối chiếu trước và sau xử lý với hiệu ứng highlight trực quan:
     * **Màu xanh lá**: Các ô dữ liệu vừa được điền bổ sung.
     * **Màu vàng**: Các ô vừa được chuẩn hóa định dạng.
   - Thống kê thời gian thực: *Số dòng hợp lệ, số dòng lỗi, số dòng đã hiệu chỉnh*.
   - Cho phép người dùng **bấm trực tiếp vào từng ô để chỉnh sửa giá trị**.
   - Nút **"Xác nhận duyệt dữ liệu"** để đóng băng dataset trước khi chuyển sang phân tích toán học.
4. **Giai đoạn 4: Phân Tích Chuyên Sâu (Deep Statistical Analysis & AI Q&A)**

   - Thống kê mô tả (Descriptive Statistics): Sum, Mean, Median, Std, Min, Max, Quantiles.
   - Ma trận tương quan Pearson giữa tất cả các biến số định lượng.
   - Phát hiện Outliers bằng thuật toán $IQR = Q_3 - Q_1$ ($[Q_1 - 1.5 \times IQR, Q_3 + 1.5 \times IQR]$).
   - **Trợ lý AI Truy vấn Dữ liệu (AI Chat Assistant)**: Cho phép người dùng gõ câu hỏi tự do để truy vấn số liệu tức thì. Tất cả kết quả đều được tính bằng code Pandas, không bịa số liệu.
5. **Giai đoạn 5: Kết Quả & Đề Xuất (Executive Dashboard)**

   - Bảng điều khiển phong cách Behance Modern SaaS: Tone màu tím pastel (`#7C3AED`, `#8B5CF6`) kết hợp xanh Cyan (`#06B6D4`) và nền xám sáng thanh lịch (`#F8FAFC`).
   - Thẻ KPI quan trọng: Tổng số bản ghi, Tổng doanh thu/giá trị, Tổng lợi nhuận, Tỷ suất lợi nhuận ròng, Sản lượng trung bình.
   - Biểu đồ tương tác thời gian thực (Chart.js): Bar chart, Doughnut chart, Timeline dynamic line chart.
   - **Bộ lọc tương tác (Slicers)** theo Ngành hàng và Khu vực: Khi thay đổi bộ lọc, biểu đồ và chỉ số KPI tự động cập nhật ngay lập tức.
   - Thẻ Insight giải thích kết quả và Danh mục đề xuất hành động thực tế dựa trên mục tiêu đã nhập ở Bước 1.
6. **Giai đoạn 6: Nghiệm Thu Kỹ Thuật & Xuất Báo Cáo**

   - **Biên bản kiểm tra nghiệm thu tự động (Automated QA Audit Gate)**:
     * ✓ Tổng số liệu trên dashboard khớp 100% với báo cáo.
     * ✓ Biểu đồ thay đổi mượt mà theo bộ lọc.
     * ✓ Công thức tính toán chạy bằng code Python chuẩn xác.
     * ✓ Nội dung Insight và Đề xuất bám sát mục tiêu.
   - Xem trước toàn văn báo cáo chuẩn chuyên nghiệp.
   - **Hỗ trợ 3 định dạng xuất báo cáo**:
     * **HTML Độc lập**: Tải file `.html` mở offline trên mọi thiết bị.
     * **PDF Executive**: Mở giao diện in chuẩn văn bản A4 sắc nét.
     * **DOCX Chỉnh sửa được**: Tải file Microsoft Word `.docx` chuẩn tạo bởi thư viện `python-docx`.

---

## II. Hướng Dẫn Khởi Chạy Trên Localhost

### Cách 1: Khởi chạy 1-Click bằng File Script

- Nhấp đúp chuột vào file [run_app.bat](file:///c:/Users/pc/OneDrive/Tài liệu/ktcm/run_app.bat)
- Hoặc mở PowerShell và chạy:
  ```powershell
  .\run_app.ps1
  ```

### Cách 2: Khởi chạy thủ công từ Terminal

```bash
python server.py
```

Sau đó mở trình duyệt và truy cập: **`http://localhost:8000`**

---

## III. Dữ Liệu Mẫu Kiểm Thử (Sample Data)

Ứng dụng đã đi kèm sẵn 2 tệp dữ liệu mẫu trong thư mục `sample_data/`:

1. `du_lieu_kinh_doanh_mau.csv`
2. `du_lieu_kinh_doanh_mau.xlsx`

Bộ dữ liệu này chứa 21 dòng giao dịch thương mại với đầy đủ các tình huống:

- 1 hàng trùng lặp hoàn toàn (DH-1008)
- 2 ô bị thiếu giá trị (`Doanh_Thu` và `Chi_Phi`)
- 1 đơn hàng ngoại lai quy mô lớn (DH-1020: 100 sản phẩm B2B)
- Đầy đủ các trường thời gian, khu vực, ngành hàng để kiểm thử toàn diện quy trình 6 bước.
