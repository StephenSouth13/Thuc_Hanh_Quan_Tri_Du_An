# 📊 Áp dụng quy trình 5 bước xử lý nhanh dữ liệu tài nguyên trong bảng Dự toán dân dụng

## 🎯 Mục tiêu
Bài tập này áp dụng **5 bước xử lý dữ liệu** trên file Excel dự toán xây dựng dân dụng, nhằm:
1. Đọc dữ liệu gốc từ file Excel phức tạp nhiều sheet.
2. Xác định sheet và dòng header chứa thông tin cần thiết.
3. Chuẩn hóa, lọc, và làm sạch dữ liệu.
4. Tổng hợp theo **Mã hiệu / Mã số** để tính:
   - Tổng khối lượng  
   - Tổng thành tiền
5. Xuất kết quả ra file Excel đẹp + biểu đồ minh họa trực quan.

---

## 📂 Cấu trúc thư mục

buoi_5/
│── buoi5.py # Script Python chính
│── check_excel.py # Script hỗ trợ dò sheet và header
│── Áp dụng quy trình 5 bước xử lý nhanh dữ liệu tài nguyên trong bảng Dự toán dân dụng (1).xlsx # File Excel gốc
│── buoi5_ketqua.xlsx # File Excel kết quả sau xử lý
│── buoi5_chart.png # Biểu đồ top 10 mã hiệu theo khối lượng
│── README.md # Tài liệu hướng dẫn


---

## ⚙️ Cách chạy

### 1. Chuẩn bị môi trường
Cài Python 3.10+ và thư viện cần thiết:
```bash
pip install pandas matplotlib openpyxl

2. Chạy script xử lý

Trong thư mục buoi_5, chạy:

py buoi5.py

3. Kết quả

buoi5_ketqua.xlsx: gồm 2 sheet

Du lieu goc: dữ liệu gốc đã làm sạch

Tong hop: dữ liệu tổng hợp theo mã hiệu/mã số

buoi5_chart.png: biểu đồ top 10 mã hiệu theo khối lượng

📊 Minh họa kết quả
File Excel đầu ra:

Có header được tô màu xanh + chữ trắng.

Số liệu căn chỉnh, có border, format #,##0.00.

Dễ đọc và dễ sử dụng cho báo cáo.

Biểu đồ:

Dạng bar chart ngang.

Thể hiện Top 10 mã hiệu có khối lượng lớn nhất.

Giúp hình dung nhanh phân bổ khối lượng.

✅ Kết luận

Quy trình này cho phép:

Tự động nhận diện sheet và header có dữ liệu hợp lệ.

Chuẩn hóa & tổng hợp dữ liệu nhanh chóng.

Xuất báo cáo đẹp + biểu đồ trực quan, tiết kiệm thời gian làm dự toán.