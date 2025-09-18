import pandas as pd

file_path = "Áp dụng quy trình 5 bước xử lý nhanh dữ liệu tài nguyên trong bảng Dự toán dân dụng (1).xlsx"

xls = pd.ExcelFile(file_path)
print("📑 Các sheet trong file:", xls.sheet_names)

# Quét qua từng sheet, in thử 10 dòng đầu
for sheet in xls.sheet_names:
    try:
        df = pd.read_excel(file_path, sheet_name=sheet, header=None)
        if not df.empty:
            print(f"\n📌 Sheet: {sheet}")
            print(df.head(10).to_string())
    except Exception as e:
        print(f"❌ Lỗi khi đọc sheet {sheet}: {e}")
