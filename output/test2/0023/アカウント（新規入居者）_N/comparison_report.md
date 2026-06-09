# Báo Cáo So Sánh File VBA .bas Mới (Sau Cập Nhật)

## 1. Tổng quan
Sau khi tệp `アカウント（新規入居者）_N (compare).bas` được bạn cập nhật trực tiếp, cấu trúc và logic của 2 file hiện tại đã khác biệt hoàn toàn:
- **`アカウント（新規入居者）_N.bas`:** Phiên bản cũ (do pipeline sinh ra tự động), chứa nhiều lỗi logic nghiêm trọng và thiếu trường.
- **`アカウント（新規入居者）_N (compare).bas`:** Phiên bản mới (sau khi bạn sửa đổi), đã tối ưu hóa và giải quyết hầu hết các nghiệp vụ phức tạp.

**Kết luận:** File **`アカウント（新規入居者）_N (compare).bas`** là file thực hiện đúng và đầy đủ mô tả nghiệp vụ nhất.

---

## 2. Bảng so sánh chi tiết giữa 2 file

| Đặc tính / Nghiệp vụ | `アカウント（新規入居者）_N.bas` (Cũ) | `アカウント（新規入居者）_N (compare).bas` (Mới) | Đánh giá file Mới |
| :--- | :--- | :--- | :--- |
| **Thư viện truy cập DB** | Sử dụng **DAO** (`DAO.Database`) | Sử dụng **ADODB** (`ADODB.Connection`) | Cả hai đều chạy được trong Access, tùy thuộc vào tiêu chuẩn chung của dự án. |
| **Xóa dữ liệu cũ** | Chỉ xóa tài khoản mới (`DELETE ... WHERE ID LIKE 'N_%'`) | Xóa sạch bảng (`DELETE FROM Account;`) | ⚠️ **Cần lưu ý:** File cũ an toàn hơn vì không làm ảnh hưởng đến tài khoản của các sheet khác (như `既存入居者` hoặc `新規オーナー`). |
| **Deduplication (DEDUP)** | Thực hiện trên bảng nguồn `新規契約更新一覧` bằng cột `FLG`. | Thực hiện trên bảng đích `Account` bằng lệnh `DELETE` sau khi `INSERT`. | Cả hai đều đúng, nhưng cách của file mới trực quan và dễ quản lý hơn. |
| **Liên kết bảng (`JOIN`)** | **Không** liên kết với `入居者管理` và `契約者情報`. | **Có** liên kết động qua mệnh đề `SWITCH` trực tiếp trong SQL. | **File mới đúng.** |
| **Thông tin `email` & `tel_mobile`** | - Thiếu trường `email`. <br>- Trường `tel_mobile` bị lỗi tham chiếu cột không tồn tại. | - Lấy từ `入居者管理` (ưu tiên). <br>- Tự động fallback sang `契約者情報` (nếu không khớp). | **File mới đúng.** |
| **Thông tin Họ tên & Công ty** | Thiếu hoàn toàn các trường `name_family`, `name_family_kana`, `company_name`, `company_name_kana`. | Điền đầy đủ, đồng thời đảo giá trị tên/kana cho trường Công ty đúng như `sheet_raw` mô tả. | **File mới đúng.** |
| **Dọn dẹp Cá nhân / Pháp nhân** | Không thực hiện. | Tự động xóa thông tin công ty nếu là Cá nhân, và xóa họ tên nếu là Pháp nhân. | **File mới đúng.** |
| **Mã hóa `kind_id`** | Gán sai giá trị tên khách hàng vào mã phân loại. | Chuyển đổi đúng: `'個人'` $\rightarrow$ `'10'`, `'法人'` $\rightarrow$ `'20'`. | **File mới đúng.** |
| **Thông tin `gender_id`** | Thiếu trường. | ❌ Vẫn **thiếu** trường `gender_id` (chưa được cập nhật từ bảng `契約者情報`). | Cần bổ sung cập nhật trường này. |

---

## 3. Các điểm cần tinh chỉnh thêm ở file `(compare).bas`

Mặc dù file `(compare).bas` đã rất tốt, bạn nên thực hiện 3 chỉnh sửa nhỏ sau để hoàn thiện 100%:

### Chỉnh sửa 1: Đảm bảo không xóa nhầm dữ liệu của sheet khác
Thay đổi câu lệnh xóa dữ liệu ở đầu hàm (dòng 19) để chỉ xóa tài khoản của sheet hiện tại:
```diff
-    db.Execute "DELETE FROM Account;"
+    db.Execute "DELETE FROM Account WHERE ID LIKE 'N_%';"
```

### Chỉnh sửa 2: Bổ sung cập nhật trường `gender_id`
Bổ sung trường `T.gender_id = T2.[性別]` vào khối cập nhật thông tin từ bảng `契約者情報` (từ dòng 227):
```diff
     SQL = ""
     SQL = SQL & "UPDATE (Account AS T "
     SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
     SQL = SQL & "INNER JOIN 契約者情報 AS T2 ON T2.[契約者No] = SWITCH( "
     SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
     SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
     SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
     SQL = SQL & "    True, T1.[契約者1No_17] ) "
-    SQL = SQL & "SET T.company_name = T2.[契約者名カナ]; "
+    SQL = SQL & "SET T.company_name = T2.[契約者名カナ], "
+    SQL = SQL & "    T.gender_id = IIF(T2.[性別] = '-', NULL, T2.[性別]); "
     db.Execute SQL
```
*(Hoặc cập nhật chung trong khối gán `kind_id` hoặc các trường thông tin cá nhân).*
