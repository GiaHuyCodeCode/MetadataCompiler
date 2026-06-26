# Agent Prompt Templates
# Danh sách câu lệnh giao việc cho AI Agent

Tài liệu này chứa các câu lệnh (prompts) chuẩn đã được tối ưu để giao việc cho AI Agent trong hệ thống MetadataCompiler_V1. Bất kỳ ai, kể cả người không hiểu sâu về hệ thống, chỉ cần **copy, thay đổi nội dung trong ngoặc vuông `[...]`**, và dán vào khung chat để Agent thực thi.

---

## 1. Lệnh xử lý cho MỘT file đơn lẻ (An toàn, có chờ duyệt)
*Sử dụng khi bạn muốn Agent làm cẩn thận từng file một, phân tích lỗi và trình bày kế hoạch (Implementation Plan) cho bạn duyệt trước khi nó tự động sửa code.*

**Copy câu lệnh sau:**
```text
Hãy generate/review code VBA cho sheet `[Tên Sheet]` ở thư mục `[Tên thư mục test]`. Yêu cầu tuân thủ nghiêm ngặt các quy tắc trong agents/sk-architect/SKILL.md, đặc biệt là các quy tắc về tối ưu Production. Bắt buộc phải lập Implementation Plan chờ duyệt, và sau khi được duyệt, phải in ra Checklist & Chứng minh đầy đủ trước khi xuất khối code VBA chất lượng ra bên ngoài.
```

**Ví dụ cách điền:**
- `[Tên Sheet]` → `アカウント（入居者）_Y`
- `[Tên thư mục test]` → `output/test1/sample`

---

## 2. Lệnh xử lý HÀNG LOẠT nhiều file (Nhanh, tự động hoàn toàn)
*Sử dụng khi bạn có một thư mục chứa nhiều file cần làm lại. Agent sẽ tự động phân tích và sửa hàng loạt file mà không bắt bạn phải Approve liên tục.*

**Copy câu lệnh sau:**
```text
Hãy rà soát và chỉnh sửa code VBA hàng loạt cho các sheet trong thư mục `[Tên thư mục test]` so với spec sheet_raw.json. Yêu cầu: BỎ QUA việc lập Implementation Plan và chờ Approve. Agent hãy tự động phân tích, sửa code thẳng vào các file `.bas` nếu có sai lệch so với spec và agents/sk-architect/SKILL.md. Bắt buộc vẫn phải tự rà soát Checklist (Bước 5.4) trước khi hoàn tất mỗi file.
```

**Ví dụ cách điền:**
- `[Tên thư mục test]` → `output/test2/0023`

---

## 3. Lệnh Terminal (Quét Spec & Sinh Code Nhanh)
*Sử dụng khi bạn muốn dùng trực tiếp trình biên dịch Native Compiler (chạy nội bộ, không cần dùng Agent AI) để quét file Excel ra JSON và tự động dịch thành file VBA.*

**Chạy từ Terminal / Command Prompt:**

- **Lệnh SCAN (Phân tích Excel thành JSON):**
  ```bash
  python3 main.py scan --[Thư mục test]
  ```
  *(Ví dụ: `python3 main.py scan --test2/0023`)*

- **Lệnh COMPILE (Biên dịch JSON thành code VBA):**
  ```bash
  python3 main.py compile --[Thư mục test]
  ```
  *(Ví dụ: `python3 main.py compile --test2/0023`)*

- **Lệnh RUN (Thực hiện tuần tự cả SCAN và COMPILE):**
  ```bash
  python3 main.py run --[Thư mục test]
  ```
  *(Ví dụ: `python3 main.py run --test2/0023`)*

> **Mẹo:** Bạn có thể truyền thêm cờ `--file [Tên Sheet]` nếu chỉ muốn quét/dịch một sheet cụ thể.
> *(Ví dụ: `python3 main.py run --test1 --file 契約_K`)*

---

### 💡 Lưu ý quan trọng:
Hai câu lệnh giao việc cho Agent (Phần 1 & 2) đã được tích hợp sẵn các chỉ thị cốt lõi của hệ thống:
1. Ép Agent phải tuân thủ chuẩn kiến trúc sản phẩm tại `agents/sk-architect/SKILL.md`.
2. Yêu cầu Agent bắt buộc phải chạy qua bước kiểm tra chất lượng (Checklist 5.4) để ngăn ngừa lỗi logic trước khi lưu file `*.bas`.

> **⚠️ CẢNH BÁO CRITICAL (SAU KHI NHỜ AI REVIEW/SỬA CODE):**
> Sau khi bạn đã giao việc cho AI Agent (sử dụng lệnh ở Phần 1 hoặc 2) và Agent đã tự động sửa trực tiếp các file `.bas`, **TUYỆT ĐỐI KHÔNG** được yêu cầu chạy lại lệnh `compile` hay `run` (ở Phần 3). 
> Hệ thống native compiler của Python sẽ ghi đè và làm mất sạch toàn bộ những đoạn code VBA phức tạp mà AI Agent vừa tốn công tinh chỉnh tay, khiến đoạn code trở về trạng thái lỗi cũ (vô nghĩa). Lệnh `compile` / `run` chỉ nên được dùng ở bước đầu tiên để tạo khung sườn!
