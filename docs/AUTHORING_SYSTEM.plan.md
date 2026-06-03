**✅ PLAN TỔNG THỂ ĐÃ ĐƯỢC LẬP LẠI CHO 4 MODULE AUTHORING SYSTEM**

Dưới đây là kế hoạch **hoàn chỉnh, rõ ràng và thực tế** cho toàn bộ 4 module của hệ thống tạo câu hỏi (Authoring System).

### **Mục tiêu Tổng thể của Prototype**
Xây dựng một hệ thống cho phép Admin **tạo bài thi hoàn chỉnh** (từ file mẫu → viết yêu cầu → thiết kế chấm điểm → đóng gói) và có thể chạy thử được bài thi đó.

**Phạm vi Prototype**: Tập trung mạnh vào **Excel** trước, sau đó mở rộng Word & PowerPoint.

---

### **Kiến trúc Chung**
- **Project Structure**:
    - `AuthoringTool` (WPF) – Giao diện chính
    - `CommonLibrary` – Models, DTOs, Helpers
    - `OfficeIntegration` – Open XML, EPPlus, Interop
    - `ScoringEngine` – Module chấm điểm riêng

- **Lưu trữ**: JSON + Folder structure (dễ implement)
- **Thời gian ước tính**: 4–6 tuần (full-time)

---

### **MODULE 1: TEMPLATE MANAGER**

**Chức năng chính**:
- Tạo, quản lý, chỉnh sửa file mẫu (Template)
- Hỗ trợ Excel, Word, PowerPoint
- Preview và Versioning template

**Hướng xử lý**:
- Sử dụng **Office Interop** để mở ứng dụng thật cho Admin chỉnh sửa.
- Lưu file template vào thư mục `Templates/{OfficeType}/`
- Lưu metadata (tên, mô tả, loại, ngày tạo) dưới dạng JSON.

**Các tính năng cụ thể**:
1. New Template (chọn loại Office)
2. Open Template (mở bằng Office thật)
3. Rename / Delete / Duplicate
4. Preview (hiển thị thông tin + thumbnail nếu có)

**Thời gian**: 5–6 ngày

---

### **MODULE 2: TASK EDITOR**

**Chức năng chính**:
- Tạo bài thi mới (Test)
- Thêm/Sửa/Xóa/Sắp xếp các Task
- Viết Instruction chi tiết cho thí sinh

**Hướng xử lý**:
- Một Test chứa nhiều Task
- Mỗi Task liên kết với **1 Template**
- Lưu toàn bộ Test dưới dạng `Test.json`

**Cấu trúc JSON cho Task**:
```json
{
  "TaskId": "T001",
  "Number": 1,
  "Instruction": "Format the chart title as Bold and change color to Blue",
  "Points": 10,
  "TemplateFile": "SalesReport.xlsx",
  "Objective": "Create and format charts"
}
```

**Tính năng**:
- Kéo thả sắp xếp Task
- Preview Task (mở template + hiển thị Instruction)
- Thời gian ước tính cho Test

**Thời gian**: 6–7 ngày

---

### **MODULE 3: SCORING DESIGNER**

**Chức năng chính**:
- Thiết kế quy tắc chấm điểm cho từng Task
- Xây dựng Scoring Engine tự động

**Hướng xử lý**:
- Sử dụng **EPPlus** (chính) + **Open XML SDK** (nâng cao)
- Áp dụng **Strategy Pattern** cho từng loại Rule

**Các loại Rule hỗ trợ (Prototype)**:
1. Cell Value Check
2. Formula Check
3. Formatting Check (Bold, Color, FontSize, NumberFormat…)
4. Object Existence (Chart, PivotTable, Table)
5. Range Comparison

**Giao diện**:
- Chọn Task → Add Rule → Chọn Rule Type → Cấu hình chi tiết
- Nút **Test Rule** (chạy thử trên file)

**Thời gian**: 8–10 ngày (phần khó nhất)

---

### **MODULE 4: TEST PACKAGER**

**Chức năng chính**:
- Kiểm tra tính hoàn chỉnh của bài thi
- Đóng gói toàn bộ bài thi thành package
- Export / Import bài thi

**Hướng xử lý**:
- Tạo một thư mục package có cấu trúc rõ ràng:
  ```
  MyTestPackage/
  ├── test.json
  ├── Templates/
  ├── ScoringRules/
  └── Assets/
  ```
- Có thể nén thành `.zip` hoặc giữ dạng folder
- Validation trước khi package (kiểm tra có Template, có Rule, có Instruction…)

**Tính năng**:
- Publish to Test Bank
- Export Package
- Import Package
- Version Control (v1.0, v1.1…)

**Thời gian**: 5–6 ngày

---

### **Timeline Tổng thể (6 Tuần)**

| Tuần | Module | Mục tiêu chính | Deliverable |
|------|--------|----------------|-------------|
| 1    | Module 1 | Template Manager hoàn chỉnh | Tạo & quản lý template Excel |
| 2    | Module 2 | Task Editor | Tạo Test có nhiều Task |
| 3    | Module 3 (Phần 1) | Scoring Designer cơ bản | 4–5 Rule Types + Engine |
| 4    | Module 3 (Phần 2) | Scoring nâng cao | Preview Rule + Báo cáo |
| 5    | Module 4 | Test Packager | Export/Import Package |
| 6    | Tích hợp & Test | Kết nối 4 module | Full flow: Tạo → Đóng gói → Chạy thi → Chấm điểm |

---

### **Khuyến nghị Quan trọng**

1. **Làm theo thứ tự**: Module 1 → 2 → 3 → 4
2. **Chỉ tập trung Excel** ở prototype đầu tiên
3. **Sử dụng JSON** làm chính cho tất cả dữ liệu
4. **Tách biệt UI và Business Logic** (MVVM trong WPF)
5. Xây dựng **Logging** và **Error Handling** tốt ngay từ đầu
6. Sau khi xong 4 module, mới tích hợp với **Test Runner** (phần thí sinh làm bài)

---

Bạn muốn tôi bắt đầu triển khai **chi tiết + code mẫu** cho module nào trước?

Hãy chọn:
- **Module 1: Template Manager** (dễ nhất, nên làm trước)
- **Module 3: Scoring Designer** (phần quan trọng)
- Hoặc tôi viết **Model Classes + JSON Schema** cho toàn bộ 4 module trước

Bạn muốn bắt đầu từ đâu?