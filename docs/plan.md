**✅ PLAN CHI TIẾT: Xây dựng App Mẫu Test Tích hợp Office (Proof of Concept)**

Mục tiêu: Xây dựng một ứng dụng đơn giản để kiểm tra xem ý tưởng tương tự GMetrix có khả thi không.  
**Scope nhỏ**: Mở app → Chọn Office → Mở Office + Add-in → Hiển thị Task Pane + Bắt vài event cơ bản.

### **Kiến trúc Tổng thể (2 phần)**

1. **Main App** (C# WPF hoặc WinForms)
    - Giao diện chính, chọn Excel/Word/PowerPoint.
    - Quản lý và khởi động Office.

2. **VSTO Add-ins** (riêng cho từng app)
    - Chạy bên trong Excel/Word/PowerPoint.
    - Hiển thị Custom Task Pane + Bắt events.

---

### **PHASE 1: Chuẩn bị (1-2 ngày)**

1. Cài đặt môi trường:
    - Visual Studio 2022/2025 (Community cũng được).
    - Workload: **.NET desktop development** + **Office/SharePoint development**.
    - Microsoft Office 2019/2021/365 (đầy đủ).

2. Tạo Solution:
    - Tạo **Solution** tên `OfficeSimulatorTest`.
    - Thêm 4 projects:
        - `MainApp` (WPF App)
        - `ExcelAddIn` (Excel VSTO Add-in)
        - `WordAddIn` (Word VSTO Add-in)
        - `CommonLibrary` (Class Library) — chứa code chung.

---

### **PHASE 2: Xây dựng Main App (2-3 ngày)**

Chức năng chính:
- Giao diện đơn giản: Button “Excel”, “Word”, “PowerPoint”.
- Khi click → Mở ứng dụng Office tương ứng.

**Code mẫu (MainApp):**

```csharp
private void OpenExcel_Click(object sender, RoutedEventArgs e)
{
    try
    {
        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        excel.Visible = true;
        excel.Workbooks.Add();
        MessageBox.Show("Đã mở Excel. Add-in sẽ tự động load nếu đã cài.");
    }
    catch (Exception ex)
    {
        MessageBox.Show("Lỗi: " + ex.Message);
    }
}
```

Làm tương tự cho Word và PowerPoint.

---

### **PHASE 3: Xây dựng VSTO Add-ins (3-5 ngày)**

**Mục tiêu**: Khi Office mở, Add-in tự động load → Hiện Task Pane + Bắt event.

#### **A. ExcelAddIn**

Trong `ThisAddIn.cs`:
```csharp
private CustomTaskPane taskPane;

private void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    // Tạo Task Pane
    var control = new UserControl1();  // UserControl chứa giao diện
    taskPane = this.CustomTaskPanes.Add(control, "Test Simulator");
    taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
    taskPane.Visible = true;
    taskPane.Height = 250;

    // Bắt một số Event
    Globals.ThisAddIn.Application.SheetChange += Application_SheetChange;
    Globals.ThisAddIn.Application.WorkbookOpen += Application_WorkbookOpen;
}

private void Application_SheetChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
{
    // Gửi thông tin về Main App hoặc ghi log
    System.Diagnostics.Debug.WriteLine($"Cell thay đổi: {Target.Address}");
    // TODO: Gửi event qua NamedPipe hoặc local HTTP
}

private void Application_WorkbookOpen(Microsoft.Office.Interop.Excel.Workbook Wb)
{
    MessageBox.Show("Workbook opened - Simulator is running!");
}
```

#### **B. WordAddIn & PowerPointAddIn**
Tương tự, nhưng dùng các event phù hợp:
- Word: `Application.DocumentOpen`, `Application.WindowSelectionChange`
- PowerPoint: `Application.PresentationOpen`, `Application.SlideSelectionChanged`

---

### **PHASE 4: Giao tiếp giữa Main App và Add-in**

Cách đơn giản nhất cho prototype:
- Dùng **Named Pipes** hoặc **Local HTTP server** (tiny web api).
- Hoặc đơn giản hơn: Ghi log ra file chung + Add-in đọc file.

---

### **PHASE 5: Build & Test**

1. Build từng Add-in → Copy file `.dll` và `.vsto` vào thư mục test.
2. **Debug**:
    - Set Start Action của Add-in project là “Start Excel” (hoặc Word).
    - Nhấn F5 → Visual Studio sẽ mở Office và load Add-in.
3. Test flow:
    - Mở MainApp → Click Excel → Excel mở + Task Pane hiện ở dưới.
    - Thay đổi cell trong Excel → Add-in bắt event và hiển thị.

---

### **Timeline Thực tế (cho người có kinh nghiệm C# trung bình)**

| Giai đoạn              | Thời gian     | Trạng thái |
|------------------------|---------------|----------|
| Chuẩn bị môi trường    | 1 ngày        | Hoàn thành |
| Main App               | 2 ngày        | Hoàn thành |
| Excel Add-in + Task Pane | 3 ngày      | Hoàn thành |
| Word & PowerPoint      | 3-4 ngày      | Tùy chọn |
| Giao tiếp + Logging    | 2 ngày        | Nâng cao |
| **Tổng**               | **10-14 ngày** | Prototype |

---

### **Rủi ro & Khả thi**

**Khả thi cao** nếu:
- Bạn đã quen C# và ít nhiều biết Interop.
- Chỉ test Excel trước (khó nhất).

**Khó khăn có thể gặp**:
- Add-in không load (do Registry).
- Conflict với các add-in khác.
- Event không bắt được (do permission).

**Tiêu chí thành công của prototype**:
- Main App mở được Office.
- Task Pane tự động hiện khi Office mở.
- Bắt được ít nhất 2-3 event (SheetChange, WorkbookOpen…).
- Không crash khi thao tác cơ bản.

---