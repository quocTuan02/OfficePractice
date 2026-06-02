# OfficePractice — Office Simulator (Proof of Concept)

Ứng dụng thử nghiệm tích hợp Microsoft Office theo phong cách GMetrix: mở Office từ app chính, hiển thị Task Pane bên trong Excel/Word, và bắt events theo thời gian thực qua Named Pipe.

---

## Kiến trúc

```
MOS/
├── CommonLibrary/        # Shared: PipeMessage, PipeClient, PipeServer
├── MainApp/              # WPF app (.NET 8): giao diện chính + event log
├── ExcelAddIn/           # VSTO Excel Add-in (.NET Framework 4.8)
├── WordAddIn_old/        # VSTO Word Add-in (code sẵn, chưa tạo VS project)
└── docs/plan.md          # Kế hoạch chi tiết
```

### Luồng hoạt động

```
MainApp (WPF)
  ├── Mở Excel/Word qua COM
  ├── Named Pipe Server ←──── ExcelAddIn (VSTO)
  │     (nhận events)              └── Named Pipe Client
  │                                      (gửi events khi user thao tác)
  └── Task Panel (hướng dẫn + event log)
```

---

## Yêu cầu

| Thành phần | Version | Ghi chú |
|---|---|---|
| Windows | 10/11 | |
| .NET 8 SDK | ≥ 8.0 | Cho MainApp |
| .NET Framework | 4.8 Runtime + Developer Pack | Cho VSTO Add-ins |
| Visual Studio | 2022/2026 Community | Workload: **Microsoft 365 development** |
| Microsoft Office | Excel/Word 2013+ | Để chạy VSTO add-in |
| JetBrains Rider | Tùy chọn | Cho MainApp/CommonLibrary |

---

## Cài đặt môi trường

### 1. .NET Framework 4.8 Developer Pack

> **Lưu ý**: Runtime 4.8 đã có sẵn trên Windows 10/11. Cần cài thêm **Developer Pack** (reference assemblies).

Tải tại: `https://dotnet.microsoft.com/download/dotnet-framework/net48`

Chọn **"Developer Pack"** (không phải Runtime hay Web Installer) → file `ndp48-devpack-enu.exe`

Kiểm tra sau khi cài:
```powershell
ls "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8"
# Phải thấy ~237 DLL files
```

### 2. Visual Studio 2022/2026 Community

Tải tại: `https://visualstudio.microsoft.com/downloads/`

Khi cài chọn **2 workloads**:
- ✅ **Microsoft 365 development** — bắt buộc cho VSTO
- ✅ **.NET desktop development** — cho WPF

### 3. Microsoft Office

Office 2013 trở lên (Excel, Word). Cần cài đầy đủ (không phải Online version).

---

## Build & Run

### Bước 1 — Clone / mở solution

```
MOS.sln
```

Mở trong **Visual Studio** (bắt buộc cho ExcelAddIn).

### Bước 2 — Build CommonLibrary + MainApp

Dùng dotnet CLI hoặc Rider:

```bash
dotnet build MainApp/MainApp.csproj -c Debug
```

Hoặc trong VS: `Build → Build Solution` (`Ctrl+Shift+B`)

### Bước 3 — Build ExcelAddIn (cần Visual Studio)

1. Mở `MOS.sln` trong **Visual Studio**
2. `Build → Build Solution` (`Ctrl+Shift+B`)
3. Kiểm tra Output window — phải thấy `ExcelAddIn.dll` được tạo

### Bước 4 — Debug ExcelAddIn

1. Right-click **ExcelAddIn** trong Solution Explorer → **Set as Startup Project**
2. Nhấn **F5**
3. Visual Studio tự mở Excel với add-in loaded
4. Task Pane "MOS Excel Simulator" xuất hiện ở dưới cửa sổ Excel

### Bước 5 — Chạy MainApp

```bash
# Chạy trực tiếp
MainApp\bin\Debug\net8.0-windows\MainApp.exe
```

Hoặc trong Rider: Set MainApp làm startup → Run.

---

## Test end-to-end

1. **Debug ExcelAddIn** trong VS (F5) → Excel mở với Task Pane
2. **Chạy MainApp** (exe hoặc Rider)
3. Thao tác trong Excel:
   - Nhập số vào cell → `SheetChange` event
   - Mở workbook mới → `WorkbookOpen` event
4. Quan sát:
   - **Task Pane** trong Excel: hiện log real-time
   - **Event Log** trong MainApp: nhận qua Named Pipe

---

## Events được bắt

### Excel Add-in
| Event | Mô tả |
|---|---|
| `AddInLoad` | Add-in khởi động thành công |
| `SheetChange` | Cell bị thay đổi (kèm tên sheet + địa chỉ cell) |
| `WorkbookOpen` | Workbook được mở |
| `WorkbookClose` | Workbook đóng |
| `WorkbookActivate` | Chuyển workbook active |

### Word Add-in *(code sẵn trong `WordAddIn_old/`)*
| Event | Mô tả |
|---|---|
| `AddInLoad` | Add-in khởi động |
| `DocumentOpen` | Document được mở |
| `DocumentClose` | Document đóng |
| `SelectionChange` | User chọn/bôi đen text |

---

## Giao tiếp Named Pipe (Phase 4)

- **Pipe name**: `MOS_Office_Pipe`
- **PipeServer**: chạy trong MainApp, lắng nghe liên tục trên background thread
- **PipeClient**: trong add-ins, gửi `PipeMessage` JSON mỗi khi có event
- **Timeout**: 300ms — nếu MainApp không chạy, add-in bỏ qua (không crash)
- **Message format**:
```json
{
  "EventType": "SheetChange",
  "Application": "Excel",
  "Detail": "Sheet: Sheet1, Cell: $A$1",
  "Timestamp": "10:30:45"
}
```

---

## Troubleshooting

| Vấn đề | Nguyên nhân | Giải pháp |
|---|---|---|
| ExcelAddIn build lỗi "VSTO targets not found" | VS chưa cài workload Office | Cài workload **Microsoft 365 development** |
| "Reference assemblies for .NETFramework v4.8 not found" | Chưa cài Developer Pack | Cài `.NET Framework 4.8 Developer Pack` |
| F5 lỗi "Unable to start debugging 0x80070057" | Excel đang chạy sẵn | Đóng Excel trước khi F5 |
| Add-in không load trong Excel | Chưa đăng ký add-in | Debug qua VS (F5) sẽ tự đăng ký tạm thời |
| MainApp không nhận events | PipeServer chưa chạy | Đảm bảo MainApp đã mở trước khi thao tác Excel |

---

## Trạng thái

| Phase | Nội dung | Trạng thái |
|---|---|---|
| 1 | Chuẩn bị môi trường | ✅ Hoàn thành |
| 2 | MainApp WPF (split-view) | ✅ Hoàn thành |
| 3 | ExcelAddIn VSTO + Task Pane + Events | ✅ Hoàn thành |
| 3 | WordAddIn VSTO | 🔧 Code sẵn, cần tạo VS project |
| 4 | Named Pipe communication | ✅ Hoàn thành |
| 5 | Excel embedded trong MainApp | ⚠️ Word OK, Excel ribbon bị ẩn khi SetParent |
