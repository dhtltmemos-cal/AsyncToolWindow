# Changelog

All notable changes to **AsyncToolWindowSample** are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased] – 2026-03-19 (fix: CS1061-document-encoding)

### Fixed

#### `src/Services/DocumentService.cs`
- **CS1061 (×2) – `'Document' does not contain a definition for 'Encoding'`:**
  `EnvDTE.Document` (DTE 8.0) không có property `Encoding`. Property này không tồn tại
  trên interface COM `Document` ở version này.
  **Fix:** Xóa toàn bộ truy cập `doc.Encoding` và field `Encoding` khỏi `DocumentInfo` DTO.
  Encoding của file có thể đọc qua `FileInfo` hoặc `StreamReader` ngoài DTE nếu thực sự cần.

#### `src/ToolWindows/SampleToolWindowControl.xaml.cs`
- Xóa dòng log `[Doc] Encoding : {info.Encoding}` tương ứng khỏi `Button_DocInfo_Click`.

---

## [Unreleased] – 2026-03-19 (feat: document-file-apis)

### Added

#### `src/Services/DocumentService.cs` *(new file)*
- `GetActiveDocumentInfo()`, `GetAllOpenDocuments()`, `SaveActiveDocument()`,
  `SaveActiveDocumentAs()`, `CloseActiveDocument()`, `UndoActiveDocument()`,
  `OpenFile()`, `NavigateUrl()`, `GetTextDocumentInfo()`, `ReadLines()`,
  `InsertAtStart()`, `ExecuteCommand()`, `FormatDocument()`, `SaveAll()`, `GoToLine()`.
- **DTOs:** `DocumentInfo`, `TextDocumentInfo`.

#### `docs/tutorials/document-file-apis_2026-03-19.md`

### Changed
- `SampleToolWindowState.cs` – thêm `DocumentService Document`.
- `MyPackage.cs` – construct + wire `DocumentService`.
- `SampleToolWindowControl.xaml` – 8 button Document & File APIs.
- `SampleToolWindowControl.xaml.cs` – 8 handler tương ứng.
- `AsyncToolWindowSample.csproj` – thêm `DocumentService.cs`.

---

## [Unreleased] – 2026-03-19 (patch: CS0122-getservice)

### Fixed
- `SelectionService.cs` – CS0122: dùng `IServiceProvider` thay `AsyncPackage.GetService`.

---

## [Unreleased] – 2026-03-19 (patch: compiler-errors)

### Fixed
- `StatusBarService.cs` – CS0165 `frozen = 0`; CS1503 `ref ulong → ref uint`.
- `SampleToolWindowControl.xaml.cs` – `uint cookie`.

---

## [1.1] – baseline
- Async Tool Window cơ bản, AsyncPackage background load.
