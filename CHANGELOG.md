# Changelog

All notable changes to **AsyncToolWindowSample** are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased] – 2026-03-19 (feat: menu-toolbar)

### Added

#### `src/PackageGuids.cs` *(new)*
- Tập trung toàn bộ GUID/ID constants: `PackageGuids` + `PackageIds`.
- `[Guid]` trên `MyPackage` giờ dùng `PackageGuids.PackageGuidString` thay magic string.

#### `src/Services/MenuService.cs` *(new)*
- `InitializeAsync()` — đăng ký 3 `OleMenuCommand` với `BeforeQueryStatus`:
  - `CmdIdToggleFeature` (0x0200): check-mark style, label "Feature ON/OFF", Ctrl+Shift+F9.
  - `CmdIdDocAction` (0x0201): label thay đổi theo `dte.ActiveDocument.Name`, disabled khi không có doc.
  - `CmdIdContextInfo` (0x0202): disabled khi không có doc, xuất hiện cả Tools menu lẫn code-editor context menu.
- `GetSnapshot()` — DTO trạng thái 3 commands cho tool window.
- `FireToggleCommand()` — kích hoạt toggle programmatically.
- `ExecuteVsCommand()` — gọi bất kỳ VS command theo tên chuỗi.
- `GetCommandName()` — tìm tên đã đăng ký của command theo GUID+ID.
- **DTO:** `DynamicCommandsSnapshot`.

#### `src/Services/ToolbarService.cs` *(new)*
- `GetAllCommandBars()` — liệt kê tất cả CommandBars trong VS.
- `GetStandardBarInfo(name)` — thông tin một toolbar cụ thể.
- `CreateCustomToolbar()` — tạo toolbar "AsyncToolWindowSample Toolbar" với 1 demo button (FaceId=59).
- `ShowCustomToolbar()`, `HideCustomToolbar()`, `DeleteCustomToolbar()`.
- `IsCustomToolbarVisible()` — kiểm tra trạng thái.
- **DTO:** `CommandBarInfo`.

#### `src/VSCommandTable.vsct`
- Thêm `<Groups>`: `DynamicCmdGroup` (parent=Tools menu), `ContextMenuGroup` (parent=code-editor context menu).
- Thêm 3 `<Button>` entries cho §6 commands.
- `CmdIdContextInfo` xuất hiện ở cả 2 groups (dual parent).
- Thêm `<KeyBinding>`: Ctrl+Shift+F9 → `CmdIdToggleFeature`.
- Thêm `<IDSymbol>` entries cho 5 IDs mới.

#### `docs/tutorials/menu-toolbar_2026-03-19.md` *(new)*

### Changed

#### `src/MyPackage.cs`
- Dùng `PackageGuids.PackageGuidString` trong `[Guid]` attribute.
- Construct và wire `MenuService`, `ToolbarService`.
- Gọi `await Menu.InitializeAsync()` sau khi switch sang UI thread.

#### `src/ToolWindows/SampleToolWindowState.cs`
- Thêm properties: `Menu`, `Toolbar`.

#### `src/ToolWindows/SampleToolWindowControl.xaml`
- Thêm section §6 với 5 button: Show Dynamic Cmd State, Fire Toggle, Fire DocAction, Fire ContextInfo, Exec FormatDoc.
- Thêm section §7 với 7 button: List All CommandBars, Show Standard Bar Info, Create/Show/Hide/Delete Custom Toolbar, Custom Toolbar Status.

#### `src/ToolWindows/SampleToolWindowControl.xaml.cs`
- Thêm 12 handler cho §6/§7 buttons.

#### `src/AsyncToolWindowSample.csproj`
- Thêm Compile entries: `PackageGuids.cs`, `MenuService.cs`, `ToolbarService.cs`.

#### `README.md`
- Cập nhật Features table và Source map cho §6/§7.

---

## [Unreleased] – 2026-03-19 (feat: project-events-options-tagger)

### Added
- `ProjectService.cs` (§5), `EventService.cs` (§9), `OptionsService.cs` + `SampleOptionsPage` (§10), `Tagging/TodoTagger.cs` (§12).

---

## [Unreleased] – 2026-03-19 (fix: CS1061-document-encoding)

### Fixed
- `DocumentService.cs` – CS1061: xóa `doc.Encoding`.

---

## [Unreleased] – 2026-03-19 (feat: document-file-apis)

### Added
- `DocumentService.cs` (§4).

---

## [1.1] – baseline
- Async Tool Window cơ bản, AsyncPackage background load.
