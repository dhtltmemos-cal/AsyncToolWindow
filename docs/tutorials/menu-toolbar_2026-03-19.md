# Hướng dẫn: Menu & Command (§6) và Toolbar & CommandBar (§7)

**Tính năng:** `MenuService` (§6) + `ToolbarService` (§7)  
**Ngày:** 2026-03-19  
**Phiên bản:** `feat: menu-toolbar`

---

## §6 — MenuService: Menu & Command

### Kiến trúc

```
VSCommandTable.vsct
  └─ GuidSymbol guidMyPackageCmdSet
       ├─ Group  DynamicCmdGroup  (parent = Tools menu)
       ├─ Group  ContextMenuGroup (parent = code-editor right-click)
       ├─ Button CmdIdToggleFeature  (0x0200) ← check-mark style
       ├─ Button CmdIdDocAction      (0x0201) ← label changes with active doc
       └─ Button CmdIdContextInfo    (0x0202) ← appears in both menus

PackageGuids.cs
  └─ const CommandSetGuid, PackageGuid, CmdIdToggleFeature, CmdIdDocAction, CmdIdContextInfo

MenuService.cs
  └─ InitializeAsync() ← registers 3 OleMenuCommand instances
       ├─ OnToggleQueryStatus   / OnToggleExecute
       ├─ OnDocActionQueryStatus / OnDocActionExecute
       └─ OnContextInfoQueryStatus / OnContextInfoExecute
```

### Luồng BeforeQueryStatus

VS gọi `BeforeQueryStatus` **ngay trước khi render menu item**. Đây là nơi duy nhất để:
- `cmd.Enabled = ...` — bật/tắt item
- `cmd.Visible = ...` — ẩn/hiện item
- `cmd.Checked = ...` — hiện dấu ✓
- `cmd.Text    = ...` — thay đổi label động

```csharp
private void OnDocActionQueryStatus(object sender, EventArgs e)
{
    ThreadHelper.ThrowIfNotOnUIThread();
    if (sender is OleMenuCommand cmd)
    {
        var dte     = GetDte();
        bool hasDoc = dte?.ActiveDocument != null;
        cmd.Enabled = hasDoc;
        cmd.Text    = hasDoc ? $"Doc: {dte.ActiveDocument.Name}" : "Doc Action (no document)";
    }
}
```

### 3 command được đăng ký

| ID | Tên | Behavior |
|----|-----|----------|
| `CmdIdToggleFeature` (0x0200) | Feature toggle | Check-mark, Ctrl+Shift+F9, label đổi ON/OFF |
| `CmdIdDocAction` (0x0201) | Doc action | Label thay đổi theo active document, disabled khi không có doc |
| `CmdIdContextInfo` (0x0202) | Context info | Có trong Tools menu VÀ code-editor right-click, disabled khi không có doc |

### Keyboard shortcut

```xml
<!-- VSCommandTable.vsct -->
<KeyBinding guid="guidMyPackageCmdSet" id="CmdIdToggleFeature"
            editor="guidVSStd97" key1="F9" mod1="Control Shift"/>
```
→ Ctrl+Shift+F9 fires the Toggle command.

### PackageGuids.cs

File mới tập trung tất cả GUID/ID constants, thay vì dùng magic strings rải rác:

```csharp
internal static class PackageGuids
{
    public const string PackageGuidString    = "6e3b2e95-...";
    public static readonly Guid PackageGuid  = new Guid(PackageGuidString);
    public const string CommandSetGuidString = "9cc1062b-...";
    public static readonly Guid CommandSetGuid = new Guid(CommandSetGuidString);
}
internal static class PackageIds
{
    public const int ShowToolWindowId    = 0x0100;
    public const int CmdIdToggleFeature  = 0x0200;
    public const int CmdIdDocAction      = 0x0201;
    public const int CmdIdContextInfo    = 0x0202;
}
```

### Buttons trong Tool Window

| Button | Hành động |
|--------|-----------|
| Show Dynamic Cmd State | In ToggleChecked, DocActionLabel, ContextInfoEnabled |
| Fire Toggle Command | Programmatically toggle (cùng logic với click menu) |
| Fire DocAction Command | `ExecuteVsCommand("AsyncToolWindowSample.DocAction")` |
| Fire ContextInfo Command | `ExecuteVsCommand("AsyncToolWindowSample.ContextInfo")` |
| Exec VS Cmd: FormatDoc | `Menu.ExecuteVsCommand("Edit.FormatDocument")` |

---

## §7 — ToolbarService: Toolbar & CommandBar

### API chính

```csharp
// Liệt kê tất cả CommandBars
IReadOnlyList<CommandBarInfo> bars = Toolbar.GetAllCommandBars();
// bar.Name, .Type, .IsVisible, .Position, .ControlCount

// Thông tin toolbar cụ thể
CommandBarInfo std = Toolbar.GetStandardBarInfo("Standard");

// Tạo custom toolbar với 1 demo button
Toolbar.CreateCustomToolbar();
// → Tạo toolbar tên "AsyncToolWindowSample Toolbar"
// → Thêm button "Async Sample" với FaceId=59 (Run icon)
// → Wire Click handler → log + MessageBox

// Show/Hide/Delete
Toolbar.ShowCustomToolbar();
Toolbar.HideCustomToolbar();
Toolbar.DeleteCustomToolbar();

// Kiểm tra trạng thái
bool vis = Toolbar.IsCustomToolbarVisible();
```

### Cách tạo button trong toolbar

```csharp
var btn = (CommandBarButton)_customBar.Controls.Add(
    MsoControlType.msoControlButton, Missing, Missing, Missing, false);

btn.Caption     = "Async Sample";
btn.TooltipText = "Demo button added by AsyncToolWindowSample";
btn.FaceId      = 59;  // Office icon ID (Run/play icon)
btn.Style       = MsoButtonStyle.msoButtonIconAndCaption;
btn.Enabled     = true;
btn.Click      += OnToolbarButtonClick;
```

### Lưu ý

- `Temporary = false` → toolbar tồn tại sau khi đóng/mở VS lại.
- `Temporary = true`  → toolbar biến mất khi VS restart.
- Luôn kiểm tra duplicate trước khi tạo mới (dùng `TryFindBar`).
- `CommandBars` là Office COM API — một số internal bars có thể throw khi access properties.

### Buttons trong Tool Window

| Button | Hành động |
|--------|-----------|
| List All CommandBars | In toàn bộ bars: Type, Visible, ControlCount, Position |
| Show 'Standard' Bar Info | Thông tin toolbar Standard |
| Create Custom Toolbar | Tạo "AsyncToolWindowSample Toolbar" + 1 button |
| Show Custom Toolbar | `_customBar.Visible = true` |
| Hide Custom Toolbar | `_customBar.Visible = false` |
| Delete Custom Toolbar | `_customBar.Delete()` |
| Custom Toolbar Status | VISIBLE / HIDDEN |

---

## Thread Safety

Tất cả methods của `MenuService` và `ToolbarService` đều yêu cầu UI thread:

```csharp
ThreadHelper.ThrowIfNotOnUIThread();
// hoặc
await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
```

`MenuService.InitializeAsync()` tự switch sang UI thread bên trong.

---

## Tóm tắt file thay đổi

| File | Thay đổi |
|------|----------|
| `src/PackageGuids.cs` | **Mới** – tập trung GUID/ID |
| `src/Services/MenuService.cs` | **Mới** – §6 OleMenuCommand |
| `src/Services/ToolbarService.cs` | **Mới** – §7 CommandBar |
| `src/VSCommandTable.vsct` | Thêm Groups (DynamicCmdGroup, ContextMenuGroup), 3 Button §6, KeyBinding |
| `src/MyPackage.cs` | +`Menu`, +`Toolbar`, `await Menu.InitializeAsync()`, dùng `PackageGuids.PackageGuidString` |
| `src/ToolWindows/SampleToolWindowState.cs` | +`Menu`, +`Toolbar` |
| `src/ToolWindows/SampleToolWindowControl.xaml` | +5 button §6, +7 button §7 |
| `src/ToolWindows/SampleToolWindowControl.xaml.cs` | +12 handler §6/§7 |
| `src/AsyncToolWindowSample.csproj` | +`PackageGuids.cs`, +`MenuService.cs`, +`ToolbarService.cs` |
