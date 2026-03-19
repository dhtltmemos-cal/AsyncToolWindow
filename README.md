# Async Tool Window Sample

**Applies to Visual Studio 2017 Update 6 (v15.6) and newer**

This sample shows how to build a VS 2017 extension using the `AsyncPackage` pattern.
It covers Output Window, Status Bar, and **Selection / Caret APIs** (both DTE and MEF tiers).

Clone the repo and open in Visual Studio 2017 to run:

```
git clone https://github.com/madskristensen/AsyncToolWindowSample
```

Press **F5** — VS launches an Experimental Instance with the extension loaded.
Open the tool window via **View › Other Windows › Sample Tool Window**.

![Tool Window](art/tool-window.png)

---

## Minimum supported version

```xml
<InstallationTarget Id="Microsoft.VisualStudio.Community" Version="[15.0.27413, 16.0)" />
```

*15.0.27413* is the build number of Visual Studio 2017 Update 6.

---

## Features demonstrated

### 1. AsyncPackage infrastructure
- Background loading (`AllowsBackgroundLoading = true`)
- `GetServiceAsync` (non-blocking)
- `JoinableTaskFactory.SwitchToMainThreadAsync` for UI-thread operations
- `ProvideToolWindow` + async factory (`IVsAsyncToolWindowFactory`)

### 2. Output Window (`OutputWindowService`)
| Button | What it shows |
|--------|---------------|
| Write to Output | `WriteLine` / timestamped `Log` to a custom pane |
| Clear Output Pane | `Clear()` the pane |

### 3. Status Bar (`StatusBarService`)
| Button | What it shows |
|--------|---------------|
| Set Status Text | `SetText()` with current time |
| Animate (3 s) | `StartAnimation` / `StopAnimation` with background work |
| Show Progress Bar | 5-step `ReportProgress` / `ClearProgress` loop |

### 4. Selection APIs — Tier 1: DTE (`SelectionService`)
Simple COM-based `TextSelection` API. Positions are **1-based**.

| Button | What it shows |
|--------|---------------|
| Show Caret Info (DTE) | `CurrentLine`, `CurrentColumn`, Anchor/Active points, `Mode` |
| Select Current Line (DTE) | `SelectLine()` |
| Find 'TODO' (DTE) | `FindText("TODO")` → found/not found |
| Collapse Selection (DTE) | `Collapse()` |

### 5. Selection APIs — Tier 2: MEF / IWpfTextView (`SelectionService`)
Managed MEF API. Positions are **0-based**. Supports multi-caret and `ITextEdit` transactions.

| Button | What it shows |
|--------|---------------|
| Show Caret Info (MEF) | Offset, Line, Col, TotalChars, TotalLines, ContentType |
| Show Selected Spans (MEF) | All `SnapshotSpan` items: Start/End/Length/Text/Lines |
| Insert Text at Caret (MEF) | `ITextEdit.Insert` — inserts a comment placeholder |
| Replace Selection (MEF) | `ITextEdit.Replace` — wraps selection in a comment |
| Buffer Char Count (MEF) | `snapshot.GetText().Length` + line count |

---

## Source map

```
src/
├── MyPackage.cs                          ← AsyncPackage entry point
├── VSCommandTable.vsct                   ← Menu command definition
├── Commands/
│   └── ShowToolWindow.cs                 ← Opens the tool window
├── Services/
│   ├── OutputWindowService.cs            ← Custom Output pane wrapper
│   ├── StatusBarService.cs               ← IVsStatusbar wrapper
│   └── SelectionService.cs               ← DTE + MEF selection wrapper  ← NEW
├── ToolWindows/
│   ├── SampleToolWindow.cs               ← ToolWindowPane subclass
│   ├── SampleToolWindowControl.xaml      ← WPF UI
│   ├── SampleToolWindowControl.xaml.cs   ← Button handlers
│   └── SampleToolWindowState.cs          ← State bag passed via async factory
└── Properties/
    └── AssemblyInfo.cs
```

---

## Key concepts

### Thread safety cheat-sheet

```csharp
// Switch to UI thread before any VS COM call
await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

// Assert UI thread (throws if violated)
ThreadHelper.ThrowIfNotOnUIThread();

// Fire-and-forget async work without blocking UI
package.JoinableTaskFactory.RunAsync(async () => { ... });
```

### DTE vs MEF selection — when to use which

| Concern | DTE (Tier 1) | MEF / IWpfTextView (Tier 2) |
|---------|-------------|---------------------------|
| Offset style | 1-based | 0-based |
| Multi-caret | ✗ | ✓ |
| Box selection detection | via `sel.Mode` | via `selection.Mode` |
| Edit buffer | `EditPoint.Insert/Replace` | `ITextEdit` transaction |
| Requires MEF setup | ✗ | ✓ (`IVsEditorAdaptersFactoryService`) |

---

## Further reading

- [VSCT Schema Reference](https://docs.microsoft.com/en-us/visualstudio/extensibility/vsct-xml-schema-reference)
- [Use AsyncPackage with background load](https://docs.microsoft.com/en-us/visualstudio/extensibility/how-to-use-asyncpackage-to-load-vspackages-in-the-background)
- [IVsTextView and IWpfTextView](https://docs.microsoft.com/en-us/visualstudio/extensibility/editor-and-language-service-extensions)
- [Custom command sample](https://github.com/madskristensen/CustomCommandSample)
