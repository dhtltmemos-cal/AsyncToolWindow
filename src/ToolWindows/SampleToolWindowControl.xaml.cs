using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using AsyncToolWindowSample.Services;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace AsyncToolWindowSample.ToolWindows
{
    public partial class SampleToolWindowControl : UserControl
    {
        private readonly SampleToolWindowState _state;
        private OutputWindowService OutputWindow => _state.OutputWindow;
        private StatusBarService    StatusBar    => _state.StatusBar;
        private SelectionService    Selection    => _state.Selection;
        private DocumentService     Document     => _state.Document;
        private ProjectService      Project      => _state.Project;
        private EventService        Events       => _state.Events;
        private OptionsService      Options      => _state.Options;
        private MenuService         Menu         => _state.Menu;
        private ToolbarService      Toolbar      => _state.Toolbar;

        public SampleToolWindowControl(SampleToolWindowState state)
        {
            _state = state;
            InitializeComponent();
        }

        // ------------------------------------------------------------------ //
        //  Original                                                            //
        // ------------------------------------------------------------------ //

        private void Button_ShowVsLocation_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string location = _state.DTE?.FullName ?? "(DTE not available)";
            MessageBox.Show($"Visual Studio is located here:\n'{location}'",
                "VS Location", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // ------------------------------------------------------------------ //
        //  Output Window                                                       //
        // ------------------------------------------------------------------ //

        private void Button_WriteOutput_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            OutputWindow.Activate();
            OutputWindow.Log($"Button clicked – VS path: {_state.DTE?.FullName ?? "N/A"}");
            OutputWindow.WriteLine("You can write arbitrary text here.");
            StatusBar.SetText("Written to Output Window.");
        }

        private void Button_ClearOutput_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            OutputWindow.Clear();
            OutputWindow.Log("Output pane cleared.");
            StatusBar.SetText("Output pane cleared.");
        }

        // ------------------------------------------------------------------ //
        //  Status Bar                                                          //
        // ------------------------------------------------------------------ //

        private void Button_SetStatus_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            StatusBar.SetText($"Hello from Async Tool Window Sample – {DateTime.Now:T}");
            OutputWindow.Log("Status bar text updated.");
        }

        private void Button_Animate_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                OutputWindow.Log("Starting 3-second animation…");
                await StatusBar.RunWithAnimationAsync(
                    async () => await Task.Delay(3000), "Processing… please wait");
                OutputWindow.Log("Animation finished.");
            });
        }

        private void Button_Progress_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                OutputWindow.Log("Starting progress bar demo…");
                uint cookie = 0;
                const uint total = 5;
                for (uint i = 1; i <= total; i++)
                {
                    StatusBar.ReportProgress(ref cookie, "Demo progress", i, total);
                    OutputWindow.Log($"  Step {i}/{total}");
                    await Task.Delay(600);
                }
                StatusBar.ClearProgress(ref cookie);
                StatusBar.SetText("Progress complete.");
                OutputWindow.Log("Progress bar demo finished.");
            });
        }

        // ================================================================== //
        //  SELECTION TIER 1 — DTE                                             //
        // ================================================================== //

        private void Button_DteCaretInfo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var info = Selection.GetDteCaretInfo();
            if (info == null) { LogNoDoc("[DTE]"); return; }
            OutputWindow.Activate();
            OutputWindow.Log($"[DTE Caret] Line={info.Line}, Col={info.Column}");
            OutputWindow.Log($"  Anchor  : Line={info.AnchorLine}, Col={info.AnchorDisplayColumn}, AbsOffset={info.AnchorAbsOffset}");
            OutputWindow.Log($"  Active  : Line={info.ActiveLine}, AbsOffset={info.ActiveAbsOffset}");
            OutputWindow.Log($"  TopLine={info.TopLine}, BottomLine={info.BottomLine}");
            OutputWindow.Log($"  Mode={info.Mode}, IsEmpty={info.IsEmpty}");
            if (!info.IsEmpty) OutputWindow.Log($"  Selected: \"{Truncate(info.SelectedText, 80)}\"");
            StatusBar.SetText($"DTE Caret – Line {info.Line}, Col {info.Column}");
        }

        private void Button_DteSelectLine_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Selection.SelectCurrentLine();
            var info = Selection.GetDteCaretInfo();
            string msg = info != null ? $"[DTE] Selected line {info.AnchorLine}." : "[DTE] SelectLine called.";
            OutputWindow.Log(msg); StatusBar.SetText(msg);
        }

        private void Button_DteFindTodo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool found = Selection.FindText("TODO", matchCase: false);
            string msg = found ? "[DTE] Found 'TODO'." : "[DTE] 'TODO' not found.";
            OutputWindow.Log(msg); StatusBar.SetText(msg);
        }

        private void Button_DteCollapse_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Selection.CollapseSelection();
            OutputWindow.Log("[DTE] Selection collapsed."); StatusBar.SetText("DTE selection collapsed.");
        }

        // ================================================================== //
        //  SELECTION TIER 2 — MEF                                             //
        // ================================================================== //

        private void Button_MefCaretInfo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var info = Selection.GetMefCaretInfo();
            if (info == null) { OutputWindow.Log("[MEF] No focused editor pane."); StatusBar.SetText("No focused editor pane."); return; }
            OutputWindow.Activate();
            OutputWindow.Log($"[MEF Caret] Offset={info.Offset0Based}  Line={info.LineNumber0Based}  Col={info.Column0Based}");
            OutputWindow.Log($"  Buffer: {info.TotalChars} chars, {info.TotalLines} lines  ContentType={info.ContentType}");
            StatusBar.SetText($"MEF Caret – Line {info.LineNumber0Based + 1}, Col {info.Column0Based}");
        }

        private void Button_MefSelectedSpans_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var spans = Selection.GetMefSelectedSpans();
            if (spans.Count == 0) { OutputWindow.Log("[MEF] No selection."); StatusBar.SetText("No selection."); return; }
            OutputWindow.Activate();
            OutputWindow.Log($"[MEF] {spans.Count} span(s):");
            for (int i = 0; i < spans.Count; i++)
            {
                var s = spans[i];
                OutputWindow.Log($"  [{i}] {s.Start}–{s.End} Len={s.Length} Lines {s.StartLine}–{s.EndLine}");
                OutputWindow.Log($"       \"{Truncate(s.Text, 60)}\"");
            }
            StatusBar.SetText($"MEF: {spans.Count} span(s).");
        }

        private void Button_MefInsert_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            const string ins = "/* inserted by AsyncToolWindowSample */ ";
            Selection.InsertAtCaret(ins);
            OutputWindow.Log($"[MEF] Inserted: \"{ins}\""); StatusBar.SetText("MEF: inserted.");
        }

        private void Button_MefReplace_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var spans = Selection.GetMefSelectedSpans();
            if (spans.Count == 0 || spans[0].Length == 0) { OutputWindow.Log("[MEF] Select text first."); return; }
            string orig = spans[0].Text;
            string rep  = $"/* replaced: {orig} */";
            Selection.ReplaceSelection(rep);
            OutputWindow.Log($"[MEF] Replaced \"{Truncate(orig, 40)}\""); StatusBar.SetText("MEF: replaced.");
        }

        private void Button_MefBufferCount_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string text = Selection.GetBufferText();
            if (text == null) { OutputWindow.Log("[MEF] No focused editor."); return; }
            string msg = $"[MEF] {text.Length} chars, {text.Split('\n').Length} lines.";
            OutputWindow.Log(msg); StatusBar.SetText(msg);
        }

        // ================================================================== //
        //  DOCUMENT & FILE APIs                                               //
        // ================================================================== //

        private void Button_DocInfo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var info = Document.GetActiveDocumentInfo();
            if (info == null) { LogNoDoc("[Doc]"); return; }
            OutputWindow.Activate();
            OutputWindow.Log($"[Doc] Name={info.Name}  Language={info.Language}  Saved={info.Saved}  RO={info.ReadOnly}");
            OutputWindow.Log($"[Doc] FullName={info.FullName}");
            if (info.ProjectName != null) OutputWindow.Log($"[Doc] Project={info.ProjectName}");
            StatusBar.SetText($"Doc: {info.Name} | {info.Language}");
        }

        private void Button_DocListAll_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var docs = Document.GetAllOpenDocuments();
            OutputWindow.Activate();
            OutputWindow.Log($"[Doc] {docs.Count} open document(s):");
            foreach (var d in docs)
                OutputWindow.Log($"  [{(d.Saved ? "✓" : "*")}] {d.Language,-12} {d.Name}");
            StatusBar.SetText($"Doc: {docs.Count} open.");
        }

        private void Button_TextDocInfo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var info = Document.GetTextDocumentInfo();
            if (info == null) { OutputWindow.Log("[TextDoc] Not a text document."); return; }
            OutputWindow.Log($"[TextDoc] Lines={info.LastLine}  Chars={info.TotalChars}");
            OutputWindow.Log($"  Preview: \"{Truncate(info.Preview, 160)}\"");
            StatusBar.SetText($"TextDoc: {info.LastLine} lines.");
        }

        private void Button_DocReadLines_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string lines = Document.ReadLines(1, 6);
            if (lines == null) { LogNoDoc("[TextDoc]"); return; }
            OutputWindow.Log("[TextDoc] Lines 1–5:"); OutputWindow.Log(Truncate(lines, 400));
            StatusBar.SetText("Lines 1–5 read.");
        }

        private void Button_DocSave_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool ok = Document.SaveActiveDocument();
            string msg = ok ? "[Doc] Saved." : "[Doc] No active document.";
            OutputWindow.Log(msg); StatusBar.SetText(msg);
        }

        private void Button_DocFormat_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var info = Document.GetActiveDocumentInfo();
            if (info == null) { LogNoDoc("[Doc]"); return; }
            Document.FormatDocument();
            string msg = $"[Doc] FormatDocument on {info.Name}.";
            OutputWindow.Log(msg); StatusBar.SetText(msg);
        }

        private void Button_DocSaveAll_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Document.SaveAll();
            OutputWindow.Log("[Doc] SaveAll."); StatusBar.SetText("All saved.");
        }

        private void Button_DocGoToLine_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Document.GoToLine(1);
            OutputWindow.Log("[Doc] GoToLine 1."); StatusBar.SetText("Line 1.");
        }

        // ================================================================== //
        //  §5 PROJECT & SOLUTION                                              //
        // ================================================================== //

        private void Button_SolInfo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var info = Project.GetSolutionInfo();
            if (info == null) { OutputWindow.Log("[Solution] No solution open."); return; }
            OutputWindow.Activate();
            OutputWindow.Log($"[Solution] {info.FullName}");
            OutputWindow.Log($"  IsDirty={info.IsDirty}  Projects={info.ProjectCount}  Config={info.ActiveConfig}");
            OutputWindow.Log($"  LastBuild={( info.LastBuildInfo == 0 ? "SUCCESS" : info.LastBuildInfo == -1 ? "(n/a)" : $"FAILED ({info.LastBuildInfo})" )}");
            StatusBar.SetText($"Solution: {info.ProjectCount} project(s)");
        }

        private void Button_ProjList_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var projects = Project.GetAllProjects();
            OutputWindow.Activate();
            OutputWindow.Log($"[Projects] {projects.Count} project(s):");
            foreach (var p in projects)
                OutputWindow.Log($"  • {p.Name}  [{p.AssemblyName ?? "?"}]  {p.TargetFw ?? ""}  {p.OutputType ?? ""}");
            StatusBar.SetText($"Projects: {projects.Count}.");
        }

        private void Button_ProjActiveDoc_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var info = Project.GetActiveDocumentProject();
            if (info == null) { OutputWindow.Log("[Project] No project found for active doc."); return; }
            OutputWindow.Activate();
            OutputWindow.Log($"[Project] {info.Name}  Assembly={info.AssemblyName}  FW={info.TargetFw}");
            StatusBar.SetText($"Project: {info.Name}");
        }

        private void Button_ProjRefs_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var refs = Project.GetReferencesOfActiveProject();
            if (refs.Count == 0) { OutputWindow.Log("[Refs] No references (or not C#/VB project)."); return; }
            OutputWindow.Activate();
            OutputWindow.Log($"[Refs] {refs.Count} reference(s):");
            foreach (var r in refs)
                OutputWindow.Log($"  {r.Name,-40} v{r.Version}  [{r.Type}]");
            StatusBar.SetText($"Refs: {refs.Count}.");
        }

        private void Button_ProjFiles_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var items = Project.GetProjectItemsOfActiveProject();
            if (items.Count == 0) { OutputWindow.Log("[Files] No items."); return; }
            OutputWindow.Activate();
            OutputWindow.Log($"[Files] {items.Count} item(s):");
            int n = 0;
            foreach (var item in items)
            {
                if (n++ > 50) { OutputWindow.Log("  … (truncated)"); break; }
                string indent = new string(' ', item.Depth * 2);
                OutputWindow.Log($"  {indent}{(item.IsFolder ? "[DIR]" : "[FILE]")} {item.Name}  BA={item.BuildAction ?? "-"}");
            }
            StatusBar.SetText($"Files: {items.Count}.");
        }

        private void Button_ProjConfig_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var cfg = Project.GetActiveBuildConfig();
            if (cfg == null) { OutputWindow.Log("[Config] Not available."); return; }
            OutputWindow.Activate();
            OutputWindow.Log($"[Config] {cfg.ConfigName}|{cfg.Platform}  Build={cfg.IsBuildable}  Out={cfg.OutputPath}  Opt={cfg.Optimize}");
            StatusBar.SetText($"Config: {cfg.ConfigName}|{cfg.Platform}");
        }

        private void Button_SolBuild_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var info = Project.GetSolutionInfo();
            if (info == null || !info.IsOpen) { OutputWindow.Log("[Build] No solution open."); return; }
            OutputWindow.Log("[Build] Triggered."); StatusBar.SetText("Building…");
            Project.BuildSolution(waitForFinish: false);
        }

        // ================================================================== //
        //  §6 MENU & COMMAND                                                  //
        // ================================================================== //

        private void Button_MenuCmdState_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var snap = Menu.GetSnapshot();
            OutputWindow.Activate();
            OutputWindow.Log("[Menu] Dynamic command state:");
            OutputWindow.Log($"  ToggleChecked    : {snap.ToggleChecked}");
            OutputWindow.Log($"  DocActionLabel   : {snap.DocActionLabel}");
            OutputWindow.Log($"  ContextInfoEnabled: {snap.ContextInfoEnabled}");
            StatusBar.SetText($"Menu: Toggle={snap.ToggleChecked}");
        }

        private void Button_MenuToggle_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Menu.FireToggleCommand();
            var snap = Menu.GetSnapshot();
            string msg = $"[Menu] Toggle fired → Checked={snap.ToggleChecked}";
            OutputWindow.Log(msg);
            StatusBar.SetText(msg);
        }

        private void Button_MenuDocAction_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Directly call execute logic (same as clicking menu item)
            string err1 = Menu.ExecuteVsCommand("AsyncToolWindowSample.DocAction");
            if (err1 != null) OutputWindow.Log($"[Menu] DocAction: {err1}");
            else OutputWindow.Log("[Menu] DocAction command fired.");
            StatusBar.SetText(err1 == null ? "Menu: DocAction fired." : $"Menu error: {err1}");
        }

        private void Button_MenuContextInfo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string err2 = Menu.ExecuteVsCommand("AsyncToolWindowSample.ContextInfo");
            if (err2 != null) OutputWindow.Log($"[Menu] ContextInfo: {err2}");
            else OutputWindow.Log("[Menu] ContextInfo command fired.");
            StatusBar.SetText(err2 == null ? "Menu: ContextInfo fired." : $"Menu error: {err2}");
        }

        private void Button_MenuVsCmd_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var docInfo = Document.GetActiveDocumentInfo();
            if (docInfo == null) { LogNoDoc("[Menu]"); return; }
            string fmtErr = Menu.ExecuteVsCommand("Edit.FormatDocument");
            string msg = fmtErr == null
                ? $"[Menu] Executed Edit.FormatDocument on {docInfo.Name}"
                : $"[Menu] FormatDocument error: {fmtErr}";
            OutputWindow.Log(msg);
            StatusBar.SetText(msg);
        }

        // ================================================================== //
        //  §7 TOOLBAR & COMMANDBAR                                            //
        // ================================================================== //

        private void Button_BarList_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var bars = Toolbar.GetAllCommandBars();
            OutputWindow.Activate();
            OutputWindow.Log($"[Toolbar] {bars.Count} CommandBar(s) found:");
            int shown = 0;
            foreach (var bar in bars)
            {
                if (shown++ > 40) { OutputWindow.Log("  … (truncated)"); break; }
                string vis = bar.IsVisible ? "visible" : "hidden";
                OutputWindow.Log($"  [{vis}] [{bar.Type,-14}] {bar.Name}  ({bar.ControlCount} controls)  pos={bar.Position}");
            }
            StatusBar.SetText($"CommandBars: {bars.Count} total.");
        }

        private void Button_BarStandard_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var info = Toolbar.GetStandardBarInfo("Standard");
            if (info == null) { OutputWindow.Log("[Toolbar] 'Standard' bar not found."); return; }
            OutputWindow.Activate();
            OutputWindow.Log($"[Toolbar] Standard bar: Type={info.Type}  Visible={info.IsVisible}  Controls={info.ControlCount}  Pos={info.Position}");
            StatusBar.SetText("Standard toolbar info shown.");
        }

        private void Button_BarCreate_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Toolbar.CreateCustomToolbar();
            StatusBar.SetText("Custom toolbar created/shown.");
        }

        private void Button_BarShow_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Toolbar.ShowCustomToolbar();
            StatusBar.SetText("Custom toolbar shown.");
        }

        private void Button_BarHide_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Toolbar.HideCustomToolbar();
            StatusBar.SetText("Custom toolbar hidden.");
        }

        private void Button_BarDelete_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Toolbar.DeleteCustomToolbar();
            StatusBar.SetText("Custom toolbar deleted.");
        }

        private void Button_BarStatus_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool vis = Toolbar.IsCustomToolbarVisible();
            string msg = $"[Toolbar] Custom toolbar: {(vis ? "VISIBLE" : "HIDDEN / NOT CREATED")}";
            OutputWindow.Log(msg);
            StatusBar.SetText(msg);
        }

        // ================================================================== //
        //  §9 DTE EVENTS                                                      //
        // ================================================================== //

        private void Button_EventSubscribe_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (Events.IsSubscribed) { OutputWindow.Log("[Events] Already subscribed."); return; }
            Events.Subscribe(); StatusBar.SetText("Events: subscribed.");
        }

        private void Button_EventUnsubscribe_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!Events.IsSubscribed) { OutputWindow.Log("[Events] Not subscribed."); return; }
            Events.Unsubscribe(); StatusBar.SetText("Events: unsubscribed.");
        }

        private void Button_EventStatus_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string msg = $"[Events] Status: {(Events.IsSubscribed ? "ACTIVE" : "INACTIVE")}";
            OutputWindow.Log(msg); StatusBar.SetText(msg);
        }

        // ================================================================== //
        //  §10 SETTINGS / OPTIONS                                             //
        // ================================================================== //

        private void Button_OptShow_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            OutputWindow.Activate();
            OutputWindow.Log("[Options] Current settings:");
            OutputWindow.Log($"  ServerURL          : {Options.ServerUrl}");
            OutputWindow.Log($"  AutoFormat         : {Options.AutoFormat}");
            OutputWindow.Log($"  MaxLogItems        : {Options.MaxLogItems}");
            OutputWindow.Log($"  EnableSelectionLog : {Options.EnableSelectionLog}");
            StatusBar.SetText("Options shown.");
        }

        private void Button_OptOpen_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Options.OpenOptionsDialog();
        }

        // ================================================================== //
        //  §12 TEXT ADORNMENT / TAGGER                                        //
        // ================================================================== //

        private void Button_TaggerStatus_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var view = Selection.GetActiveWpfTextView();
            if (view == null) { OutputWindow.Log("[Tagger] No active editor."); return; }
            OutputWindow.Log("[Tagger] TodoTagger ACTIVE on current buffer.");
            OutputWindow.Log($"  ContentType={view.TextBuffer.ContentType.TypeName}  Size={view.TextBuffer.CurrentSnapshot.Length} chars");
            StatusBar.SetText("Tagger: active.");
        }

        private void Button_TaggerCount_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var view = Selection.GetActiveWpfTextView();
            if (view == null) { OutputWindow.Log("[Tagger] No active editor."); return; }
            string text  = view.TextBuffer.CurrentSnapshot.GetText();
            int    count = 0, idx = 0;
            while ((idx = text.IndexOf("TODO", idx, StringComparison.OrdinalIgnoreCase)) >= 0)
            { count++; idx += 4; }
            string msg = $"[Tagger] {count} TODO token(s) in buffer.";
            OutputWindow.Log(msg); StatusBar.SetText(msg);
        }

        // ------------------------------------------------------------------ //
        //  Helpers                                                             //
        // ------------------------------------------------------------------ //

        private void LogNoDoc(string prefix)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string msg = $"{prefix} No active document.";
            OutputWindow.Log(msg); StatusBar.SetText(msg);
        }

        private static string Truncate(string s, int max)
            => s != null && s.Length > max ? s.Substring(0, max) + "…" : s ?? "";
    }
}
