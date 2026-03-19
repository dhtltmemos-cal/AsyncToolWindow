using System;
using System.Linq;
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

        public SampleToolWindowControl(SampleToolWindowState state)
        {
            _state = state;
            InitializeComponent();
        }

        // ------------------------------------------------------------------ //
        //  Original button                                                     //
        // ------------------------------------------------------------------ //

        private void Button_ShowVsLocation_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string location = _state.DTE?.FullName ?? "(DTE not available)";
            MessageBox.Show($"Visual Studio is located here:\n'{location}'",
                "VS Location", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // ------------------------------------------------------------------ //
        //  Output Window buttons                                               //
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
        //  Status Bar buttons                                                  //
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
                    async () => await Task.Delay(3000),
                    "Processing… please wait");

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
        //  SELECTION TIER 1 — DTE buttons                                     //
        // ================================================================== //

        private void Button_DteCaretInfo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var info = Selection.GetDteCaretInfo();
            if (info == null)
            {
                OutputWindow.Log("[DTE] No active text document.");
                StatusBar.SetText("No active text document.");
                return;
            }

            OutputWindow.Activate();
            OutputWindow.Log($"[DTE Caret] Line={info.Line}, Col={info.Column}");
            OutputWindow.Log($"  Anchor  : Line={info.AnchorLine}, Col={info.AnchorDisplayColumn}, AbsOffset={info.AnchorAbsOffset}");
            OutputWindow.Log($"  Active  : Line={info.ActiveLine}, AbsOffset={info.ActiveAbsOffset}");
            OutputWindow.Log($"  TopLine={info.TopLine}, BottomLine={info.BottomLine}");
            OutputWindow.Log($"  Mode={info.Mode}, IsEmpty={info.IsEmpty}");

            if (!info.IsEmpty)
                OutputWindow.Log($"  Selected: \"{Truncate(info.SelectedText, 80)}\"");

            StatusBar.SetText($"DTE Caret – Line {info.Line}, Col {info.Column}");
        }

        private void Button_DteSelectLine_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Selection.SelectCurrentLine();

            var info = Selection.GetDteCaretInfo();
            string msg = info != null
                ? $"[DTE] Selected line {info.AnchorLine}."
                : "[DTE] SelectLine called (no caret info).";

            OutputWindow.Log(msg);
            StatusBar.SetText(msg);
        }

        private void Button_DteFindTodo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            bool found = Selection.FindText("TODO", matchCase: false);
            string msg = found
                ? "[DTE] Found 'TODO' in active document."
                : "[DTE] 'TODO' not found in active document.";

            OutputWindow.Log(msg);
            StatusBar.SetText(msg);
        }

        private void Button_DteCollapse_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Selection.CollapseSelection();
            OutputWindow.Log("[DTE] Selection collapsed.");
            StatusBar.SetText("DTE selection collapsed.");
        }

        // ================================================================== //
        //  SELECTION TIER 2 — MEF / IWpfTextView buttons                      //
        // ================================================================== //

        private void Button_MefCaretInfo_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var info = Selection.GetMefCaretInfo();
            if (info == null)
            {
                OutputWindow.Log("[MEF] No focused editor pane.");
                StatusBar.SetText("No focused editor pane.");
                return;
            }

            OutputWindow.Activate();
            OutputWindow.Log($"[MEF Caret] Offset={info.Offset0Based} (0-based)");
            OutputWindow.Log($"  Line={info.LineNumber0Based} (0-based) → DTE Line={info.LineNumber0Based + 1}");
            OutputWindow.Log($"  Col={info.Column0Based} (0-based)");
            OutputWindow.Log($"  Buffer: {info.TotalChars} chars, {info.TotalLines} lines");
            OutputWindow.Log($"  ContentType: {info.ContentType}");

            StatusBar.SetText($"MEF Caret – Line {info.LineNumber0Based + 1}, Col {info.Column0Based}");
        }

        private void Button_MefSelectedSpans_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var spans = Selection.GetMefSelectedSpans();
            if (spans.Count == 0)
            {
                OutputWindow.Log("[MEF] No selection (or no focused editor).");
                StatusBar.SetText("No selection.");
                return;
            }

            OutputWindow.Activate();
            OutputWindow.Log($"[MEF] {spans.Count} selected span(s):");
            for (int i = 0; i < spans.Count; i++)
            {
                var s = spans[i];
                OutputWindow.Log($"  [{i}] Start={s.Start}, End={s.End}, Len={s.Length}");
                OutputWindow.Log($"       Lines {s.StartLine}–{s.EndLine} (0-based)");
                OutputWindow.Log($"       Text: \"{Truncate(s.Text, 60)}\"");
            }

            StatusBar.SetText($"MEF: {spans.Count} span(s) selected.");
        }

        private void Button_MefInsert_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            const string inserted = "/* inserted by AsyncToolWindowSample */ ";
            Selection.InsertAtCaret(inserted);

            OutputWindow.Log($"[MEF] Inserted text at caret: \"{inserted}\"");
            StatusBar.SetText("MEF: text inserted at caret.");
        }

        private void Button_MefReplace_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var spans = Selection.GetMefSelectedSpans();
            if (spans.Count == 0 || spans[0].Length == 0)
            {
                OutputWindow.Log("[MEF] ReplaceSelection: nothing selected.");
                StatusBar.SetText("MEF: select some text first.");
                return;
            }

            string original    = spans[0].Text;
            string replacement = $"/* replaced: {original} */";
            Selection.ReplaceSelection(replacement);

            OutputWindow.Log($"[MEF] Replaced \"{Truncate(original, 40)}\" → \"{Truncate(replacement, 60)}\"");
            StatusBar.SetText("MEF: selection replaced.");
        }

        private void Button_MefBufferCount_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string text = Selection.GetBufferText();
            if (text == null)
            {
                OutputWindow.Log("[MEF] No focused editor pane.");
                StatusBar.SetText("No focused editor pane.");
                return;
            }

            string msg = $"[MEF] Buffer: {text.Length} chars, {text.Split('\n').Length} lines.";
            OutputWindow.Log(msg);
            StatusBar.SetText(msg);
        }

        // ------------------------------------------------------------------ //
        //  Helpers                                                             //
        // ------------------------------------------------------------------ //

        private static string Truncate(string s, int max)
            => s != null && s.Length > max ? s.Substring(0, max) + "…" : s ?? "";
    }
}
