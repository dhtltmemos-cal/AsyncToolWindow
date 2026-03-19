using System;
using System.ComponentModel.Design;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace AsyncToolWindowSample.Services
{
    /// <summary>
    /// Demonstrates programmatic interaction with the VS menu/command system (Section 6).
    ///
    /// The actual command registrations (GUIDs, IDs, parent groups) live in
    /// <c>VSCommandTable.vsct</c>.  This service provides:
    ///   • helpers to enable/disable/check/rename existing OleMenuCommands at runtime
    ///   • helpers to execute any VS built-in command by name
    ///   • a query API so the tool window can reflect current command state
    ///
    /// All public methods must be called on the UI thread.
    /// </summary>
    public sealed class MenuService
    {
        private readonly AsyncPackage  _package;
        private readonly IServiceProvider _serviceProvider;

        // The three dynamic commands registered by DynamicCommandsInitializer
        private OleMenuCommand _cmdToggle;
        private OleMenuCommand _cmdDocAction;
        private OleMenuCommand _cmdContextInfo;

        // ── Internal state reflected by the dynamic commands ──────────────
        private bool   _toggleChecked;
        private string _docActionLabel = "Doc Action";

        public MenuService(AsyncPackage package)
        {
            _package         = package ?? throw new ArgumentNullException(nameof(package));
            _serviceProvider = package;
        }

        // ================================================================== //
        //  Initialization — registers the three dynamic OleMenuCommands       //
        // ================================================================== //

        /// <summary>
        /// Registers dynamic OleMenuCommands.
        /// Must be called on the UI thread after <c>GetServiceAsync(IMenuCommandService)</c>.
        /// </summary>
        public async Task InitializeAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var cs = await _package.GetServiceAsync(typeof(IMenuCommandService))
                         as OleMenuCommandService;
            if (cs == null) return;

            // ── 1. Toggle command (check-mark style) ──────────────────────
            _cmdToggle = RegisterCommand(cs,
                PackageIds.CmdIdToggleFeature,
                OnToggleExecute,
                OnToggleQueryStatus);

            // ── 2. Doc-Action command (label changes with active doc) ──────
            _cmdDocAction = RegisterCommand(cs,
                PackageIds.CmdIdDocAction,
                OnDocActionExecute,
                OnDocActionQueryStatus);

            // ── 3. Context-info command (only enabled when doc is open) ───
            _cmdContextInfo = RegisterCommand(cs,
                PackageIds.CmdIdContextInfo,
                OnContextInfoExecute,
                OnContextInfoQueryStatus);
        }

        // ================================================================== //
        //  Public API — called from the tool-window buttons                   //
        // ================================================================== //

        /// <summary>Returns a snapshot of the three dynamic commands' current state.</summary>
        public DynamicCommandsSnapshot GetSnapshot()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return new DynamicCommandsSnapshot
            {
                ToggleChecked    = _toggleChecked,
                DocActionLabel   = _docActionLabel,
                ContextInfoEnabled = _cmdContextInfo?.Enabled ?? false
            };
        }

        /// <summary>
        /// Programmatically fires the Toggle command (same as clicking it in the menu).
        /// </summary>
        public void FireToggleCommand()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            OnToggleExecute(this, EventArgs.Empty);
        }

        /// <summary>
        /// Executes any VS built-in or registered command by name.
        /// Returns <c>null</c> on success, or an error message string on failure.
        /// Never throws — COM exceptions and invalid-command errors are caught internally.
        /// </summary>
        public string ExecuteVsCommand(string commandName, string args = "")
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrWhiteSpace(commandName))
                return "Command name is empty.";

            var dte = _serviceProvider.GetService(typeof(DTE)) as DTE2;
            if (dte == null)
                return "DTE not available.";

            try
            {
                dte.ExecuteCommand(commandName, args ?? string.Empty);
                return null; // success
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                return $"COMException 0x{ex.HResult:X8}: {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"{ex.GetType().Name}: {ex.Message}";
            }
        }

        /// <summary>
        /// Returns the display name of a registered VS command, or null if not found.
        /// </summary>
        public string GetCommandName(Guid cmdSetGuid, int cmdId)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var dte = _serviceProvider.GetService(typeof(DTE)) as DTE2;
            if (dte == null) return null;

            try
            {
                foreach (Command cmd in dte.Commands)
                {
                    if (cmd.Guid == cmdSetGuid.ToString("B").ToUpperInvariant() &&
                        cmd.ID   == cmdId)
                        return cmd.Name;
                }
            }
            catch { /* Commands collection may throw on certain items */ }
            return null;
        }

        // ================================================================== //
        //  Toggle command handlers                                             //
        // ================================================================== //

        private void OnToggleQueryStatus(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (sender is OleMenuCommand cmd)
            {
                cmd.Enabled = true;
                cmd.Visible = true;
                cmd.Checked = _toggleChecked;
                cmd.Text    = _toggleChecked ? "✓ Feature ON" : "Feature OFF";
            }
        }

        private void OnToggleExecute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _toggleChecked = !_toggleChecked;
            // Force VS to re-query the command status immediately
            if (_cmdToggle != null)
                _cmdToggle.Checked = _toggleChecked;
        }

        // ================================================================== //
        //  Doc-Action command handlers                                         //
        // ================================================================== //

        private void OnDocActionQueryStatus(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (sender is OleMenuCommand cmd)
            {
                var dte     = _serviceProvider.GetService(typeof(DTE)) as DTE2;
                bool hasDoc = dte?.ActiveDocument != null;
                cmd.Enabled = hasDoc;
                cmd.Visible = true;
                _docActionLabel = hasDoc
                    ? $"Doc: {dte.ActiveDocument.Name}"
                    : "Doc Action (no document)";
                cmd.Text = _docActionLabel;
            }
        }

        private void OnDocActionExecute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var dte = _serviceProvider.GetService(typeof(DTE)) as DTE2;
            string name = dte?.ActiveDocument?.Name ?? "(none)";
            System.Windows.MessageBox.Show(
                $"Doc-Action fired!\nActive document: {name}",
                "Menu Command Demo",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }

        // ================================================================== //
        //  Context-Info command handlers                                       //
        // ================================================================== //

        private void OnContextInfoQueryStatus(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (sender is OleMenuCommand cmd)
            {
                var dte     = _serviceProvider.GetService(typeof(DTE)) as DTE2;
                bool hasDoc = dte?.ActiveDocument != null;
                cmd.Enabled = hasDoc;
                cmd.Visible = true;
                cmd.Text    = "Show Context Info";
            }
        }

        private void OnContextInfoExecute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var dte = _serviceProvider.GetService(typeof(DTE)) as DTE2;
            var doc = dte?.ActiveDocument;
            if (doc == null)
            {
                System.Windows.MessageBox.Show("No active document.",
                    "Context Info", System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }
            var sel = doc.Selection as TextSelection;
            string selText = sel?.Text;
            System.Windows.MessageBox.Show(
                $"File    : {doc.Name}\n" +
                $"Language: {doc.Language}\n" +
                $"Saved   : {doc.Saved}\n" +
                $"Selected: {(string.IsNullOrEmpty(selText) ? "(none)" : $"\"{selText}\"")}",
                "Context Info",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }

        // ================================================================== //
        //  Private helper                                                      //
        // ================================================================== //

        private static OleMenuCommand RegisterCommand(
            OleMenuCommandService cs,
            int cmdId,
            EventHandler execute,
            EventHandler queryStatus)
        {
            var id  = new CommandID(PackageGuids.CommandSetGuid, cmdId);
            var cmd = new OleMenuCommand(execute, id);
            if (queryStatus != null)
                cmd.BeforeQueryStatus += queryStatus;
            cs.AddCommand(cmd);
            return cmd;
        }
    }

    // ====================================================================== //
    //  DTO                                                                     //
    // ====================================================================== //

    public sealed class DynamicCommandsSnapshot
    {
        public bool   ToggleChecked       { get; set; }
        public string DocActionLabel      { get; set; }
        public bool   ContextInfoEnabled  { get; set; }
    }
}
