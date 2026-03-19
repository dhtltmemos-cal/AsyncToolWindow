using System;
using System.Collections.Generic;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace AsyncToolWindowSample.Services
{
    /// <summary>
    /// Demonstrates CommandBar / Toolbar APIs (Section 7).
    ///
    /// Provides:
    ///   • List all toolbars visible in VS
    ///   • Create a custom toolbar at runtime
    ///   • Add a button to the custom toolbar
    ///   • Show / hide / delete the custom toolbar
    ///
    /// Uses <c>EnvDTE.CommandBars</c> (Microsoft.VisualStudio.CommandBars).
    /// All public methods must be called on the UI thread.
    /// </summary>
    public sealed class ToolbarService
    {
        private readonly AsyncPackage _package;
        private readonly IServiceProvider _serviceProvider;
        private readonly OutputWindowService _outputWindow;

        // The toolbar we create at runtime
        private CommandBar _customBar;
        private const string CustomBarName = "AsyncToolWindowSample Toolbar";

        public ToolbarService(AsyncPackage package, OutputWindowService outputWindow)
        {
            _package         = package      ?? throw new ArgumentNullException(nameof(package));
            _outputWindow    = outputWindow  ?? throw new ArgumentNullException(nameof(outputWindow));
            _serviceProvider = package;
        }

        private DTE2 GetDte()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return _serviceProvider.GetService(typeof(DTE)) as DTE2;
        }

        private CommandBars GetCommandBars()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return GetDte()?.CommandBars as CommandBars;
        }

        // ================================================================== //
        //  List toolbars                                                       //
        // ================================================================== //

        /// <summary>
        /// Returns info for every CommandBar currently registered in VS.
        /// </summary>
        public IReadOnlyList<CommandBarInfo> GetAllCommandBars()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var result = new List<CommandBarInfo>();
            var bars   = GetCommandBars();
            if (bars == null) return result;

            foreach (CommandBar bar in bars)
            {
                try
                {
                    result.Add(new CommandBarInfo
                    {
                        Name      = bar.Name,
                        Type      = bar.Type.ToString(),
                        IsVisible = bar.Visible,
                        Position  = bar.Position.ToString(),
                        ControlCount = bar.Controls.Count
                    });
                }
                catch { /* some internal bars throw on property access */ }
            }
            return result;
        }

        // ================================================================== //
        //  Custom toolbar — Create / Show / Hide / Delete                     //
        // ================================================================== //

        /// <summary>
        /// Creates (or re-uses) the extension's custom toolbar and adds one
        /// demo button to it.  The toolbar appears docked at the top.
        /// </summary>
        public void CreateCustomToolbar()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var bars = GetCommandBars();
            if (bars == null) return;

            // Avoid duplicates: try to find an existing one first
            _customBar = TryFindBar(bars, CustomBarName);
            if (_customBar != null)
            {
                _customBar.Visible = true;
                _outputWindow.Log($"[Toolbar] '{CustomBarName}' already exists – made visible.");
                return;
            }

            // Add() signature: name, position, MenuBar, Temporary
            _customBar = (CommandBar)bars.Add(
                CustomBarName,
                MsoBarPosition.msoBarTop,
                System.Type.Missing,   // MenuBar = missing (not a menu bar)
                false);                // Temporary = false → persists across sessions

            _customBar.Visible = true;

            // Add a demo button
            var btn = (CommandBarButton)_customBar.Controls.Add(
                MsoControlType.msoControlButton,
                System.Type.Missing,
                System.Type.Missing,
                System.Type.Missing,
                false); // Temporary

            btn.Caption     = "Async Sample";
            btn.TooltipText = "Demo button added by AsyncToolWindowSample";
            //btn.FaceId      = 59;  // standard "run" icon from Office icon set
            btn.Style       = MsoButtonStyle.msoButtonIconAndCaption;
            btn.Enabled     = true;
            btn.Visible     = true;

            // Wire click handler
            btn.Click += OnToolbarButtonClick;

            _outputWindow.Log($"[Toolbar] Created '{CustomBarName}' with 1 button.");
        }

        /// <summary>Shows the custom toolbar (creates it if needed).</summary>
        public void ShowCustomToolbar()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_customBar == null) CreateCustomToolbar();
            else
            {
                _customBar.Visible = true;
                _outputWindow.Log($"[Toolbar] '{CustomBarName}' shown.");
            }
        }

        /// <summary>Hides the custom toolbar without deleting it.</summary>
        public void HideCustomToolbar()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_customBar == null)
            {
                _outputWindow.Log("[Toolbar] Custom toolbar does not exist yet.");
                return;
            }
            _customBar.Visible = false;
            _outputWindow.Log($"[Toolbar] '{CustomBarName}' hidden.");
        }

        /// <summary>Permanently deletes the custom toolbar.</summary>
        public void DeleteCustomToolbar()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_customBar == null)
            {
                // Try to find and delete a left-over bar from a previous session
                var bars = GetCommandBars();
                var found = bars != null ? TryFindBar(bars, CustomBarName) : null;
                if (found == null)
                {
                    _outputWindow.Log("[Toolbar] Custom toolbar not found.");
                    return;
                }
                found.Delete();
                _outputWindow.Log($"[Toolbar] '{CustomBarName}' deleted.");
                return;
            }

            try   { _customBar.Delete(); }
            catch { /* already gone */ }
            _customBar = null;
            _outputWindow.Log($"[Toolbar] '{CustomBarName}' deleted.");
        }

        /// <summary>Returns whether the custom toolbar currently exists and is visible.</summary>
        public bool IsCustomToolbarVisible()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_customBar == null)
            {
                var bars = GetCommandBars();
                var found = bars != null ? TryFindBar(bars, CustomBarName) : null;
                return found?.Visible ?? false;
            }
            return _customBar.Visible;
        }

        // ================================================================== //
        //  Standard toolbar helpers                                            //
        // ================================================================== //

        /// <summary>
        /// Returns info about a named standard toolbar (e.g. "Standard", "Debug").
        /// Returns <c>null</c> when not found.
        /// </summary>
        public CommandBarInfo GetStandardBarInfo(string barName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var bars = GetCommandBars();
            if (bars == null) return null;
            var bar = TryFindBar(bars, barName);
            if (bar == null) return null;
            return new CommandBarInfo
            {
                Name         = bar.Name,
                Type         = bar.Type.ToString(),
                IsVisible    = bar.Visible,
                Position     = bar.Position.ToString(),
                ControlCount = bar.Controls.Count
            };
        }

        // ================================================================== //
        //  Private helpers                                                     //
        // ================================================================== //

        private static CommandBar TryFindBar(CommandBars bars, string name)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                foreach (CommandBar bar in bars)
                {
                    try { if (bar.Name == name) return bar; }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        private void OnToolbarButtonClick(CommandBarButton ctrl, ref bool cancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _outputWindow.Log("[Toolbar] Demo button clicked!");
            System.Windows.MessageBox.Show(
                "Toolbar button clicked!\nThis button was added programmatically by AsyncToolWindowSample.",
                "Toolbar Demo",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }
    }

    // ====================================================================== //
    //  DTO                                                                     //
    // ====================================================================== //

    public sealed class CommandBarInfo
    {
        public string Name         { get; set; }
        public string Type         { get; set; }
        public bool   IsVisible    { get; set; }
        public string Position     { get; set; }
        public int    ControlCount { get; set; }
    }
}
