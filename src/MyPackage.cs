using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using AsyncToolWindowSample.Services;
using AsyncToolWindowSample.ToolWindows;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace AsyncToolWindowSample
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("Async Tool Window Sample",
        "Shows how to use an Async Tool Window in Visual Studio 15.6+", "1.0")]
    [ProvideToolWindow(typeof(SampleToolWindow),
        Style = VsDockStyle.Tabbed, DockedWidth = 300,
        Window = "DocumentWell", Orientation = ToolWindowOrientation.Left)]
    [Guid(PackageGuids.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    // §10: register the Options page under Tools › Options
    [ProvideOptionPage(typeof(SampleOptionsPage),
        "Async Tool Window Sample", "General",
        categoryResourceID: 0, pageNameResourceID: 0,
        supportsAutomation: true)]
    public sealed class MyPackage : AsyncPackage
    {
        public OutputWindowService OutputWindow { get; private set; }
        public StatusBarService    StatusBar    { get; private set; }
        public SelectionService    Selection    { get; private set; }
        public DocumentService     Document     { get; private set; }
        public ProjectService      Project      { get; private set; }
        public EventService        Events       { get; private set; }
        public OptionsService      Options      { get; private set; }
        public MenuService         Menu         { get; private set; }
        public ToolbarService      Toolbar      { get; private set; }

        protected override async Task InitializeAsync(
            CancellationToken cancellationToken,
            IProgress<ServiceProgressData> progress)
        {
            // ── Construct all services (background thread) ────────────────
            OutputWindow = new OutputWindowService(this);
            StatusBar    = new StatusBarService(this);
            Selection    = new SelectionService(this);
            Document     = new DocumentService(this);
            Project      = new ProjectService(this);
            Options      = new OptionsService(this);

            // EventService needs OutputWindow already constructed
            Events  = new EventService(this, OutputWindow);

            // ToolbarService needs OutputWindow
            Toolbar = new ToolbarService(this, OutputWindow);

            // MenuService registers OleMenuCommands — must be called later on UI thread
            Menu = new MenuService(this);

            // ── Async init (may still be on background thread) ────────────
            await OutputWindow.InitializeAsync();
            await StatusBar.InitializeAsync();
            await Selection.InitializeAsync();

            // ── Switch to UI thread for everything that needs COM ─────────
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // Show Tool Window command (existing)
            await ShowToolWindow.InitializeAsync(this);

            // §6: register the three dynamic OleMenuCommands
            await Menu.InitializeAsync();

            OutputWindow.Log("AsyncToolWindowSample loaded successfully.");
            StatusBar.SetText("Async Tool Window Sample loaded.");
        }

        public override IVsAsyncToolWindowFactory GetAsyncToolWindowFactory(Guid toolWindowType)
        {
            return toolWindowType.Equals(Guid.Parse(SampleToolWindow.WindowGuidString))
                ? this
                : null;
        }

        protected override string GetToolWindowTitle(Type toolWindowType, int id)
        {
            return toolWindowType == typeof(SampleToolWindow)
                ? SampleToolWindow.Title
                : base.GetToolWindowTitle(toolWindowType, id);
        }

        protected override async Task<object> InitializeToolWindowAsync(
            Type toolWindowType, int id, CancellationToken cancellationToken)
        {
            var dte = await GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE80.DTE2;

            return new SampleToolWindowState
            {
                DTE          = dte,
                OutputWindow = OutputWindow,
                StatusBar    = StatusBar,
                Selection    = Selection,
                Document     = Document,
                Project      = Project,
                Events       = Events,
                Options      = Options,
                Menu         = Menu,
                Toolbar      = Toolbar
            };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Events?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
