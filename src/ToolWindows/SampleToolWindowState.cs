using AsyncToolWindowSample.Services;

namespace AsyncToolWindowSample.ToolWindows
{
    /// <summary>
    /// State object passed from MyPackage.InitializeToolWindowAsync
    /// into SampleToolWindow constructor. Carries DTE and all services.
    /// </summary>
    public class SampleToolWindowState
    {
        public EnvDTE80.DTE2         DTE          { get; set; }
        public OutputWindowService   OutputWindow { get; set; }
        public StatusBarService      StatusBar    { get; set; }
        public SelectionService      Selection    { get; set; }
        public DocumentService       Document     { get; set; }
        public ProjectService        Project      { get; set; }
        public EventService          Events       { get; set; }
        public OptionsService        Options      { get; set; }
        public MenuService           Menu         { get; set; }
        public ToolbarService        Toolbar      { get; set; }
        public ConfigurationService  Config       { get; set; }

        /// <summary>JSON-based configuration service (§JsonSettings).</summary>
        public JsonConfigService     JsonConfig   { get; set; }
    }
}
