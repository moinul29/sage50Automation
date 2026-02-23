using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Sage50Automation.Utilities;

namespace Sage50Automation.Pages
{
    /// <summary>
    /// Base class for all Page Objects.
    /// Provides shared access to the automation session, desktop, and logger.
    /// 
    /// Every page in the application extends this class:
    ///   SageMainPage → ReportsMenuPage → ReportViewerPage → ModifyReportPage
    ///                                                     → ExcelSavePage
    /// </summary>
    public abstract class BasePage
    {
        protected readonly Application App;
        protected readonly UIA3Automation Automation;
        protected readonly Logger Log;

        protected BasePage(Application app, UIA3Automation automation, Logger logger)
        {
            App = app;
            Automation = automation;
            Log = logger;
        }

        /// <summary>
        /// Gets the desktop automation element (root of all windows)
        /// </summary>
        protected AutomationElement Desktop => Automation.GetDesktop();
    }
}
