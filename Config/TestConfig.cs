namespace Sage50Automation.Config
{
    /// <summary>
    /// Centralized configuration for all test settings.
    /// ┌─────────────────────────────────────────────────────┐
    /// │  UPDATE PATHS AND TIMEOUTS HERE                     │
    /// │  instead of modifying test or page object code.     │
    /// └─────────────────────────────────────────────────────┘
    /// </summary>
    public static class TestConfig
    {
        // ===== Application Paths =====
        /// <summary>Sage 50 Actian 2026 executable path</summary>
        public static readonly string Sage2026AppPath = @"C:\Program Files (x86)\Sage\Peachtree\Peachw.exe";

        /// <summary>Sage 50 Actian 2024 executable path</summary>
        public static readonly string Sage2024AppPath = @"C:\Program Files (x86)\Sage\Peachtree2\Peachw.exe";

        // ===== Output Paths =====
        /// <summary>Folder where exported CSV reports are saved</summary>
        public static readonly string ReportsFolderPath = @"C:\Users\moinul.islam\Downloads\Reports";

        /// <summary>Project root directory for HTML report output</summary>
        public static readonly string ProjectDirectory = @"C:\Users\moinul.islam\Downloads\desktop_automation\sabir_winAppDriver Automation\winAppDriver Automation";

        /// <summary>Temp file name for cross-process log sharing</summary>
        public static readonly string ExecutionLogsFileName = "sage50_execution_logs.txt";

        // ===== Company Settings =====
        /// <summary>Window title of the opened company</summary>
        public static readonly string CompanyWindowName = "* Bellwether Garden Supply - Sage 50 Accounting";

        // ===== Timeouts (milliseconds) =====
        /// <summary>Wait after launching the Sage 50 application</summary>
        public static readonly int AppLaunchWaitMs = 10000;

        /// <summary>Timeout for finding the main window (seconds)</summary>
        public static readonly int MainWindowTimeoutSeconds = 10;

        /// <summary>Wait after opening a company file</summary>
        public static readonly int CompanyOpenWaitMs = 20000;

        /// <summary>Short pause for UI elements to appear (500ms)</summary>
        public static readonly int ShortWaitMs = 500;

        /// <summary>Medium pause for dialogs/menus (2000ms)</summary>
        public static readonly int MediumWaitMs = 2000;

        /// <summary>Long pause for heavy operations (5000ms)</summary>
        public static readonly int LongWaitMs = 5000;

        // ===== Report Settings =====
        /// <summary>Base file name for HTML comparison reports</summary>
        public static readonly string HtmlReportBaseFileName = "reportAnalysisResult";

        /// <summary>Max number of values to select in "One or More" filter options</summary>
        public static readonly int MaxFilterValuesToSelect = 3;

        /// <summary>Max wait time for Excel export progress (seconds)</summary>
        public static readonly int ExportMaxWaitSeconds = 120;

        /// <summary>JSON file name for sharing filter selections between Act26 and Act24</summary>
        public static readonly string SelectionsFileName = "filter_selections.json";

        /// <summary>File name for the execution log HTML report</summary>
        public static readonly string ExecutionLogHtmlFileName = "executionLog.html";

        /// <summary>Temp folder name for failure screenshots (stored in %TEMP%)</summary>
        public static readonly string ScreenshotsTempFolder = "sage50_screenshots";
    }
}
