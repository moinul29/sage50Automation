using Microsoft.VisualStudio.TestTools.UnitTesting;
using FlaUI.Core;
using FlaUI.Core.Capturing;
using FlaUI.UIA3;
using Sage50Automation.Config;
using Sage50Automation.Data;
using Sage50Automation.Pages;
using Sage50Automation.Utilities;

// Run tests sequentially (1 worker) to prevent parallel Sage 50 launches
[assembly: Parallelize(Workers = 1, Scope = ExecutionScope.ClassLevel)]

namespace Sage50Automation.Tests
{
    /// <summary>
    /// Base test class providing shared setup, teardown, and page objects.
    /// 
    /// ┌──────────────────────────────────────────────────────────────────┐
    /// │  ARCHITECTURE OVERVIEW (Page Object Model)                        │
    /// │                                                                  │
    /// │  Tests/                                                          │
    /// │    BaseTest.cs ──────── Shared setup, logger, page creation      │
    /// │    Actian2026...cs ──── First test (creates act01.csv)           │
    /// │    Actian2024...cs ──── Second test (creates act02.csv)          │
    /// │                                                                  │
    /// │  Pages/                                                          │
    /// │    SageMainPage ─────── Launch app, open company, close app      │
    /// │    ReportsMenuPage ──── Navigate: Reports > [Report Group] > [Report]   │
    /// │    ReportViewerPage ─── Report window: options, export, filters  │
    /// │    ModifyReportPage ─── Filter dialog: discover, select, apply   │
    /// │    ExcelSavePage ────── Excel: save as CSV                       │
    /// │                                                                  │
    /// │  Data/                                                           │
    /// │    ReportTestData ───── Parameterized: report group, report, index, csv  │
    /// │                                                                  │
    /// │  Config/                                                         │
    /// │    TestConfig ────────── All paths, timeouts, settings           │
    /// │                                                                  │
    /// │  Utilities/                                                      │
    /// │    Logger ────────────── Console + file logging                  │
    /// │    CsvComparer ──────── CSV comparison + report generation       │
    /// │    HtmlReportGenerator ─ Markdown → HTML conversion              │
    /// │                                                                  │
    /// │  Models/                                                         │
    /// │    FilterModels ──────── FilterInfo, OptionInfo, etc.            │
    /// └──────────────────────────────────────────────────────────────────┘
    /// 
    /// To run tests:
    ///   dotnet test --filter "FullyQualifiedName~Test_1_Actian2026" -v d
    ///   dotnet test --filter "FullyQualifiedName~Test_2_Actian2024" -v d
    /// 
    /// Or both sequentially:
    ///   dotnet test --filter "FullyQualifiedName~Test_1_Actian2026" -v d; if ($?) { dotnet test --filter "FullyQualifiedName~Test_2_Actian2024" -v d }
    /// </summary>
    public abstract class BaseTest
    {
        // ===== Core Objects =====
        protected Application App = null!;
        protected UIA3Automation Automation = null!;
        protected Logger Log = null!;

        // ===== Page Objects =====
        protected SageMainPage MainPage = null!;
        protected ReportsMenuPage MenuPage = null!;
        protected ReportViewerPage? ViewerPage;
        protected ExcelSavePage? ExcelPage;

        // ===== Utilities =====
        protected CsvComparer Comparer = null!;
        protected HtmlReportGenerator ReportGenerator = null!;

        /// <summary>MSTest context — used to detect test outcome for failure screenshots</summary>
        public TestContext TestContext { get; set; } = null!;

        /// <summary>
        /// Initialize the test session: create logger, launch app, create page objects.
        /// Called by derived class [TestInitialize] methods.
        /// </summary>
        protected void InitializeSession(string appPath, bool clearLogs = false)
        {
            // Create logger
            string logFilePath = Path.Combine(Path.GetTempPath(), TestConfig.ExecutionLogsFileName);
            Log = new Logger(logFilePath);

            if (clearLogs)
                Log.Clear();

            // Launch application
            Log.Info($"Launching Sage 50 from: {appPath}");
            Automation = new UIA3Automation();
            App = Application.Launch(appPath);

            // Create page objects
            MainPage = new SageMainPage(App, Automation, Log);
            MainPage.WaitForMainWindow();

            MenuPage = new ReportsMenuPage(App, Automation, Log);

            // Create utilities
            Comparer = new CsvComparer(Log);
            ReportGenerator = new HtmlReportGenerator();
        }

        /// <summary>
        /// Create a ReportViewerPage for the given report window title
        /// </summary>
        protected ReportViewerPage CreateViewerPage(string reportWindowTitle)
        {
            ViewerPage = new ReportViewerPage(App, Automation, Log, reportWindowTitle);
            return ViewerPage;
        }

        /// <summary>
        /// Create an ExcelSavePage for CSV export operations
        /// </summary>
        protected ExcelSavePage CreateExcelPage()
        {
            ExcelPage = new ExcelSavePage(App, Automation, Log);
            return ExcelPage;
        }

        /// <summary>
        /// Clean up: close app and dispose resources.
        /// Called by derived class [TestCleanup] methods.
        /// </summary>
        protected void CleanupSession()
        {
            MainPage?.CloseApp();
        }

        // ==================================================
        //  SCREENSHOT + EXECUTION LOG HELPERS
        // ==================================================

        /// <summary>
        /// Capture a screenshot of the primary monitor and log a marker.
        /// Screenshots are saved to a temp folder and embedded in the execution log HTML.
        /// </summary>
        protected void CaptureScreenshot(string context)
        {
            try
            {
                string screenshotsDir = Path.Combine(Path.GetTempPath(), TestConfig.ScreenshotsTempFolder);
                Directory.CreateDirectory(screenshotsDir);

                string timestamp = DateTime.Now.ToString("HHmmss_fff");
                string fileName = $"screenshot_{timestamp}.png";
                string filePath = Path.Combine(screenshotsDir, fileName);

                using var image = Capture.MainScreen();
                image.ToFile(filePath);

                Log?.Info($"[SCREENSHOT:{fileName}] {context}");
            }
            catch (Exception ex)
            {
                Log?.Info($"WARNING: Failed to capture screenshot: {ex.Message}");
            }
        }

        /// <summary>
        /// Generate (or regenerate) the executionLog.html from accumulated log content.
        /// Each call produces the HTML with all log entries up to this point.
        /// </summary>
        protected void GenerateExecutionLogHtml()
        {
            try
            {
                string logContent = Log?.GetAllLogs() ?? "";
                if (!string.IsNullOrWhiteSpace(logContent))
                {
                    var generator = new ExecutionLogHtmlGenerator();
                    string fileName = generator.SaveToFile(logContent);
                    Log?.Info($"Execution log HTML generated: {fileName}");
                }
            }
            catch (Exception ex)
            {
                Log?.Info($"WARNING: Failed to generate execution log HTML: {ex.Message}");
            }
        }

        /// <summary>
        /// Delete the temporary screenshots folder.
        /// Called at the start of Act26 test to clean up from previous runs.
        /// </summary>
        protected void CleanScreenshots()
        {
            try
            {
                string screenshotsDir = Path.Combine(Path.GetTempPath(), TestConfig.ScreenshotsTempFolder);
                if (Directory.Exists(screenshotsDir))
                {
                    Directory.Delete(screenshotsDir, true);
                    Log?.Info("Cleaned screenshots temp folder");
                }
            }
            catch { }
        }
    }
}
