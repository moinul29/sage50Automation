using Microsoft.VisualStudio.TestTools.UnitTesting;
using Sage50Automation.Config;
using Sage50Automation.Data;
using Sage50Automation.Models;
using Sage50Automation.Pages;
using Sage50Automation.Utilities;

namespace Sage50Automation.Tests
{
    /// <summary>
    /// Test class for Actian 2026 (Sage 50 2026).
    /// This runs FIRST and creates the Reports folder.
    /// 
    /// Test flow:
    ///   1. Manage Reports folder (delete old files, recreate)
    ///   2. Open company
    ///   3. Navigate to report (via ReportTestData)
    ///   4. For each filter+option combination:
    ///      a. Select filter → option → values → OK
    ///      b. Export to Excel → Save CSV as "act26_FilterName_OptionName.csv"
    ///   5. Save all selections to JSON for Act24 to replay
    /// </summary>
    [TestClass]
    public class Actian2026CustomerLedgerTests : BaseTest
    {
        // ════════════════════════════════════════════════════════
        //  TEST DATA — Change these 2 lines to test a different report
        //  (Act24 reads these same values so both tests stay in sync)
        // ════════════════════════════════════════════════════════
        public static readonly string ReportGroupName = MenuCategory.AccountsReceivable;
        public static readonly ReportInfo Report = ReportList.CustomerLedger;

        [TestInitialize]
        public void Setup()
        {
            try
            {
                InitializeSession(TestConfig.Sage2026AppPath, clearLogs: true);
                Log.Info("=== Starting Actian 2026 Test ===");
            }
            catch (Exception ex)
            {
                Log?.Info($"FATAL ERROR in Actian 2026 Setup: {ex.Message}");
                Log?.Info($"Stack Trace: {ex.StackTrace}");
                CaptureScreenshot("FATAL: Actian 2026 Setup failed");
                Log?.Info("=== Actian 2026 Test FAILED in Setup ===");
                throw;
            }
        }

        [TestMethod]
        [Priority(1)]
        public void Test_1_Actian2026_CustomerLedger()
        {
            string reportGroup = ReportGroupName;
            ReportInfo report = Report;

            try
            {
                // Step 1: Manage Reports folder
                Log.Info("Step 1: Managing Reports folder...");
                ManageReportsFolder();

                // Step 2: Open company
                MainPage.OpenCompany();

                // Step 3: Navigate — report group + report (separate calls)
                MenuPage.OpenReportsMenu();
                MenuPage.ClickReportGroup(reportGroup);
                MenuPage.SelectReport(report);

                // Step 4: Explore all filters + Export each as act26_Filter_Option.csv
                var viewerPage = CreateViewerPage(report.ReportWindowTitle);

                // Capture DEFAULT (unmodified) report data first
                viewerPage.DefaultDataExport("act26");

                var modifyPage = viewerPage.OpenModifyReportDialog();
                viewerPage.ExploreAllFiltersAndExportAct26("act26", modifyPage);

                Thread.Sleep(2000);
                Log.Info("=== Actian 2026 Test Completed Successfully! ===");
            }
            catch (Exception ex)
            {
                Log.Info($"ERROR in Actian 2026 test: {ex.Message}");
                Log.Info($"Stack Trace: {ex.StackTrace}");
                CaptureScreenshot("ERROR: Actian 2026 test failed");
                Log.Info("=== Actian 2026 Test FAILED ===");
                throw;
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            Log?.Info("Actian 2026 Cleanup started");
            GenerateExecutionLogHtml();
            CleanupSession();
            Log?.Info("Actian 2026 Cleanup completed");
        }

        /// <summary>
        /// Delete and recreate the Reports folder for a clean test run
        /// </summary>
        private void ManageReportsFolder()
        {
            string path = TestConfig.ReportsFolderPath;

            if (Directory.Exists(path))
            {
                Log.Info($"Reports folder exists, attempting to delete: {path}");
                try
                {
                    Directory.Delete(path, true);
                    Log.Info("Successfully deleted Reports folder");
                }
                catch (IOException)
                {
                    Log.Info("WARNING: Folder is in use. Deleting files inside...");
                    try
                    {
                        foreach (var file in Directory.GetFiles(path))
                            File.Delete(file);
                        Log.Info("Deleted existing files in Reports folder");
                    }
                    catch
                    {
                        Log.Info("WARNING: Could not delete some files. Continuing...");
                    }
                }
            }

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
                Log.Info($"Created Reports folder: {path}");
            }
            else
            {
                Log.Info($"Reports folder already exists: {path}");
            }

            // Delete old HTML report files from project directory
            DeleteOldHtmlReports();

            // Clean up screenshots from previous runs
            CleanScreenshots();
        }

        /// <summary>
        /// Delete all existing HTML report files from the project directory
        /// </summary>
        private void DeleteOldHtmlReports()
        {
            string projectDir = TestConfig.ProjectDirectory;
            string baseFileName = TestConfig.HtmlReportBaseFileName;

            var htmlFiles = Directory.GetFiles(projectDir, $"{baseFileName}*.html");
            if (htmlFiles.Length > 0)
            {
                Log.Info($"Deleting {htmlFiles.Length} old HTML report(s)...");
                foreach (var file in htmlFiles)
                {
                    try
                    {
                        File.Delete(file);
                        Log.Info($"  Deleted: {Path.GetFileName(file)}");
                    }
                    catch
                    {
                        Log.Info($"  WARNING: Could not delete {Path.GetFileName(file)}");
                    }
                }
            }
            else
            {
                Log.Info("No old HTML reports to delete");
            }

            // Delete old execution log HTML
            string execLogPath = Path.Combine(projectDir, TestConfig.ExecutionLogHtmlFileName);
            if (File.Exists(execLogPath))
            {
                try
                {
                    File.Delete(execLogPath);
                    Log.Info($"  Deleted: {TestConfig.ExecutionLogHtmlFileName}");
                }
                catch
                {
                    Log.Info($"  WARNING: Could not delete {TestConfig.ExecutionLogHtmlFileName}");
                }
            }
        }

    }
}
