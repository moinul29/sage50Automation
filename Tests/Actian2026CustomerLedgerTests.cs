using Microsoft.VisualStudio.TestTools.UnitTesting;
using Sage50Automation.Config;
using Sage50Automation.Data;

namespace Sage50Automation.Tests
{
    /// <summary>
    /// Test class for Actian 2026 (Sage 50 2026).
    /// This runs FIRST and creates the Reports folder with act01.csv.
    /// 
    /// Test flow:
    ///   1. Manage Reports folder (delete old files, recreate)
    ///   2. Open company
    ///   3. Navigate to report (via ReportTestData)
    ///   4. Explore all filters and options
    ///   5. Export to Excel → Save as act01.csv
    /// </summary>
    [TestClass]
    public class Actian2026CustomerLedgerTests : BaseTest
    {
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
                Log?.Info("=== Actian 2026 Test FAILED in Setup ===");
                throw;
            }
        }

        [TestMethod]
        [Priority(1)]
        public void Test_1_Actian2026_CustomerLedger()
        {
            // ══════════════════════════════════════════════════════════
            //  TEST DATA — Change these 3 lines to test a different report
            // ══════════════════════════════════════════════════════════
            string menuCategory = MenuCategory.AccountsReceivable;
            ReportInfo report = ReportList.CustomerLedger;
            string csvFileName = "act01.csv";

            try
            {
                // Step 1: Manage Reports folder
                Log.Info("Step 1: Managing Reports folder...");
                ManageReportsFolder();

                // Step 2: Open company
                MainPage.OpenCompany();

                // Step 3: Navigate — menu category + report (separate calls)
                MenuPage.OpenReportsMenu();
                MenuPage.ClickMenuCategory(menuCategory);
                MenuPage.SelectReport(report);

                // Step 4: Explore all filters and options
                var viewerPage = CreateViewerPage(report.ReportWindowTitle);
                viewerPage.ExploreAllFiltersAndOptions();

                // TODO: Uncomment when ready to export
                // // Step 5: Export to Excel
                // viewerPage.ExportToExcel();

                // // Step 6: Save as CSV
                // var excelPage = CreateExcelPage();
                // excelPage.SaveAsCSV(csvFileName);

                Thread.Sleep(2000);
                Log.Info("=== Actian 2026 Test Completed Successfully! ===");
            }
            catch (Exception ex)
            {
                Log.Info($"ERROR in Actian 2026 test: {ex.Message}");
                Log.Info($"Stack Trace: {ex.StackTrace}");
                Log.Info("=== Actian 2026 Test FAILED ===");
                throw;
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            Log?.Info("Actian 2026 Cleanup started");
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
        }
    }
}
