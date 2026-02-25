using Microsoft.VisualStudio.TestTools.UnitTesting;
using Sage50Automation.Config;
using Sage50Automation.Data;

namespace Sage50Automation.Tests
{
    /// <summary>
    /// Test class for Actian 2024 (Sage 50 2024).
    /// This runs SECOND and uses the existing Reports folder to save act02.csv.
    /// 
    /// Test flow:
    ///   1. Verify Reports folder exists (from Actian 2026)
    ///   2. Open company
    ///   3. Navigate to report (via ReportTestData)
    ///   4. Explore all filters and options
    ///   5. Export to Excel → Save as act02.csv
    ///   6. Compare act01.csv vs act02.csv → Generate HTML report
    /// </summary>
    [TestClass]
    public class Actian2024CustomerLedgerTests : BaseTest
    {
        [TestInitialize]
        public void Setup()
        {
            try
            {
                // Do NOT clear logs — we need Actian 2026 logs for the HTML report
                InitializeSession(TestConfig.Sage2024AppPath, clearLogs: false);
                Log.Info("=== Starting Actian 2024 Test ===");
            }
            catch (Exception ex)
            {
                Log?.Info($"FATAL ERROR in Actian 2024 Setup: {ex.Message}");
                Log?.Info($"Stack Trace: {ex.StackTrace}");
                Log?.Info("=== Actian 2024 Test FAILED in Setup ===");
                throw;
            }
        }

        [TestMethod]
        [Priority(2)]
        public void Test_2_Actian2024_CustomerLedger()
        {
            // ══════════════════════════════════════════════════════════
            //  TEST DATA — Change these 3 lines to test a different report
            // ══════════════════════════════════════════════════════════
            string menuCategory = MenuCategory.AccountsReceivable;
            ReportInfo report = ReportList.CustomerLedger;
            string csvFileName = "act02.csv";

            try
            {
                // Step 1: Verify Reports folder exists
                Log.Info("Step 1: Verifying Reports folder exists...");
                Assert.IsTrue(Directory.Exists(TestConfig.ReportsFolderPath),
                    $"Reports folder should exist from Actian 2026 test: {TestConfig.ReportsFolderPath}");
                Log.Info($"Reports folder verified: {TestConfig.ReportsFolderPath}");

                // Step 2: Open company
                MainPage.OpenCompany();

                // Step 3: Navigate — menu category + report (separate calls)
                MenuPage.OpenReportsMenu();
                MenuPage.ClickMenuCategory(menuCategory);
                MenuPage.SelectReport(report);

                // Step 4: Explore all filters and options
                var viewerPage = CreateViewerPage(report.ReportWindowTitle);
                var modifyPage = viewerPage.OpenModifyReportDialog();
                modifyPage.ExploreAllFiltersAndOptions(viewerPage.OpenModifyReportDialog);

                // TODO: Uncomment when ready to export and compare
                // // Step 5: Export to Excel
                // viewerPage.ExportToExcel();

                // // Step 6: Delete existing act02.csv if it exists
                // string act02Path = Path.Combine(TestConfig.ReportsFolderPath, csvFileName);
                // if (File.Exists(act02Path))
                // {
                //     Log.Info($"Deleting existing {csvFileName}...");
                //     File.Delete(act02Path);
                //     Log.Info($"Deleted {csvFileName}");
                // }

                // // Step 7: Save as CSV
                // var excelPage = CreateExcelPage();
                // excelPage.SaveAsCSV(csvFileName);

                // Thread.Sleep(2000);

                // // Step 8: Compare CSV files and generate HTML report
                // Log.Info("=== Step 8: Comparing act01.csv and act02.csv ===");
                // CompareAndReport();

                Log.Info("=== Actian 2024 Test Completed Successfully! ===");
            }
            catch (Exception ex)
            {
                Log.Info($"ERROR in Actian 2024 test: {ex.Message}");
                Log.Info($"Stack Trace: {ex.StackTrace}");
                Log.Info("=== Actian 2024 Test FAILED ===");
                throw;
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            Log?.Info("Actian 2024 Cleanup started");
            CleanupSession();
            Log?.Info("Actian 2024 Cleanup completed");
        }

        /// <summary>
        /// Compare act01.csv vs act02.csv and generate HTML report.
        /// Asserts failure if files don't match.
        /// </summary>
        private void CompareAndReport()
        {
            string act01Path = Path.Combine(TestConfig.ReportsFolderPath, "act01.csv");
            string act02Path = Path.Combine(TestConfig.ReportsFolderPath, "act02.csv");

            Assert.IsTrue(File.Exists(act01Path), $"act01.csv not found at: {act01Path}");
            Assert.IsTrue(File.Exists(act02Path), $"act02.csv not found at: {act02Path}");

            // Compare CSV files
            var result = Comparer.Compare(act01Path, act02Path);

            // Generate HTML report
            string reportFileName = ReportGenerator.Save(result.MarkdownReport);
            Log.Info($"HTML report saved: {reportFileName}");

            // Assert
            if (result.HasMismatch)
                Assert.Fail(result.FailureMessage);
        }
    }
}
