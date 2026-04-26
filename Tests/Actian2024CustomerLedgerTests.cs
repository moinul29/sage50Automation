using Microsoft.VisualStudio.TestTools.UnitTesting;
using Sage50Automation.Config;
using Sage50Automation.Data;
using Sage50Automation.Models;
using Sage50Automation.Pages;
using Sage50Automation.Utilities;

namespace Sage50Automation.Tests
{
    /// <summary>
    /// Test class for Actian 2024 (Sage 50 2024).
    /// This runs SECOND and replays the SAME filter selections from Act26.
    /// 
    /// Test flow:
    ///   1. Verify Reports folder exists (from Actian 2026)
    ///   2. Load filter selections JSON from Act26
    ///   3. Open company
    ///   4. Navigate to report
    ///   5. For each filter+option combination (from JSON):
    ///      a. Select same filter → option → values → OK
    ///      b. Export to Excel → Save CSV as "act24_FilterName_OptionName.csv"
    ///   6. Compare each act26_*.csv vs act24_*.csv pair → separate HTML reports
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
                CaptureScreenshot("FATAL: Actian 2024 Setup failed");
                Log?.Info("=== Actian 2024 Test FAILED in Setup ===");
                throw;
            }
        }

        [TestMethod]
        [Priority(2)]
        public void Test_2_Actian2024_CustomerLedger()
        {
            // Test data is sourced from Actian2026CustomerLedgerTests so both tests stay in sync.
            string reportGroup = Actian2026CustomerLedgerTests.ReportGroupName;
            ReportInfo report = Actian2026CustomerLedgerTests.Report;

            try
            {
                // Step 1: Verify Reports folder exists
                Log.Info("Step 1: Verifying Reports folder exists...");
                Assert.IsTrue(Directory.Exists(TestConfig.ReportsFolderPath),
                    $"Reports folder should exist from Actian 2026 test: {TestConfig.ReportsFolderPath}");
                Log.Info($"Reports folder verified: {TestConfig.ReportsFolderPath}");

                // Step 2: Load filter selections from Act26
                Log.Info("Step 2: Loading filter selections from Act26 JSON...");
                var selectionsData = SelectionsPersistence.Load();
                Assert.IsNotNull(selectionsData,
                    "filter_selections.json not found — Act26 must run first!");
                Log.Info($"Loaded {selectionsData.Selections.Count} filter+option combinations from Act26");

                // Step 3: Open company
                MainPage.OpenCompany();

                // Step 4: Navigate — report group + report (separate calls)
                MenuPage.OpenReportsMenu();
                MenuPage.ClickReportGroup(reportGroup);
                MenuPage.SelectReport(report);

                // Step 5: Replay all filter selections + Export each as act24_Filter_Option.csv
                var viewerPage = CreateViewerPage(report.ReportWindowTitle);

                // Capture DEFAULT (unmodified) report data first
                viewerPage.DefaultDataExport("act24");

                var modifyPage = viewerPage.OpenModifyReportDialog();
                viewerPage.ReplayAndExploreAllFiltersAndExportAct24("act24", selectionsData, modifyPage);

                Thread.Sleep(2000);

                // Step 6: Compare each act26_*.csv vs act24_*.csv pair
                Log.Info("=== Step 6: Comparing CSV file pairs ===");
                var comparisonRunner = new CsvComparisonRunner(Log, Comparer, ReportGenerator);
                comparisonRunner.CompareAllPairs(selectionsData);

                Log.Info("=== Actian 2024 Test Completed Successfully! ===");
            }
            catch (Exception ex)
            {
                Log.Info($"ERROR in Actian 2024 test: {ex.Message}");
                Log.Info($"Stack Trace: {ex.StackTrace}");
                CaptureScreenshot("ERROR: Actian 2024 test failed");
                Log.Info("=== Actian 2024 Test FAILED ===");
                throw;
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            Log?.Info("Actian 2024 Cleanup started");
            GenerateExecutionLogHtml();
            CleanupSession();
            Log?.Info("Actian 2024 Cleanup completed");
        }

    }
}
