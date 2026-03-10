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
                Log?.Info("=== Actian 2024 Test FAILED in Setup ===");
                throw;
            }
        }

        [TestMethod]
        [Priority(2)]
        public void Test_2_Actian2024_CustomerLedger()
        {
            // ══════════════════════════════════════════════════════════
            //  TEST DATA — Change these 2 lines to test a different report
            // ══════════════════════════════════════════════════════════
            string menuCategory = MenuCategory.AccountsReceivable;
            ReportInfo report = ReportList.CustomerLedger;

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

                // Step 4: Navigate — menu category + report (separate calls)
                MenuPage.OpenReportsMenu();
                MenuPage.ClickMenuCategory(menuCategory);
                MenuPage.SelectReport(report);

                // Step 5: Replay all filter selections + Export each as act24_Filter_Option.csv
                var viewerPage = CreateViewerPage(report.ReportWindowTitle);
                var modifyPage = viewerPage.OpenModifyReportDialog();

                ReplayAndExportAll("act24", selectionsData, viewerPage, modifyPage);

                Thread.Sleep(2000);

                // Step 6: Compare each act26_*.csv vs act24_*.csv pair
                Log.Info("=== Step 6: Comparing CSV file pairs ===");
                CompareAllPairs(selectionsData);

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
        /// Compare each act26_Filter_Option.csv vs act24_Filter_Option.csv pair.
        /// Generates a separate HTML report for each comparison.
        /// </summary>
        private void CompareAllPairs(FilterSelectionsData selectionsData)
        {
            int totalPairs = selectionsData.Selections.Count;
            int passCount = 0;
            int failCount = 0;
            var failedPairs = new List<string>();

            for (int i = 0; i < totalPairs; i++)
            {
                var sel = selectionsData.Selections[i];

                string act26Csv = SelectionsPersistence.BuildCsvFileName("act26", sel.FilterName, sel.OptionName);
                string act24Csv = SelectionsPersistence.BuildCsvFileName("act24", sel.FilterName, sel.OptionName);

                string act26Path = Path.Combine(TestConfig.ReportsFolderPath, act26Csv);
                string act24Path = Path.Combine(TestConfig.ReportsFolderPath, act24Csv);

                Log.Info($"\n--- Comparing pair [{i + 1}/{totalPairs}] ---");
                Log.Info($"  Filter: {sel.FilterName} | Option: {sel.OptionName}");
                Log.Info($"  Act26: {act26Csv}");
                Log.Info($"  Act24: {act24Csv}");

                if (!File.Exists(act26Path))
                {
                    Log.Info($"  SKIP: Act26 file not found: {act26Path}");
                    failCount++;
                    failedPairs.Add($"{sel.FilterName}/{sel.OptionName} (act26 file missing)");
                    continue;
                }
                if (!File.Exists(act24Path))
                {
                    Log.Info($"  SKIP: Act24 file not found: {act24Path}");
                    failCount++;
                    failedPairs.Add($"{sel.FilterName}/{sel.OptionName} (act24 file missing)");
                    continue;
                }

                // Compare with filter context
                var result = Comparer.Compare(act26Path, act24Path, sel);

                // Generate separate HTML report for this pair
                string reportFile = ReportGenerator.Save(result.MarkdownReport);
                Log.Info($"  HTML report: {reportFile}");

                if (result.HasMismatch)
                {
                    failCount++;
                    failedPairs.Add($"{sel.FilterName}/{sel.OptionName}");
                    Log.Info($"  RESULT: MISMATCH — {result.FailureMessage}");
                }
                else
                {
                    passCount++;
                    Log.Info($"  RESULT: MATCH");
                }
            }

            // Summary
            Log.Info($"\n========================================");
            Log.Info($"  COMPARISON SUMMARY");
            Log.Info($"  Total pairs: {totalPairs}");
            Log.Info($"  Passed: {passCount}");
            Log.Info($"  Failed: {failCount}");
            Log.Info($"========================================");

            if (failCount > 0)
            {
                string failMsg = $"CSV comparison failed for {failCount}/{totalPairs} pairs: " +
                                 string.Join(", ", failedPairs);
                Log.Info(failMsg);
                Assert.Fail(failMsg);
            }
        }

        /// <summary>
        /// Replay the filter selections from Act26 JSON, and for each combination:
        ///   1. Select same filter → option → values → OK
        ///   2. Export to Excel → Save CSV as "act24_FilterName_OptionName.csv"
        ///   3. Re-open options → next combination
        /// </summary>
        private void ReplayAndExportAll(
            string filePrefix,
            FilterSelectionsData selectionsData,
            ReportViewerPage viewerPage,
            ModifyReportPage modifyPage)
        {
            Log.Info($"=== Starting REPLAY + EXPORT ALL (prefix: {filePrefix}) ===");
            Log.Info($"Replaying {selectionsData.Selections.Count} combinations from Act26...");

            for (int i = 0; i < selectionsData.Selections.Count; i++)
            {
                var sel = selectionsData.Selections[i];
                Log.Info($"\n======================================================");
                Log.Info($"  REPLAY [{i + 1}/{selectionsData.Selections.Count}]");
                Log.Info($"  FILTER: {sel.FilterName}");
                Log.Info($"  OPTION: {sel.OptionName} ({sel.OptionType})");
                Log.Info($"======================================================");

                // A: Click on filter
                var freshFilter = modifyPage.RefindFilter(sel.FilterName);
                if (freshFilter != null)
                {
                    freshFilter.Click();
                    Thread.Sleep(TestConfig.ShortWaitMs);
                }
                else
                {
                    Log.Info($"      WARNING: Could not find filter '{sel.FilterName}'");
                }

                // B: Select the option
                var optionElement = modifyPage.Window.FindFirstDescendant(cf => cf.ByName(sel.OptionName));
                if (optionElement != null)
                {
                    optionElement.Click();
                    Thread.Sleep(TestConfig.ShortWaitMs);
                    Log.Info($"      Selected option: {sel.OptionName}");
                }
                else
                {
                    Log.Info($"      WARNING: Could not find option '{sel.OptionName}'");
                }

                // C: Replay value selection (no values needed for 'All')
                switch (sel.OptionType)
                {
                    case OptionType.All:
                        Log.Info($"      'All' option - no values to configure");
                        break;
                    case OptionType.Range:
                        modifyPage.SelectRangeValues();
                        break;
                    case OptionType.OneOrMore:
                        modifyPage.SelectSpecificValues(sel.SelectedValues);
                        break;
                }

                // D: Click OK
                modifyPage.ClickOK();
                Thread.Sleep(1000);

                // E: Export to Excel
                Log.Info($"      Exporting to Excel...");
                viewerPage.ExportToExcel();

                // F: Save as CSV with act24 prefix
                string csvName = SelectionsPersistence.BuildCsvFileName(filePrefix, sel.FilterName, sel.OptionName);
                Log.Info($"      Saving as CSV: {csvName}");
                var excelPage = CreateExcelPage();
                excelPage.SaveAsCSV(csvName);
                Thread.Sleep(1000);

                Log.Info($"  Completed: {csvName}");

                // G: Re-open options for next combination
                if (i < selectionsData.Selections.Count - 1)
                {
                    Log.Info($"      Re-opening Options for next combination...");
                    modifyPage = viewerPage.OpenModifyReportDialog();
                    modifyPage.ClearAllFilters();
                }
            }

            Log.Info($"\n=== REPLAY + EXPORT ALL completed ({selectionsData.Selections.Count} files) ===");
        }
    }
}
