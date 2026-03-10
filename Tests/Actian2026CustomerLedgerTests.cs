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
            //  TEST DATA — Change these 2 lines to test a different report
            // ══════════════════════════════════════════════════════════
            string menuCategory = MenuCategory.AccountsReceivable;
            ReportInfo report = ReportList.CustomerLedger;

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

                // Step 4: Explore all filters + Export each as act26_Filter_Option.csv
                var viewerPage = CreateViewerPage(report.ReportWindowTitle);
                var modifyPage = viewerPage.OpenModifyReportDialog();

                ExploreAndExportAll("act26", viewerPage, modifyPage);

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

            // Delete old HTML report files from project directory
            DeleteOldHtmlReports();
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
        }

        /// <summary>
        /// Explore all filters/options dynamically, and for each combination:
        ///   1. Select filter → option → values → OK
        ///   2. Export to Excel → Save CSV
        ///   3. Re-open options → next combination
        ///   4. Save all selections to JSON for Act24 to replay
        /// </summary>
        private FilterSelectionsData ExploreAndExportAll(
            string filePrefix,
            ReportViewerPage viewerPage,
            ModifyReportPage modifyPage)
        {
            Log.Info($"=== Starting EXPLORE + EXPORT ALL (prefix: {filePrefix}) ===");

            var selectionsData = new FilterSelectionsData();

            // STEP 1: Discover all filters
            Log.Info("\n=== STEP 1: Discovering all filters from UI ===");
            var allFilters = modifyPage.DiscoverFilters();
            Log.Info($"Total filters discovered: {allFilters.Count}");

            Log.Info("\n--- Discovered Filters List ---");
            for (int i = 0; i < allFilters.Count; i++)
            {
                Log.Info($"  Filter [{i + 1}]: {allFilters[i].Name}");
                selectionsData.FilterNames.Add(allFilters[i].Name);
            }

            // STEP 2: Discover options for each filter
            Log.Info("\n=== STEP 2: Discovering options for each filter ===");
            var filterOptionsMap = new Dictionary<string, List<OptionInfo>>();

            foreach (var filter in allFilters)
            {
                Log.Info($"\n--- Discovering options for filter: {filter.Name} ---");
                filter.Element.Click();
                Thread.Sleep(TestConfig.ShortWaitMs);

                var options = modifyPage.DiscoverFilterOptions(filter);
                filterOptionsMap[filter.Name] = options;

                var optionNames = new List<string>();
                Log.Info($"  Options discovered for '{filter.Name}':");
                foreach (var opt in options)
                {
                    Log.Info($"    - {opt.Name} (Type: {opt.OptionType})");
                    optionNames.Add(opt.Name);
                }
                selectionsData.FilterOptionsMap[filter.Name] = optionNames;
            }

            // STEP 3: For each filter+option: select → OK → export → save CSV
            Log.Info("\n=== STEP 3: Selecting, exporting, and saving for each combination ===");

            int filterIndex = 0;
            foreach (var filter in allFilters)
            {
                filterIndex++;
                var options = filterOptionsMap[filter.Name];
                int optionIndex = 0;

                foreach (var option in options)
                {
                    optionIndex++;
                    Log.Info($"\n======================================================");
                    Log.Info($"  FILTER [{filterIndex}/{allFilters.Count}]: {filter.Name}");
                    Log.Info($"  OPTION [{optionIndex}/{options.Count}]: {option.Name}");
                    Log.Info($"======================================================");

                    var selection = new FilterSelection
                    {
                        FilterName = filter.Name,
                        OptionName = option.Name,
                        OptionType = option.OptionType
                    };

                    // A: Click on filter
                    var freshFilter = modifyPage.RefindFilter(filter.Name);
                    if (freshFilter != null)
                    {
                        freshFilter.Click();
                        Thread.Sleep(TestConfig.ShortWaitMs);
                    }
                    else
                    {
                        try { filter.Element.Click(); Thread.Sleep(TestConfig.ShortWaitMs); } catch { }
                    }

                    // B: Select the option
                    modifyPage.SelectOption(option);
                    Thread.Sleep(TestConfig.ShortWaitMs);

                    // C: Handle value selection (no values needed for 'All')
                    switch (option.OptionType)
                    {
                        case OptionType.All:
                            Log.Info($"      'All' option - no values to configure");
                            break;
                        case OptionType.Range:
                            modifyPage.SelectRangeValues();
                            var (rangeFrom, rangeTo) = modifyPage.ExtractRangeFromSetToRow();
                            selection.RangeFrom = rangeFrom;
                            selection.RangeTo = rangeTo;
                            break;
                        case OptionType.OneOrMore:
                            var selectedVals = modifyPage.SelectMultipleValues(TestConfig.MaxFilterValuesToSelect);
                            selection.SelectedValues = selectedVals;
                            break;
                    }

                    // D: Click OK
                    modifyPage.ClickOK();
                    Thread.Sleep(1000);

                    // E: Export to Excel
                    Log.Info($"      Exporting to Excel...");
                    viewerPage.ExportToExcel();

                    // F: Save as CSV with dynamic name
                    string csvName = SelectionsPersistence.BuildCsvFileName(filePrefix, filter.Name, option.Name);
                    selection.CsvFileName = csvName;
                    Log.Info($"      Saving as CSV: {csvName}");
                    var excelPage = CreateExcelPage();
                    excelPage.SaveAsCSV(csvName);
                    Thread.Sleep(1000);

                    // Record this selection
                    selectionsData.Selections.Add(selection);
                    Log.Info($"  Completed: {csvName}");

                    // G: Re-open options for next combination
                    bool hasMore = optionIndex < options.Count || filterIndex < allFilters.Count;
                    if (hasMore)
                    {
                        Log.Info($"      Re-opening Options for next combination...");
                        modifyPage = viewerPage.OpenModifyReportDialog();
                        modifyPage.ClearAllFilters();

                        string nextFilterName = optionIndex < options.Count
                            ? filter.Name
                            : (filterIndex < allFilters.Count ? allFilters[filterIndex].Name : filter.Name);

                        var refreshed = modifyPage.RefindFilter(nextFilterName);
                        if (refreshed != null)
                        {
                            if (optionIndex < options.Count)
                                filter.Element = refreshed;
                            else if (filterIndex < allFilters.Count)
                                allFilters[filterIndex].Element = refreshed;
                        }
                    }
                }
            }

            // Save selections to JSON for Act24
            Log.Info("\n=== Saving filter selections to JSON ===");
            SelectionsPersistence.Save(selectionsData);
            Log.Info($"Saved {selectionsData.Selections.Count} filter+option combinations to JSON");

            Log.Info($"\n=== EXPLORE + EXPORT ALL completed ({selectionsData.Selections.Count} files) ===");
            return selectionsData;
        }
    }
}
