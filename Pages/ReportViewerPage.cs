using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Sage50Automation.Config;
using Sage50Automation.Models;
using Sage50Automation.Utilities;

namespace Sage50Automation.Pages
{
    /// <summary>
    /// Page Object for the report viewer window (e.g., "Customer Ledgers").
    /// Handles: Opening Modify Report dialog, Export to Excel, and Filter exploration.
    /// 
    /// Usage:
    ///   var viewerPage = new ReportViewerPage(app, automation, logger, "Customer Ledgers");
    ///   var modifyPage = viewerPage.OpenModifyReportDialog();
    ///   modifyPage.ExploreAllFiltersAndOptions(viewerPage.OpenModifyReportDialog);
    ///   viewerPage.ExportToExcel();
    /// </summary>
    public class ReportViewerPage : BasePage
    {
        private readonly string _reportWindowTitle;

        public ReportViewerPage(Application app, UIA3Automation automation, Logger logger, string reportWindowTitle)
            : base(app, automation, logger)
        {
            _reportWindowTitle = reportWindowTitle;
        }

        /// <summary>
        /// Find and return the report viewer window
        /// </summary>
        public AutomationElement GetReportWindow()
        {
            var window = Desktop.FindFirstDescendant(cf => cf.ByName(_reportWindowTitle));

            // Fallback: try partial match
            if (window == null)
            {
                var allWindows = Desktop.FindAllChildren(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));
                foreach (var win in allWindows)
                {
                    if (win.Name != null && win.Name.Contains(_reportWindowTitle.Split(' ')[0]))
                    {
                        window = win;
                        Log.Info($"Found report window by partial match: '{win.Name}'");
                        break;
                    }
                }
            }

            Assert.IsNotNull(window, $"Report window '{_reportWindowTitle}' should be found");
            return window;
        }

        /// <summary>
        /// Open the Modify Report dialog and return a ModifyReportPage instance
        /// </summary>
        public ModifyReportPage OpenModifyReportDialog()
        {
            var reportWindow = GetReportWindow();

            Log.Info("Finding Options icon (AutomationId: 127)...");
            var optionsIcon = reportWindow.FindFirstDescendant(cf => cf.ByAutomationId("127"));
            Assert.IsNotNull(optionsIcon, "Options icon should be found");
            Log.Info("Found Options icon, clicking...");
            optionsIcon.Click();
            Thread.Sleep(TestConfig.MediumWaitMs);

            return new ModifyReportPage(App, Automation, Log, _reportWindowTitle);
        }

        /// <summary>
        /// Export the current report to Excel
        /// </summary>
        public void ExportToExcel()
        {
            Log.Info("Exporting to Excel...");
            var reportWindow = GetReportWindow();

            // Click Excel tool
            Log.Info("Finding Excel tool...");
            var excelTool = reportWindow.FindFirstDescendant(cf => cf.ByAutomationId("122"));
            Assert.IsNotNull(excelTool, "Excel tool should be found");
            Log.Info($"Found Excel tool, clicking...");
            excelTool.Click();

            // Handle "Copy Report to Excel" dialog
            Thread.Sleep(TestConfig.MediumWaitMs);
            Log.Info("Finding Copy Report to Excel window...");
            var copyDialog = Desktop.FindFirstDescendant(cf => cf.ByName("Copy Report to Excel"));
            Assert.IsNotNull(copyDialog, "Copy Report to Excel window should be found");
            Log.Info($"Found window: {copyDialog.Name}");

            var okButton = copyDialog.FindFirstDescendant(cf => cf.ByAutomationId("1"));
            Assert.IsNotNull(okButton, "OK button should be found");
            okButton.Click();
            Log.Info("Clicked OK to start export");

            // Wait for export to complete
            WaitForExportProgress();
            Log.Info("Excel export completed");
        }

        /// <summary>
        /// Wait for the Excel export progress window to close
        /// </summary>
        private void WaitForExportProgress()
        {
            Log.Info("Looking for export progress window...");

            AutomationElement? progressWindow = null;
            int attempts = 0;

            // Try to find the progress window
            while (progressWindow == null && attempts < 10)
            {
                Thread.Sleep(500);
                attempts++;

                progressWindow = Desktop.FindFirstDescendant(
                    cf => cf.ByName($"Copying {_reportWindowTitle} to Excel"));

                if (progressWindow == null)
                {
                    var allWindows = Desktop.FindAllDescendants(
                        cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));
                    foreach (var win in allWindows)
                    {
                        if (win.Name != null && win.Name.Contains("Copying") && win.Name.Contains("Excel"))
                        {
                            progressWindow = win;
                            Log.Info($"Found progress window: '{win.Name}'");
                            break;
                        }
                    }
                }

                if (progressWindow == null)
                {
                    var allPanes = Desktop.FindAllDescendants(
                        cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Pane));
                    foreach (var pane in allPanes)
                    {
                        var cancelBtn = pane.FindFirstDescendant(cf => cf.ByName("Cancel"));
                        if (cancelBtn != null && pane.Name != null && pane.Name.Contains("Copying"))
                        {
                            progressWindow = pane;
                            break;
                        }
                    }
                }

                if (attempts % 2 == 0)
                    Log.Info($"Searching for progress window... (attempt {attempts}/10)");
            }

            if (progressWindow != null)
            {
                Log.Info($"Progress window found, waiting for completion...");
                int waited = 0;

                while (progressWindow != null && waited < TestConfig.ExportMaxWaitSeconds)
                {
                    Thread.Sleep(1000);
                    waited++;

                    progressWindow = Desktop.FindFirstDescendant(
                        cf => cf.ByName($"Copying {_reportWindowTitle} to Excel"));

                    if (progressWindow == null)
                    {
                        var allWindows = Desktop.FindAllDescendants(
                            cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));
                        foreach (var win in allWindows)
                        {
                            if (win.Name != null && win.Name.Contains("Copying") && win.Name.Contains("Excel"))
                            {
                                progressWindow = win;
                                break;
                            }
                        }
                    }

                    if (waited % 5 == 0)
                        Log.Info($"Still exporting... ({waited}s elapsed)");
                }

                Log.Info(progressWindow == null
                    ? $"Export completed after {waited}s"
                    : $"WARNING: Progress window still present after {TestConfig.ExportMaxWaitSeconds}s");
            }
            else
            {
                Log.Info("Progress window not detected, waiting a bit...");
                Thread.Sleep(TestConfig.LongWaitMs);
            }
        }

        /// <summary>
        /// Capture the report's DEFAULT data (no filter modifications) and save as CSV.
        /// Clicks the Excel icon on the report viewer window directly and saves the
        /// CSV as "{prefix}_default_data.csv".
        /// </summary>
        public void DefaultDataExport(string filePrefix)
        {
            Log.Info($"=== Capturing DEFAULT data export (prefix: {filePrefix}) ===");

            // Click Excel icon on the report viewer and complete the export
            ExportToExcel();

            // Save as "{prefix}_default_data.csv"
            string csvName = $"{filePrefix}_default_data.csv";
            Log.Info($"      Saving default data as CSV: {csvName}");
            var excelPage = new ExcelSavePage(App, Automation, Log);
            excelPage.SaveAsCSV(csvName);
            Thread.Sleep(1000);

            Log.Info($"=== DEFAULT data export completed: {csvName} ===");
        }

        /// <summary>
        /// Explore all filters/options dynamically, and for each combination:
        ///   1. Select filter → option → values → OK
        ///   2. Export to Excel → Save CSV
        ///   3. Re-open options → next combination
        ///   4. Save all selections to JSON for Act24 to replay
        /// </summary>
        public FilterSelectionsData ExploreAllFiltersAndExportAct26(
            string filePrefix,
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
                    ExportToExcel();

                    // F: Save as CSV with dynamic name
                    string csvName = SelectionsPersistence.BuildCsvFileName(filePrefix, filter.Name, option.Name);
                    selection.CsvFileName = csvName;
                    Log.Info($"      Saving as CSV: {csvName}");
                    var excelPage = new ExcelSavePage(App, Automation, Log);
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
                        modifyPage = OpenModifyReportDialog();
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

        /// <summary>
        /// Replay the filter selections from Act26 JSON, and for each combination:
        ///   1. Select same filter → option → values → OK
        ///   2. Export to Excel → Save CSV as "act24_FilterName_OptionName.csv"
        ///   3. Re-open options → next combination
        /// </summary>
        public void ReplayAndExploreAllFiltersAndExportAct24(
            string filePrefix,
            FilterSelectionsData selectionsData,
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
                ExportToExcel();

                // F: Save as CSV with act24 prefix
                string csvName = SelectionsPersistence.BuildCsvFileName(filePrefix, sel.FilterName, sel.OptionName);
                Log.Info($"      Saving as CSV: {csvName}");
                var excelPage = new ExcelSavePage(App, Automation, Log);
                excelPage.SaveAsCSV(csvName);
                Thread.Sleep(1000);

                Log.Info($"  Completed: {csvName}");

                // G: Re-open options for next combination
                if (i < selectionsData.Selections.Count - 1)
                {
                    Log.Info($"      Re-opening Options for next combination...");
                    modifyPage = OpenModifyReportDialog();
                    modifyPage.ClearAllFilters();
                }
            }

            Log.Info($"\n=== REPLAY + EXPORT ALL completed ({selectionsData.Selections.Count} files) ===");
        }
    }
}
