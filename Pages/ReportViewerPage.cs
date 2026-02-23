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
    ///   viewerPage.ExploreAllFiltersAndOptions();   // or
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
        /// Explore ALL filters and their options dynamically.
        /// 
        /// For each filter:
        ///   1. Discover available options (All, Range, One or More)
        ///   2. Select each option and fill in values
        ///   3. Click OK to apply
        ///   4. Re-open and clear for next combination
        /// </summary>
        public void ExploreAllFiltersAndOptions()
        {
            Log.Info("=== Starting Options and Filters exploration ===");

            var modifyPage = OpenModifyReportDialog();

            // STEP 1: Discover all filters
            Log.Info("\n=== STEP 1: Discovering all filters from UI ===");
            var allFilters = modifyPage.DiscoverFilters();
            Log.Info($"Total filters discovered: {allFilters.Count}");

            Log.Info("\n--- Discovered Filters List ---");
            for (int i = 0; i < allFilters.Count; i++)
                Log.Info($"  Filter [{i + 1}]: {allFilters[i].Name}");

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

                Log.Info($"  Options discovered for '{filter.Name}':");
                foreach (var opt in options)
                    Log.Info($"    - {opt.Name} (Type: {opt.OptionType})");
            }

            // STEP 3: Test all filter-option combinations
            Log.Info("\n=== STEP 3: Testing all filters, options and values ===");

            int filterIndex = 0;
            foreach (var filter in allFilters)
            {
                filterIndex++;
                Log.Info($"\n========================================");
                Log.Info($"FILTER [{filterIndex}/{allFilters.Count}]: {filter.Name}");
                Log.Info($"========================================");

                var options = filterOptionsMap[filter.Name];
                int optionIndex = 0;

                foreach (var option in options)
                {
                    optionIndex++;
                    Log.Info($"\n  --- Option [{optionIndex}/{options.Count}]: {option.Name} ---");

                    // A: Click on filter to show its options
                    Log.Info($"      Clicking on filter '{filter.Name}' to show options...");
                    var freshFilter = modifyPage.RefindFilter(filter.Name);
                    if (freshFilter != null)
                    {
                        freshFilter.Click();
                        Thread.Sleep(TestConfig.ShortWaitMs);
                    }
                    else
                    {
                        Log.Info($"      WARNING: Could not find filter '{filter.Name}', trying stored element...");
                        try { filter.Element.Click(); Thread.Sleep(TestConfig.ShortWaitMs); } catch { }
                    }

                    // B: Select the option (radio button)
                    modifyPage.SelectOption(option);
                    Thread.Sleep(TestConfig.ShortWaitMs);

                    // C: Handle value selection based on option type
                    switch (option.OptionType)
                    {
                        case OptionType.All:
                            Log.Info($"      'All' option selected - no additional values needed");
                            break;
                        case OptionType.Range:
                            modifyPage.SelectRangeValues();
                            break;
                        case OptionType.OneOrMore:
                            modifyPage.SelectMultipleValues(TestConfig.MaxFilterValuesToSelect);
                            break;
                    }

                    // D: Click OK to apply
                    modifyPage.ClickOK();
                    Thread.Sleep(1000);
                    Log.Info($"  Completed: Filter='{filter.Name}', Option='{option.Name}'");

                    // E: Re-open Options and Clear All Filters for next iteration
                    bool hasMoreOptions = optionIndex < options.Count;
                    bool hasMoreFilters = filterIndex < allFilters.Count;

                    if (hasMoreOptions || hasMoreFilters)
                    {
                        Log.Info($"      [STEP E] Re-opening Options and clearing filters...");
                        modifyPage = OpenModifyReportDialog();
                        modifyPage.ClearAllFilters();

                        // Re-find the filter element for the next iteration
                        string nextFilterName = hasMoreOptions
                            ? filter.Name
                            : (filterIndex < allFilters.Count ? allFilters[filterIndex].Name : filter.Name);

                        Log.Info($"      Re-finding filter '{nextFilterName}' in new window...");
                        var refreshed = modifyPage.RefindFilter(nextFilterName);
                        if (refreshed != null)
                        {
                            if (hasMoreOptions)
                                filter.Element = refreshed;
                            else if (filterIndex < allFilters.Count)
                                allFilters[filterIndex].Element = refreshed;
                            Log.Info($"      Successfully re-found filter '{nextFilterName}'");
                        }
                        else
                        {
                            Log.Info($"      WARNING: Could not re-find filter '{nextFilterName}'");
                        }
                    }
                }
            }

            Log.Info("\n=== Options and Filters exploration completed ===");
            Log.Info($"Total combinations tested: {allFilters.Count} filters with their respective options");
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
    }
}
