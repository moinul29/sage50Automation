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
    }
}
