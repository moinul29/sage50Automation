using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Sage50Automation.Config;
using Sage50Automation.Data;
using Sage50Automation.Utilities;

namespace Sage50Automation.Pages
{
    /// <summary>
    /// Page Object for navigating Reports &amp; Forms menu.
    /// 
    /// ┌──────────────────────────────────────────────────────────────────┐
    /// │  3 INDEPENDENT METHODS (call each from your test file):           │
    /// │                                                                  │
    /// │  1. OpenReportsMenu()          → Opens "Reports & Forms"        │
    /// │  2. ClickReportGroup(string)   → Clicks any report group        │
    /// │  3. SelectReport(ReportInfo)   → Selects report from the list   │
    /// │                                                                  │
    /// │  This gives you full control to mix & match:                     │
    /// │    Accounts Receivable + Customer Ledger                         │
    /// │    Accounts Receivable + Customer List                           │
    /// │    Account Payable + Vendor Ledger                               │
    /// └──────────────────────────────────────────────────────────────────┘
    /// 
    /// Usage (each step called independently from your test file):
    ///   MenuPage.OpenReportsMenu();                              // Step 1
    ///   MenuPage.ClickReportGroup(MenuCategory.AccountsReceivable);  // Step 2
    ///   MenuPage.SelectReport(ReportList.CustomerLedger);        // Step 3
    /// </summary>
    public class ReportsMenuPage : BasePage
    {
        public ReportsMenuPage(Application app, UIA3Automation automation, Logger logger)
            : base(app, automation, logger) { }

        // ═══════════════════════════════════════════════════════════════
        //  STEP 1: Open "Reports & Forms" menu
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Open the "Reports &amp; Forms" menu from the main menu bar.
        /// Call this FIRST before ClickReportGroup.
        /// </summary>
        public void OpenReportsMenu()
        {
            Log.Info("Opening 'Reports & Forms' menu...");

            var companyWindow = Desktop.FindFirstDescendant(cf => cf.ByName(TestConfig.CompanyWindowName));
            Assert.IsNotNull(companyWindow, $"Company window '{TestConfig.CompanyWindowName}' should be found");

            var menuBar = companyWindow.FindFirstChild(cf => cf.ByAutomationId("MenuBar"));
            Assert.IsNotNull(menuBar, "MenuBar should be found");

            var reportMenu = menuBar.FindFirstChild(cf => cf.ByName("Reports && Forms"));
            Assert.IsNotNull(reportMenu, "Reports & Forms menu should be found");
            Log.Info($"Found Reports menu: {reportMenu.Name}, clicking...");
            reportMenu.Click();
            Thread.Sleep(TestConfig.MediumWaitMs);

            Log.Info("Reports & Forms menu opened");
        }

        // ═══════════════════════════════════════════════════════════════
        //  STEP 2: Click a report group (Accounts Receivable, etc.)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Click a report group from the Reports &amp; Forms dropdown.
        /// 
        /// Example:
        ///   MenuPage.ClickReportGroup(MenuCategory.AccountsReceivable);
        ///   MenuPage.ClickReportGroup(MenuCategory.AccountPayable);
        /// </summary>
        public void ClickReportGroup(string reportGroup)
        {
            Log.Info($"Clicking report group: '{reportGroup}'...");

            var companyWindow = Desktop.FindFirstDescendant(cf => cf.ByName(TestConfig.CompanyWindowName));
            Assert.IsNotNull(companyWindow, $"Company window '{TestConfig.CompanyWindowName}' should be found");

            var menuItem = companyWindow.FindFirstDescendant(cf => cf.ByName(reportGroup));
            Assert.IsNotNull(menuItem, $"'{reportGroup}' menu item should be found");
            Log.Info($"Found: {menuItem.Name}, clicking...");
            menuItem.Click();
            Log.Info($"Successfully clicked '{reportGroup}'");
            Thread.Sleep(3000);
        }

        // ═══════════════════════════════════════════════════════════════
        //  STEP 3: Select a report from the list
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Select a report from the "Select a Report or Form" dialog.
        /// 
        /// Example:
        ///   MenuPage.SelectReport(ReportList.CustomerLedger);
        ///   MenuPage.SelectReport(ReportList.VendorLedger);
        /// </summary>
        public void SelectReport(ReportInfo report)
        {
            Log.Info($"Selecting report: '{report.ReportName}' (index {report.ReportListIndex})...");

            var selectWindow = Desktop.FindFirstDescendant(cf => cf.ByName("Select a Report or Form"));
            Assert.IsNotNull(selectWindow, "'Select a Report or Form' window should be found");
            Log.Info($"Found window: {selectWindow.Name}");

            // Find the report list (use second ListBox — first is Form Types)
            Log.Info("Finding ListBox...");
            var allListBoxes = selectWindow.FindAllDescendants(
                cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.List));
            Log.Info($"Found {allListBoxes.Length} ListBoxes");

            AutomationElement? listBox = null;
            if (allListBoxes.Length > 1)
            {
                listBox = allListBoxes[1];
                Log.Info("Using the second ListBox (report list)");
            }
            else if (allListBoxes.Length == 1)
            {
                listBox = allListBoxes[0];
                Log.Info("Only one ListBox found, using it");
            }
            Assert.IsNotNull(listBox, "Report ListBox should be found");

            // Double-click the report at the specified index
            var allItems = listBox.FindAllDescendants(
                cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.ListItem));
            Log.Info($"Found {allItems.Length} list items, selecting index {report.ReportListIndex} ({report.ReportName})...");

            Assert.IsTrue(allItems.Length > report.ReportListIndex,
                $"Report list should have at least {report.ReportListIndex + 1} items, found only {allItems.Length}");

            allItems[report.ReportListIndex].DoubleClick();
            Log.Info($"Double-clicked '{report.ReportName}' (index {report.ReportListIndex})");
            Thread.Sleep(3000);
        }
    }
}
