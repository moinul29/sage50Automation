namespace Sage50Automation.Data
{
    // ╔══════════════════════════════════════════════════════════════════╗
    // ║  HOW TO USE (3 independent pieces — full control from test)     ║
    // ║                                                                 ║
    // ║  In your test file:                                             ║
    // ║                                                                 ║
    // ║  // Step 1: Pick the menu category                              ║
    // ║  string menu = MenuCategory.AccountsReceivable;                 ║
    // ║                                                                 ║
    // ║  // Step 2: Pick the report                                     ║
    // ║  ReportInfo report = ReportList.CustomerLedger;                 ║
    // ║                                                                 ║
    // ║  // Step 3: Pick the CSV file name                              ║
    // ║  string csvFile = "act01.csv";                                  ║
    // ║                                                                 ║
    // ║  // Use them independently:                                     ║
    // ║  MenuPage.ClickMenuCategory(menu);                              ║
    // ║  MenuPage.SelectReport(report);                                 ║
    // ║  var viewer = CreateViewerPage(report.ReportWindowTitle);        ║
    // ║  excelPage.SaveAsCSV(csvFile);                                  ║
    // ║                                                                 ║
    // ║  Mix & match freely:                                            ║
    // ║    MenuCategory.AccountsReceivable + ReportList.CustomerList     ║
    // ║    MenuCategory.AccountPayable + ReportList.VendorLedger         ║
    // ╚══════════════════════════════════════════════════════════════════╝

    // ═══════════════════════════════════════════════════════════════════
    //  PART 1: MENU CATEGORIES (under "Reports & Forms")
    // ═══════════════════════════════════════════════════════════════════

    /// <summary>
    /// Menu categories under "Reports &amp; Forms".
    /// Use these in your test file to control which menu to click.
    /// 
    /// Example: MenuPage.ClickMenuCategory(MenuCategory.AccountsReceivable);
    /// </summary>
    public static class MenuCategory
    {
        public const string AccountsReceivable = "Accounts Receivable...";
        public const string AccountPayable = "Account Payable...";
        public const string GeneralLedger = "General Ledger...";
        public const string FinancialStatements = "Financial Statements...";
        public const string Inventory = "Inventory...";
        public const string JobReports = "Job Reports...";
        public const string Payroll = "Payroll...";
        public const string AccountReconciliation = "Account Reconciliation...";

        // ===== ADD MORE MENU CATEGORIES HERE =====
    }

    // ═══════════════════════════════════════════════════════════════════
    //  PART 2: REPORT INFO (name, window title, list index)
    // ═══════════════════════════════════════════════════════════════════

    /// <summary>
    /// Information about a specific report in the "Select a Report or Form" list.
    /// Contains the report name, window title, and list index.
    /// </summary>
    public class ReportInfo
    {
        /// <summary>
        /// Display name of the report (used for logging).
        /// Examples: "Customer Ledger", "Vendor Ledger", "Customer List"
        /// </summary>
        public string ReportName { get; set; } = "";

        /// <summary>
        /// Window title that appears when the report opens.
        /// Examples: "Customer Ledgers", "Vendor Ledgers"
        /// </summary>
        public string ReportWindowTitle { get; set; } = "";

        /// <summary>
        /// Zero-based index of the report in the "Select a Report or Form" list.
        /// Example: Customer Ledger = 3 (4th item in the list)
        /// </summary>
        public int ReportListIndex { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════════
    //  PART 3: PREDEFINED REPORTS (pick one in your test file)
    // ═══════════════════════════════════════════════════════════════════

    /// <summary>
    /// Predefined report definitions.
    /// Use these in your test file to control which report to select.
    /// 
    /// Example: MenuPage.SelectReport(ReportList.CustomerLedger);
    /// </summary>
    public static class ReportList
    {
        // ── Accounts Receivable Reports ──

        public static ReportInfo CustomerLedger => new()
        {
            ReportName = "Customer Ledger",
            ReportWindowTitle = "Customer Ledgers",
            ReportListIndex = 3
        };

        // public static ReportInfo CustomerList => new()
        // {
        //     ReportName = "Customer List",
        //     ReportWindowTitle = "Customer Lists",
        //     ReportListIndex = 0   // Update with correct index
        // };

        // ── Account Payable Reports ──

        // public static ReportInfo VendorLedger => new()
        // {
        //     ReportName = "Vendor Ledger",
        //     ReportWindowTitle = "Vendor Ledgers",
        //     ReportListIndex = 2   // Update with correct index
        // };

        // ── General Ledger Reports ──

        // public static ReportInfo GeneralLedger => new()
        // {
        //     ReportName = "General Ledger",
        //     ReportWindowTitle = "General Ledger",
        //     ReportListIndex = 0   // Update with correct index
        // };

        // ===== ADD MORE REPORTS HERE =====
    }
}
