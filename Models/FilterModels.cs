using FlaUI.Core.AutomationElements;

namespace Sage50Automation.Models
{
    /// <summary>
    /// Types of filter options available in Sage 50 report dialogs
    /// </summary>
    public enum OptionType
    {
        /// <summary>Select all items (no additional input needed)</summary>
        All,

        /// <summary>Select a From-To range using dropdowns</summary>
        Range,

        /// <summary>Select one or more specific items via checkboxes</summary>
        OneOrMore
    }

    /// <summary>
    /// Represents a filter discovered in the Modify Report dialog.
    /// Example filters: "Customer ID", "Date", "Transaction Type"
    /// </summary>
    public class FilterInfo
    {
        public string Name { get; set; } = "";
        public AutomationElement Element { get; set; } = null!;
    }

    /// <summary>
    /// Represents an option (radio button) for a specific filter.
    /// Example options: "All", "Range", "One or More"
    /// </summary>
    public class OptionInfo
    {
        public string Name { get; set; } = "";
        public AutomationElement Element { get; set; } = null!;
        public OptionType OptionType { get; set; } = OptionType.All;
    }

    /// <summary>
    /// Stores range values (From/To) for range-type filters
    /// </summary>
    public class RangeValues
    {
        public List<string> FromValues { get; set; } = new();
        public List<string> ToValues { get; set; } = new();
    }

    /// <summary>
    /// Represents a selectable value (checkbox) for one-or-more type filters.
    /// Example: Customer names like "ALDRED", "ARMSTRONG"
    /// </summary>
    public class SelectableValue
    {
        public string Name { get; set; } = "";
        public AutomationElement Element { get; set; } = null!;
        public bool IsCheckBox { get; set; } = false;
    }

    // ╔══════════════════════════════════════════════════════════════╗
    // ║  PERSISTENCE MODELS — saved to JSON for Act24 to replay      ║
    // ╚══════════════════════════════════════════════════════════════╝

    /// <summary>
    /// Records what was selected for ONE filter+option combination.
    /// Serialized to JSON so Act24 can replay the exact same selections.
    /// 
    /// Example:
    ///   FilterName = "Customer ID"
    ///   OptionName = "Range"
    ///   OptionType = Range
    ///   CsvFileName = "act26_CustomerID_Range.csv"
    ///   RangeFrom = "ALDRED"
    ///   RangeTo = "BENSON"
    /// </summary>
    public class FilterSelection
    {
        /// <summary>Filter name (e.g., "Customer ID")</summary>
        public string FilterName { get; set; } = "";

        /// <summary>Option name (e.g., "All", "Range", "One or More")</summary>
        public string OptionName { get; set; } = "";

        /// <summary>Option type for replay logic</summary>
        public OptionType OptionType { get; set; } = OptionType.All;

        /// <summary>CSV file name generated for this combination</summary>
        public string CsvFileName { get; set; } = "";

        /// <summary>From value for Range selections</summary>
        public string RangeFrom { get; set; } = "";

        /// <summary>To value for Range selections</summary>
        public string RangeTo { get; set; } = "";

        /// <summary>Selected values for OneOrMore selections</summary>
        public List<string> SelectedValues { get; set; } = new();
    }

    /// <summary>
    /// Container for all filter selections from one test run.
    /// Saved as JSON file to share between Act26 and Act24 runs.
    /// </summary>
    public class FilterSelectionsData
    {
        /// <summary>All filter+option combinations explored</summary>
        public List<FilterSelection> Selections { get; set; } = new();

        /// <summary>All discovered filter names (in order)</summary>
        public List<string> FilterNames { get; set; } = new();

        /// <summary>Map of filter name → list of option names (in order)</summary>
        public Dictionary<string, List<string>> FilterOptionsMap { get; set; } = new();
    }

    /// <summary>
    /// Holds the result of a CSV file comparison
    /// </summary>
    public class ComparisonResult
    {
        public bool HasMismatch { get; set; }
        public int MismatchCount { get; set; }
        public int TotalMismatchedColumns { get; set; }
        public int EmptyRowCount { get; set; }
        public List<int> EmptyRowNumbers { get; set; } = new();
        public int File1RowCount { get; set; }
        public int File2RowCount { get; set; }

        /// <summary>Filter context info for the comparison report</summary>
        public string FilterName { get; set; } = "";
        public string OptionName { get; set; } = "";
        public string OptionTypeName { get; set; } = "";
        public string SelectedValuesDescription { get; set; } = "";

        /// <summary>Markdown-formatted detailed report content</summary>
        public string MarkdownReport { get; set; } = "";

        /// <summary>Summary message for Assert.Fail</summary>
        public string FailureMessage { get; set; } = "";
    }
}
