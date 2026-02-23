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

        /// <summary>Markdown-formatted detailed report content</summary>
        public string MarkdownReport { get; set; } = "";

        /// <summary>Summary message for Assert.Fail</summary>
        public string FailureMessage { get; set; } = "";
    }
}
