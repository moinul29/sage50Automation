using System.Text.Json;
using Sage50Automation.Config;
using Sage50Automation.Models;

namespace Sage50Automation.Utilities
{
    /// <summary>
    /// Saves and loads filter selections to/from JSON file.
    /// Act26 saves its selections → Act24 reads and replays them.
    /// 
    /// File location: Reports folder / filter_selections.json
    /// </summary>
    public static class SelectionsPersistence
    {
        private static readonly JsonSerializerOptions _jsonOptions = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        /// <summary>
        /// Save filter selections to JSON file in the Reports folder
        /// </summary>
        public static void Save(FilterSelectionsData data)
        {
            string filePath = Path.Combine(TestConfig.ReportsFolderPath, TestConfig.SelectionsFileName);
            string json = JsonSerializer.Serialize(data, _jsonOptions);
            File.WriteAllText(filePath, json);
        }

        /// <summary>
        /// Load filter selections from JSON file in the Reports folder.
        /// Returns null if file doesn't exist.
        /// </summary>
        public static FilterSelectionsData? Load()
        {
            string filePath = Path.Combine(TestConfig.ReportsFolderPath, TestConfig.SelectionsFileName);
            if (!File.Exists(filePath))
                return null;

            string json = File.ReadAllText(filePath);
            return JsonSerializer.Deserialize<FilterSelectionsData>(json, _jsonOptions);
        }

        /// <summary>
        /// Generate a safe CSV file name from filter name, option name, and prefix.
        /// Example: "act26_CustomerID_Range.csv"
        /// </summary>
        public static string BuildCsvFileName(string prefix, string filterName, string optionName)
        {
            // Remove special characters, replace spaces with empty
            string safeFilter = MakeSafe(filterName);
            string safeOption = MakeSafe(optionName);
            return $"{prefix}_{safeFilter}_{safeOption}.csv";
        }

        /// <summary>
        /// Make a string safe for file names
        /// </summary>
        private static string MakeSafe(string input)
        {
            // Remove dots, colons, slashes etc. Keep only letters, digits, underscores
            var safe = new string(input
                .Where(c => char.IsLetterOrDigit(c) || c == '_' || c == ' ')
                .ToArray());
            // Replace spaces, collapse multiple underscores
            safe = safe.Replace(" ", "");
            return safe;
        }
    }
}
