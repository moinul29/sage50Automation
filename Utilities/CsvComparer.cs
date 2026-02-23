using Sage50Automation.Config;
using Sage50Automation.Models;

namespace Sage50Automation.Utilities
{
    /// <summary>
    /// Compares two CSV files and produces a detailed markdown report.
    /// 
    /// Usage:
    ///   var comparer = new CsvComparer(logger);
    ///   var result = comparer.Compare("act01.csv", "act02.csv");
    ///   if (result.HasMismatch) { /* handle failure */ }
    /// </summary>
    public class CsvComparer
    {
        private readonly Logger _log;

        public CsvComparer(Logger logger)
        {
            _log = logger;
        }

        /// <summary>
        /// Compare two CSV files and generate a detailed comparison result
        /// with markdown-formatted report content
        /// </summary>
        public ComparisonResult Compare(string file1Path, string file2Path)
        {
            var result = new ComparisonResult();

            // Read both CSV files
            string[] file1Lines = File.ReadAllLines(file1Path);
            string[] file2Lines = File.ReadAllLines(file2Path);
            result.File1RowCount = file1Lines.Length;
            result.File2RowCount = file2Lines.Length;

            // ===== Console Summary (3 lines) =====
            PrintConsoleSummary(file1Lines, file2Lines, result);

            // ===== Build Markdown Report =====
            var log = new System.Text.StringBuilder();
            BuildReportHeader(log, file1Path, file2Path, file1Lines, file2Lines);
            BuildEmptyRowsSection(log, file1Lines, file2Lines, result);

            // Row count check (early exit if different)
            if (file1Lines.Length != file2Lines.Length)
            {
                log.AppendLine("### ❌ Row Count Mismatch");
                log.AppendLine($"- act01.csv has **{file1Lines.Length}** rows");
                log.AppendLine($"- act02.csv has **{file2Lines.Length}** rows");
                log.AppendLine();
                result.HasMismatch = true;
                result.MarkdownReport = log.ToString();
                result.FailureMessage = $"CSV files have different number of rows. act01: {file1Lines.Length}, act02: {file2Lines.Length}";
                return result;
            }

            log.AppendLine("### ✓ Row Count Match");
            log.AppendLine($"Both files have **{file1Lines.Length}** rows");
            log.AppendLine();

            // Row-by-row comparison
            BuildRowComparison(log, file1Lines, file2Lines, result);

            // Execution logs section
            BuildExecutionLogsSection(log);

            // Comparison summary section
            BuildSummarySection(log, file1Lines, file2Lines, result);

            // Final test status
            BuildTestResultSection(log, result);

            result.MarkdownReport = log.ToString();
            return result;
        }

        /// <summary>
        /// Print 3-line console summary
        /// </summary>
        private void PrintConsoleSummary(string[] file1Lines, string[] file2Lines, ComparisonResult result)
        {
            // Row count
            int rowDiff = Math.Abs(file1Lines.Length - file2Lines.Length);
            Console.WriteLine(rowDiff > 0
                ? $"❌ Row count mismatch: {rowDiff} (act01: {file1Lines.Length}, act02: {file2Lines.Length})"
                : $"✓ Row count matches: {file1Lines.Length} rows");

            // Mismatched columns
            int minLines = Math.Min(file1Lines.Length, file2Lines.Length);
            for (int i = 0; i < minLines; i++)
            {
                if (file1Lines[i] != file2Lines[i])
                {
                    string[] cols1 = file1Lines[i].Split(',');
                    string[] cols2 = file2Lines[i].Split(',');
                    if (cols1.Length != cols2.Length)
                        result.TotalMismatchedColumns += Math.Max(cols1.Length, cols2.Length);
                    else
                        for (int j = 0; j < cols1.Length; j++)
                            if (cols1[j] != cols2[j])
                                result.TotalMismatchedColumns++;
                }
            }

            Console.WriteLine(result.TotalMismatchedColumns > 0
                ? $"❌ Total mismatched columns: {result.TotalMismatchedColumns}"
                : $"✓ Total mismatched columns: 0");

            // Empty rows
            CountEmptyRows(file1Lines, file2Lines, result);
            Console.WriteLine(result.EmptyRowCount > 0
                ? $"❌ Empty rows found: {result.EmptyRowCount}"
                : $"✓ No empty rows");
            Console.WriteLine();
        }

        /// <summary>
        /// Count empty rows across both files
        /// </summary>
        private void CountEmptyRows(string[] file1Lines, string[] file2Lines, ComparisonResult result)
        {
            for (int i = 0; i < file1Lines.Length; i++)
                if (string.IsNullOrWhiteSpace(file1Lines[i]))
                {
                    result.EmptyRowCount++;
                    result.EmptyRowNumbers.Add(i + 1);
                }

            for (int i = 0; i < file2Lines.Length; i++)
                if (string.IsNullOrWhiteSpace(file2Lines[i]) && !result.EmptyRowNumbers.Contains(i + 1))
                {
                    result.EmptyRowCount++;
                    result.EmptyRowNumbers.Add(i + 1);
                }
        }

        private void BuildReportHeader(System.Text.StringBuilder log, string file1Path, string file2Path,
            string[] file1Lines, string[] file2Lines)
        {
            log.AppendLine("# Report Comparison Analysis Result");
            log.AppendLine($"**Date:** {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            log.AppendLine();
            log.AppendLine("## Files Compared");
            log.AppendLine($"- **File 1:** `{file1Path}`");
            log.AppendLine($"- **File 2:** `{file2Path}`");
            log.AppendLine();
            log.AppendLine("## Row Count Analysis");
            log.AppendLine($"- **act01.csv:** {file1Lines.Length} rows");
            log.AppendLine($"- **act02.csv:** {file2Lines.Length} rows");
            log.AppendLine();
        }

        private void BuildEmptyRowsSection(System.Text.StringBuilder log, string[] file1Lines,
            string[] file2Lines, ComparisonResult result)
        {
            log.AppendLine("## Empty Rows Analysis");
            if (result.EmptyRowCount > 0)
            {
                log.AppendLine($"- **Total empty rows found:** {result.EmptyRowCount}");
                log.AppendLine($"- **Row numbers:** {string.Join(", ", result.EmptyRowNumbers)}");
            }
            else
            {
                log.AppendLine("- ✓ No empty rows found");
            }
            log.AppendLine();
        }

        private void BuildRowComparison(System.Text.StringBuilder log, string[] file1Lines,
            string[] file2Lines, ComparisonResult result)
        {
            log.AppendLine("## Row-by-Row Comparison");
            log.AppendLine();

            for (int i = 0; i < file1Lines.Length; i++)
            {
                if (file1Lines[i] != file2Lines[i])
                {
                    result.HasMismatch = true;
                    result.MismatchCount++;

                    log.AppendLine($"### ❌ Mismatch at Row {i + 1}");
                    log.AppendLine($"**act01.csv:** `{file1Lines[i]}`");
                    log.AppendLine();
                    log.AppendLine($"**act02.csv:** `{file2Lines[i]}`");
                    log.AppendLine();

                    string[] cols1 = file1Lines[i].Split(',');
                    string[] cols2 = file2Lines[i].Split(',');

                    if (cols1.Length != cols2.Length)
                    {
                        log.AppendLine($"**Column Count Mismatch:**");
                        log.AppendLine($"- act01: {cols1.Length} columns");
                        log.AppendLine($"- act02: {cols2.Length} columns");
                        log.AppendLine();
                    }
                    else
                    {
                        log.AppendLine("**Column Differences:**");
                        for (int j = 0; j < cols1.Length; j++)
                        {
                            if (cols1[j] != cols2[j])
                            {
                                log.AppendLine($"- **Column {j + 1}:**");
                                log.AppendLine($"  - act01: `{cols1[j]}`");
                                log.AppendLine($"  - act02: `{cols2[j]}`");
                            }
                        }
                        log.AppendLine();
                    }
                }
            }

            if (result.MismatchCount > 0)
                result.FailureMessage = $"CSV files do not match! Found {result.MismatchCount} line(s) with differences.";
        }

        private void BuildExecutionLogsSection(System.Text.StringBuilder log)
        {
            log.AppendLine("---");
            log.AppendLine();
            log.AppendLine("## Execution Logs");
            log.AppendLine();

            string allLogs = _log.GetAllLogs();

            // Split into Actian 2026 and 2024 sections
            string[] sections = allLogs.Split(
                new[] { "=== Starting Actian 2024 Test ===" },
                StringSplitOptions.None);

            if (sections.Length >= 1 && !string.IsNullOrWhiteSpace(sections[0]))
            {
                log.AppendLine("### Actian 2026 Test Execution");
                log.AppendLine();
                log.AppendLine("```");
                log.AppendLine(sections[0].TrimEnd());
                log.AppendLine("```");
                log.AppendLine();
            }

            if (sections.Length >= 2 && !string.IsNullOrWhiteSpace(sections[1]))
            {
                log.AppendLine("### Actian 2024 Test Execution");
                log.AppendLine();
                log.AppendLine("```");
                log.AppendLine("=== Starting Actian 2024 Test ===" + sections[1].TrimEnd());
                log.AppendLine("```");
                log.AppendLine();
            }
        }

        private void BuildSummarySection(System.Text.StringBuilder log, string[] file1Lines,
            string[] file2Lines, ComparisonResult result)
        {
            log.AppendLine("---");
            log.AppendLine();
            log.AppendLine("## Comparison Summary");
            log.AppendLine();

            // Row count
            if (file1Lines.Length == file2Lines.Length)
                log.AppendLine($"✓ Row count matches: {file1Lines.Length} rows");
            else
            {
                log.AppendLine($"❌ Row count mismatch: {Math.Abs(file1Lines.Length - file2Lines.Length)}");
                log.AppendLine($"   - act01.csv: {file1Lines.Length} rows");
                log.AppendLine($"   - act02.csv: {file2Lines.Length} rows");
            }

            // Columns
            log.AppendLine(result.TotalMismatchedColumns > 0
                ? $"❌ Total mismatched columns: {result.TotalMismatchedColumns}"
                : $"✓ Total mismatched columns: 0");

            // Empty rows
            log.AppendLine(result.EmptyRowCount > 0
                ? $"❌ Empty rows found: {result.EmptyRowCount}"
                : $"✓ No empty rows");
        }

        private void BuildTestResultSection(System.Text.StringBuilder log, ComparisonResult result)
        {
            log.AppendLine();
            log.AppendLine("---");
            log.AppendLine();
            log.AppendLine("## Test Result");
            log.AppendLine();

            if (result.HasMismatch)
            {
                _log.Info("=== Actian 2024 Test FAILED - CSV files do not match ===");
                log.AppendLine("❌ **Actian 2024 Test FAILED** - CSV files do not match");
            }
            else
            {
                _log.Info("=== Actian 2024 Test Completed Successfully! ===");
                log.AppendLine("✓ **Actian 2024 Test Completed Successfully!**");
            }
        }
    }
}
