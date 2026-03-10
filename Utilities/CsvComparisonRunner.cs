using Microsoft.VisualStudio.TestTools.UnitTesting;
using Sage50Automation.Config;
using Sage50Automation.Models;

namespace Sage50Automation.Utilities
{
    /// <summary>
    /// Runs pairwise CSV comparisons for all filter+option combinations.
    /// Compares act26_*.csv vs act24_*.csv and generates individual HTML reports.
    /// 
    /// Usage:
    ///   var runner = new CsvComparisonRunner(log, comparer, reportGenerator);
    ///   runner.CompareAllPairs(selectionsData);
    /// </summary>
    public class CsvComparisonRunner
    {
        private readonly Logger _log;
        private readonly CsvComparer _comparer;
        private readonly HtmlReportGenerator _reportGenerator;

        public CsvComparisonRunner(Logger log, CsvComparer comparer, HtmlReportGenerator reportGenerator)
        {
            _log = log;
            _comparer = comparer;
            _reportGenerator = reportGenerator;
        }

        /// <summary>
        /// Compare each act26_Filter_Option.csv vs act24_Filter_Option.csv pair.
        /// Generates a separate HTML report for each comparison.
        /// Asserts failure if any pair mismatches.
        /// </summary>
        public void CompareAllPairs(FilterSelectionsData selectionsData)
        {
            int totalPairs = selectionsData.Selections.Count;
            int passCount = 0;
            int failCount = 0;
            var failedPairs = new List<string>();

            for (int i = 0; i < totalPairs; i++)
            {
                var sel = selectionsData.Selections[i];

                string act26Csv = SelectionsPersistence.BuildCsvFileName("act26", sel.FilterName, sel.OptionName);
                string act24Csv = SelectionsPersistence.BuildCsvFileName("act24", sel.FilterName, sel.OptionName);

                string act26Path = Path.Combine(TestConfig.ReportsFolderPath, act26Csv);
                string act24Path = Path.Combine(TestConfig.ReportsFolderPath, act24Csv);

                _log.Info($"\n--- Comparing pair [{i + 1}/{totalPairs}] ---");
                _log.Info($"  Filter: {sel.FilterName} | Option: {sel.OptionName}");
                _log.Info($"  Act26: {act26Csv}");
                _log.Info($"  Act24: {act24Csv}");

                if (!File.Exists(act26Path))
                {
                    _log.Info($"  SKIP: Act26 file not found: {act26Path}");
                    failCount++;
                    failedPairs.Add($"{sel.FilterName}/{sel.OptionName} (act26 file missing)");
                    continue;
                }
                if (!File.Exists(act24Path))
                {
                    _log.Info($"  SKIP: Act24 file not found: {act24Path}");
                    failCount++;
                    failedPairs.Add($"{sel.FilterName}/{sel.OptionName} (act24 file missing)");
                    continue;
                }

                // Compare with filter context
                var result = _comparer.Compare(act26Path, act24Path, sel);

                // Generate separate HTML report for this pair
                string reportFile = _reportGenerator.Save(result.MarkdownReport);
                _log.Info($"  HTML report: {reportFile}");

                if (result.HasMismatch)
                {
                    failCount++;
                    failedPairs.Add($"{sel.FilterName}/{sel.OptionName}");
                    _log.Info($"  RESULT: MISMATCH — {result.FailureMessage}");
                }
                else
                {
                    passCount++;
                    _log.Info($"  RESULT: MATCH");
                }
            }

            // Summary
            _log.Info($"\n========================================");
            _log.Info($"  COMPARISON SUMMARY");
            _log.Info($"  Total pairs: {totalPairs}");
            _log.Info($"  Passed: {passCount}");
            _log.Info($"  Failed: {failCount}");
            _log.Info($"========================================");

            if (failCount > 0)
            {
                string failMsg = $"CSV comparison failed for {failCount}/{totalPairs} pairs: " +
                                string.Join(", ", failedPairs);
                _log.Info(failMsg);
                Assert.Fail(failMsg);
            }
        }
    }
}
