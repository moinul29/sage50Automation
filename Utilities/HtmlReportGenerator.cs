using Sage50Automation.Config;

namespace Sage50Automation.Utilities
{
    /// <summary>
    /// Converts markdown-formatted report content into styled HTML files.
    /// 
    /// Usage:
    ///   var generator = new HtmlReportGenerator();
    ///   string fileName = generator.Save(markdownContent);
    /// </summary>
    public class HtmlReportGenerator
    {
        /// <summary>
        /// Save markdown content as a styled HTML file.
        /// Automatically increments filename if file already exists
        /// (reportAnalysisResult.html → _01.html → _02.html).
        /// Returns the generated file name.
        /// </summary>
        public string Save(string markdownContent)
        {
            string htmlFileName = GetNextFileName();
            string filePath = Path.Combine(TestConfig.ProjectDirectory, htmlFileName);
            string htmlContent = ConvertMarkdownToHtml(markdownContent);

            Console.WriteLine($"Saving HTML comparison result to: {filePath}");
            File.WriteAllText(filePath, htmlContent);
            Console.WriteLine($"HTML comparison result saved successfully: {htmlFileName}");

            return htmlFileName;
        }

        /// <summary>
        /// Determine the next available file name with sequential numbering
        /// </summary>
        private string GetNextFileName()
        {
            string baseName = TestConfig.HtmlReportBaseFileName;
            string fileName = $"{baseName}.html";

            if (File.Exists(Path.Combine(TestConfig.ProjectDirectory, fileName)))
            {
                int counter = 1;
                while (true)
                {
                    fileName = $"{baseName}_{counter:D2}.html";
                    if (!File.Exists(Path.Combine(TestConfig.ProjectDirectory, fileName)))
                        break;
                    counter++;
                }
            }

            return fileName;
        }

        /// <summary>
        /// Convert markdown-styled content to HTML with professional styling
        /// </summary>
        private string ConvertMarkdownToHtml(string markdownContent)
        {
            var html = new System.Text.StringBuilder();

            // HTML header with CSS
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("    <meta charset=\"UTF-8\">");
            html.AppendLine("    <title>Report Comparison Analysis Result</title>");
            html.AppendLine("    <style>");
            html.AppendLine("        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }");
            html.AppendLine("        .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 30px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }");
            html.AppendLine("        h1 { color: #333; border-bottom: 3px solid #4CAF50; padding-bottom: 10px; }");
            html.AppendLine("        h2 { color: #555; margin-top: 30px; border-bottom: 2px solid #ddd; padding-bottom: 8px; }");
            html.AppendLine("        h3 { color: #666; margin-top: 20px; }");
            html.AppendLine("        .success { color: #4CAF50; font-weight: bold; }");
            html.AppendLine("        .error { color: #f44336; font-weight: bold; }");
            html.AppendLine("        .info { background-color: #e3f2fd; padding: 10px; border-left: 4px solid #2196F3; margin: 10px 0; }");
            html.AppendLine("        .mismatch { background-color: #ffebee; padding: 15px; border-left: 4px solid #f44336; margin: 15px 0; }");
            html.AppendLine("        .match { background-color: #e8f5e9; padding: 10px; border-left: 4px solid #4CAF50; margin: 10px 0; }");
            html.AppendLine("        code { background-color: #f5f5f5; padding: 2px 6px; border-radius: 3px; font-family: 'Courier New', monospace; }");
            html.AppendLine("        ul { line-height: 1.8; }");
            html.AppendLine("        .timestamp { color: #888; font-size: 0.9em; }");
            html.AppendLine("        table { border-collapse: collapse; width: 100%; margin: 15px 0; }");
            html.AppendLine("        th, td { padding: 12px; text-align: left; border: 1px solid #ddd; }");
            html.AppendLine("        th { background-color: #4CAF50; color: white; }");
            html.AppendLine("        hr { border: 0; border-top: 1px solid #ddd; margin: 30px 0; }");
            html.AppendLine("        pre { background-color: #f8f8f8; padding: 15px; border-radius: 5px; overflow-x: auto; border: 1px solid #e0e0e0; }");
            html.AppendLine("    </style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");
            html.AppendLine("    <div class=\"container\">");

            // Parse and convert each markdown line to HTML
            string[] lines = markdownContent.Split('\n');
            bool inCodeBlock = false;

            foreach (string line in lines)
            {
                string trimmed = line.Trim();

                // Handle code blocks
                if (trimmed.StartsWith("```"))
                {
                    if (inCodeBlock)
                    {
                        html.AppendLine("</pre>");
                        inCodeBlock = false;
                    }
                    else
                    {
                        html.AppendLine("<pre>");
                        inCodeBlock = true;
                    }
                    continue;
                }

                if (inCodeBlock)
                {
                    html.AppendLine($"        {System.Net.WebUtility.HtmlEncode(trimmed)}");
                    continue;
                }

                // Markdown to HTML conversions
                if (trimmed.StartsWith("# "))
                    html.AppendLine($"        <h1>{trimmed.Substring(2)}</h1>");
                else if (trimmed.StartsWith("## "))
                    html.AppendLine($"        <h2>{trimmed.Substring(3)}</h2>");
                else if (trimmed.StartsWith("### ✓"))
                    html.AppendLine($"        <h3 class=\"success\">{trimmed.Substring(4)}</h3>");
                else if (trimmed.StartsWith("### ❌"))
                    html.AppendLine($"        <h3 class=\"error\">{trimmed.Substring(4)}</h3>");
                else if (trimmed.StartsWith("### "))
                    html.AppendLine($"        <h3>{trimmed.Substring(4)}</h3>");
                else if (trimmed.StartsWith("**Date:**"))
                    html.AppendLine($"        <p class=\"timestamp\">{trimmed.Replace("**", "").Replace("Date:", "<strong>Date:</strong>")}</p>");
                else if (trimmed.StartsWith("- ✓"))
                    html.AppendLine($"        <div class=\"match\">✓ {trimmed.Substring(4)}</div>");
                else if (trimmed.StartsWith("- ❌"))
                    html.AppendLine($"        <div class=\"error\">{trimmed}</div>");
                else if (trimmed.StartsWith("- **"))
                {
                    string content = trimmed.Substring(2).Replace("**", "");
                    if (trimmed.Contains("File 1:") || trimmed.Contains("File 2:") ||
                        trimmed.Contains("act01.csv:") || trimmed.Contains("act02.csv:"))
                        html.AppendLine($"        <div class=\"info\">{content}</div>");
                    else
                        html.AppendLine($"        <p>{content}</p>");
                }
                else if (trimmed.StartsWith("- "))
                    html.AppendLine($"        <ul><li>{trimmed.Substring(2)}</li></ul>");
                else if (trimmed.StartsWith("✓"))
                    html.AppendLine($"        <div class=\"match\">{trimmed}</div>");
                else if (trimmed.StartsWith("❌"))
                    html.AppendLine($"        <div class=\"error\">{trimmed}</div>");
                else if (trimmed.StartsWith("**") && trimmed.Contains("**"))
                {
                    string content = trimmed.Replace("**", "<strong>").Replace("**", "</strong>");
                    if (trimmed.Contains("act01.csv:") || trimmed.Contains("act02.csv:"))
                        html.AppendLine($"        <div class=\"mismatch\">{content}</div>");
                    else
                        html.AppendLine($"        <p>{content}</p>");
                }
                else if (trimmed.StartsWith("---"))
                    html.AppendLine("        <hr>");
                else if (!string.IsNullOrWhiteSpace(trimmed))
                {
                    string content = trimmed.Replace("`", "<code>").Replace("`", "</code>");
                    html.AppendLine($"        <p>{content}</p>");
                }
            }

            if (inCodeBlock)
                html.AppendLine("</pre>");

            html.AppendLine("    </div>");
            html.AppendLine("</body>");
            html.AppendLine("</html>");

            return html.ToString();
        }
    }
}
