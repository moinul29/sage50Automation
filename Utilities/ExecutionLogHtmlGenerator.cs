using Sage50Automation.Config;

namespace Sage50Automation.Utilities
{
    /// <summary>
    /// Generates a single styled HTML file from the automation execution log.
    /// 
    /// Features:
    ///   - Color-coded log lines (errors, warnings, successes, sections)
    ///   - Screenshots embedded inline as base64 images (captured on failure)
    ///   - Separate sections for Act26 and Act24 test phases
    ///   - Professional dark header with metadata
    /// 
    /// Usage:
    ///   var generator = new ExecutionLogHtmlGenerator();
    ///   generator.SaveToFile(logContent);
    /// </summary>
    public class ExecutionLogHtmlGenerator
    {
        /// <summary>
        /// Generate execution log HTML and save to the project directory.
        /// Returns the generated file name.
        /// </summary>
        public string SaveToFile(string logContent)
        {
            string htmlContent = Generate(logContent);
            string fileName = TestConfig.ExecutionLogHtmlFileName;
            string filePath = Path.Combine(TestConfig.ProjectDirectory, fileName);

            File.WriteAllText(filePath, htmlContent);
            Console.WriteLine($"Execution log HTML saved: {filePath}");

            return fileName;
        }

        /// <summary>
        /// Convert raw log text into a styled HTML document with embedded screenshots.
        /// </summary>
        public string Generate(string logContent)
        {
            var html = new System.Text.StringBuilder();
            string screenshotsDir = Path.Combine(Path.GetTempPath(), TestConfig.ScreenshotsTempFolder);

            BuildHtmlHeader(html);

            string[] lines = logContent.Split('\n');
            foreach (string line in lines)
            {
                string trimmed = line.TrimEnd('\r');

                if (string.IsNullOrWhiteSpace(trimmed))
                {
                    html.AppendLine("        <div class=\"log-blank\">&nbsp;</div>");
                    continue;
                }

                // Screenshot marker: [SCREENSHOT:filename.png] context text
                if (trimmed.Contains("[SCREENSHOT:"))
                {
                    RenderScreenshot(html, trimmed, screenshotsDir);
                    continue;
                }

                string cssClass = ClassifyLine(trimmed);
                string encoded = System.Net.WebUtility.HtmlEncode(trimmed);
                html.AppendLine($"        <div class=\"log-line {cssClass}\">{encoded}</div>");
            }

            BuildHtmlFooter(html);
            return html.ToString();
        }

        // ==================================================
        //              LINE CLASSIFICATION
        // ==================================================

        /// <summary>
        /// Classify a log line to determine its CSS style class.
        /// </summary>
        private string ClassifyLine(string line)
        {
            // Errors
            if (line.Contains("ERROR") || line.Contains("FATAL") ||
                line.Contains("FAILED") || line.Contains("Assert."))
                return "log-error";

            // Warnings
            if (line.Contains("WARNING"))
                return "log-warning";

            // Success markers
            if (line.Contains("Completed Successfully") || line.Contains("SUCCESS") ||
                line.Contains("MATCH") || line.Contains("Passed:"))
                return "log-success";

            // Section headers (=== ... ===)
            if (line.Contains("===") && line.TrimStart().StartsWith("==="))
                return "log-section";

            // Step headers (STEP, Step, FILTER, OPTION, REPLAY)
            if (line.Contains("STEP ") || line.Contains("Step ") ||
                line.TrimStart().StartsWith("FILTER [") || line.TrimStart().StartsWith("OPTION [") ||
                line.TrimStart().StartsWith("REPLAY ["))
                return "log-step";

            // Separator lines
            if (line.TrimStart().StartsWith("======"))
                return "log-separator";

            // Completed items
            if (line.TrimStart().StartsWith("Completed:") || line.Contains("Completed:"))
                return "log-complete";

            return "";
        }

        // ==================================================
        //              SCREENSHOT EMBEDDING
        // ==================================================

        /// <summary>
        /// Parse a screenshot marker line and embed the image as base64 in the HTML.
        /// Marker format: [SCREENSHOT:filename.png] context description
        /// </summary>
        private void RenderScreenshot(System.Text.StringBuilder html, string line, string screenshotsDir)
        {
            try
            {
                // Parse: ... [SCREENSHOT:filename.png] context ...
                int startIdx = line.IndexOf("[SCREENSHOT:");
                int endIdx = line.IndexOf(']', startIdx);

                if (startIdx < 0 || endIdx < 0) return;

                string fileName = line.Substring(startIdx + 12, endIdx - startIdx - 12);
                string context = line.Substring(endIdx + 1).Trim();
                string filePath = Path.Combine(screenshotsDir, fileName);

                html.AppendLine("        <div class=\"screenshot-container\">");
                html.AppendLine($"            <div class=\"screenshot-label\">&#128247; Screenshot: {System.Net.WebUtility.HtmlEncode(context)}</div>");

                if (File.Exists(filePath))
                {
                    byte[] imageBytes = File.ReadAllBytes(filePath);
                    string base64 = Convert.ToBase64String(imageBytes);
                    html.AppendLine($"            <img src=\"data:image/png;base64,{base64}\" alt=\"Screenshot: {System.Net.WebUtility.HtmlEncode(context)}\" />");
                }
                else
                {
                    html.AppendLine($"            <div class=\"screenshot-missing\">Screenshot file not found: {System.Net.WebUtility.HtmlEncode(fileName)}</div>");
                }

                html.AppendLine("        </div>");
            }
            catch
            {
                string encoded = System.Net.WebUtility.HtmlEncode(line);
                html.AppendLine($"        <div class=\"log-line log-warning\">{encoded}</div>");
            }
        }

        // ==================================================
        //              HTML STRUCTURE
        // ==================================================

        private void BuildHtmlHeader(System.Text.StringBuilder html)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("    <meta charset=\"UTF-8\">");
            html.AppendLine("    <title>Sage 50 Automation - Execution Log</title>");
            html.AppendLine("    <style>");
            html.AppendLine("        * { box-sizing: border-box; margin: 0; padding: 0; }");
            html.AppendLine("        body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; }");
            html.AppendLine();
            html.AppendLine("        /* Header */");
            html.AppendLine("        .header { background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%); color: white; padding: 25px 30px; }");
            html.AppendLine("        .header h1 { font-size: 22px; font-weight: 600; margin-bottom: 6px; }");
            html.AppendLine("        .header .meta { color: #94a3b8; font-size: 13px; }");
            html.AppendLine();
            html.AppendLine("        /* Container */");
            html.AppendLine("        .container { max-width: 1400px; margin: 20px auto; padding: 0 20px 40px; }");
            html.AppendLine();
            html.AppendLine("        /* Log lines */");
            html.AppendLine("        .log-line { font-family: 'Cascadia Code', 'Consolas', 'Courier New', monospace; font-size: 12.5px; padding: 3px 12px; border-left: 3px solid transparent; line-height: 1.6; white-space: pre-wrap; word-break: break-word; }");
            html.AppendLine("        .log-line:hover { background: #e2e8f0; }");
            html.AppendLine("        .log-blank { height: 8px; }");
            html.AppendLine();
            html.AppendLine("        /* Color classes */");
            html.AppendLine("        .log-error { color: #dc2626; background: #fef2f2; border-left-color: #ef4444; font-weight: 600; }");
            html.AppendLine("        .log-warning { color: #d97706; background: #fffbeb; border-left-color: #f59e0b; }");
            html.AppendLine("        .log-success { color: #16a34a; background: #f0fdf4; border-left-color: #22c55e; font-weight: 600; }");
            html.AppendLine("        .log-section { color: #1e40af; background: #eff6ff; padding: 8px 12px; margin-top: 12px; font-weight: 700; font-size: 13px; border-left-color: #3b82f6; border-left-width: 4px; }");
            html.AppendLine("        .log-step { color: #7c3aed; font-weight: 600; }");
            html.AppendLine("        .log-separator { color: #94a3b8; font-size: 11px; }");
            html.AppendLine("        .log-complete { color: #059669; }");
            html.AppendLine();
            html.AppendLine("        /* Screenshots */");
            html.AppendLine("        .screenshot-container { margin: 12px 0; padding: 16px; background: #fef2f2; border: 2px solid #fca5a5; border-radius: 8px; }");
            html.AppendLine("        .screenshot-label { color: #b91c1c; font-weight: 700; font-size: 14px; margin-bottom: 12px; font-family: 'Segoe UI', Arial, sans-serif; }");
            html.AppendLine("        .screenshot-container img { max-width: 100%; border: 1px solid #e5e7eb; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }");
            html.AppendLine("        .screenshot-container img:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.2); }");
            html.AppendLine("        .screenshot-missing { color: #9ca3af; font-style: italic; font-family: 'Segoe UI', Arial, sans-serif; }");
            html.AppendLine("    </style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");
            html.AppendLine("    <div class=\"header\">");
            html.AppendLine("        <h1>Sage 50 Automation - Execution Log</h1>");
            html.AppendLine($"        <div class=\"meta\">Generated: {timestamp}</div>");
            html.AppendLine("    </div>");
            html.AppendLine("    <div class=\"container\">");
        }

        private void BuildHtmlFooter(System.Text.StringBuilder html)
        {
            html.AppendLine("    </div>");
            html.AppendLine("</body>");
            html.AppendLine("</html>");
        }
    }
}
