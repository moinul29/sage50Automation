using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Sage50Automation.Config;
using Sage50Automation.Utilities;

namespace Sage50Automation.Pages
{
    /// <summary>
    /// Page Object for Excel window operations.
    /// Handles: Bringing Excel to foreground, closing, Save As CSV.
    /// 
    /// Usage:
    ///   var excelPage = new ExcelSavePage(app, automation, logger);
    ///   excelPage.SaveAsCSV("act01.csv");
    /// </summary>
    public class ExcelSavePage : BasePage
    {
        public ExcelSavePage(Application app, UIA3Automation automation, Logger logger)
            : base(app, automation, logger) { }

        /// <summary>
        /// Save the currently open Excel file as CSV.
        /// Handles: Find Excel → Bring to front → Close → More Options → Save As → CSV
        /// </summary>
        public void SaveAsCSV(string fileName)
        {
            Log.Info($"Saving as CSV: {fileName}");

            // Step 1: Find Excel window
            var excelWindow = FindExcelWindow();
            Assert.IsNotNull(excelWindow, "Excel window should be found");
            Log.Info("Found Excel window, closing it...");

            // Step 2: Bring Excel to foreground and close with Alt+F4
            BringToForegroundAndClose(excelWindow);

            // Step 3: Click "More options..." button
            ClickMoreOptions();

            // Step 4: Handle Save As dialog
            HandleSaveAsDialog(fileName);

            // Step 5: Verify file was created
            VerifyFileCreated(fileName);
        }

        /// <summary>
        /// Find the Excel window (with retry logic and taskbar fallback)
        /// </summary>
        private AutomationElement? FindExcelWindow()
        {
            Log.Info("Finding Excel window...");
            Thread.Sleep(TestConfig.MediumWaitMs);

            var excelWindow = Desktop.FindFirstDescendant(cf => cf.ByClassName("XLMAIN"));

            if (excelWindow == null)
            {
                Log.Info("Excel window not found, waiting longer...");
                Thread.Sleep(TestConfig.MediumWaitMs);
                excelWindow = Desktop.FindFirstDescendant(cf => cf.ByClassName("XLMAIN"));
            }

            // Fallback: Try clicking Excel taskbar icon
            if (excelWindow == null)
            {
                Log.Info("Excel window not visible, trying taskbar icon...");
                var taskbar = Desktop.FindFirstDescendant(cf => cf.ByClassName("Shell_TrayWnd"));
                if (taskbar != null)
                {
                    var excelIcon = taskbar.FindFirstDescendant(
                        cf => cf.ByAutomationId("Appid: Microsoft.Office.EXCEL.EXE.15"));
                    if (excelIcon != null)
                    {
                        Log.Info("Found Excel taskbar icon, clicking...");
                        excelIcon.Click();
                        Thread.Sleep(3000);
                        excelWindow = Desktop.FindFirstDescendant(cf => cf.ByClassName("XLMAIN"));
                    }
                }
            }

            return excelWindow;
        }

        /// <summary>
        /// Aggressively bring Excel window to foreground and send Alt+F4.
        /// Uses multiple activation methods to ensure the window is focused.
        /// </summary>
        private void BringToForegroundAndClose(AutomationElement excelWindow)
        {
            Log.Info("Bringing Excel to foreground...");

            // First: restore if minimized
            try
            {
                var windowPattern = excelWindow.Patterns.Window.PatternOrDefault;
                if (windowPattern != null &&
                    windowPattern.WindowVisualState == FlaUI.Core.Definitions.WindowVisualState.Minimized)
                {
                    Log.Info("Window is minimized, restoring...");
                    windowPattern.SetWindowVisualState(FlaUI.Core.Definitions.WindowVisualState.Normal);
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                Log.Info($"Warning during restore: {ex.Message}");
            }

            // Aggressive activation with retries
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    Log.Info($"Activation attempt {attempt}...");
                    excelWindow.Click();
                    Thread.Sleep(300);
                    excelWindow.SetForeground();
                    Thread.Sleep(300);
                    excelWindow.Focus();
                    Thread.Sleep(300);
                    Log.Info("Window activated");
                    break;
                }
                catch (Exception ex)
                {
                    Log.Info($"Warning during attempt {attempt}: {ex.Message}");
                    if (attempt < 3) Thread.Sleep(500);
                }
            }

            // Send Alt+F4
            Log.Info("Sending Alt+F4 to close Excel...");
            FlaUI.Core.Input.Keyboard.TypeSimultaneously(
                FlaUI.Core.WindowsAPI.VirtualKeyShort.ALT,
                FlaUI.Core.WindowsAPI.VirtualKeyShort.F4);
            Log.Info("Pressed Alt+F4 to close Excel");
            Thread.Sleep(TestConfig.MediumWaitMs);
        }

        /// <summary>
        /// Click "More options..." button in the Excel save prompt
        /// </summary>
        private void ClickMoreOptions()
        {
            Log.Info("Looking for 'More options' button...");
            var moreOptionsButton = Desktop.FindFirstDescendant(cf => cf.ByName("More options…"));
            Assert.IsNotNull(moreOptionsButton, "More options button should be found");
            Log.Info("Found 'More options' button, clicking...");
            moreOptionsButton.Click();
            Thread.Sleep(TestConfig.MediumWaitMs);
        }

        /// <summary>
        /// Handle the Save As dialog: set file name, change type to CSV, click Save
        /// </summary>
        private void HandleSaveAsDialog(string fileName)
        {
            Log.Info("Handling Save As dialog...");
            var saveAsDialog = Desktop.FindFirstDescendant(cf => cf.ByName("Save As"));
            if (saveAsDialog == null)
            {
                Thread.Sleep(TestConfig.MediumWaitMs);
                saveAsDialog = Desktop.FindFirstDescendant(cf => cf.ByName("Save As"));
            }
            Assert.IsNotNull(saveAsDialog, "Save As dialog should be found");
            Log.Info("Found Save As dialog");

            // Set file path
            string csvFilePath = Path.Combine(TestConfig.ReportsFolderPath, fileName);
            Log.Info($"Setting file path to: {csvFilePath}");

            var fileNameBox = saveAsDialog.FindFirstDescendant(cf => cf.ByAutomationId("FileNameControlHost"))
                            ?? saveAsDialog.FindFirstDescendant(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit));
            Assert.IsNotNull(fileNameBox, "File name box should be found");
            fileNameBox.AsTextBox().Text = csvFilePath;
            Log.Info("Entered file path");
            Thread.Sleep(1000);

            // Change file type to CSV
            SetFileTypeToCSV(saveAsDialog);

            // Click Save
            Log.Info("Clicking Save button...");
            var saveButton = saveAsDialog.FindFirstDescendant(cf => cf.ByName("Save"))
                            ?? saveAsDialog.FindFirstDescendant(cf => cf.ByAutomationId("1"));
            Assert.IsNotNull(saveButton, "Save button should be found");
            saveButton.Click();
            Log.Info("Clicked Save button");
            Thread.Sleep(TestConfig.MediumWaitMs);
        }

        /// <summary>
        /// Change the file type dropdown to CSV
        /// </summary>
        private void SetFileTypeToCSV(AutomationElement saveAsDialog)
        {
            Log.Info("Setting file type to CSV...");
            var fileTypeCombo = saveAsDialog.FindFirstDescendant(cf => cf.ByAutomationId("FileTypeControlHost"));
            if (fileTypeCombo == null)
            {
                var allCombos = saveAsDialog.FindAllDescendants(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.ComboBox));
                if (allCombos.Length > 0)
                    fileTypeCombo = allCombos[allCombos.Length - 1];
            }

            if (fileTypeCombo != null)
            {
                Log.Info("Found file type dropdown, expanding...");
                fileTypeCombo.AsComboBox().Expand();
                Thread.Sleep(1000);

                var csvOption = saveAsDialog.FindFirstDescendant(
                    cf => cf.ByName("CSV (Comma delimited) (*.csv)"));
                if (csvOption != null)
                {
                    csvOption.Click();
                    Log.Info("Selected CSV file type");
                    Thread.Sleep(TestConfig.ShortWaitMs);
                }
            }
        }

        /// <summary>
        /// Verify the CSV file was created successfully
        /// </summary>
        private void VerifyFileCreated(string fileName)
        {
            string csvFilePath = Path.Combine(TestConfig.ReportsFolderPath, fileName);
            Log.Info("Verifying CSV file...");

            if (File.Exists(csvFilePath))
            {
                var fileInfo = new FileInfo(csvFilePath);
                Log.Info($"SUCCESS: CSV file created at: {csvFilePath}");
                Log.Info($"File size: {fileInfo.Length} bytes");
            }
            else
            {
                Log.Info($"WARNING: CSV file not found at: {csvFilePath}");
            }
        }
    }
}
