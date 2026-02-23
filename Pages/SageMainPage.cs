using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Sage50Automation.Config;
using Sage50Automation.Utilities;

namespace Sage50Automation.Pages
{
    /// <summary>
    /// Page Object for the Sage 50 main window.
    /// Handles: Application launch → Company opening → Application close.
    /// 
    /// Usage:
    ///   var mainPage = new SageMainPage(app, automation, logger);
    ///   mainPage.WaitForMainWindow();
    ///   mainPage.OpenCompany();
    ///   // ... do test work ...
    ///   mainPage.CloseApp();
    /// </summary>
    public class SageMainPage : BasePage
    {
        public Window MainWindow { get; private set; } = null!;

        public SageMainPage(Application app, UIA3Automation automation, Logger logger)
            : base(app, automation, logger) { }

        /// <summary>
        /// Wait for the main window to load after application launch
        /// </summary>
        public void WaitForMainWindow()
        {
            Log.Info("Wait for main window to load");
            Thread.Sleep(TestConfig.AppLaunchWaitMs);

            MainWindow = App.GetMainWindow(Automation, TimeSpan.FromSeconds(TestConfig.MainWindowTimeoutSeconds));
            Assert.IsNotNull(MainWindow, "Main window should be found");
            Log.Info("Sage 50 application launched successfully");
        }

        /// <summary>
        /// Open existing company (BCS - Bellwether Garden Supply)
        /// </summary>
        public void OpenCompany()
        {
            Log.Info("Opening existing company...");
            Thread.Sleep(TestConfig.LongWaitMs);

            // Click "Open Existing Company" button
            var openButton = MainWindow.FindFirstDescendant(cf => cf.ByAutomationId("pictureBoxOpen"));
            Assert.IsNotNull(openButton, "Open Existing Company button should be found");
            openButton.Click();
            Thread.Sleep(3000);

            // Handle the "Open an Existing Company" dialog
            var dialog = MainWindow.FindFirstDescendant(cf => cf.ByName("Open an Existing Company"));
            Assert.IsNotNull(dialog, "Open an Existing Company dialog should be found");

            var okButton = dialog.FindFirstChild(cf => cf.ByAutomationId("btnOK"));
            Assert.IsNotNull(okButton, "OK button should be found");
            okButton.Click();

            Thread.Sleep(TestConfig.CompanyOpenWaitMs);
            Log.Info("Company opened successfully");
        }

        /// <summary>
        /// Close the Sage application and dispose automation resources
        /// </summary>
        public void CloseApp()
        {
            Log.Info("Closing Sage application...");
            App?.Close();
            Automation?.Dispose();
            Log.Info("Application closed");
        }
    }
}
