using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Sage50Automation.Config;
using Sage50Automation.Models;
using Sage50Automation.Utilities;

namespace Sage50Automation.Pages
{
    /// <summary>
    /// Page Object for the "Modify Report" dialog (filters and options).
    /// Handles all filter discovery, option selection, range values, and multi-select operations.
    /// 
    /// This page is created by ReportViewerPage.OpenModifyReportDialog().
    /// </summary>
    public class ModifyReportPage : BasePage
    {
        private AutomationElement? _window;
        private readonly string _reportWindowTitle;

        public ModifyReportPage(Application app, UIA3Automation automation, Logger logger, string reportWindowTitle)
            : base(app, automation, logger)
        {
            _reportWindowTitle = reportWindowTitle;
            FindWindow();
        }

        // ╔══════════════════════════════════════════════════╗
        // ║              WINDOW MANAGEMENT                    ║
        // ╚══════════════════════════════════════════════════╝

        /// <summary>
        /// Find the Modify Report dialog window
        /// </summary>
        private void FindWindow()
        {
            Log.Info("Finding Options/Filters window...");

            _window = Desktop.FindFirstDescendant(cf => cf.ByName($"Modify Report - {_reportWindowTitle}"));

            if (_window == null)
                _window = Desktop.FindFirstDescendant(cf => cf.ByName($"{_reportWindowTitle} - Filter"));
            if (_window == null)
                _window = Desktop.FindFirstDescendant(cf => cf.ByName("Options"));

            if (_window == null)
            {
                var allWindows = Desktop.FindAllDescendants(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));
                foreach (var win in allWindows)
                {
                    if (win.Name != null &&
                        (win.Name.Contains("Filter") || win.Name.Contains("Options") || win.Name.Contains("Modify")))
                    {
                        _window = win;
                        Log.Info($"Found options window by search: '{win.Name}'");
                        break;
                    }
                }
            }

            if (_window != null)
                Log.Info($"Found Options window: {_window.Name}");
        }

        /// <summary>
        /// Get the window element (asserts it exists)
        /// </summary>
        private AutomationElement Window
        {
            get
            {
                if (_window == null)
                    FindWindow();
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNotNull(
                    _window, "Options/Filters window should be found");
                return _window!;
            }
        }

        // ╔══════════════════════════════════════════════════╗
        // ║              FILTER DISCOVERY                     ║
        // ╚══════════════════════════════════════════════════╝

        /// <summary>
        /// Discover all filter items dynamically from the UI.
        /// Tries multiple approaches: ListBox → TreeView → DataGrid → Table
        /// </summary>
        public List<FilterInfo> DiscoverFilters()
        {
            var filters = new List<FilterInfo>();
            Log.Info("  Searching for filters in UI...");

            // Method 1: ListBox
            var filterList = Window.FindFirstDescendant(
                cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.List));
            if (filterList != null)
            {
                Log.Info("  Found ListBox, extracting items...");
                var items = filterList.FindAllDescendants(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.ListItem));
                foreach (var item in items)
                {
                    if (!string.IsNullOrEmpty(item.Name))
                    {
                        filters.Add(new FilterInfo { Name = item.Name, Element = item });
                        Log.Info($"    Extracted filter: '{item.Name}'");
                    }
                }
            }

            // Method 2: TreeView
            if (filters.Count == 0)
            {
                var tree = Window.FindFirstDescendant(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Tree));
                if (tree != null)
                {
                    Log.Info("  Found TreeView, extracting items...");
                    var items = tree.FindAllDescendants(
                        cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.TreeItem));
                    foreach (var item in items)
                        if (!string.IsNullOrEmpty(item.Name))
                        {
                            filters.Add(new FilterInfo { Name = item.Name, Element = item });
                            Log.Info($"    Extracted filter: '{item.Name}'");
                        }
                }
            }

            // Method 3: DataGrid
            if (filters.Count == 0)
            {
                var grid = Window.FindFirstDescendant(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.DataGrid));
                if (grid != null)
                {
                    Log.Info("  Found DataGrid, extracting rows...");
                    var rows = grid.FindAllDescendants(
                        cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.DataItem));
                    foreach (var row in rows)
                        if (!string.IsNullOrEmpty(row.Name))
                        {
                            filters.Add(new FilterInfo { Name = row.Name, Element = row });
                            Log.Info($"    Extracted filter: '{row.Name}'");
                        }
                }
            }

            // Method 4: Table
            if (filters.Count == 0)
            {
                var table = Window.FindFirstDescendant(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Table));
                if (table != null)
                {
                    Log.Info("  Found Table, extracting rows...");
                    var rows = table.FindAllChildren();
                    foreach (var row in rows)
                        if (!string.IsNullOrEmpty(row.Name) &&
                            row.ControlType != FlaUI.Core.Definitions.ControlType.Header)
                        {
                            filters.Add(new FilterInfo { Name = row.Name, Element = row });
                            Log.Info($"    Extracted filter: '{row.Name}'");
                        }
                }
            }

            Log.Info($"  Total filters extracted: {filters.Count}");
            return filters;
        }

        /// <summary>
        /// Re-find a filter element by name (to avoid stale element references)
        /// </summary>
        public AutomationElement? RefindFilter(string filterName)
        {
            var filterList = Window.FindFirstDescendant(
                cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.List));
            if (filterList != null)
            {
                var items = filterList.FindAllDescendants(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.ListItem));
                foreach (var item in items)
                    if (item.Name == filterName)
                        return item;
            }
            return Window.FindFirstDescendant(cf => cf.ByName(filterName));
        }

        // ╔══════════════════════════════════════════════════╗
        // ║              OPTION DISCOVERY & SELECTION         ║
        // ╚══════════════════════════════════════════════════╝

        /// <summary>
        /// Discover all options (radio buttons) available for a specific filter
        /// </summary>
        public List<OptionInfo> DiscoverFilterOptions(FilterInfo filter)
        {
            var options = new List<OptionInfo>();
            Log.Info($"    Searching for options for filter '{filter.Name}'...");

            // Method 1: Radio buttons
            var radioButtons = Window.FindAllDescendants(
                cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.RadioButton));
            foreach (var radio in radioButtons)
            {
                if (string.IsNullOrEmpty(radio.Name)) continue;
                try
                {
                    if (!radio.IsEnabled || radio.IsOffscreen) continue;
                }
                catch { }

                var type = DetermineOptionType(radio.Name);
                options.Add(new OptionInfo { Name = radio.Name, Element = radio, OptionType = type });
                Log.Info($"      Found radio option: '{radio.Name}' (Type: {type})");
            }

            // Method 2: ComboBox options
            if (options.Count == 0)
            {
                var combos = Window.FindAllDescendants(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.ComboBox));
                foreach (var combo in combos)
                {
                    try
                    {
                        combo.AsComboBox().Expand();
                        Thread.Sleep(300);
                        var items = combo.FindAllDescendants(
                            cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.ListItem));
                        foreach (var item in items)
                        {
                            if (!string.IsNullOrEmpty(item.Name) && !options.Any(o => o.Name == item.Name))
                            {
                                var type = DetermineOptionType(item.Name);
                                options.Add(new OptionInfo { Name = item.Name, Element = item, OptionType = type });
                                Log.Info($"      Found combo option: '{item.Name}' (Type: {type})");
                            }
                        }
                        combo.AsComboBox().Collapse();
                    }
                    catch { }
                }
            }

            Log.Info($"    Total options for '{filter.Name}': {options.Count}");
            return options;
        }

        /// <summary>
        /// Determine option type from its name
        /// </summary>
        private OptionType DetermineOptionType(string optionName)
        {
            string lower = optionName.ToLower();
            if (lower.Contains("range")) return OptionType.Range;
            if (lower.Contains("one or more") || lower.Contains("selected") || lower.Contains("specific"))
                return OptionType.OneOrMore;
            return OptionType.All;
        }

        /// <summary>
        /// Select an option by clicking its element
        /// </summary>
        public void SelectOption(OptionInfo option)
        {
            Log.Info($"      Selecting option: {option.Name}");
            try
            {
                var fresh = Window.FindFirstDescendant(cf => cf.ByName(option.Name));
                if (fresh != null)
                {
                    fresh.Click();
                    Thread.Sleep(300);
                    Log.Info($"      Clicked option: {option.Name}");
                }
                else
                {
                    Log.Info($"      WARNING: Could not find element '{option.Name}'");
                }
            }
            catch (Exception ex)
            {
                Log.Info($"      ERROR selecting '{option.Name}': {ex.Message}");
            }
        }

        // ╔══════════════════════════════════════════════════╗
        // ║              RANGE VALUE SELECTION                ║
        // ╚══════════════════════════════════════════════════╝

        /// <summary>
        /// Discover and select range values (From and To dropdowns).
        /// From pane: m_sde_From or m_autofill_From
        /// To pane: m_sde_To or m_autofill_To
        /// </summary>
        public RangeValues SelectRangeValues()
        {
            var rangeValues = new RangeValues();
            Log.Info("      Discovering range values (From/To)...");

            var (fromPane, toPane) = FindFromToPanes();

            AutomationElement? fromDropdown = fromPane?.FindFirstDescendant(
                cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button));
            AutomationElement? toDropdown = toPane?.FindFirstDescendant(
                cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button));

            Log.Info($"        From dropdown: {fromDropdown != null}, To dropdown: {toDropdown != null}");

            // Select From value (first item)
            if (fromDropdown != null)
            {
                Log.Info("        Selecting From value (1st item)...");
                fromDropdown.Click();
                Thread.Sleep(500);
                FlaUI.Core.Input.Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.DOWN);
                Thread.Sleep(200);
                FlaUI.Core.Input.Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.ENTER);
                Thread.Sleep(500);
                Log.Info("        From value selected");
            }

            // Select To value (second item)
            if (toDropdown != null)
            {
                Log.Info("        Selecting To value (2nd item)...");
                toDropdown.Click();
                Thread.Sleep(500);
                FlaUI.Core.Input.Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.DOWN);
                Thread.Sleep(200);
                FlaUI.Core.Input.Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.DOWN);
                Thread.Sleep(200);
                FlaUI.Core.Input.Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.ENTER);
                Thread.Sleep(500);
                Log.Info("        To value selected");
            }
            else
            {
                Log.Info("        WARNING: To dropdown not found");
            }

            // Extract and display selected values
            var (fromValue, toValue) = ExtractRangeFieldValues();
            rangeValues.FromValues.Add(fromValue);
            rangeValues.ToValues.Add(toValue);

            Log.Info("        Range selection completed");
            return rangeValues;
        }

        /// <summary>
        /// Find From and To panes in the dialog
        /// </summary>
        private (AutomationElement? fromPane, AutomationElement? toPane) FindFromToPanes()
        {
            var allPanes = Window.FindAllDescendants(
                cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Pane));

            AutomationElement? fromPane = null;
            AutomationElement? toPane = null;

            foreach (var pane in allPanes)
            {
                var id = pane.AutomationId ?? "";
                if ((id == "m_sde_From" || id == "m_autofill_From") && fromPane == null)
                {
                    fromPane = pane;
                    break;
                }
            }

            if (fromPane != null)
            {
                var parent = fromPane.Parent;
                if (parent != null)
                {
                    var siblings = parent.FindAllChildren(
                        cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Pane));
                    foreach (var sib in siblings)
                    {
                        var id = sib.AutomationId ?? "";
                        if ((id == "m_sde_To" || id == "m_autofill_To") && toPane == null)
                        {
                            toPane = sib;
                            break;
                        }
                    }
                }
            }

            return (fromPane, toPane);
        }

        /// <summary>
        /// Extract the actual text values from From/To input fields
        /// </summary>
        private (string fromValue, string toValue) ExtractRangeFieldValues()
        {
            string fromValue = "";
            string toValue = "";

            Log.Info("        Extracting From/To field values...");

            var (fromPane, toPane) = FindFromToPanes();

            fromValue = ExtractPaneValue(fromPane);
            toValue = ExtractPaneValue(toPane);

            Log.Info($"        ========================================");
            Log.Info($"        SELECTED FROM VALUE: '{fromValue}'");
            Log.Info($"        SELECTED TO VALUE: '{toValue}'");
            Log.Info($"        ========================================");

            return (fromValue, toValue);
        }

        /// <summary>
        /// Extract text value from a pane using multiple strategies
        /// </summary>
        private string ExtractPaneValue(AutomationElement? pane)
        {
            if (pane == null) return "";

            string value = "";

            // Try 1: ValuePattern
            var vp = pane.Patterns.Value.PatternOrDefault;
            if (vp != null)
                try { value = vp.Value.Value ?? ""; } catch { }

            // Try 2: LegacyIAccessible
            if (string.IsNullOrEmpty(value))
            {
                var lp = pane.Patterns.LegacyIAccessible.PatternOrDefault;
                if (lp != null)
                    try { value = lp.Value.Value ?? ""; } catch { }
            }

            // Try 3: Child elements
            if (string.IsNullOrEmpty(value))
            {
                var children = pane.FindAllChildren();
                foreach (var child in children)
                {
                    var cvp = child.Patterns.Value.PatternOrDefault;
                    if (cvp != null)
                        try
                        {
                            var v = cvp.Value.Value ?? "";
                            if (!string.IsNullOrEmpty(v)) { value = v; break; }
                        }
                        catch { }

                    if (string.IsNullOrEmpty(value))
                    {
                        var clp = child.Patterns.LegacyIAccessible.PatternOrDefault;
                        if (clp != null)
                            try
                            {
                                var v = clp.Value.Value ?? "";
                                if (!string.IsNullOrEmpty(v)) { value = v; break; }
                            }
                            catch { }
                    }

                    if (string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(child.Name))
                    {
                        value = child.Name;
                        break;
                    }
                }
            }

            // Try 4: Pane name
            if (string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(pane.Name))
                value = pane.Name;

            return value;
        }

        // ╔══════════════════════════════════════════════════╗
        // ║              MULTI-SELECT VALUES                  ║
        // ╚══════════════════════════════════════════════════╝

        /// <summary>
        /// Discover and select multiple values (for "One or More" option).
        /// Finds checkboxes in the m_multiIDSelector pane and selects up to maxToSelect.
        /// </summary>
        public List<string> SelectMultipleValues(int maxToSelect)
        {
            var selectableValues = DiscoverSelectableValues();
            var selectedValues = new List<string>();
            int count = 0;

            Log.Info($"      Selecting up to {maxToSelect} values...");

            foreach (var val in selectableValues)
            {
                if (count >= maxToSelect) break;
                try
                {
                    if (val.IsCheckBox)
                        val.Element.AsCheckBox().IsChecked = true;
                    else
                        val.Element.Click();

                    selectedValues.Add(val.Name);
                    Log.Info($"        Selected: '{val.Name}'");
                    count++;
                }
                catch (Exception ex)
                {
                    Log.Info($"        WARNING: Could not select '{val.Name}': {ex.Message}");
                }
            }

            Log.Info($"      ========================================");
            Log.Info($"      SELECTED VALUES (One or More):");
            foreach (var v in selectedValues)
                Log.Info($"        - '{v}'");
            Log.Info($"      ========================================");

            return selectedValues;
        }

        /// <summary>
        /// Discover selectable values from the m_multiIDSelector pane
        /// </summary>
        private List<SelectableValue> DiscoverSelectableValues()
        {
            var values = new List<SelectableValue>();
            Log.Info("      Discovering selectable values from m_multiIDSelector...");

            var selector = Window.FindFirstDescendant(cf => cf.ByAutomationId("m_multiIDSelector"));

            if (selector == null)
            {
                Log.Info("      [WARNING] m_multiIDSelector not found, searching alternatives...");
                var allPanes = Window.FindAllDescendants(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Pane));
                foreach (var pane in allPanes)
                {
                    var id = pane.AutomationId ?? "";
                    if (id.Contains("multiID") || id.Contains("Selector"))
                    {
                        selector = pane;
                        Log.Info($"      Found alternative pane: '{id}'");
                        break;
                    }
                }
            }

            if (selector != null)
            {
                Log.Info("      Found m_multiIDSelector pane");
                var checkboxes = selector.FindAllDescendants(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.CheckBox));
                Log.Info($"      Found {checkboxes.Length} checkboxes");

                foreach (var cb in checkboxes)
                {
                    string name = cb.Name ?? "";
                    if (name.ToLower().Contains("select all")) continue;

                    // Checkbox has meaningful name
                    if (!string.IsNullOrEmpty(name) && !name.ToLower().Contains("check row"))
                    {
                        values.Add(new SelectableValue { Name = name, Element = cb, IsCheckBox = true });
                        continue;
                    }

                    // Look for actual value in parent/siblings
                    string actualValue = FindActualValueFromParent(cb);
                    if (!string.IsNullOrEmpty(actualValue))
                        values.Add(new SelectableValue { Name = actualValue, Element = cb, IsCheckBox = true });
                }
            }
            else
            {
                // Fallback: any checkboxes with meaningful names
                Log.Info("      [WARNING] Using fallback checkbox search");
                var allCbs = Window.FindAllDescendants(
                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.CheckBox));
                foreach (var cb in allCbs)
                {
                    string name = cb.Name ?? "";
                    if (!string.IsNullOrEmpty(name) &&
                        !name.ToLower().Contains("select all") &&
                        !name.ToLower().Contains("check row"))
                    {
                        values.Add(new SelectableValue { Name = name, Element = cb, IsCheckBox = true });
                        Log.Info($"        Found checkbox: '{name}'");
                    }
                }
            }

            Log.Info($"      Total selectable values: {values.Count}");
            return values;
        }

        /// <summary>
        /// Find the actual display value from a checkbox's parent/siblings
        /// </summary>
        private string FindActualValueFromParent(AutomationElement checkbox)
        {
            var parent = checkbox.Parent;
            if (parent == null) return "";

            // Check siblings
            var siblings = parent.FindAllChildren();
            foreach (var sib in siblings)
            {
                if (sib.Equals(checkbox)) continue;

                string sibName = sib.Name ?? "";
                string sibType = "Unknown";
                try { sibType = sib.ControlType.ToString(); } catch { }

                if (sibType == "Text" && !string.IsNullOrEmpty(sibName))
                    return sibName;

                try
                {
                    var vp = sib.Patterns.Value.PatternOrDefault;
                    if (vp != null)
                    {
                        var v = vp.Value.Value ?? "";
                        if (!string.IsNullOrEmpty(v)) return v;
                    }
                }
                catch { }

                try
                {
                    var lp = sib.Patterns.LegacyIAccessible.PatternOrDefault;
                    if (lp != null)
                    {
                        var v = lp.Value.Value ?? "";
                        if (!string.IsNullOrEmpty(v)) return v;
                    }
                }
                catch { }

                if (!string.IsNullOrEmpty(sibName) &&
                    !sibName.ToLower().Contains("check") &&
                    !sibName.ToLower().Contains("row"))
                    return sibName;
            }

            // Try grandchildren
            var grandChildren = parent.FindAllDescendants();
            foreach (var gc in grandChildren)
            {
                if (gc.Equals(checkbox)) continue;
                string gcName = gc.Name ?? "";
                if (!string.IsNullOrEmpty(gcName) &&
                    !gcName.ToLower().Contains("check") &&
                    !gcName.ToLower().Contains("row") &&
                    !gcName.ToLower().Contains("select"))
                    return gcName;
            }

            return "";
        }

        // ╔══════════════════════════════════════════════════╗
        // ║              DIALOG BUTTONS                       ║
        // ╚══════════════════════════════════════════════════╝

        /// <summary>
        /// Click the OK button to apply the current filter selection
        /// </summary>
        public void ClickOK()
        {
            Log.Info("      Clicking OK button...");

            var okButton = Window.FindFirstDescendant(cf => cf.ByName("OK"))
                        ?? Window.FindFirstDescendant(cf => cf.ByAutomationId("1"));

            if (okButton != null)
            {
                var invoke = okButton.Patterns.Invoke.PatternOrDefault;
                if (invoke != null)
                {
                    invoke.Invoke();
                    Log.Info("      Clicked OK (via Invoke)");
                }
                else
                {
                    okButton.Click();
                    Log.Info("      Clicked OK (via Click)");
                }
            }
            else
            {
                Log.Info("      WARNING: OK button not found");
            }
        }

        // ╔══════════════════════════════════════════════════╗
        // ║         FILTER EXPLORATION (full orchestration)     ║
        // ╚══════════════════════════════════════════════════╝

        /// <summary>
        /// Explore ALL filters and their options dynamically.
        /// 
        /// For each filter:
        ///   1. Discover available options (All, Range, One or More)
        ///   2. Select each option and fill in values
        ///   3. Click OK to apply
        ///   4. Re-open and clear for next combination
        /// 
        /// Usage from test file:
        ///   var viewerPage = CreateViewerPage(report.ReportWindowTitle);
        ///   var modifyPage = viewerPage.OpenModifyReportDialog();
        ///   modifyPage.ExploreAllFiltersAndOptions(viewerPage.OpenModifyReportDialog);
        /// </summary>
        public void ExploreAllFiltersAndOptions(Func<ModifyReportPage> reopenDialog)
        {
            Log.Info("=== Starting Options and Filters exploration ===");

            // Use 'this' as the current ModifyReportPage instance
            var modifyPage = this;

            // STEP 1: Discover all filters
            Log.Info("\n=== STEP 1: Discovering all filters from UI ===");
            var allFilters = modifyPage.DiscoverFilters();
            Log.Info($"Total filters discovered: {allFilters.Count}");

            Log.Info("\n--- Discovered Filters List ---");
            for (int i = 0; i < allFilters.Count; i++)
                Log.Info($"  Filter [{i + 1}]: {allFilters[i].Name}");

            // STEP 2: Discover options for each filter
            Log.Info("\n=== STEP 2: Discovering options for each filter ===");
            var filterOptionsMap = new Dictionary<string, List<OptionInfo>>();

            foreach (var filter in allFilters)
            {
                Log.Info($"\n--- Discovering options for filter: {filter.Name} ---");
                filter.Element.Click();
                Thread.Sleep(TestConfig.ShortWaitMs);

                var options = modifyPage.DiscoverFilterOptions(filter);
                filterOptionsMap[filter.Name] = options;

                Log.Info($"  Options discovered for '{filter.Name}':");
                foreach (var opt in options)
                    Log.Info($"    - {opt.Name} (Type: {opt.OptionType})");
            }

            // STEP 3: Test all filter-option combinations
            Log.Info("\n=== STEP 3: Testing all filters, options and values ===");

            int filterIndex = 0;
            foreach (var filter in allFilters)
            {
                filterIndex++;
                Log.Info($"\n========================================");
                Log.Info($"FILTER [{filterIndex}/{allFilters.Count}]: {filter.Name}");
                Log.Info($"========================================");

                var options = filterOptionsMap[filter.Name];
                int optionIndex = 0;

                foreach (var option in options)
                {
                    optionIndex++;
                    Log.Info($"\n  --- Option [{optionIndex}/{options.Count}]: {option.Name} ---");

                    // A: Click on filter to show its options
                    Log.Info($"      Clicking on filter '{filter.Name}' to show options...");
                    var freshFilter = modifyPage.RefindFilter(filter.Name);
                    if (freshFilter != null)
                    {
                        freshFilter.Click();
                        Thread.Sleep(TestConfig.ShortWaitMs);
                    }
                    else
                    {
                        Log.Info($"      WARNING: Could not find filter '{filter.Name}', trying stored element...");
                        try { filter.Element.Click(); Thread.Sleep(TestConfig.ShortWaitMs); } catch { }
                    }

                    // B: Select the option (radio button)
                    modifyPage.SelectOption(option);
                    Thread.Sleep(TestConfig.ShortWaitMs);

                    // C: Handle value selection based on option type
                    switch (option.OptionType)
                    {
                        case OptionType.All:
                            Log.Info($"      'All' option selected - no additional values needed");
                            break;
                        case OptionType.Range:
                            modifyPage.SelectRangeValues();
                            break;
                        case OptionType.OneOrMore:
                            modifyPage.SelectMultipleValues(TestConfig.MaxFilterValuesToSelect);
                            break;
                    }

                    // D: Click OK to apply
                    modifyPage.ClickOK();
                    Thread.Sleep(1000);
                    Log.Info($"  Completed: Filter='{filter.Name}', Option='{option.Name}'");

                    // E: Re-open Options and Clear All Filters for next iteration
                    bool hasMoreOptions = optionIndex < options.Count;
                    bool hasMoreFilters = filterIndex < allFilters.Count;

                    if (hasMoreOptions || hasMoreFilters)
                    {
                        Log.Info($"      [STEP E] Re-opening Options and clearing filters...");
                        modifyPage = reopenDialog();
                        modifyPage.ClearAllFilters();

                        // Re-find the filter element for the next iteration
                        string nextFilterName = hasMoreOptions
                            ? filter.Name
                            : (filterIndex < allFilters.Count ? allFilters[filterIndex].Name : filter.Name);

                        Log.Info($"      Re-finding filter '{nextFilterName}' in new window...");
                        var refreshed = modifyPage.RefindFilter(nextFilterName);
                        if (refreshed != null)
                        {
                            if (hasMoreOptions)
                                filter.Element = refreshed;
                            else if (filterIndex < allFilters.Count)
                                allFilters[filterIndex].Element = refreshed;
                            Log.Info($"      Successfully re-found filter '{nextFilterName}'");
                        }
                        else
                        {
                            Log.Info($"      WARNING: Could not re-find filter '{nextFilterName}'");
                        }
                    }
                }
            }

            Log.Info("\n=== Options and Filters exploration completed ===");
            Log.Info($"Total combinations tested: {allFilters.Count} filters with their respective options");
        }

        // ╔══════════════════════════════════════════════════╗
        // ║              FILTER RESET                         ║
        // ╚══════════════════════════════════════════════════╝

        /// <summary>
        /// Click "Clear All Filters" to reset before next iteration
        /// </summary>
        public void ClearAllFilters()
        {
            var clearButton = Window.FindFirstDescendant(cf => cf.ByAutomationId("m_btnClearAllFilters"));

            if (clearButton != null)
            {
                var invoke = clearButton.Patterns.Invoke.PatternOrDefault;
                if (invoke != null)
                {
                    invoke.Invoke();
                    Log.Info("      Clicked Clear All Filters (via Invoke)");
                }
                else
                {
                    clearButton.Click();
                    Log.Info("      Clicked Clear All Filters (via Click)");
                }
                Thread.Sleep(TestConfig.ShortWaitMs);
            }
            else
            {
                Log.Info("      WARNING: Clear All Filters button not found");
            }
        }
    }
}
