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

        // ==================================================
        //              WINDOW MANAGEMENT
        // ==================================================

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
        public AutomationElement Window
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

        // ==================================================
        //              FILTER DISCOVERY
        // ==================================================

        /// <summary>
        /// Discover all filter items dynamically from the UI.
        /// Tries multiple approaches: ListBox -> TreeView -> DataGrid -> Table
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

        // ==================================================
        //              OPTION DISCOVERY & SELECTION
        // ==================================================

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

        // ==================================================
        //              RANGE VALUE SELECTION
        // ==================================================

        /// <summary>
        /// Discover and select range values (From and To dropdowns).
        /// From pane: m_sde_From or m_autofill_From
        /// To pane: m_sde_To or m_autofill_To
        /// </summary>
        public void SelectRangeValues()
        {
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

            Log.Info("        Range selection completed");
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
        /// Extract range values from the UI after SelectRangeValues() has set them.
        /// Uses multiple strategies:
        ///   1. Find "Set To Row 0" element and read its Value/LegacyIAccessible/children
        ///   2. Scan all descendants for any element name matching "Range ... - ..."
        ///   3. Read From/To pane text values directly
        /// Parses text like "Range ALDRED - ARCHER" to extract from/to values.
        /// </summary>
        public (string from, string to) ExtractRangeFromSetToRow()
        {
            string from = "", to = "";

            try
            {
                Log.Info("      Extracting range values...");
                string rangeText = "";

                // Strategy 1: Find "Set To Row 0" and read its properties
                var setToRow = Window.FindFirstDescendant(cf => cf.ByName("Set To Row 0"));
                if (setToRow != null)
                {
                    rangeText = TryReadRangeText(setToRow);
                    Log.Info($"      Set To Row 0 found, text='{rangeText}'");
                }

                // Strategy 2: Search all descendants for element name starting with "Range "
                if (!IsValidRangeText(rangeText))
                {
                    Log.Info("      Scanning all elements for range text...");
                    var allElements = Window.FindAllDescendants();
                    foreach (var elem in allElements)
                    {
                        string name = elem.Name ?? "";
                        if (IsValidRangeText(name))
                        {
                            rangeText = name;
                            Log.Info($"      Found range in element name: '{rangeText}'");
                            break;
                        }
                    }
                }

                // Strategy 3: Read From/To pane values directly
                if (!IsValidRangeText(rangeText))
                {
                    Log.Info("      Trying From/To pane values...");
                    var (fromPane, toPane) = FindFromToPanes();
                    string fromVal = ReadPaneText(fromPane);
                    string toVal = ReadPaneText(toPane);

                    if (!string.IsNullOrEmpty(fromVal) || !string.IsNullOrEmpty(toVal))
                    {
                        from = fromVal;
                        to = toVal;
                        Log.Info($"      From pane='{from}', To pane='{to}'");
                        Log.Info($"      Extracted range: From='{from}', To='{to}'");
                        return (from, to);
                    }
                }

                // Parse "Range FROM - TO" format
                if (IsValidRangeText(rangeText))
                {
                    string body = rangeText.Substring(rangeText.IndexOf("Range ", StringComparison.OrdinalIgnoreCase) + 6).Trim();
                    var parts = body.Split(" - ", 2, StringSplitOptions.TrimEntries);
                    if (parts.Length == 2)
                    {
                        from = parts[0];
                        to = parts[1];
                    }
                }

                Log.Info($"      Extracted range: From='{from}', To='{to}'");
            }
            catch (Exception ex)
            {
                Log.Info($"      WARNING: ExtractRangeFromSetToRow failed: {ex.Message}");
            }

            return (from, to);
        }

        /// <summary>
        /// Check if a string looks like valid range text (contains "Range " and " - ")
        /// </summary>
        private static bool IsValidRangeText(string text)
        {
            return !string.IsNullOrEmpty(text)
                && text.Contains("Range ", StringComparison.OrdinalIgnoreCase)
                && text.Contains(" - ");
        }

        /// <summary>
        /// Try to read range text from an element using Name, Value, LegacyIAccessible, and children
        /// </summary>
        private string TryReadRangeText(AutomationElement element)
        {
            // Try Name
            string text = element.Name ?? "";
            if (IsValidRangeText(text)) return text;

            // Try Value pattern
            var vp = element.Patterns.Value.PatternOrDefault;
            if (vp != null)
                try { text = vp.Value.Value ?? ""; if (IsValidRangeText(text)) return text; } catch { }

            // Try LegacyIAccessible
            var lp = element.Patterns.LegacyIAccessible.PatternOrDefault;
            if (lp != null)
            {
                try { text = lp.Value.Value ?? ""; if (IsValidRangeText(text)) return text; } catch { }
                try { text = lp.Description.Value ?? ""; if (IsValidRangeText(text)) return text; } catch { }
            }

            // Try children
            var children = element.FindAllDescendants();
            foreach (var child in children)
            {
                string childName = child.Name ?? "";
                if (IsValidRangeText(childName)) return childName;

                var cvp = child.Patterns.Value.PatternOrDefault;
                if (cvp != null)
                    try { text = cvp.Value.Value ?? ""; if (IsValidRangeText(text)) return text; } catch { }
            }

            return "";
        }

        /// <summary>
        /// Read text value from a From/To pane (used as fallback in range extraction)
        /// </summary>
        private string ReadPaneText(AutomationElement? pane)
        {
            if (pane == null) return "";

            // Value pattern
            var vp = pane.Patterns.Value.PatternOrDefault;
            if (vp != null)
                try { var v = vp.Value.Value ?? ""; if (!string.IsNullOrEmpty(v)) return v; } catch { }

            // LegacyIAccessible
            var lp = pane.Patterns.LegacyIAccessible.PatternOrDefault;
            if (lp != null)
                try { var v = lp.Value.Value ?? ""; if (!string.IsNullOrEmpty(v)) return v; } catch { }

            // Children
            foreach (var child in pane.FindAllChildren())
            {
                var cvp = child.Patterns.Value.PatternOrDefault;
                if (cvp != null)
                    try { var v = cvp.Value.Value ?? ""; if (!string.IsNullOrEmpty(v)) return v; } catch { }

                var clp = child.Patterns.LegacyIAccessible.PatternOrDefault;
                if (clp != null)
                    try { var v = clp.Value.Value ?? ""; if (!string.IsNullOrEmpty(v)) return v; } catch { }

                if (!string.IsNullOrEmpty(child.Name))
                    return child.Name;
            }

            return pane.Name ?? "";
        }

        // ==================================================
        //              MULTI-SELECT VALUES
        // ==================================================

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

        // ==================================================
        //              DIALOG BUTTONS
        // ==================================================

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

        // ==================================================
        //              SELECT SPECIFIC VALUES (for replay)
        // ==================================================

        /// <summary>
        /// Select specific values by name (for replaying OneOrMore selections from Act26).
        /// Finds checkboxes matching the provided names and checks them.
        /// </summary>
        public void SelectSpecificValues(List<string> valueNames)
        {
            if (valueNames.Count == 0)
            {
                Log.Info("      No specific values to select");
                return;
            }

            Log.Info($"      Selecting {valueNames.Count} specific values...");

            var selectableValues = DiscoverSelectableValues();

            foreach (var targetName in valueNames)
            {
                var match = selectableValues.FirstOrDefault(v =>
                    v.Name.Equals(targetName, StringComparison.OrdinalIgnoreCase));

                if (match != null)
                {
                    try
                    {
                        if (match.IsCheckBox)
                            match.Element.AsCheckBox().IsChecked = true;
                        else
                            match.Element.Click();

                        Log.Info($"        Selected: '{targetName}'");
                    }
                    catch (Exception ex)
                    {
                        Log.Info($"        WARNING: Could not select '{targetName}': {ex.Message}");
                    }
                }
                else
                {
                    Log.Info($"        WARNING: Value '{targetName}' not found in list");
                }
            }
        }

        // ==================================================
        //              FILTER RESET
        // ==================================================

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
