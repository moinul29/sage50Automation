# WinAppDriver Automation with C#

This project demonstrates Windows Desktop application automation using WinAppDriver and C#.

## Prerequisites

### 1. Install WinAppDriver
- Download from [GitHub Releases](https://github.com/Microsoft/WinAppDriver/releases)
- Run the installer `WindowsApplicationDriver.msi`
- Default installation path: `C:\Program Files (x86)\Windows Application Driver\`

### 2. Enable Developer Mode
- Open **Settings** → **Update & Security** → **For developers**
- Select **Developer mode**

### 3. Install .NET SDK
- Download from [dotnet.microsoft.com](https://dotnet.microsoft.com/download)
- Verify installation: `dotnet --version`

## Project Setup

This project is already configured with:
- ✅ MSTest framework
- ✅ Appium.WebDriver package (v8.0.1)
- ✅ Sample test files for Calculator and Notepad

## Running Tests

### Step 1: Start WinAppDriver
Open a terminal and run:
```powershell
& "C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe"
```

WinAppDriver will start listening on `http://127.0.0.1:4723`

### Step 2: Run Tests
In a new terminal, navigate to the project directory and run:

**Run all tests:**
```powershell
dotnet test
```

**Run specific test class:**
```powershell
dotnet test --filter ClassName=WinAppDriverTests.CalculatorTests
```

**Run specific test method:**
```powershell
dotnet test --filter Name=TestAddition
```

**Run with detailed output:**
```powershell
dotnet test --logger "console;verbosity=detailed"
```

## Test Files

### CalculatorTests.cs
Tests for Windows Calculator application:
- `TestAddition()` - Tests 5 + 7 = 12
- `TestSubtraction()` - Tests 9 - 3 = 6
- `TestMultiplication()` - Tests 8 × 4 = 32

### NotepadTests.cs
Tests for Notepad application:
- `TestTypeText()` - Types text into Notepad
- `TestMenuNavigation()` - Navigates File menu
- `TestClearText()` - Clears text using keyboard shortcuts

## Project Structure

```
WinAppDriver Automation/
├── CalculatorTests.cs          # Calculator automation tests
├── NotepadTests.cs              # Notepad automation tests
├── WinAppDriverTests.csproj     # Project file
├── README.md                    # This file
└── obj/                         # Build output
```

## Common Element Locators

WinAppDriver supports multiple locator strategies:

```csharp
// By Name
session.FindElementByName("Button Name");

// By Accessibility ID
session.FindElementByAccessibilityId("AutomationId");

// By Class Name
session.FindElementByClassName("Edit");

// By XPath
session.FindElementByXPath("//Button[@Name='Submit']");
```

## Capabilities

Common Appium capabilities for Windows apps:

```csharp
var appiumOptions = new AppiumOptions();

// For Microsoft Store apps
appiumOptions.AddAdditionalCapability("app", "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App");

// For classic desktop apps
appiumOptions.AddAdditionalCapability("app", "C:\\Path\\To\\App.exe");

// Device name
appiumOptions.AddAdditionalCapability("deviceName", "WindowsPC");

// Platform name (optional)
appiumOptions.AddAdditionalCapability("platformName", "Windows");
```

## Troubleshooting

### WinAppDriver not found
Make sure WinAppDriver is installed and the path is correct.

### App not launching
- Verify the app ID or path is correct
- For Store apps, use the Application User Model ID
- For classic apps, use the full executable path

### Elements not found
- Increase implicit wait time
- Use **Inspect.exe** (Windows SDK tool) to find correct element properties
- Try different locator strategies

### Tests failing randomly
- Add explicit waits for dynamic elements
- Ensure WinAppDriver is running before tests start
- Close app instances before running tests

## Find Element Properties

Use **Inspect.exe** from Windows SDK:
1. Open Inspect.exe
2. Hover over the UI element
3. Check properties: Name, AutomationId, ClassName, etc.

## Additional Resources

- [WinAppDriver GitHub](https://github.com/Microsoft/WinAppDriver)
- [Appium Documentation](http://appium.io/docs/en/about-appium/intro/)
- [MSTest Documentation](https://docs.microsoft.com/en-us/dotnet/core/testing/unit-testing-with-mstest)

## License

This is a sample project for learning purposes.
