using System.Diagnostics;

namespace TestLauncher;

/// <summary>
/// Sage 50 Automation Test Launcher
/// Provides a simple UI to execute Act26, Act24, or both test suites
/// without needing command-line access.
/// </summary>
public class MainForm : Form
{
    // ===== UI Controls =====
    private Button btnAct26 = null!;
    private Button btnAct24 = null!;
    private Button btnBoth = null!;
    private Button btnStop = null!;
    private Button btnClear = null!;
    private RichTextBox rtbOutput = null!;
    private Label lblStatus = null!;
    private Panel headerPanel = null!;
    private ProgressBar progressBar = null!;

    // ===== State =====
    private Process? _runningProcess;
    private CancellationTokenSource? _cts;
    private bool _isRunning;

    /// <summary>
    /// Path to the test project (.csproj) — one directory up from TestLauncher
    /// </summary>
    private readonly string _testProjectPath;

    public MainForm()
    {
        // Resolve test project path — exe is in TestLauncher\, csproj is one level up
        string launcherDir = AppDomain.CurrentDomain.BaseDirectory;
        string workspaceDir = Path.GetFullPath(Path.Combine(launcherDir, ".."));
        _testProjectPath = Path.Combine(workspaceDir, "WinAppDriverTests.csproj");

        InitializeUI();
    }

    // ==================================================
    //  UI CONSTRUCTION
    // ==================================================

    private void InitializeUI()
    {
        // Form settings
        Text = "Sage 50 Automation - Test Launcher";
        Size = new Size(900, 680);
        MinimumSize = new Size(750, 550);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(245, 247, 250);
        Font = new Font("Segoe UI", 9.5f);

        // Header panel
        headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 70,
            BackColor = Color.FromArgb(26, 26, 46),
            Padding = new Padding(20, 0, 20, 0)
        };

        var lblTitle = new Label
        {
            Text = "Sage 50 Automation — Test Launcher",
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 16f, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(20, 12)
        };

        var lblSubtitle = new Label
        {
            Text = "Select a test suite to execute",
            ForeColor = Color.FromArgb(148, 163, 184),
            Font = new Font("Segoe UI", 9.5f),
            AutoSize = true,
            Location = new Point(22, 42)
        };

        headerPanel.Controls.Add(lblTitle);
        headerPanel.Controls.Add(lblSubtitle);

        // Buttons panel — FlowLayoutPanel wraps buttons when the window is narrow
        var buttonsPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            MinimumSize = new Size(0, 60),
            WrapContents = true,
            Padding = new Padding(15, 10, 15, 5)
        };

        btnAct26 = CreateActionButton("act26CustomerLedger", Color.FromArgb(37, 99, 235));
        btnAct26.Click += (s, e) => RunTest("Actian2026");

        btnAct24 = CreateActionButton("act24CustomerLedger", Color.FromArgb(5, 150, 105));
        btnAct24.Click += (s, e) => RunTest("Actian2024");

        btnBoth = CreateActionButton("act26CustomerLedger && act24CustomerLedger", Color.FromArgb(124, 58, 237));
        btnBoth.Click += (s, e) => RunBothTests();

        btnStop = CreateActionButton("Stop", Color.FromArgb(220, 38, 38));
        btnStop.Enabled = false;
        btnStop.Click += (s, e) => StopExecution();

        btnClear = CreateActionButton("Clear", Color.FromArgb(107, 114, 128));
        btnClear.Click += (s, e) => { rtbOutput.Clear(); lblStatus.Text = "Ready"; };

        buttonsPanel.Controls.AddRange(new Control[] { btnAct26, btnAct24, btnBoth, btnStop, btnClear });

        // Status bar
        var statusPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 30,
            Padding = new Padding(20, 5, 20, 5)
        };

        lblStatus = new Label
        {
            Text = "Ready",
            ForeColor = Color.FromArgb(100, 116, 139),
            Font = new Font("Segoe UI", 9f),
            AutoSize = true,
            Location = new Point(20, 5)
        };

        progressBar = new ProgressBar
        {
            Style = ProgressBarStyle.Marquee,
            MarqueeAnimationSpeed = 30,
            Height = 4,
            Dock = DockStyle.Bottom,
            Visible = false
        };

        statusPanel.Controls.Add(lblStatus);
        statusPanel.Controls.Add(progressBar);

        // Output area
        rtbOutput = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            BackColor = Color.FromArgb(15, 23, 42),
            ForeColor = Color.FromArgb(226, 232, 240),
            Font = new Font("Cascadia Code", 9.5f, FontStyle.Regular, GraphicsUnit.Point, 0,
                            IsFontInstalled("Cascadia Code") ? false : true),
            BorderStyle = BorderStyle.None,
            Margin = new Padding(20, 5, 20, 20),
            WordWrap = true,
            DetectUrls = false
        };

        // Fix font fallback
        if (!IsFontInstalled("Cascadia Code"))
            rtbOutput.Font = new Font("Consolas", 9.5f);

        // Assemble the form (order matters — reverse dock order)
        var outputContainer = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(20, 5, 20, 20)
        };
        outputContainer.Controls.Add(rtbOutput);

        Controls.Add(outputContainer);
        Controls.Add(statusPanel);
        Controls.Add(buttonsPanel);
        Controls.Add(headerPanel);
    }

    private Button CreateActionButton(string text, Color bgColor)
    {
        var btn = new Button
        {
            Text = text,
            FlatStyle = FlatStyle.Flat,
            BackColor = bgColor,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(14, 8, 14, 8),
            Margin = new Padding(5),
            MinimumSize = new Size(0, 40),
            Cursor = Cursors.Hand,
            FlatAppearance =
            {
                BorderSize = 0,
                MouseOverBackColor = ControlPaint.Light(bgColor, 0.15f),
                MouseDownBackColor = ControlPaint.Dark(bgColor, 0.1f)
            }
        };
        return btn;
    }

    private static bool IsFontInstalled(string fontName)
    {
        using var testFont = new Font(fontName, 10f);
        return testFont.Name.Equals(fontName, StringComparison.OrdinalIgnoreCase);
    }

    // ==================================================
    //  TEST EXECUTION
    // ==================================================

    private async void RunTest(string filter)
    {
        if (_isRunning)
        {
            AppendOutput("A test is already running. Please wait or stop it first.", Color.Orange);
            return;
        }

        SetRunningState(true);
        string displayName = filter.Contains("2026") ? "Actian 2026 Customer Ledger" : "Actian 2024 Customer Ledger";
        lblStatus.Text = $"Running: {displayName}...";

        AppendOutput($"{'=',-60}", Color.FromArgb(100, 116, 139));
        AppendOutput($"  Starting: {displayName}", Color.FromArgb(96, 165, 250));
        AppendOutput($"  Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}", Color.FromArgb(148, 163, 184));
        AppendOutput($"{'=',-60}", Color.FromArgb(100, 116, 139));
        AppendOutput("", Color.White);

        bool success = await ExecuteDotnetTest(filter);

        AppendOutput("", Color.White);
        if (success)
        {
            AppendOutput($"  {displayName} — PASSED", Color.FromArgb(34, 197, 94));
            lblStatus.Text = $"Completed: {displayName} — PASSED";
        }
        else
        {
            AppendOutput($"  {displayName} — FAILED", Color.FromArgb(248, 113, 113));
            lblStatus.Text = $"Completed: {displayName} — FAILED";
        }
        AppendOutput($"{'=',-60}", Color.FromArgb(100, 116, 139));
        AppendOutput("", Color.White);

        SetRunningState(false);
    }

    private async void RunBothTests()
    {
        if (_isRunning)
        {
            AppendOutput("A test is already running. Please wait or stop it first.", Color.Orange);
            return;
        }

        SetRunningState(true);
        lblStatus.Text = "Running: Both Tests (Act 2026 → Act 2024)...";

        AppendOutput($"{'=',-60}", Color.FromArgb(100, 116, 139));
        AppendOutput("  Running Both Tests Sequentially", Color.FromArgb(167, 139, 250));
        AppendOutput($"  Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}", Color.FromArgb(148, 163, 184));
        AppendOutput($"{'=',-60}", Color.FromArgb(100, 116, 139));
        AppendOutput("", Color.White);

        // Run Act26 first
        AppendOutput("  [1/2] Actian 2026 Customer Ledger", Color.FromArgb(96, 165, 250));
        AppendOutput("", Color.White);
        bool act26Success = await ExecuteDotnetTest("Actian2026");

        AppendOutput("", Color.White);
        if (act26Success)
            AppendOutput("  Act 2026 — PASSED", Color.FromArgb(34, 197, 94));
        else
            AppendOutput("  Act 2026 — FAILED", Color.FromArgb(248, 113, 113));

        // Check if cancelled
        if (_cts?.IsCancellationRequested == true)
        {
            AppendOutput("  Execution stopped by user.", Color.Orange);
            SetRunningState(false);
            lblStatus.Text = "Stopped by user";
            return;
        }

        AppendOutput("", Color.White);
        AppendOutput("  [2/2] Actian 2024 Customer Ledger", Color.FromArgb(96, 165, 250));
        AppendOutput("", Color.White);

        bool act24Success = await ExecuteDotnetTest("Actian2024");

        AppendOutput("", Color.White);
        if (act24Success)
            AppendOutput("  Act 2024 — PASSED", Color.FromArgb(34, 197, 94));
        else
            AppendOutput("  Act 2024 — FAILED", Color.FromArgb(248, 113, 113));

        AppendOutput("", Color.White);
        AppendOutput($"{'=',-60}", Color.FromArgb(100, 116, 139));

        string overall = (act26Success && act24Success) ? "ALL PASSED" : "SOME FAILED";
        var overallColor = (act26Success && act24Success)
            ? Color.FromArgb(34, 197, 94)
            : Color.FromArgb(248, 113, 113);
        AppendOutput($"  Final Result: {overall}", overallColor);
        AppendOutput($"{'=',-60}", Color.FromArgb(100, 116, 139));

        lblStatus.Text = $"Completed: Both Tests — {overall}";
        SetRunningState(false);
    }

    /// <summary>
    /// Execute a dotnet test command and stream output to the UI in real time.
    /// Returns true if the test passed (exit code 0).
    /// </summary>
    private async Task<bool> ExecuteDotnetTest(string filter)
    {
        _cts = new CancellationTokenSource();

        string args = $"test \"{_testProjectPath}\" --filter \"FullyQualifiedName~{filter}\" --logger \"console;verbosity=detailed\"";

        var psi = new ProcessStartInfo
        {
            FileName = "dotnet",
            Arguments = args,
            WorkingDirectory = Path.GetDirectoryName(_testProjectPath) ?? ".",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        try
        {
            _runningProcess = new Process { StartInfo = psi, EnableRaisingEvents = true };
            _runningProcess.Start();

            // Stream stdout and stderr concurrently
            var stdoutTask = StreamOutputAsync(_runningProcess.StandardOutput, _cts.Token);
            var stderrTask = StreamOutputAsync(_runningProcess.StandardError, _cts.Token, isError: true);

            await Task.WhenAll(stdoutTask, stderrTask);
            await _runningProcess.WaitForExitAsync(_cts.Token);

            int exitCode = _runningProcess.ExitCode;
            _runningProcess = null;

            return exitCode == 0;
        }
        catch (OperationCanceledException)
        {
            KillProcess();
            return false;
        }
        catch (Exception ex)
        {
            AppendOutput($"ERROR: {ex.Message}", Color.FromArgb(248, 113, 113));
            _runningProcess = null;
            return false;
        }
    }

    private async Task StreamOutputAsync(StreamReader reader, CancellationToken ct, bool isError = false)
    {
        try
        {
            while (!ct.IsCancellationRequested)
            {
                string? line = await reader.ReadLineAsync(ct);
                if (line == null) break;

                Color color = isError
                    ? Color.FromArgb(248, 113, 113)
                    : ClassifyOutputColor(line);

                BeginInvoke(() => AppendOutput(line, color));
            }
        }
        catch (OperationCanceledException) { }
    }

    /// <summary>
    /// Color-code dotnet test output lines for readability.
    /// </summary>
    private static Color ClassifyOutputColor(string line)
    {
        if (line.Contains("Passed") || line.Contains("SUCCESS") || line.Contains("Completed Successfully"))
            return Color.FromArgb(34, 197, 94);     // green
        if (line.Contains("Failed") || line.Contains("ERROR") || line.Contains("FATAL"))
            return Color.FromArgb(248, 113, 113);    // red
        if (line.Contains("WARNING") || line.Contains("Warning"))
            return Color.FromArgb(251, 191, 36);     // yellow
        if (line.Contains("==="))
            return Color.FromArgb(96, 165, 250);     // blue
        if (line.Contains("Step ") || line.Contains("FILTER [") || line.Contains("OPTION ["))
            return Color.FromArgb(167, 139, 250);    // purple
        return Color.FromArgb(226, 232, 240);        // default light gray
    }

    // ==================================================
    //  UI STATE MANAGEMENT
    // ==================================================

    private void SetRunningState(bool running)
    {
        _isRunning = running;
        btnAct26.Enabled = !running;
        btnAct24.Enabled = !running;
        btnBoth.Enabled = !running;
        btnStop.Enabled = running;
        progressBar.Visible = running;
    }

    private void StopExecution()
    {
        _cts?.Cancel();
        KillProcess();
        AppendOutput("  >> Execution stopped by user", Color.Orange);
        lblStatus.Text = "Stopped";
        SetRunningState(false);
    }

    private void KillProcess()
    {
        try
        {
            if (_runningProcess != null && !_runningProcess.HasExited)
            {
                _runningProcess.Kill(entireProcessTree: true);
                _runningProcess = null;
            }
        }
        catch { }
    }

    private void AppendOutput(string text, Color color)
    {
        rtbOutput.SelectionStart = rtbOutput.TextLength;
        rtbOutput.SelectionLength = 0;
        rtbOutput.SelectionColor = color;
        rtbOutput.AppendText(text + Environment.NewLine);
        rtbOutput.ScrollToCaret();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (_isRunning)
        {
            var result = MessageBox.Show(
                "A test is still running. Do you want to stop it and close?",
                "Test Running",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }

            StopExecution();
        }

        base.OnFormClosing(e);
    }
}
