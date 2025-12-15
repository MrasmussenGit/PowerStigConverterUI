using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Windows.Documents;
using System.Windows.Input;

namespace PowerStigConverterUI
{
    public partial class ConvertStigWindow : Window
    {
        private const string LastModulePathRegKey = @"Software\PowerStigConverterUI";
        private const string LastModulePathRegValue = "LastModulePath";

        // Track rule IDs that failed during conversion
        private readonly System.Collections.Generic.HashSet<string> _failedRuleIds =
            new(System.StringComparer.OrdinalIgnoreCase);
        private readonly System.Collections.Generic.Dictionary<string, string> _failedRuleErrors =
            new(System.StringComparer.OrdinalIgnoreCase);

        private System.Collections.Generic.List<string>? _lastSuccessIds;
        private System.Windows.Threading.DispatcherTimer? _busyTimer;
        private DateTime _busyStart;

        public ConvertStigWindow()
        {
            InitializeComponent();
            // Ensure the auto-discovery runs after the visual tree is ready
            this.Loaded += OnLoaded;
        }

        // Centering requires Owner to be set by the caller:
        // var win = new ConvertStigWindow { Owner = this }; win.ShowDialog();

        private void OnLoaded(object? sender, RoutedEventArgs e)
        {
            // If already set (e.g., from previous run), keep it.
            if (!string.IsNullOrWhiteSpace(ModulePathTextBox.Text) && File.Exists(ModulePathTextBox.Text))
                return;

            // Try registry (last successful path)
            var last = ReadLastModulePath();
            if (!string.IsNullOrWhiteSpace(last) && File.Exists(last))
            {
                ModulePathTextBox.Text = last;
                return;
            }

            // Discover typical locations
            var discovered = FindPowerStigConvertModule();
            if (!string.IsNullOrWhiteSpace(discovered))
            {
                ModulePathTextBox.Text = discovered!;
                WriteLastModulePath(discovered!);
                return;
            }

            // Alert user if not found
            ModulePathTextBox.Text = string.Empty;
            System.Windows.MessageBox.Show(
                "PowerStig.Convert module was not found in PSModulePath or standard module locations.\n\n" +
                "Please install PowerStig and PowerStig.Convert (Windows PowerShell 5.x) under:\n" +
                "  - %ProgramFiles%\\WindowsPowerShell\\Modules\n" +
                "  - %UserProfile%\\Documents\\WindowsPowerShell\\Modules\n\n" +
                "Or browse to the module path using the Change… button.",
                "PowerStig module not found",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private void RenderSuccessIds()
        {
            if (_lastSuccessIds is null || _lastSuccessIds.Count == 0)
            {
                AppendInfo("No successfully converted rule IDs.", System.Windows.Media.Brushes.DarkGreen, null);
                return;
            }

            AppendInfo($"Successfully converted rule IDs ({_lastSuccessIds.Count}):", System.Windows.Media.Brushes.DarkGreen, null);
            foreach (var id in _lastSuccessIds)
                AppendInfo($" - {id}", System.Windows.Media.Brushes.DarkGreen, null);
        }

        private static string? FindPowerStigConvertModule()
        {
            // Known repo path under user profile (optional, non-hardcoded)
            var repoCandidate = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "git", "PowerStig", "source", "PowerStig.Convert", "PowerStig.Convert.psm1");
            if (File.Exists(repoCandidate)) return repoCandidate;

            // Enumerate PSModulePath entries for PowerStig.Convert and PowerSTIG
            var psModulePath = Environment.GetEnvironmentVariable("PSModulePath");
            if (!string.IsNullOrWhiteSpace(psModulePath))
            {
                foreach (var dir in psModulePath.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
                {
                    try
                    {
                        // Direct Convert module root
                        var convertRoot = Path.Combine(dir, "PowerStig.Convert");
                        var found = FindPsm1InModuleRoot(convertRoot, "PowerStig.Convert.psm1");
                        if (found is not null) return found;

                        // Newer installs: PowerSTIG\<version>\... (or legacy PowerStig)
                        var foundLatest =
                            FindInLatestPowerStigVersion(Path.Combine(dir, "PowerSTIG"), "PowerStig.Convert.psm1") ??
                            FindInLatestPowerStigVersion(Path.Combine(dir, "PowerStig"), "PowerStig.Convert.psm1");
                        if (foundLatest is not null) return foundLatest;
                    }
                    catch { /* ignore */ }
                }
            }

            // Typical WindowsPowerShell module locations
            var programFilesModules = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "WindowsPowerShell", "Modules");
            var documentsModules = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WindowsPowerShell", "Modules");

            // Try direct Convert under Program Files
            var foundTypical = FindPsm1InModuleRoot(Path.Combine(programFilesModules, "PowerStig.Convert"), "PowerStig.Convert.psm1");
            if (foundTypical is not null) return foundTypical;

            // Try PowerSTIG latest version under Program Files
            var foundLatestPf =
                FindInLatestPowerStigVersion(Path.Combine(programFilesModules, "PowerSTIG"), "PowerStig.Convert.psm1") ??
                FindInLatestPowerStigVersion(Path.Combine(programFilesModules, "PowerStig"), "PowerStig.Convert.psm1");
            if (foundLatestPf is not null) return foundLatestPf;

            // Try direct Convert under Documents
            var foundTypicalDocs = FindPsm1InModuleRoot(Path.Combine(documentsModules, "PowerStig.Convert"), "PowerStig.Convert.psm1");
            if (foundTypicalDocs is not null) return foundTypicalDocs;

            // Try PowerSTIG latest version under Documents
            var foundLatestDocs =
                FindInLatestPowerStigVersion(Path.Combine(documentsModules, "PowerSTIG"), "PowerStig.Convert.psm1") ??
                FindInLatestPowerStigVersion(Path.Combine(documentsModules, "PowerStig"), "PowerStig.Convert.psm1");
            if (foundLatestDocs is not null) return foundLatestDocs;

            // App directory fallback
            var appLocal = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PowerStig.Convert.psm1");
            if (File.Exists(appLocal)) return appLocal;

            return null;
        }

        private static string? FindInLatestPowerStigVersion(string powerStigRoot, string psm1Name)
        {
            try
            {
                if (!Directory.Exists(powerStigRoot))
                    return null;

                var versionDirs = Directory.GetDirectories(powerStigRoot);
                Array.Sort(versionDirs, (a, b) =>
                {
                    var va = ParseVersion(Path.GetFileName(a));
                    var vb = ParseVersion(Path.GetFileName(b));
                    return Comparer<Version>.Default.Compare(vb, va); // descending
                });

                foreach (var vdir in versionDirs)
                {
                    // Common layouts under PowerSTIG\<version>\...
                    // 1) ...\PowerStig.Convert\PowerStig.Convert.psm1
                    var underConvert = Path.Combine(vdir, "PowerStig.Convert", psm1Name);
                    if (File.Exists(underConvert)) return underConvert;

                    // 2) ...\Modules\PowerStig.Convert\PowerStig.Convert.psm1
                    var underModules = Path.Combine(vdir, "Modules", "PowerStig.Convert", psm1Name);
                    if (File.Exists(underModules)) return underModules;

                    // 3) ...\<version>\PowerStig.Convert.psm1 (directly under version folder)
                    var direct = Path.Combine(vdir, psm1Name);
                    if (File.Exists(direct)) return direct;

                    // 4) Fallback: deep recursive search (handles uncommon layouts)
                    var candidate = FindFileRecursive(vdir, psm1Name);
                    if (candidate is not null)
                        return candidate;
                }
            }
            catch { /* ignore */ }
            return null;
        }

        private static Version ParseVersion(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return new Version(0, 0, 0, 0);
            if (Version.TryParse(s, out var v)) return v;
            var digits = new string((s ?? "").Trim().Where(ch => char.IsDigit(ch) || ch == '.').ToArray());
            return Version.TryParse(digits, out v) ? v : new Version(0, 0, 0, 0);
        }

        private static string? FindFileRecursive(string root, string fileName)
        {
            try
            {
                var direct = Path.Combine(root, fileName);
                if (File.Exists(direct)) return direct;

                foreach (var dir in Directory.GetDirectories(root))
                {
                    var candidate = FindFileRecursive(dir, fileName);
                    if (candidate is not null) return candidate;
                }
            }
            catch { /* ignore */ }
            return null;
        }

        private static string? FindPsm1InModuleRoot(string moduleRoot, string psm1Name)
        {
            try
            {
                if (Directory.Exists(moduleRoot))
                {
                    var direct = Path.Combine(moduleRoot, psm1Name);
                    if (File.Exists(direct))
                        return direct;

                    foreach (var sub in Directory.GetDirectories(moduleRoot))
                    {
                        var candidate = Path.Combine(sub, psm1Name);
                        if (File.Exists(candidate))
                            return candidate;
                    }
                }
            }
            catch
            {
                // Ignore IO/permission errors; return null when not found
            }
            return null;
        }

        private static string? ReadLastModulePath()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(LastModulePathRegKey);
                return key?.GetValue(LastModulePathRegValue) as string;
            }
            catch { return null; }
        }

        private static void WriteLastModulePath(string path)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(LastModulePathRegKey);
                key?.SetValue(LastModulePathRegValue, path);
            }
            catch { /* ignore */ }
        }

        private static string ResolveConvertedFullPath(string destination, string convertedFileName)
        {
            // Primary path (as generated)
            var primary = Path.Combine(destination, convertedFileName);

            if (File.Exists(primary))
                return primary;

            // If the filename contains "-<major>.0-", also try "-<major>-"
            // Example: DotNetFramework-4.0-2.7.xml -> DotNetFramework-4-2.7.xml
            var altWithMissingMinor = Regex.Replace(
                convertedFileName,
                @"-(\d+)\.0-",
                m => $"-{m.Groups[1].Value}-",
                RegexOptions.IgnoreCase);

            var altMissingMinorPath = Path.Combine(destination, altWithMissingMinor);
            if (File.Exists(altMissingMinorPath))
                return altMissingMinorPath;

            // Conversely, if the filename contains "-<major>-", also try "-<major>.0-"
            // Example: DotNetFramework-4-2.7.xml -> DotNetFramework-4.0-2.7.xml
            var altWithMinorZero = Regex.Replace(
                convertedFileName,
                @"-(\d+)-",
                m => $"-{m.Groups[1].Value}.0-",
                RegexOptions.IgnoreCase);

            var altMinorZeroPath = Path.Combine(destination, altWithMinorZero);
            if (File.Exists(altMinorZeroPath))
                return altMinorZeroPath;

            // As a final fallback, return the primary path (caller may handle missing file)
            return primary;
        }

        private void SetBusy(bool isBusy, string? status = null)
        {
            BusyProgressBar.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;
            BusyStatusText.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;
            BusyStatusText.Text = isBusy ? (status ?? "Converting…") : string.Empty;
            this.Cursor = isBusy ? System.Windows.Input.Cursors.Wait : System.Windows.Input.Cursors.Arrow;

            if (isBusy)
            {
                _busyStart = DateTime.Now;
                _busyTimer ??= new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(2)
                };
                _busyTimer.Tick -= BusyTimer_Tick;
                _busyTimer.Tick += BusyTimer_Tick;
                _busyTimer.Start();
            }
            else
            {
                if (_busyTimer is not null) _busyTimer.Stop();
            }
        }

        private void BusyTimer_Tick(object? sender, EventArgs e)
        {
            var elapsed = DateTime.Now - _busyStart;
            BusyStatusText.Text = $"Converting… elapsed {elapsed:mm\\:ss}";
            // Optional lightweight heartbeat in the log without spamming:
            // AppendInfo($"Still working… {elapsed:mm\\:ss}", System.Windows.Media.Brushes.DarkSlateGray, null);
        }

        private async void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            InfoRichTextBox.Document.Blocks.Clear();
            _failedRuleIds.Clear();
            _failedRuleErrors.Clear();

            var modulePath = ModulePathTextBox.Text?.Trim();
            var xccdfPath = XccdfPathTextBox.Text?.Trim();
            var destination = DestinationTextBox.Text?.Trim();
            var createOrgSettings = CreateOrgSettingsCheckBox.IsChecked == true;

            if (string.IsNullOrWhiteSpace(modulePath) || !File.Exists(modulePath))
            {
                AppendInfo("Invalid module path. Please select PowerStig.Convert.psm1.", System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
                return;
            }
            if (string.IsNullOrWhiteSpace(xccdfPath) || !File.Exists(xccdfPath))
            {
                AppendInfo("Invalid XCCDF file path. Please select a valid XML.", System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
                return;
            }
            if (string.IsNullOrWhiteSpace(destination))
            {
                AppendInfo("Invalid destination folder.", System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
                return;
            }
            try
            {
                Directory.CreateDirectory(destination);
            }
            catch (Exception ex)
            {
                AppendInfo($"Failed to ensure destination folder: {ex.Message}", System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
                return;
            }

            ConvertButton.IsEnabled = false;
            SetBusy(true, "Converting…");
            AppendInfo("Starting conversion with Windows PowerShell 5.x...", System.Windows.Media.Brushes.DarkGreen, null);

            try
            {
                // Setup ProcessStartInfo for PowerShell execution

                var moduleRoot = Path.GetDirectoryName(modulePath)!; // if psm1 is under module root
                var tempScript = Path.Combine(Path.GetTempPath(), $"PowerStigConvert_{Guid.NewGuid():N}.ps1");
                var scriptContent = @"
param(
    [string]$Psm1,
    [string]$XccdfPath,
    [string]$Destination
)
$ErrorActionPreference = 'Stop'
Import-Module -Force -ErrorAction Stop -Name $Psm1
ConvertTo-PowerStigXml -Destination $Destination -Path $XccdfPath -CreateOrgSettingsFile:$" + (createOrgSettings ? "true" : "false") + @"
";
                File.WriteAllText(tempScript, scriptContent, new UTF8Encoding(false));

                // Build arguments for -File without brittle quoting
                var psArgs = $"-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File \"{tempScript}\" \"{modulePath}\" \"{xccdfPath}\" \"{destination}\" \"Windows\"";

                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe", // Windows PowerShell 5.x
                    Arguments = psArgs,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WorkingDirectory = moduleRoot
                };

                string stdOut = string.Empty, stdErr = string.Empty;

                await System.Threading.Tasks.Task.Run(() =>
                {
                    using var proc = new Process { StartInfo = psi, EnableRaisingEvents = true };
                    var outSb = new StringBuilder();
                    var errSb = new StringBuilder();

                    proc.OutputDataReceived += (s, ea) =>
                    {
                        var line = ea.Data;
                        if (string.IsNullOrWhiteSpace(line)) return;
                        outSb.AppendLine(line);

                        if (TryFormatRuleFailure(line, out var compact))
                        {
                            var rid = ExtractRuleIdFromFailure(compact);
                            if (!string.IsNullOrWhiteSpace(rid))
                            {
                                _failedRuleIds.Add(rid!);
                                // Store detailed error text per rule
                                // compact format: "Rule V-XXXXX failed: <errorText>"
                                var sepIdx = compact.IndexOf(" failed:", StringComparison.OrdinalIgnoreCase);
                                string error = sepIdx >= 0 ? compact.Substring(sepIdx + " failed:".Length).Trim() : compact;
                                _failedRuleErrors[rid!] = error;
                            }
                            // No immediate UI output; we print a grouped summary after conversion
                            return;
                        }

                        var isError = line.StartsWith("ERROR:", StringComparison.OrdinalIgnoreCase) ||
                                      line.IndexOf("Exception", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                      line.IndexOf("CategoryInfo", StringComparison.OrdinalIgnoreCase) >= 0;

                        if (isError)
                            AppendInfoThreadSafe(line, System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
                    };

                    proc.ErrorDataReceived += (s, ea) =>
                    {
                        var line = ea.Data;
                        if (string.IsNullOrWhiteSpace(line)) return;
                        errSb.AppendLine(line);
                        AppendInfoThreadSafe(line, System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
                    };

                    proc.Start();
                    proc.BeginOutputReadLine();
                    proc.BeginErrorReadLine();
                    proc.WaitForExit();

                    stdOut = outSb.ToString();
                    stdErr = errSb.ToString();

                    try { File.Delete(tempScript); } catch { /* ignore */ }

                    if (proc.ExitCode != 0)
                        throw new InvalidOperationException($"Windows PowerShell exited with code {proc.ExitCode}. First error:\n{stdErr.Trim()}");
                });

                // Collect converted V- IDs to emit success/failure messages
                string? convertedFileName = GetConvertedFileName(xccdfPath!);
                convertedFileName = ResolveConvertedFullPath(destination!, convertedFileName!);

                var convertedIds = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var id in ExtractRuleIdsFromConverted(convertedFileName))
                    convertedIds.Add(id);

                // Emit failures sorted by numeric portion of V- IDs
                if (_failedRuleIds.Count > 0)
                {
                    var sortedFailures = _failedRuleIds
                        .Select(id => id.Trim())
                        .Where(id => id.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(id => ExtractNumericKey(id, prefix: "V-"))
                        .ThenBy(id => id, StringComparer.OrdinalIgnoreCase);

                    AppendInfo($"Failed rule conversions ({_failedRuleIds.Count}):", System.Windows.Media.Brushes.OrangeRed, null);
                    foreach (var vId in sortedFailures)
                    {
                        _failedRuleErrors.TryGetValue(vId, out var err);
                        var msg = string.IsNullOrWhiteSpace(err) ? $"{vId} failed." : $"{vId} failed: {err}";
                        AppendInfo(msg, System.Windows.Media.Brushes.OrangeRed, null);
                    }
                }

                // Always show successful conversions, regardless of Skip Compare
                var successIds = convertedIds
                    .Where(id => !_failedRuleIds.Contains(id))
                    .Select(id => id.Trim())
                    .Where(id => id.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(id => ExtractNumericKey(id, prefix: "V-"))
                    .ThenBy(id => id, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (successIds.Count > 0)
                {
                    AppendInfo($"Successfully converted rule IDs ({successIds.Count}):", System.Windows.Media.Brushes.DarkGreen, null);
                    foreach (var id in successIds)
                    {
                        AppendInfo($" - {id}", System.Windows.Media.Brushes.DarkGreen, null);
                    }
                }
                else
                {
                    AppendInfo("No successfully converted rule IDs.", System.Windows.Media.Brushes.DarkGreen, null);
                }

                // Only perform "missing" comparison when Skip Compare is unchecked
                var skipCompare = SkipCompareCheckBox?.IsChecked == true;
                if (!skipCompare)
                {
                    // NEW: make it clear in the Messages panel that comparison is starting
                    AppendInfo("Comparing converted output against XCCDF (this may take a moment)…", System.Windows.Media.Brushes.DarkSlateBlue, null);

                    var missing = CompareRuleIds(xccdfPath!, convertedFileName!);
                    if (missing.Count > 0)
                    {
                        AppendInfo($"Missing rule IDs ({missing.Count}) — present in XCCDF but not in converted output:", System.Windows.Media.Brushes.Firebrick, null);
                        foreach (var mid in missing
                            .Select(id => id.Trim())
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .OrderBy(id => ExtractNumericKey(id, prefix: id.StartsWith("SV-", StringComparison.OrdinalIgnoreCase) ? "SV-" :
                                                   id.StartsWith("V-", StringComparison.OrdinalIgnoreCase) ? "V-" : string.Empty))
                            .ThenBy(id => id, StringComparer.OrdinalIgnoreCase))
                        {
                            AppendInfo($" - {mid}", System.Windows.Media.Brushes.Firebrick, null);
                        }
                    }
                    else
                    {
                        AppendInfo("No missing rule IDs detected.", System.Windows.Media.Brushes.DarkGreen, null);
                    }
                }

                AppendInfo("Conversion completed.", System.Windows.Media.Brushes.DarkGreen, System.Windows.Media.Brushes.LightGreen);
                InfoRichTextBox.ScrollToHome();
            }
            catch (Exception ex)
            {
                AppendInfo($"Conversion failed: {ex.Message}", System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
            }
            finally
            {
                SetBusy(false);
                ConvertButton.IsEnabled = true;
            }
        }

        private static string? GetConvertedFileName(string originalFileName)
        {
            if (string.IsNullOrWhiteSpace(originalFileName)) return null;

            // Accept fully qualified paths or plain filenames
            var name = Path.GetFileName(originalFileName);

            // Examples:
            // U_MS_Windows_Server_2022_MS_STIG_V2R6_Manual-xccdf.xml -> WindowsServer-2022-MS-2.6.xml
            // U_MS_IIS_10-0_Site_STIG_V2R13_Manual-xccdf.xml         -> IISSite-10.0-2.13.xml
            // U_MS_DotNet_Framework_4-0_STIG_V2R7_Manual-xccdf.xml   -> DotNetFramework-4.0-2.7.xml
            var m = Regex.Match(
                name,
                @"^U_MS_(?<product>.+?)_(?:(?<edition>MS|DC)_)?STIG_(?<ver>V\d+R\d+)_Manual-xccdf\.xml$",
                RegexOptions.IgnoreCase);
            if (!m.Success) return null;

            var productRaw = m.Groups["product"].Value;
            var verRaw = m.Groups["ver"].Value;
            var editionToken = m.Groups["edition"].Success ? m.Groups["edition"].Value.ToUpperInvariant() : null;

            // Convert V#R# -> #.#
            var vm = Regex.Match(verRaw, @"^V(?<major>\d+)R(?<rev>\d+)$", RegexOptions.IgnoreCase);
            if (!vm.Success) return null;
            var verShort = $"{vm.Groups["major"].Value}.{vm.Groups["rev"].Value}";

            // Split product by underscores and normalize
            var tokens = productRaw.Split('_', StringSplitOptions.RemoveEmptyEntries)
                                   .Select(t => t.Trim())
                                   .ToList();

            // Handle optional "Site" suffix (append to base name)
            var siteSuffix = tokens.Contains("Site", StringComparer.OrdinalIgnoreCase) ? "Site" : null;
            tokens = tokens.Where(t => !string.Equals(t, "Site", StringComparison.OrdinalIgnoreCase)).ToList();

            // Identify version-like numeric token (e.g., 2022, 10-0, 4-0)
            string? numericToken = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^\d+(-\d+)?$"));
            if (numericToken != null)
                tokens = tokens.Where(t => !string.Equals(t, numericToken, StringComparison.OrdinalIgnoreCase)).ToList();

            // Build PascalCase product name (e.g., WindowsServer, DotNetFramework, IIS)
            var baseName = string.Concat(tokens.Select(t => char.ToUpperInvariant(t[0]) + t.Substring(1)));

            // Normalize numeric token: always preserve minor (e.g., "10-0" -> "10.0")
            string? normalizedNumber = null;
            if (numericToken != null)
            {
                var nm = Regex.Match(numericToken, @"^(?<a>\d+)(?:-(?<b>\d+))?$");
                if (nm.Success)
                {
                    var a = nm.Groups["a"].Value;
                    var b = nm.Groups["b"].Success ? nm.Groups["b"].Value : null;
                    normalizedNumber = b is null ? a : $"{a}.{b}";
                }
            }

            var parts = new System.Collections.Generic.List<string> { baseName + (siteSuffix ?? string.Empty) };
            if (!string.IsNullOrWhiteSpace(normalizedNumber)) parts.Add(normalizedNumber);
            if (!string.IsNullOrWhiteSpace(editionToken)) parts.Add(editionToken); // Add MS/DC when present
            parts.Add(verShort);

            return string.Join("-", parts) + ".xml";
        }

        // Add this helper method inside the ConvertStigWindow class (anywhere in the class, e.g., near other static helpers)
        private static string NormalizeVId(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return string.Empty;
            // Normalize "V-123456.a" or "V-123456" to "V-123456"
            var m = Regex.Match(id, @"^(V-\d+)", RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value : id;
        }

        // Helper: extract rule id from a compact failure message like "Rule V-278953 failed: ..."
        private static string? ExtractRuleIdFromFailure(string compactMessage)
        {
            var start = compactMessage.IndexOf("Rule ", StringComparison.OrdinalIgnoreCase);
            if (start < 0) return null;
            start += 5; // after "Rule "
            var failedIdx = compactMessage.IndexOf(" failed", StringComparison.OrdinalIgnoreCase);
            if (failedIdx < 0 || failedIdx <= start) return null;
            var id = compactMessage.Substring(start, failedIdx - start).Trim();
            return string.IsNullOrWhiteSpace(id) ? null : id;
        }

        private static System.Collections.Generic.HashSet<string> CompareRuleIds(string xccdfPath, string convertedPath)
        {
            // Normalize converted IDs: "V-123456.a" -> "V-123456"
            var convertedIdsRaw = ExtractRuleIdsFromConverted(convertedPath);
            var convertedIds = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var id in convertedIdsRaw)
            {
                var trimmed = NormalizeVId(id);
                if (!string.IsNullOrWhiteSpace(trimmed))
                    convertedIds.Add(trimmed);
            }

            // Extract XCCDF IDs (was missing)
            var xccdfIds = ExtractRuleIdsFromXccdfAndConvert(xccdfPath);

            // Missing = present in XCCDF but not in converted output
            var missing = new System.Collections.Generic.HashSet<string>(xccdfIds, StringComparer.OrdinalIgnoreCase);
            missing.ExceptWith(convertedIds);
            return missing;
        }

        private static System.Collections.Generic.HashSet<string> ExtractRuleIdsFromXccdfAndConvert(string xccdfPath)
        {
            var ids = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using var reader = System.Xml.XmlReader.Create(xccdfPath, new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });
                while (reader.Read())
                {
                    if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                        reader.LocalName.Equals("Rule", StringComparison.OrdinalIgnoreCase))
                    {
                        var id = reader.GetAttribute("id");
                        if (!string.IsNullOrWhiteSpace(id))
                            ids.Add(id.Trim());
                    }
                }
            }
            catch
            {
                // Fallback: text search for id="V-..."
                foreach (var line in File.ReadLines(xccdfPath))
                {
                    var i = line.IndexOf("id=\"V-", StringComparison.OrdinalIgnoreCase);
                    if (i >= 0)
                    {
                        var start = i + 4; // after id="
                        var end = line.IndexOf('"', start);
                        if (end > start)
                        {
                            var id = line.Substring(start, end - start);
                            if (!string.IsNullOrWhiteSpace(id))
                                ids.Add(id.Trim());
                        }
                    }
                }
            }
            // Normalize any SV-... identifiers to V-<digits>
            // Examples:
            // "SV-225223r961038_rule" -> "V-225223"
            // "SV-123456" -> "V-123456"
            var normalized = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var svRegex = new Regex(@"SV-(\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

            foreach (var id in ids)
            {
                if (id.StartsWith("SV-", StringComparison.OrdinalIgnoreCase))
                {
                    var m = svRegex.Match(id);
                    if (m.Success)
                    {
                        normalized.Add($"V-{m.Groups[1].Value}");
                        continue;
                    }
                }
                normalized.Add(id);
            }

            return normalized;
        }

        private static System.Collections.Generic.IEnumerable<string> ExtractRuleIdsFromConverted(string convertedXmlPath)
        {
            var ids = new System.Collections.Generic.List<string>();
            try
            {
                var doc = new System.Xml.XmlDocument();
                doc.Load(convertedXmlPath);

                // IDs as attributes like id="V-12345"
                var attrNodes = doc.SelectNodes("//*[@id]");
                if (attrNodes is not null)
                {
                    foreach (System.Xml.XmlNode n in attrNodes)
                    {
                        var id = n.Attributes?["id"]?.Value;
                        if (!string.IsNullOrWhiteSpace(id) && id.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                            ids.Add(id.Trim());
                    }
                }

                // IDs as elements, common names
                var ruleIdNodes = doc.SelectNodes("//*[local-name()='RuleId' or local-name()='VulnId' or local-name()='BenchmarkId']");
                if (ruleIdNodes is not null)
                {
                    foreach (System.Xml.XmlNode n in ruleIdNodes)
                    {
                        var val = n.InnerText;
                        if (!string.IsNullOrWhiteSpace(val) && val.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                            ids.Add(val.Trim());
                    }
                }
            }
            catch
            {
                // ignore parse errors
            }
            return ids;
        }

        // Helper: format "Conversion for V-#### failed. Error: ..." into compact UI text
        private static bool TryFormatRuleFailure(string line, out string formatted)
        {
            formatted = string.Empty;

            var markerIdx = line.IndexOf("Conversion for V-", StringComparison.OrdinalIgnoreCase);
            if (markerIdx < 0) return false;

            var msg = line.Substring(markerIdx).Trim();
            var ruleStart = msg.IndexOf("V-", StringComparison.OrdinalIgnoreCase);
            var failedIdx = msg.IndexOf("failed", StringComparison.OrdinalIgnoreCase);
            var errorIdx = msg.IndexOf("Error:", StringComparison.OrdinalIgnoreCase);

            if (ruleStart >= 0 && failedIdx >= 0 && errorIdx >= 0)
            {
                var rule = msg.Substring(ruleStart, failedIdx - ruleStart).Trim();
                var errorText = msg.Substring(errorIdx + "Error:".Length).Trim();

                // Clean common CLIXML encoding artifacts
                errorText = errorText.Replace("_x000D__x000A_", " ").Trim().Trim('"');

                formatted = $"Rule {rule} failed: {errorText}";
                return true;
            }

            return false;
        }

        // Helper to escape script for cmd/powershell.exe -Command
        private static string EscapeForCmd(string script)
        {
            var oneLine = script.Replace("\r", "").Replace("\n", "; ");
            return oneLine.Replace("\"", "`\"");
        }

        private void BrowseDestination_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new FolderBrowserDialog
            {
                Description = "Select destination folder"
            };
            if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DestinationTextBox.Text = folderDialog.SelectedPath;
            }
        }

        private void BrowseXccdf_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
                Title = "Select XCCDF XML File"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                XccdfPathTextBox.Text = openFileDialog.FileName;
            }
        }

        private void BrowseModule_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "PowerStig.Convert.psm1|PowerStig.Convert.psm1|All files (*.*)|*.*",
                Title = "Select PowerStig.Convert.psm1"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                ModulePathTextBox.Text = openFileDialog.FileName;
                WriteLastModulePath(openFileDialog.FileName);
            }
        }

        private void AppendInfo(string message)
        {
            AppendInfo(message, foreground: null, background: null);
        }

        private void AppendInfo(string message, System.Windows.Media.Brush? foreground, System.Windows.Media.Brush? background)
        {
            var para = new System.Windows.Documents.Paragraph();
            var run = new System.Windows.Documents.Run($"[{DateTime.Now:HH:mm:ss}] {message}");
            if (foreground != null) run.Foreground = foreground;
            if (background != null) para.Background = background;
            para.Inlines.Add(run);
            InfoRichTextBox.Document.Blocks.Add(para);
            InfoRichTextBox.ScrollToEnd();
        }

        private void AppendInfoThreadSafe(string? message, System.Windows.Media.Brush? foreground = null, System.Windows.Media.Brush? background = null)
        {
            if (message is null) return;
            Dispatcher.Invoke(() => AppendInfo(message, foreground, background));
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private static int ExtractNumericKey(string id, string prefix)
        {
            if (string.IsNullOrWhiteSpace(id)) return int.MaxValue;
            var s = id.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? id.Substring(prefix.Length) : id;
            int i = 0;
            while (i < s.Length && char.IsDigit(s[i])) i++;
            var digits = (i > 0) ? s.Substring(0, i) : string.Empty;
            return int.TryParse(digits, out var n) ? n : int.MaxValue;
        }

        // Map an SV- ident to its owning Rule's V- id (if any)
        private static string? ResolveVIdForSv(string svId, string xccdfPath)
        {
            if (string.IsNullOrWhiteSpace(svId) || !svId.StartsWith("SV-", StringComparison.OrdinalIgnoreCase))
                return null;

            try
            {
                using var reader = System.Xml.XmlReader.Create(
                    xccdfPath,
                    new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });

                while (reader.Read())
                {
                    if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                        reader.LocalName.Equals("Rule", StringComparison.OrdinalIgnoreCase))
                    {
                        var vId = reader.GetAttribute("id")?.Trim();
                        using var subTree = reader.ReadSubtree();
                        while (subTree.Read())
                        {
                            if (subTree.NodeType == System.Xml.XmlNodeType.Element &&
                                subTree.LocalName.Equals("ident", StringComparison.OrdinalIgnoreCase))
                            {
                                var identText = subTree.ReadElementContentAsString()?.Trim();
                                if (string.Equals(identText, svId, StringComparison.OrdinalIgnoreCase))
                                    return vId;
                            }
                        }
                    }
                }
            }
            catch { /* ignore */ }
            return null;
        }

        // Double-click handler on InfoRichTextBox to open rule info
        private void InfoRichTextBox_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var lineText = GetLineUnderMouse(InfoRichTextBox, e) ?? GetWordUnderMouse(InfoRichTextBox, e);
            if (string.IsNullOrWhiteSpace(lineText))
                return;

            var m = Regex.Match(lineText, @"\b(SV-\d+|V-\d+)\b", RegexOptions.IgnoreCase);
            if (!m.Success)
            {
                m = Regex.Match(lineText, @"\b(SV-\d+|V-\d+)[^0-9]", RegexOptions.IgnoreCase);
                if (!m.Success) return;
            }
            var rawId = Regex.Match(m.Value, @"\b(SV-\d+|V-\d+)\b", RegexOptions.IgnoreCase).Value;

            var xccdfPath = XccdfPathTextBox.Text?.Trim();
            var destination = DestinationTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(xccdfPath) || !File.Exists(xccdfPath))
            {
                System.Windows.MessageBox.Show("XCCDF path is not set or invalid.", "Rule Info", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // SV rules: only use the original XCCDF; do not pass converted folder
            var isSv = rawId.StartsWith("SV-", StringComparison.OrdinalIgnoreCase);
            var info = RuleInfoExtractor.TryExtractRuleInfo(rawId, xccdfPath!, isSv ? null : destination);

            if (isSv)
            {
                // Normalize and ensure SV metadata is present; avoid converted artifacts
                info.SvId ??= Regex.Match(rawId, @"SV-\d+", RegexOptions.IgnoreCase).Value;
                info.ConvertedFile = null;
                info.ConvertedSnippet = null;
            }

            var win = new RuleInfoWindow { Owner = this };
            win.SetRuleInfo(info);
            win.ShowDialog();
            e.Handled = true;
        }

        private static string? GetWordUnderMouse(System.Windows.Controls.RichTextBox rtb, MouseButtonEventArgs e)
        {
            var pos = rtb.GetPositionFromPoint(e.GetPosition(rtb), true);
            if (pos == null) return null;

            var wordRange = GetWordRange(pos);
            return wordRange?.Text;
        }

        private static string? GetLineUnderMouse(System.Windows.Controls.RichTextBox rtb, MouseButtonEventArgs e)
        {
            var pos = rtb.GetPositionFromPoint(e.GetPosition(rtb), true);
            if (pos == null) return null;

            // Expand to paragraph
            var para = pos.Paragraph;
            return para?.ContentStart != null && para?.ContentEnd != null
                ? new TextRange(para.ContentStart, para.ContentEnd).Text
                : null;
        }

        private static TextRange? GetWordRange(TextPointer position)
        {
            var wordStart = GetPositionAtWordBoundary(position, LogicalDirection.Backward);
            var wordEnd = GetPositionAtWordBoundary(position, LogicalDirection.Forward);
            if (wordStart == null || wordEnd == null)
                return null;
            return new TextRange(wordStart, wordEnd);
        }

        private static TextPointer? GetPositionAtWordBoundary(TextPointer position, LogicalDirection direction)
        {
            var navigator = position;
            while (navigator != null && !IsWordBoundary(navigator, direction))
            {
                navigator = navigator.GetPositionAtOffset(direction == LogicalDirection.Forward ? 1 : -1, LogicalDirection.Forward);
            }
            return navigator;
        }

        private static bool IsWordBoundary(TextPointer position, LogicalDirection direction)
        {
            var charType = GetCharType(position, direction);
            var adjacentPos = position.GetPositionAtOffset(direction == LogicalDirection.Forward ? 1 : -1, LogicalDirection.Forward);
            var adjacentCharType = GetCharType(adjacentPos, direction);

            return charType == CharType.Character && adjacentCharType == CharType.WhiteSpace
                   || charType == CharType.WhiteSpace && adjacentCharType == CharType.Character
                   || adjacentPos == null;
        }

        private enum CharType { None, WhiteSpace, Character }

        private static CharType GetCharType(TextPointer? position, LogicalDirection direction)
        {
            if (position == null) return CharType.None;
            char[] buffer = new char[1];
            var tr = position.GetTextInRun(direction, buffer, 0, 1);
            if (tr == 0) return CharType.None;
            var ch = buffer[0];
            if (char.IsWhiteSpace(ch)) return CharType.WhiteSpace;
            return CharType.Character;
        }

        // ... existing methods ...
    }

    internal static class RuleInfoExtractor
    {
        // Map an SV- ident to its owning Rule's V- id (if any)
        // Map an SV- ident to its owning Rule's V- id (if any)
        private static string? ResolveVIdForSv(string svId, string xccdfPath)
        {
            if (string.IsNullOrWhiteSpace(svId) || !svId.StartsWith("SV-", StringComparison.OrdinalIgnoreCase))
                return null;

            // Normalize to base "SV-<digits>" to handle inputs like "SV-225223r961038_rule"
            var baseSv = Regex.Match(svId, @"SV-\d+", RegexOptions.IgnoreCase).Value;
            if (string.IsNullOrWhiteSpace(baseSv)) return null;

            try
            {
                using var reader = System.Xml.XmlReader.Create(
                    xccdfPath,
                    new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });

                while (reader.Read())
                {
                    if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                        reader.LocalName.Equals("Rule", StringComparison.OrdinalIgnoreCase))
                    {
                        var vId = reader.GetAttribute("id")?.Trim();
                        using var subTree = reader.ReadSubtree();
                        while (subTree.Read())
                        {
                            if (subTree.NodeType == System.Xml.XmlNodeType.Element &&
                                subTree.LocalName.Equals("ident", StringComparison.OrdinalIgnoreCase))
                            {
                                var identText = subTree.ReadElementContentAsString()?.Trim();
                                var systemAttr = subTree.GetAttribute("system");
                                if ((!string.IsNullOrWhiteSpace(identText) && identText.StartsWith(baseSv, StringComparison.OrdinalIgnoreCase)) ||
                                    (!string.IsNullOrWhiteSpace(systemAttr) && systemAttr.IndexOf(baseSv, StringComparison.OrdinalIgnoreCase) >= 0))
                                {
                                    return vId;
                                }
                            }
                        }
                    }
                }
            }
            catch { /* ignore */ }
            return null;
        }

        public static RuleInfo TryExtractRuleInfo(string ruleId, string xccdfPath, string? convertedFolder)
        {
            var info = new RuleInfo { RuleId = ruleId };

            // Always fill from XCCDF (original file)
            TryFillFromXccdf(ruleId, xccdfPath, info);

            // Only use converted outputs for V- rules
            if (ruleId.StartsWith("V-", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(convertedFolder) && Directory.Exists(convertedFolder))
            {
                TryFillFromConverted(info.RuleId, convertedFolder!, info);
            }
            else
            {
                // Guard: ensure SV rules do not carry converted artifacts
                info.ConvertedFile = null;
                info.ConvertedSnippet = null;
            }

            return info;
        }

        private static void TryFillFromXccdf(string ruleId, string xccdfPath, RuleInfo info)
        {
            try
            {
                var doc = new System.Xml.XmlDocument { PreserveWhitespace = true };
                var settings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore };
                using var reader = System.Xml.XmlReader.Create(xccdfPath, settings);
                doc.Load(reader);

                System.Xml.XmlNode? ruleNode = null;

                if (ruleId.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                {
                    ruleNode = doc.SelectSingleNode($"//*[local-name()='Rule' and @id='{ruleId}']");
                }
                else if (ruleId.StartsWith("SV-", StringComparison.OrdinalIgnoreCase))
                {
                    var baseSv = Regex.Match(ruleId, @"SV-\d+", RegexOptions.IgnoreCase).Value;
                    ruleNode = doc.SelectSingleNode(
                        $"//*[local-name()='Rule'][.//*[local-name()='ident' and (starts-with(normalize-space(text()),'{baseSv}') or contains(@system,'{baseSv}'))]]");

                    // If found via SV ident, keep SvId and set RuleId to owning V- id if available
                    if (ruleNode is not null)
                    {
                        info.SvId ??= baseSv;
                        var vIdAttr = ruleNode.Attributes?["id"]?.Value?.Trim();
                        if (!string.IsNullOrWhiteSpace(vIdAttr)) info.RuleId = vIdAttr;
                    }
                    else
                    {
                        // Try resolving to V- to fetch the rule node
                        var vId = ResolveVIdForSv(ruleId, xccdfPath);
                        if (!string.IsNullOrWhiteSpace(vId))
                            ruleNode = doc.SelectSingleNode($"//*[local-name()='Rule' and @id='{vId}']");
                    }
                }

                if (ruleNode is null)
                    return;

                var vIdAttr2 = ruleNode.Attributes?["id"]?.Value?.Trim();
                if (!string.IsNullOrWhiteSpace(vIdAttr2)) info.RuleId = vIdAttr2;

                var severityAttr = ruleNode.Attributes?["severity"]?.Value;
                if (!string.IsNullOrWhiteSpace(severityAttr)) info.Severity = severityAttr;

                var titleNode = ruleNode.SelectSingleNode(".//*[local-name()='title']");
                if (titleNode is not null) info.Title = titleNode.InnerText.Trim();

                var descNodes = ruleNode.SelectNodes(".//*[local-name()='description']");
                if (descNodes is not null && descNodes.Count > 0)
                {
                    var sb = new StringBuilder();
                    foreach (System.Xml.XmlNode dn in descNodes)
                        sb.AppendLine(dn.InnerText.Trim());
                    info.Description = sb.ToString().Trim();
                }

                var fixNode = ruleNode.SelectSingleNode(".//*[local-name()='fixtext']");
                if (fixNode is not null) info.FixText = fixNode.InnerText.Trim();

                var identNode = ruleNode.SelectSingleNode(".//*[local-name()='ident' and (starts-with(normalize-space(text()),'SV-') or contains(@system,'SV-'))]");
                if (identNode is not null)
                {
                    var identText = identNode.InnerText?.Trim();
                    var systemAttr = identNode.Attributes?["system"]?.Value;
                    if (!string.IsNullOrWhiteSpace(identText) && identText.StartsWith("SV-", StringComparison.OrdinalIgnoreCase))
                        info.SvId ??= identText;
                    else if (!string.IsNullOrWhiteSpace(systemAttr))
                    {
                        var m2 = Regex.Match(systemAttr, @"SV-\d+", RegexOptions.IgnoreCase);
                        if (m2.Success) info.SvId ??= m2.Value;
                    }
                }

                var refNodes = ruleNode.SelectNodes(".//*[local-name()='reference']");
                if (refNodes is not null && refNodes.Count > 0)
                {
                    var sb = new StringBuilder();
                    foreach (System.Xml.XmlNode rn in refNodes)
                        sb.AppendLine(rn.OuterXml);
                    info.ReferencesXml = sb.ToString().Trim();
                }
            }
            catch
            {
                // ignore parse errors
            }
        }

        private static void TryFillFromConverted(string ruleId, string folder, RuleInfo info)
        {
            try
            {
                foreach (var file in Directory.EnumerateFiles(folder, "*.xml", SearchOption.TopDirectoryOnly))
                {
                    var doc = new System.Xml.XmlDocument();
                    doc.Load(file);

                    // Try attributes and common element names
                    var candidates = doc.SelectNodes($"//*[@id='{ruleId}']");
                    if (candidates is not null && candidates.Count > 0)
                    {
                        info.ConvertedFile = file;
                        info.ConvertedSnippet = GetNodeSummary(candidates[0]);
                        return;
                    }

                    var node = doc.SelectSingleNode(
                        $"//*[local-name()='RuleId' or local-name()='VulnId' or local-name()='BenchmarkId'][text()='{ruleId}']");
                    if (node is not null)
                    {
                        info.ConvertedFile = file;
                        info.ConvertedSnippet = GetNodeSummary(node);
                        return;
                    }
                }
            }
            catch
            {
                // ignore
            }
        }

        private static string GetNodeSummary(System.Xml.XmlNode node)
        {
            var sb = new StringBuilder();
            sb.AppendLine(node.OuterXml.Length <= 4000 ? node.OuterXml : node.OuterXml.Substring(0, 4000) + "...");
            return sb.ToString();
        }
    }

}