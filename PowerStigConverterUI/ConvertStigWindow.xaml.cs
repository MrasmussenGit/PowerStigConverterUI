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

        private System.Collections.Generic.List<string>? _lastSuccessIds = new();
        private System.Windows.Threading.DispatcherTimer? _busyTimer;
        private DateTime _busyStart;
        private readonly System.Collections.Generic.HashSet<string> _skippedRuleIds =
    new(System.StringComparer.OrdinalIgnoreCase);
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
            static string Combine(string dest, string n) => Path.Combine(dest, n);

            // Start with the primary name
            var candidates = new System.Collections.Generic.List<string> { convertedFileName };

            // If the primary contains an edition token (MS/DC), also try without it and at the end
            var nameNoExt = Path.GetFileNameWithoutExtension(convertedFileName);
            var ext = Path.GetExtension(convertedFileName);
            var parts = nameNoExt.Split('-', StringSplitOptions.RemoveEmptyEntries).ToList();
            var edIdx = parts.FindIndex(p => p.Equals("MS", StringComparison.OrdinalIgnoreCase) || p.Equals("DC", StringComparison.OrdinalIgnoreCase));
            if (edIdx >= 0)
            {
                var edition = parts[edIdx];

                // a) Without edition
                var noEd = parts.Where((p, i) => i != edIdx).ToArray();
                candidates.Add(string.Join("-", noEd) + ext);

                // b) Edition at the very end
                var edAtEnd = parts.Where((p, i) => i != edIdx).Concat(new[] { edition }).ToArray();
                candidates.Add(string.Join("-", edAtEnd) + ext);
            }

            // For each candidate, also consider 4 vs 4.0 variants (module differences)
            static System.Collections.Generic.IEnumerable<string> NumberVariants(string fileName)
            {
                yield return fileName;

                var a = Regex.Replace(fileName, @"-(\d+)\.0-", m => $"-{m.Groups[1].Value}-", RegexOptions.IgnoreCase);
                if (!string.Equals(a, fileName, StringComparison.OrdinalIgnoreCase)) yield return a;

                var b = Regex.Replace(fileName, @"-(\d+)-", m => $"-{m.Groups[1].Value}.0-", RegexOptions.IgnoreCase);
                if (!string.Equals(b, fileName, StringComparison.OrdinalIgnoreCase)) yield return b;
            }

            foreach (var cand in candidates.SelectMany(NumberVariants).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var path = Combine(destination, cand);
                if (File.Exists(path))
                    return path;
            }

            // Fallback: search by product + version while ignoring edition placement
            try
            {
                if (parts.Count >= 2)
                {
                    var version = parts.Last();
                    var product = string.Join("-", parts.Take(parts.Count - 1).Where((p, i) => i != edIdx));
                    var match = Directory.EnumerateFiles(destination, "*.xml", SearchOption.TopDirectoryOnly)
                                         .FirstOrDefault(f =>
                                             Path.GetFileName(f).StartsWith(product + "-", StringComparison.OrdinalIgnoreCase) &&
                                             Path.GetFileNameWithoutExtension(f).EndsWith("-" + version, StringComparison.OrdinalIgnoreCase));
                    if (match != null) return match;
                }
            }
            catch { /* ignore */ }

            // As a final fallback, return the primary path
            return Combine(destination, convertedFileName);
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

        private static (System.Collections.Generic.HashSet<string> Skips, System.Collections.Generic.HashSet<string> HardCoded) ParseLogFile(string logPath)
        {
            var skips = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var hardCoded = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (!File.Exists(logPath))
                return (skips, hardCoded);

            try
            {
                foreach (var line in File.ReadLines(logPath))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var parts = line.Split(new[] { "::" }, StringSplitOptions.None);
                    if (parts.Length < 3 || !parts[0].StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var ruleId = parts[0].Trim();
                    var field2 = parts[1].Trim();
                    var field3 = parts[2].Trim();

                    // SKIP: exact format V-XXXXX::*::.
                    if (field2 == "*" && field3 == ".")
                    {
                        skips.Add(ruleId);
                    }
                    else
                    {
                        // Everything else is a Hard Coded rule (manual override)
                        // This includes:
                        // - V-XXXXX::*::HardCodedRule(...)
                        // - V-XXXXX::custom text::custom text
                        hardCoded.Add(ruleId);
                    }
                }
            }
            catch { /* ignore parse errors */ }

            return (skips, hardCoded);
        }

        // Update ParseExistingLogEntries to only return SKIPs (for duplicate detection when writing)
        private static System.Collections.Generic.HashSet<string> ParseExistingLogEntries(string logPath)
        {
            var (skips, _) = ParseLogFile(logPath);
            return skips;
        }

        // Append failed rules as SKIP entries: V-XXXXX::*::.
        private static string? DeriveLogFilePathFromXccdf(string? xccdfPath)
        {
            if (string.IsNullOrWhiteSpace(xccdfPath)) return null;

            var directory = Path.GetDirectoryName(xccdfPath);
            var fileNameNoExt = Path.GetFileNameWithoutExtension(xccdfPath); // U_MS_IIS_8-5_Server_STIG_V2R4_Manual-xccdf

            if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(fileNameNoExt))
                return null;

            // Result: U_MS_IIS_8-5_Server_STIG_V2R4_Manual-xccdf.log in the same directory
            return Path.Combine(directory, fileNameNoExt + ".log");
        }
        private static void AppendFailedRulesToLog(string logPath, System.Collections.Generic.IEnumerable<string> failedRuleIds)
        {
            try
            {
                // Load existing entries to avoid duplicates
                var existingRules = ParseExistingLogEntries(logPath);

                var newEntries = failedRuleIds
                    .Where(ruleId => !existingRules.Contains(ruleId))
                    .OrderBy(id => ExtractNumericKey(id, "V-"))
                    .ThenBy(id => id, StringComparer.OrdinalIgnoreCase)
                    .Select(ruleId => $"{ruleId}::*::.")
                    .ToList();

                if (newEntries.Count == 0) return;

                // Ensure directory exists
                var directory = Path.GetDirectoryName(logPath);
                if (!string.IsNullOrWhiteSpace(directory))
                    Directory.CreateDirectory(directory);

                // Append new SKIP entries
                File.AppendAllLines(logPath, newEntries, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to write to log file '{logPath}': {ex.Message}", ex);
            }
        }

        private static System.Collections.Generic.HashSet<string> ExtractRulesWithNoDscResource(string convertedXmlPath)
        {
            var ids = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var doc = new System.Xml.XmlDocument();
                doc.Load(convertedXmlPath);

                // Find all Rule elements with dscresource="None"
                var noDscNodes = doc.SelectNodes("//*[local-name()='Rule' and @dscresource='None']");
                if (noDscNodes is not null)
                {
                    foreach (System.Xml.XmlNode n in noDscNodes)
                    {
                        var id = n.Attributes?["id"]?.Value;
                        if (!string.IsNullOrWhiteSpace(id) && id.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                        {
                            // Normalize to base ID (V-123456.a → V-123456)
                            var normalized = NormalizeVId(id.Trim());
                            if (!string.IsNullOrWhiteSpace(normalized))
                                ids.Add(normalized);
                        }
                    }
                }
            }
            catch
            {
                // ignore parse errors
            }
            return ids;
        }
        private async void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            InfoRichTextBox.Document.Blocks.Clear();
            _failedRuleIds.Clear();
            _failedRuleErrors.Clear();
            _skippedRuleIds.Clear();

            var modulePath = ModulePathTextBox.Text?.Trim();
            var xccdfPath = XccdfPathTextBox.Text?.Trim();
            var destination = DestinationTextBox.Text?.Trim();
            var createOrgSettings = CreateOrgSettingsCheckBox.IsChecked == true;
            var addFailedRulesToLog = AddFailedRulesToLogCheckBox.IsChecked == true;

            // Read existing skips and hard coded rules from log file BEFORE conversion starts
            var logPath = DeriveLogFilePathFromXccdf(xccdfPath);
            System.Collections.Generic.HashSet<string> hardCodedRuleIds = new(StringComparer.OrdinalIgnoreCase);
            bool logFileExistedBefore = !string.IsNullOrWhiteSpace(logPath) && File.Exists(logPath);
            string logFileStatus = "not applicable";

            if (logFileExistedBefore)
            {
                var (skips, hardCoded) = ParseLogFile(logPath!);
                foreach (var skip in skips)
                    _skippedRuleIds.Add(skip);
                foreach (var hc in hardCoded)
                    hardCodedRuleIds.Add(hc);

                if (_skippedRuleIds.Count > 0 || hardCodedRuleIds.Count > 0)
                {
                    AppendInfo($"Loaded from log file: {_skippedRuleIds.Count} SKIP rule(s), {hardCodedRuleIds.Count} Hard Coded rule(s). Path: {logPath}",
                        System.Windows.Media.Brushes.DarkOrange, null);
                }
            }

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
            if (!logFileExistedBefore)
            {
                AppendInfo("No log file loaded, log file missing from the xccdf file location.", System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
            }
            SetBusy(true, "Converting…");
            AppendInfo("Starting conversion with Windows PowerShell 5.x...", System.Windows.Media.Brushes.DarkGreen, null);

            try
            {
                var moduleRoot = Path.GetDirectoryName(modulePath)!;
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

                var psArgs = $"-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File \"{tempScript}\" \"{modulePath}\" \"{xccdfPath}\" \"{destination}\" \"Windows\"";

                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = psArgs,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WorkingDirectory = moduleRoot
                };

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
                                var sepIdx = compact.IndexOf(" failed:", StringComparison.OrdinalIgnoreCase);
                                string error = sepIdx >= 0 ? compact.Substring(sepIdx + " failed:".Length).Trim() : compact;
                                _failedRuleErrors[rid!] = error;
                            }
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

                    var stdOut = outSb.ToString();
                    var stdErr = errSb.ToString();

                    try { File.Delete(tempScript); } catch { /* ignore */ }

                    if (proc.ExitCode != 0)
                        throw new InvalidOperationException($"Windows PowerShell exited with code {proc.ExitCode}. First error:\n{stdErr.Trim()}");
                });

                string? expectedFileName = GetConvertedFileName(xccdfPath!);
                string resolvedPath = ResolveConvertedFullPath(destination!, expectedFileName!);

                if (!string.IsNullOrWhiteSpace(expectedFileName))
                {
                    var hasEditionToken = Regex.IsMatch(expectedFileName, @"-(MS|DC)-", RegexOptions.IgnoreCase);
                    var expectedPath = Path.Combine(destination!, expectedFileName);
                    if (hasEditionToken && !string.Equals(resolvedPath, expectedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
                            File.Move(resolvedPath, expectedPath, overwrite: true);
                            AppendInfo($"Renamed converted output to preserve edition token: {expectedPath}");
                            resolvedPath = expectedPath;
                        }
                        catch (Exception renameEx)
                        {
                            AppendInfo($"Warning: could not rename converted file to include edition token: {renameEx.Message}", System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
                        }
                    }

                    if (createOrgSettings && hasEditionToken)
                    {
                        var expectedOrgFileName = Path.GetFileNameWithoutExtension(expectedFileName) + ".org.default.xml";
                        var expectedOrgPath = Path.Combine(destination!, expectedOrgFileName);

                        var resolvedOrgPath = ResolveConvertedFullPath(destination!, expectedOrgFileName);
                        if (!string.Equals(resolvedOrgPath, expectedOrgPath, StringComparison.OrdinalIgnoreCase))
                        {
                            if (File.Exists(resolvedOrgPath))
                            {
                                try
                                {
                                    File.Move(resolvedOrgPath, expectedOrgPath, overwrite: true);
                                    AppendInfo($"Renamed org settings to preserve edition token: {expectedOrgPath}");
                                }
                                catch (Exception orgRenameEx)
                                {
                                    AppendInfo($"Warning: could not rename org settings file to include edition token: {orgRenameEx.Message}", System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
                                }
                            }
                            else
                            {
                                AppendInfo($"Warning: org settings file not found to rename. Expected to resolve from: {resolvedOrgPath}", System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
                            }
                        }
                    }
                }

                var convertedFilePath = resolvedPath;

                // Collect all IDs from converted output
                var convertedIds = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var id in ExtractRuleIdsFromConverted(convertedFilePath))
                    convertedIds.Add(id);

                // Extract rules with dscresource="None" (won't be applied)
                var noDscResourceIds = ExtractRulesWithNoDscResource(convertedFilePath);

                // Normalize failed IDs
                var normalizedFailedIds = new System.Collections.Generic.HashSet<string>(
                    _failedRuleIds.Select(id => NormalizeVId(id.Trim())).Where(id => id.StartsWith("V-", StringComparison.OrdinalIgnoreCase)),
                    StringComparer.OrdinalIgnoreCase);

                // Normalize converted IDs to base form (V-123456.a → V-123456) and track variants
                var normalizedConvertedIds = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var convertedIdsByBase = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>>(StringComparer.OrdinalIgnoreCase);

                foreach (var id in convertedIds)
                {
                    var normalized = NormalizeVId(id.Trim());
                    if (string.IsNullOrWhiteSpace(normalized) || !normalized.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                        continue;

                    normalizedConvertedIds.Add(normalized);

                    // Track variants for display
                    if (!convertedIdsByBase.ContainsKey(normalized))
                        convertedIdsByBase[normalized] = new System.Collections.Generic.List<string>();
                    convertedIdsByBase[normalized].Add(id);
                }

                // Calculate true successes: exclude failed, skipped, hard coded, AND no DSC resource rules
                var successIds = normalizedConvertedIds
                    .Where(id => !normalizedFailedIds.Contains(id))
                    .Where(id => !_skippedRuleIds.Contains(id))
                    .Where(id => !hardCodedRuleIds.Contains(id))
                    .Where(id => !noDscResourceIds.Contains(id))
                    .OrderBy(id => ExtractNumericKey(id, prefix: "V-"))
                    .ThenBy(id => id, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // ===== DISPLAY RESULTS IN ORDER =====

                // 1. FAILED CONVERSIONS (new failures from THIS conversion) - SHOW FIRST
                bool logFileWasWritten = false;
                if (normalizedFailedIds.Count > 0)
                {
                    var sortedFailures = normalizedFailedIds
                        .OrderBy(id => ExtractNumericKey(id, prefix: "V-"))
                        .ThenBy(id => id, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    AppendInfo($"Failed rule conversions ({sortedFailures.Count}):", System.Windows.Media.Brushes.OrangeRed, null);
                    foreach (var vId in sortedFailures)
                    {
                        _failedRuleErrors.TryGetValue(vId, out var err);
                        var msg = string.IsNullOrWhiteSpace(err) ? $"{vId} failed." : $"{vId} failed: {err}";
                        AppendInfo(msg, System.Windows.Media.Brushes.OrangeRed, null);
                    }

                    // Write NEW failed rules to log file as SKIP entries
                    if (addFailedRulesToLog && !string.IsNullOrWhiteSpace(logPath))
                    {
                        try
                        {
                            AppendFailedRulesToLog(logPath, sortedFailures);
                            logFileWasWritten = true;
                            AppendInfo($"Added {sortedFailures.Count} failed rule(s) as SKIPs to log file: {logPath}",
                                System.Windows.Media.Brushes.DarkOrange, null);

                            // Show message box if log file was just created (didn't exist before)
                            if (!logFileExistedBefore)
                            {
                                System.Windows.MessageBox.Show(
                                    "Errors during conversion, a log file was created. Run convert again to get a clean conversion.",
                                    "Log File Created",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Information);
                            }
                        }
                        catch (Exception logEx)
                        {
                            AppendInfo($"Warning: Could not write to log file: {logEx.Message}",
                                System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
                        }
                    }
                }

                // 2. SKIPPED RULES from log file (::*::.)
                if (_skippedRuleIds.Count > 0)
                {
                    var sortedSkips = _skippedRuleIds
                        .OrderBy(id => ExtractNumericKey(id, prefix: "V-"))
                        .ThenBy(id => id, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    AppendInfo($"Skipped rules from log file ({sortedSkips.Count}):", System.Windows.Media.Brushes.DarkOrange, null);
                    foreach (var vId in sortedSkips)
                    {
                        AppendInfo($" - {vId} (skipped)", System.Windows.Media.Brushes.DarkOrange, null);
                    }
                }

                // 3. HARD CODED RULES from log file (all non-skip entries)
                if (hardCodedRuleIds.Count > 0)
                {
                    var sortedHardCoded = hardCodedRuleIds
                        .OrderBy(id => ExtractNumericKey(id, prefix: "V-"))
                        .ThenBy(id => id, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    AppendInfo($"Hard Coded rules from log file ({sortedHardCoded.Count}):", System.Windows.Media.Brushes.DarkCyan, null);
                    foreach (var vId in sortedHardCoded)
                    {
                        AppendInfo($" - {vId} (hard coded)", System.Windows.Media.Brushes.DarkCyan, null);
                    }
                }

                // 4. NO DSC RESOURCE rules (converted but won't be applied)
                if (noDscResourceIds.Count > 0)
                {
                    var sortedNoDsc = noDscResourceIds
                        .OrderBy(id => ExtractNumericKey(id, prefix: "V-"))
                        .ThenBy(id => id, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    AppendInfo($"Rules with no DSC resource - will not be applied ({sortedNoDsc.Count}):", System.Windows.Media.Brushes.DarkMagenta, null);
                    foreach (var vId in sortedNoDsc)
                    {
                        AppendInfo($" - {vId} (dscresource=\"None\")", System.Windows.Media.Brushes.DarkMagenta, null);
                    }
                }

                // Determine log file status
                if (!string.IsNullOrWhiteSpace(logPath))
                {
                    if (!logFileExistedBefore && logFileWasWritten)
                        logFileStatus = "created";
                    else if (logFileExistedBefore && logFileWasWritten)
                        logFileStatus = "updated";
                    else if (logFileExistedBefore && !logFileWasWritten)
                        logFileStatus = "unchanged";
                }

                // 5. SUCCESSFUL CONVERSIONS with variant info
                if (successIds.Count > 0)
                {
                    AppendInfo($"Successfully converted rule IDs ({successIds.Count}):", System.Windows.Media.Brushes.DarkGreen, null);
                    foreach (var baseId in successIds)
                    {
                        var variants = convertedIdsByBase.ContainsKey(baseId) ? convertedIdsByBase[baseId] : new System.Collections.Generic.List<string> { baseId };
                        if (variants.Count > 1)
                        {
                            // Show base ID with variant count
                            var variantList = string.Join(", ", variants.OrderBy(v => v, StringComparer.OrdinalIgnoreCase));
                            AppendInfo($" - {baseId} ({variants.Count} variants: {variantList})", System.Windows.Media.Brushes.DarkGreen, null);
                        }
                        else
                        {
                            AppendInfo($" - {baseId}", System.Windows.Media.Brushes.DarkGreen, null);
                        }
                    }
                }
                else
                {
                    AppendInfo("No successfully converted rule IDs.", System.Windows.Media.Brushes.DarkGreen, null);
                }

                // 6. SUMMARY SECTION
                var totalVariants = convertedIds.Count;

                // Manual intervention needed: skipped + no DSC resource + hard coded (NOT failed - those were never created)
                var manualHandlingRequired = _skippedRuleIds.Count + noDscResourceIds.Count + hardCodedRuleIds.Count;

                // Rules that will be automatically handled (total created minus manual intervention)
                var rulesAutoHandled = totalVariants - manualHandlingRequired;

                // Total rules created (everything except failed, since failed rules were never created in the XML)
                var totalRulesCreated = totalVariants;

                AppendInfo($"Total: {successIds.Count} successful rule conversions created {totalVariants} total rules including rule variants (.a, .b, .c, etc.), {normalizedFailedIds.Count} new failed, {_skippedRuleIds.Count} skipped, {hardCodedRuleIds.Count} hard coded, {noDscResourceIds.Count} rules set to DSCResource=\"None\", which means they won't be applied to the endpoint.",
                    System.Windows.Media.Brushes.DarkBlue, null);

                AppendInfo($"Summary:", System.Windows.Media.Brushes.DarkBlue, null);
                AppendInfo($"\t{totalRulesCreated} total rules created", System.Windows.Media.Brushes.DarkGreen, null);
                AppendInfo($"\t{rulesAutoHandled} rules automatically handled", System.Windows.Media.Brushes.DarkGreen, null);
                AppendInfo($"\t{manualHandlingRequired} manual handling required", System.Windows.Media.Brushes.OrangeRed, null);

                // Show failed rules separately if any exist
                if (normalizedFailedIds.Count > 0)
                {
                    var failedText = normalizedFailedIds.Count == 1
                        ? "1 rule failed"
                        : $"{normalizedFailedIds.Count} rules failed";
                    AppendInfo($"\t{failedText}", System.Windows.Media.Brushes.OrangeRed, null);
                }

                //AppendInfo($"\t{totalManual} rules that need to be handled manually", System.Windows.Media.Brushes.OrangeRed, null);

                if (!string.IsNullOrWhiteSpace(logPath))
                {
                    // Choose color based on log file status
                    var logColor = logFileStatus switch
                    {
                        "created" => System.Windows.Media.Brushes.LimeGreen,
                        "updated" => System.Windows.Media.Brushes.DodgerBlue,
                        "unchanged" => System.Windows.Media.Brushes.Goldenrod,
                        _ => System.Windows.Media.Brushes.Gray
                    };

                    AppendInfo($"\tLog file status: {logFileStatus} ({logPath})", logColor, null);
                }

                // 7. COMPARE STEP (optional)
                var skipCompare = SkipCompareCheckBox?.IsChecked == true;
                if (!skipCompare)
                {
                    AppendInfo($"Comparing newly converted output against the original XCCDF for missing rules during conversion (this may take a moment)…{Environment.NewLine}Converted file: {convertedFilePath}",
                        System.Windows.Media.Brushes.DarkSlateBlue, null);

                    var missing = CompareRuleIds(xccdfPath!, convertedFilePath);
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

                // ===== 8. GENERATE HTML REPORT =====
                try
                {
                    // For successful rules, we want to list each VARIANT separately, not base IDs
                    var successVariants = successIds
                        .SelectMany(baseId => convertedIdsByBase.ContainsKey(baseId)
                            ? convertedIdsByBase[baseId]
                            : new System.Collections.Generic.List<string> { baseId })
                        .OrderBy(id => ExtractNumericKey(NormalizeVId(id), "V-"))
                        .ThenBy(id => id, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    var reportData = new ConversionReportData
                    {
                        StigName = Path.GetFileNameWithoutExtension(xccdfPath!),
                        Timestamp = DateTime.Now,
                        TotalRulesCreated = totalRulesCreated,
                        RulesAutoHandled = rulesAutoHandled,
                        ManualHandlingRequired = manualHandlingRequired,
                        FailedCount = normalizedFailedIds.Count,
                        IndividualDISARulesAutomated = successIds.Count, // Number of unique base IDs
                        LogFileStatus = logFileStatus,
                        FailedRules = normalizedFailedIds.OrderBy(id => ExtractNumericKey(id, "V-")).ThenBy(id => id).ToList(),
                        SkippedRules = _skippedRuleIds.OrderBy(id => ExtractNumericKey(id, "V-")).ThenBy(id => id).ToList(),
                        HardCodedRules = hardCodedRuleIds.OrderBy(id => ExtractNumericKey(id, "V-")).ThenBy(id => id).ToList(),
                        NoDscResourceRules = noDscResourceIds.OrderBy(id => ExtractNumericKey(id, "V-")).ThenBy(id => id).ToList(),
                        SuccessfulRules = successVariants
                    };

                    // Populate failed rule details with error messages and XCCDF info
                    foreach (var ruleId in normalizedFailedIds)
                    {
                        var ruleInfo = RuleInfoExtractor.TryExtractRuleInfo(ruleId, xccdfPath!, null);

                        reportData.FailedRuleDetails[ruleId] = new RuleDetail
                        {
                            ErrorMessage = _failedRuleErrors.GetValueOrDefault(ruleId, "Unknown error"),
                            VariantCount = 1,
                            Variants = new System.Collections.Generic.List<string> { ruleId },
                            SvId = ruleInfo.SvId,
                            Title = ruleInfo.Title,
                            Severity = ruleInfo.Severity,
                            Description = ruleInfo.Description,
                            FixText = ruleInfo.FixText,
                            CheckText = null
                        };
                    }

                    // Populate successful rule details - ONE ENTRY PER VARIANT with its specific converted snippet
                    foreach (var variantId in successVariants)
                    {
                        var ruleInfo = RuleInfoExtractor.TryExtractRuleInfo(variantId, xccdfPath!, destination);

                        reportData.SuccessfulRuleDetails[variantId] = new RuleDetail
                        {
                            VariantCount = 1, // Each variant is its own entry now
                            Variants = new System.Collections.Generic.List<string> { variantId },
                            SvId = ruleInfo.SvId,
                            Title = ruleInfo.Title,
                            Severity = ruleInfo.Severity,
                            Description = ruleInfo.Description,
                            FixText = ruleInfo.FixText,
                            ConvertedSnippet = ruleInfo.ConvertedSnippet, // This will be SPECIFIC to this variant
                            DscResource = null
                        };
                    }

                    // Populate skipped rule details
                    foreach (var ruleId in _skippedRuleIds)
                    {
                        var ruleInfo = RuleInfoExtractor.TryExtractRuleInfo(ruleId, xccdfPath!, null);
                        reportData.SkippedRuleDetails[ruleId] = new RuleDetail
                        {
                            SvId = ruleInfo.SvId,
                            Title = ruleInfo.Title,
                            Severity = ruleInfo.Severity,
                            Description = ruleInfo.Description,
                            FixText = ruleInfo.FixText
                        };
                    }

                    // Populate hard coded rule details
                    foreach (var ruleId in hardCodedRuleIds)
                    {
                        var ruleInfo = RuleInfoExtractor.TryExtractRuleInfo(ruleId, xccdfPath!, destination);
                        reportData.HardCodedRuleDetails[ruleId] = new RuleDetail
                        {
                            SvId = ruleInfo.SvId,
                            Title = ruleInfo.Title,
                            Severity = ruleInfo.Severity,
                            Description = ruleInfo.Description,
                            FixText = ruleInfo.FixText,
                            ConvertedSnippet = ruleInfo.ConvertedSnippet
                        };
                    }

                    // Populate no DSC resource rule details
                    foreach (var ruleId in noDscResourceIds)
                    {
                        var ruleInfo = RuleInfoExtractor.TryExtractRuleInfo(ruleId, xccdfPath!, destination);
                        reportData.NoDscRuleDetails[ruleId] = new RuleDetail
                        {
                            SvId = ruleInfo.SvId,
                            Title = ruleInfo.Title,
                            Severity = ruleInfo.Severity,
                            Description = ruleInfo.Description,
                            FixText = ruleInfo.FixText,
                            ConvertedSnippet = ruleInfo.ConvertedSnippet
                        };
                    }

                    // Generate and save report - same location and name as converted file, but .html extension
                    var reportHtml = ConversionReportGenerator.GenerateHtmlReport(reportData);
                    var reportPath = Path.ChangeExtension(convertedFilePath, ".html");
                    ConversionReportGenerator.SaveReport(reportHtml, reportPath);

                    AppendInfo($"Conversion report saved: {reportPath}", System.Windows.Media.Brushes.DarkGreen, System.Windows.Media.Brushes.LightGreen);
                }
                catch (Exception reportEx)
                {
                    AppendInfo($"Warning: Failed to generate report: {reportEx.Message}", System.Windows.Media.Brushes.Black, System.Windows.Media.Brushes.MistyRose);
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

        // In GetConvertedFileName, replace how endIdx is computed to avoid matching "_Visio" as the version marker.
        private static string? GetConvertedFileName(string originalFileName)
        {
            if (string.IsNullOrWhiteSpace(originalFileName)) return null;

            var name = Path.GetFileName(originalFileName);

            // Require "VxRx", "-xccdf" and ".xml"
            var m = Regex.Match(name, @".*V\d+R\d+.*-xccdf\.xml$", RegexOptions.IgnoreCase);
            if (!m.Success) return null;

            // Version short x.y from VxRy
            var ver = Regex.Match(name, @"V(?<maj>\d+)R(?<min>\d+)", RegexOptions.IgnoreCase);
            if (!ver.Success) return null;
            var verShort = $"{ver.Groups["maj"].Value}.{ver.Groups["min"].Value}";
            var verToken = ver.Value; // e.g., "V1R4"

            // Extract product segment between "U_" and one of markers
            string productRaw = string.Empty;
            int uIdx = name.IndexOf("U_", StringComparison.OrdinalIgnoreCase);
            if (uIdx >= 0)
            {
                uIdx += 2;
                int stigIdx = name.IndexOf("_STIG", uIdx, StringComparison.OrdinalIgnoreCase);

                // FIX: only treat the version marker when it is the real "_V<digits>R<digits>", not any "_V..." (e.g., "_Visio")
                int vIdx = name.IndexOf("_" + verToken, uIdx, StringComparison.OrdinalIgnoreCase);

                int manualIdx = name.IndexOf("_Manual", uIdx, StringComparison.OrdinalIgnoreCase);
                int endIdx = new[] { stigIdx, vIdx, manualIdx }.Where(i => i >= 0).DefaultIfEmpty(-1).Min();
                if (endIdx > uIdx) productRaw = name.Substring(uIdx, endIdx - uIdx);
            }

            var tokens = productRaw.Split('_', StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim()).ToList();
            if (tokens.Count == 0) return null;

            // Helpers
            static string Pascal(string s) => string.IsNullOrEmpty(s) ? s : char.ToUpperInvariant(s[0]) + s.Substring(1);
            static string JoinPascal(IEnumerable<string> parts) => string.Concat(parts.Select(Pascal));

            // Special case: combined Windows 2012 and 2012 R2 STIG (no explicit "Server" token in source)
            // Example: U_MS_Windows_2012_and_2012_R2_MS_STIG_V3R4_Manual-xccdf.xml
            if (name.IndexOf("_Windows_2012_and_2012_R2_", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Prefer explicit edition from filename pattern
                string? edition = null;
                if (name.IndexOf("_MS_STIG_", StringComparison.OrdinalIgnoreCase) >= 0) edition = "MS";
                else if (name.IndexOf("_DC_STIG_", StringComparison.OrdinalIgnoreCase) >= 0) edition = "DC";
                else
                {
                    // Fallback to tokens if present (rare)
                    if (tokens.Any(t => t.Equals("MS", StringComparison.OrdinalIgnoreCase))) edition = "MS";
                    else if (tokens.Any(t => t.Equals("DC", StringComparison.OrdinalIgnoreCase))) edition = "DC";
                }

                var baseName = "WindowsServer-2012R2";
                return edition is null
                    ? $"{baseName}-{verShort}.xml"
                    : $"{baseName}-{edition}-{verShort}.xml";
            }


            // 1) Adobe Acrobat Pro => Adobe-AcrobatPro-x.y.xml
            if (tokens.Count >= 3 &&
                tokens[0].Equals("Adobe", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("Acrobat", StringComparison.OrdinalIgnoreCase) &&
                tokens[2].Equals("Pro", StringComparison.OrdinalIgnoreCase))
            {
                return $"Adobe-AcrobatPro-{verShort}.xml";
            }

            // 2) Adobe Acrobat Reader => Adobe-AcrobatReader-x.y.xml
            if (tokens.Count >= 3 &&
                tokens[0].Equals("Adobe", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("Acrobat", StringComparison.OrdinalIgnoreCase) &&
                tokens[2].Equals("Reader", StringComparison.OrdinalIgnoreCase))
            {
                return $"Adobe-AcrobatReader-{verShort}.xml";
            }

            // 3) Firefox => FireFox-All-x.y.xml
            if (tokens.Count >= 1 &&
                tokens[0].Equals("MOZ", StringComparison.OrdinalIgnoreCase) &&
                tokens.Any(t => t.Equals("Firefox", StringComparison.OrdinalIgnoreCase)))
            {
                return $"FireFox-All-{verShort}.xml";
            }

            // 4) Chrome => Google-Chrome-x.y.xml
            if (tokens.Count >= 1 &&
                tokens[0].Equals("Google", StringComparison.OrdinalIgnoreCase) &&
                tokens.Any(t => t.Equals("Chrome", StringComparison.OrdinalIgnoreCase)))
            {
                return $"Google-Chrome-{verShort}.xml";
            }

            // 5) MS Edge => MS-Edge-x.y.xml
            if (tokens.Count >= 1 && tokens[0].Equals("MS", StringComparison.OrdinalIgnoreCase) &&
                tokens.Any(t => t.Equals("Edge", StringComparison.OrdinalIgnoreCase)))
            {
                return $"MS-Edge-{verShort}.xml";
            }

            // 6) Internet Explorer 11 => InternetExplorer-11-x.y.xml
            if (tokens.Count >= 1 && tokens[0].Equals("MS", StringComparison.OrdinalIgnoreCase) &&
                tokens.Any(t => t.Equals("IE11", StringComparison.OrdinalIgnoreCase)))
            {
                return $"InternetExplorer-11-{verShort}.xml";
            }

            // 7) DotNet Framework 4-0 => DotNetFramework-4-2.x.xml
            if (tokens.Count >= 3 && tokens[0].Equals("MS", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("DotNet", StringComparison.OrdinalIgnoreCase) &&
                tokens[2].Equals("Framework", StringComparison.OrdinalIgnoreCase))
            {
                var num = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^\d+(?:-\d+)?$")) ?? "4-0";
                var normalizedMajor = num.Replace("-0", "");
                return $"DotNetFramework-{normalizedMajor}-{verShort}.xml";
            }

            // 8) IIS Server/Site => IISServer-10.0-x.y.xml / IISSite-10.0-x.y.xml
            if (tokens.Count >= 2 && tokens[0].Equals("MS", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("IIS", StringComparison.OrdinalIgnoreCase))
            {
                var isServer = tokens.Any(t => t.Equals("Server", StringComparison.OrdinalIgnoreCase));
                var isSite = tokens.Any(t => t.Equals("Site", StringComparison.OrdinalIgnoreCase));
                var versionToken = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^\d+(?:-\d+)?$"));
                var versionNum = versionToken is null ? "" :
                    Regex.Replace(versionToken, @"^(\d+)-(\d+)$", "$1.$2");
                var baseName = isServer ? "IISServer" : isSite ? "IISSite" : "IIS";
                return string.IsNullOrEmpty(versionNum)
                    ? $"{baseName}-{verShort}.xml"
                    : $"{baseName}-{versionNum}-{verShort}.xml";
            }

            // 9) Oracle JRE 8 Windows => OracleJRE-8-x.y.xml
            if (tokens.Count >= 2 && tokens[0].Equals("Oracle", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("JRE", StringComparison.OrdinalIgnoreCase))
            {
                var v = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^\d+$")) ?? "8";
                return $"OracleJRE-{v}-{verShort}.xml";
            }

            // 10) SQL Server
            if (tokens.Count >= 3 && tokens[0].Equals("MS", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("SQL", StringComparison.OrdinalIgnoreCase) &&
                tokens[2].Equals("Server", StringComparison.OrdinalIgnoreCase))
            {
                var year = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^\d{4}$")) ?? "2012";
                var flavor = tokens.Any(t => t.Equals("Database", StringComparison.OrdinalIgnoreCase)) ? "Database" : "Instance";
                return $"SqlServer-{year}-{flavor}-{verShort}.xml";
            }

            // 11) vSphere ESXi => Vsphere-6.5-x.y.xml
            if (tokens.Count >= 2 && tokens[0].Equals("VMW", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("vSphere", StringComparison.OrdinalIgnoreCase))
            {
                var versionToken = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^\d+(?:-\d+)?$")) ?? "6-5";
                var versionNum = Regex.Replace(versionToken, @"^(\d+)-(\d+)$", "$1.$2");
                return $"Vsphere-{versionNum}-{verShort}.xml";
            }

            // 12) Oracle Linux => OracleLinux-<major>-x.y.xml
            if (tokens.Count >= 1 && tokens[0].Equals("OracleLinux", StringComparison.OrdinalIgnoreCase) ||
                (tokens.Count >= 2 && tokens[0].Equals("Oracle", StringComparison.OrdinalIgnoreCase) && tokens[1].Equals("Linux", StringComparison.OrdinalIgnoreCase)))
            {
                var major = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^\d+$")) ?? "8";
                return $"OracleLinux-{major}-{verShort}.xml";
            }

            // 13) RHEL => RHEL-<major>-x.y.xml
            if (tokens.Count >= 1 && tokens[0].Equals("RHEL", StringComparison.OrdinalIgnoreCase))
            {
                var major = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^\d+$")) ?? "7";
                return $"RHEL-{major}-{verShort}.xml";
            }

            // 14) Ubuntu => Ubuntu-18.04-x.y.xml
            if (tokens.Count >= 2 && tokens[0].Equals("CAN", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("Ubuntu", StringComparison.OrdinalIgnoreCase))
            {
                var versionToken = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^\d{2}-\d{2}$")) ?? "18-04";
                var versionNum = versionToken.Replace('-', '.');
                return $"Ubuntu-{versionNum}-{verShort}.xml";
            }

            // 15) Windows Client (10/11) => WindowsClient-<major>-x.y.xml
            if (tokens.Count >= 2 && tokens[0].Equals("MS", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("Windows", StringComparison.OrdinalIgnoreCase) &&
                (tokens.Any(t => t.Equals("10", StringComparison.OrdinalIgnoreCase)) || tokens.Any(t => t.Equals("11", StringComparison.OrdinalIgnoreCase))) &&
                !tokens.Any(t => t.Equals("Server", StringComparison.OrdinalIgnoreCase)))
            {
                var major = tokens.FirstOrDefault(t => t.Equals("10", StringComparison.OrdinalIgnoreCase) || t.Equals("11", StringComparison.OrdinalIgnoreCase)) ?? "10";
                return $"WindowsClient-{major}-{verShort}.xml";
            }

            // 16) Windows Defender Antivirus => WindowsDefender-All-x.y.xml
            if (tokens.Count >= 3 && tokens[0].Equals("MS", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("Windows", StringComparison.OrdinalIgnoreCase) &&
                tokens[2].Equals("Defender", StringComparison.OrdinalIgnoreCase))
            {
                return $"WindowsDefender-All-{verShort}.xml";
            }

            // 17) Windows Firewall => WindowsFirewall-All-x.y.xml
            if (tokens.Count >= 3 && tokens[0].Equals("MS", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("Windows", StringComparison.OrdinalIgnoreCase) &&
                tokens[2].Equals("Firewall", StringComparison.OrdinalIgnoreCase))
            {
                return $"WindowsFirewall-All-{verShort}.xml";
            }

            // 18) Windows DNS Server 2012/2012R2 => WindowsDnsServer-2012R2-x.y.xml
            if (tokens.Count >= 2
                && tokens[0].Equals("MS", StringComparison.OrdinalIgnoreCase)
                && tokens[1].Equals("Windows", StringComparison.OrdinalIgnoreCase)
                && tokens.Any(t => t.Equals("DNS", StringComparison.OrdinalIgnoreCase)))
            {
                // Prefer 2012R2 to match processed outputs; normalize plain 2012 to 2012R2
                var yearTok = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^(2012|2012R2)$", RegexOptions.IgnoreCase));
                var year = (yearTok != null && yearTok.Equals("2012", StringComparison.OrdinalIgnoreCase)) ? "2012R2" : (yearTok ?? "2012R2");
                return $"WindowsDnsServer-{year}-{verShort}.xml";
            }

            // 19) Windows Server editions and no-edition server
            if (tokens.Count >= 3 && tokens[0].Equals("MS", StringComparison.OrdinalIgnoreCase) &&
                tokens[1].Equals("Windows", StringComparison.OrdinalIgnoreCase) &&
                tokens.Any(t => t.Equals("Server", StringComparison.OrdinalIgnoreCase)))
            {
                var year = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^(2012R2|2016|2019|2022)$", RegexOptions.IgnoreCase)) ?? "2016";
                string? edition = null;
                if (name.IndexOf("_MS_STIG_", StringComparison.OrdinalIgnoreCase) >= 0) edition = "MS";
                else if (name.IndexOf("_DC_STIG_", StringComparison.OrdinalIgnoreCase) >= 0) edition = "DC";
                var baseName = $"WindowsServer-{year}";
                return (edition is null)
                    ? $"{baseName}-{verShort}.xml"
                    : $"{baseName}-{edition}-{verShort}.xml";
            }

            // 20) Office (System/365/Apps) family
            if (tokens[0].Equals("Microsoft", StringComparison.OrdinalIgnoreCase) || tokens[0].Equals("MS", StringComparison.OrdinalIgnoreCase))
            {
                if (tokens.Any(t => t.Equals("Office", StringComparison.OrdinalIgnoreCase)) && tokens.Any(t => t.Equals("365", StringComparison.OrdinalIgnoreCase)))
                    return $"Office-365ProPlus-{verShort}.xml";

                var systemYear = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^(2013|2016)$"));
                if (tokens.Any(t => t.Equals("Office", StringComparison.OrdinalIgnoreCase)) && tokens.Any(t => t.Equals("System", StringComparison.OrdinalIgnoreCase)) && systemYear != null)
                    return $"Office-System{systemYear}-{verShort}.xml";

                string? app = null;
                string? appYear = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^(2013|2016)$"));
                if (tokens.Any(t => t.Equals("Access", StringComparison.OrdinalIgnoreCase))) app = "Access";
                else if (tokens.Any(t => t.Equals("Excel", StringComparison.OrdinalIgnoreCase))) app = "Excel";
                else if (tokens.Any(t => t.Equals("OneNote", StringComparison.OrdinalIgnoreCase))) app = "OneNote";
                else if (tokens.Any(t => t.Equals("Outlook", StringComparison.OrdinalIgnoreCase))) app = "Outlook";
                else if (tokens.Any(t => t.Equals("PowerPoint", StringComparison.OrdinalIgnoreCase))) app = "PowerPoint";
                else if (tokens.Any(t => t.Equals("Publisher", StringComparison.OrdinalIgnoreCase))) app = "Publisher";
                else if (tokens.Any(t => t.Equals("Skype", StringComparison.OrdinalIgnoreCase))) app = "Skype";
                else if (tokens.Any(t => t.Equals("Visio", StringComparison.OrdinalIgnoreCase))) app = "Visio";
                else if (tokens.Any(t => t.Equals("Word", StringComparison.OrdinalIgnoreCase))) app = "Word";

                if (app != null && appYear != null)
                    return $"Office-{app}{appYear}-{verShort}.xml";

                if (tokens.Any(t => t.Equals("OfficeSystem", StringComparison.OrdinalIgnoreCase)) && appYear != null)
                    return $"Office-System{appYear}-{verShort}.xml";
            }

            // 21) McAfee VirusScan 8.8 => McAfee-8.8-VirusScan-x.y.xml
            if (tokens.Count >= 2 && tokens[0].Equals("McAfee", StringComparison.OrdinalIgnoreCase))
            {
                var versionToken = tokens.FirstOrDefault(t => Regex.IsMatch(t, @"^\d+(?:\.\d+)?$")) ?? "8.8";
                return $"McAfee-{versionToken}-VirusScan-{verShort}.xml";
            }

            // Fallback
            var cleaned = tokens.Where(t =>
                !t.Equals("Site", StringComparison.OrdinalIgnoreCase) &&
                !t.Equals("Server", StringComparison.OrdinalIgnoreCase) &&
                !t.Equals("Windows", StringComparison.OrdinalIgnoreCase) &&
                !t.Equals("STIG", StringComparison.OrdinalIgnoreCase) &&
                !t.Equals("Manual", StringComparison.OrdinalIgnoreCase)).ToList();

            var numericToken = cleaned.FirstOrDefault(t => Regex.IsMatch(t, @"^\d+(?:-\d+)?$"));
            if (numericToken != null) cleaned.Remove(numericToken);
            var basePascal = JoinPascal(cleaned);
            if (string.IsNullOrWhiteSpace(basePascal)) basePascal = JoinPascal(tokens);

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

            var parts = new System.Collections.Generic.List<string> { basePascal };
            if (!string.IsNullOrWhiteSpace(normalizedNumber)) parts.Add(normalizedNumber);
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
            var result = RuleIdAnalysis.Compare(xccdfPath, convertedPath);
            return new System.Collections.Generic.HashSet<string>(result.MissingBaseIds, System.StringComparer.OrdinalIgnoreCase);
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

        private void InfoRichTextBox_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var lineText = GetLineUnderMouse(InfoRichTextBox, e) ?? GetWordUnderMouse(InfoRichTextBox, e);
            if (string.IsNullOrWhiteSpace(lineText)) return;

            var m = Regex.Match(lineText, @"\b(SV-\d+(?:[^\s]*)?|V-\d+)\b", RegexOptions.IgnoreCase);
            if (!m.Success) return;
            var rawToken = m.Value;

            var xccdfPath = XccdfPathTextBox.Text?.Trim();
            var destination = DestinationTextBox.Text?.Trim();

            if (string.IsNullOrWhiteSpace(xccdfPath) || !File.Exists(xccdfPath))
            {
                System.Windows.MessageBox.Show("XCCDF path is not set or invalid.", "Rule Info", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string resolvedForQuery = rawToken;
            if (rawToken.StartsWith("SV-", StringComparison.OrdinalIgnoreCase))
            {
                var vId = RuleInfoExtractor.ResolveVIdForSv(rawToken, xccdfPath);
                if (!string.IsNullOrWhiteSpace(vId))
                    resolvedForQuery = vId;
                else
                {
                    var d = Regex.Match(rawToken, @"SV-(\d+)", RegexOptions.IgnoreCase);
                    if (d.Success) resolvedForQuery = $"V-{d.Groups[1].Value}";
                }
            }

            // Pass destination folder so converted snippet can be loaded for successful conversions
            var convertedFolder = !string.IsNullOrWhiteSpace(destination) && Directory.Exists(destination) ? destination : null;
            var info = RuleInfoExtractor.TryExtractRuleInfo(resolvedForQuery, xccdfPath!, convertedFolder);

            bool hasAny =
                !string.IsNullOrWhiteSpace(info.Title) ||
                !string.IsNullOrWhiteSpace(info.Severity) ||
                !string.IsNullOrWhiteSpace(info.Description) ||
                !string.IsNullOrWhiteSpace(info.FixText) ||
                !string.IsNullOrWhiteSpace(info.ReferencesXml);

            if (!hasAny)
            {
                info.RuleId = resolvedForQuery;
                info.Description = "No data found in XCCDF for the selected rule. This can happen if:\n" +
                                   "- The clicked line contains a non-rule identifier.\n" +
                                   "- The XCCDF uses only SV identifiers and the owning V-id could not be resolved.\n" +
                                   "- Namespaces or schema variations differ in this benchmark.";
            }

            if (rawToken.StartsWith("SV-", StringComparison.OrdinalIgnoreCase))
                info.SvId ??= Regex.Match(rawToken, @"SV-\d+", RegexOptions.IgnoreCase).Value;

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
        public static string? ResolveVIdForSv(string svId, string xccdfPath)
        {
            if (string.IsNullOrWhiteSpace(svId) || !svId.StartsWith("SV-", StringComparison.OrdinalIgnoreCase))
                return null;

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

        // Strengthen XCCDF lookup: load the whole document and use robust XPath fallbacks.
        // This fixes cases like V-218819 in IIS Server XCCDFs where Rule/@id is SV-... and ident carries SV digits.
        private static void TryFillFromXccdf(string ruleId, string xccdfPath, RuleInfo info)
        {
            try
            {
                // Build base tokens
                var vMatch = Regex.Match(ruleId, @"^V-(\d+)", RegexOptions.IgnoreCase);
                var svMatch = Regex.Match(ruleId, @"^SV-(\d+)", RegexOptions.IgnoreCase);
                var baseDigits = vMatch.Success ? vMatch.Groups[1].Value : svMatch.Success ? svMatch.Groups[1].Value : string.Empty;
                var baseV = string.IsNullOrEmpty(baseDigits) ? string.Empty : $"V-{baseDigits}";
                var baseSv = string.IsNullOrEmpty(baseDigits) ? string.Empty : $"SV-{baseDigits}";

                // Load whole XCCDF to avoid streaming edge cases
                var doc = new System.Xml.XmlDocument { PreserveWhitespace = true };
                var settings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore };
                using (var reader = System.Xml.XmlReader.Create(xccdfPath, settings))
                    doc.Load(reader);

                System.Xml.XmlNode? ruleNode = null;

                // 1) Direct V-id attribute (rare in DISA, but quick win)
                if (!string.IsNullOrWhiteSpace(baseV))
                    ruleNode = doc.SelectSingleNode($"//*[local-name()='Rule' and @id='{baseV}']");

                // 2) DISA style: Rule/@id is SV-... and the ident contains our base SV digits
                if (ruleNode is null && !string.IsNullOrWhiteSpace(baseSv))
                {
                    ruleNode = doc.SelectSingleNode(
                        $"//*[local-name()='Rule'][.//*[local-name()='ident' and (starts-with(normalize-space(text()),'{baseSv}') or contains(@system,'{baseSv}'))]]");
                }

                // 3) Fallback: match Rules where any ident contains base V or base SV anywhere (not just starts-with)
                if (ruleNode is null && !string.IsNullOrWhiteSpace(baseDigits))
                {
                    ruleNode = doc.SelectSingleNode(
                        $"//*[local-name()='Rule'][.//*[local-name()='ident' and (contains(normalize-space(text()),'SV-{baseDigits}') or contains(normalize-space(text()),'V-{baseDigits}') or contains(@system,'SV-{baseDigits}'))]]");
                }

                // 4) Last resort: any Rule whose @id contains our digits (handles SV-xxxxx_rNNN_rule patterns)
                if (ruleNode is null && !string.IsNullOrWhiteSpace(baseDigits))
                {
                    ruleNode = doc.SelectSingleNode($"//*[local-name()='Rule' and contains(@id,'{baseDigits}')]");
                }

                if (ruleNode is null)
                    return;

                // Populate info from the matched Rule
                var vIdAttr = ruleNode.Attributes?["id"]?.Value?.Trim();
                if (!string.IsNullOrWhiteSpace(vIdAttr) && vIdAttr.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                    info.RuleId = vIdAttr;
                else if (!string.IsNullOrWhiteSpace(baseV))
                    info.RuleId = baseV;

                var severityAttr = ruleNode.Attributes?["severity"]?.Value;
                if (!string.IsNullOrWhiteSpace(severityAttr)) info.Severity = severityAttr;

                var titleNode = ruleNode.SelectSingleNode(".//*[local-name()='title']");
                if (titleNode is not null) info.Title = titleNode.InnerText.Trim();

                var descNodes = ruleNode.SelectNodes(".//*[local-name()='description']");
                if (descNodes is not null && descNodes.Count > 0)
                {
                    var sb = new StringBuilder();
                    foreach (System.Xml.XmlNode dn in descNodes) sb.AppendLine(dn.InnerText.Trim());
                    info.Description = sb.ToString().Trim();
                }

                var fixNode = ruleNode.SelectSingleNode(".//*[local-name()='fixtext']");
                if (fixNode is not null) info.FixText = fixNode.InnerText.Trim();

                // Capture SV id for display
                var identNode = ruleNode.SelectSingleNode(".//*[local-name()='ident' and (starts-with(normalize-space(text()),'SV-') or contains(@system,'SV-'))]");
                if (identNode is not null)
                {
                    var identText = identNode.InnerText?.Trim();
                    var systemAttr = identNode.Attributes?["system"]?.Value;
                    if (!string.IsNullOrWhiteSpace(identText))
                    {
                        var m2 = Regex.Match(identText, @"SV-\d+", RegexOptions.IgnoreCase);
                        if (m2.Success) info.SvId = m2.Value;
                    }
                    if (string.IsNullOrWhiteSpace(info.SvId) && !string.IsNullOrWhiteSpace(systemAttr))
                    {
                        var m3 = Regex.Match(systemAttr, @"SV-\d+", RegexOptions.IgnoreCase);
                        if (m3.Success) info.SvId = m3.Value;
                    }
                }

                var refNodes = ruleNode.SelectNodes(".//*[local-name()='reference']");
                if (refNodes is not null && refNodes.Count > 0)
                {
                    var sb = new StringBuilder();
                    foreach (System.Xml.XmlNode rn in refNodes) sb.AppendLine(rn.OuterXml);
                    info.ReferencesXml = sb.ToString().Trim();
                }
            }
            catch
            {
                // ignore parse errors
            }
        }
        // 2) Fix CS8604: guard against null when passing XmlNode to GetNodeSummary.
        private static void TryFillFromConverted(string ruleId, string folder, RuleInfo info)
        {
            try
            {
                foreach (var file in Directory.EnumerateFiles(folder, "*.xml", SearchOption.TopDirectoryOnly))
                {
                    var doc = new System.Xml.XmlDocument();
                    doc.Load(file);

                    // For variants (V-123456.a), search for exact match on Rule/@id
                    var ruleNode = doc.SelectSingleNode($"//*[local-name()='Rule' and @id='{ruleId}']");

                    if (ruleNode is not null)
                    {
                        info.ConvertedFile = file;
                        // Return the FULL Rule XML (no truncation for HTML reports)
                        info.ConvertedSnippet = ruleNode.OuterXml;
                        return;
                    }

                    // Fallback: search by child elements (for base IDs without variant suffix)
                    var node = doc.SelectSingleNode(
                        $"//*[local-name()='RuleId' or local-name()='VulnId' or local-name()='BenchmarkId'][text()='{ruleId}']");
                    if (node is not null && node.ParentNode is not null)
                    {
                        info.ConvertedFile = file;
                        // Get the parent Rule element
                        var parentRule = node.ParentNode;
                        while (parentRule is not null && parentRule.LocalName != "Rule")
                            parentRule = parentRule.ParentNode;

                        info.ConvertedSnippet = parentRule?.OuterXml ?? node.OuterXml;
                        return;
                    }
                }
            }
            catch
            {
                // ignore
            }
        }

    }

}