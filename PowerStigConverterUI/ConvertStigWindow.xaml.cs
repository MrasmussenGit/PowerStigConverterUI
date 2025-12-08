using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Forms;

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

        public ConvertStigWindow()
        {
            InitializeComponent();
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
ConvertTo-PowerStigXml -Destination $Destination -Path $XccdfPath -CreateOrgSettingsFile
";
                File.WriteAllText(tempScript, scriptContent, new UTF8Encoding(false));

                // Build arguments for -File without brittle quoting
                var psArgs = $"-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File \"{tempScript}\" \"{modulePath}\" \"{xccdfPath}\" \"{destination}\"";

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
                var convertedIds = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                try
                {
                    foreach (var file in Directory.EnumerateFiles(destination, "*.xml", SearchOption.TopDirectoryOnly))
                    {
                        foreach (var id in ExtractRuleIdsFromConverted(file))
                            convertedIds.Add(id);
                    }
                }
                catch
                {
                    // Ignore folder read errors
                }

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

                // Skip compare when requested
                var skipCompare = SkipCompareCheckBox?.IsChecked == true;
                if (!skipCompare)
                {
                    // Successes
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
                        AppendInfo($"Successfully converted rule IDs ({successIds.Count}):", System.Windows.Media.Brushes.DarkGoldenrod, null);
                        foreach (var id in successIds.Take(100))
                        {
                            AppendInfo($" - {id}", System.Windows.Media.Brushes.DarkGoldenrod, null);
                        }
                        if (successIds.Count > 100)
                            AppendInfo($" ...and {successIds.Count - 100} more", System.Windows.Media.Brushes.DarkGoldenrod, null);
                    }
                    else
                    {
                        AppendInfo("No successfully converted rule IDs.", System.Windows.Media.Brushes.DarkGoldenrod, null);
                    }

                    // Missing (present in XCCDF but not in converted output)
                    var missing = CompareRuleIds(xccdfPath!, destination!);
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
                ConvertButton.IsEnabled = true;
            }
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

        // Helper: compare rule IDs between original XCCDF and converted output
        private static System.Collections.Generic.HashSet<string> CompareRuleIds(string xccdfPath, string destinationFolder)
        {
            var xccdfIds = ExtractRuleIdsFromXccdf(xccdfPath);
            var convertedIds = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                foreach (var file in Directory.EnumerateFiles(destinationFolder, "*.xml", SearchOption.TopDirectoryOnly))
                {
                    foreach (var id in ExtractRuleIdsFromConverted(file))
                        convertedIds.Add(id);
                }
            }
            catch
            {
                // Ignore folder read errors
            }

            var missing = new System.Collections.Generic.HashSet<string>(xccdfIds, StringComparer.OrdinalIgnoreCase);
            missing.ExceptWith(convertedIds);
            return missing;
        }

        private static System.Collections.Generic.HashSet<string> ExtractRuleIdsFromXccdf(string xccdfPath)
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
            return ids;
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
    }
}