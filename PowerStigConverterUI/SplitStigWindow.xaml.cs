using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using Brushes = System.Windows.Media.Brushes;
using Microsoft.Win32;

namespace PowerStigConverterUI
{
    public partial class SplitStigWindow : Window
    {
        private const string LastFunctionsPathRegKey = @"Software\PowerStigConverterUI";
        private const string LastFunctionsPathRegValue = "LastFunctionsPath";

        private static readonly Regex WindowsOsStigName = new(
            // Examples supported:
            // U_MS_Windows_Server_2022_STIG_V2R6_Manual-xccdf.xml
            // U_MS_Windows_11_STIG_V2R6_Manual-xccdf.xml
            // U_MS_Windows_10_STIG_V2R6_Manual-xccdf.xml
            @"^U_MS_(Windows(?:_Server)?(?:_\d{2,4}|_\d{1,2})?_STIG)_V\d+R\d+_Manual-xccdf$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private string? _tempExtractPath = null;
        private string? _actualXccdfPath = null; // Track actual XCCDF path (from ZIP or direct)
        private string? _lastDestinationDirectory = null; // Track last directory used by destination folder browse button

        public SplitStigWindow()
        {
            InitializeComponent();
            this.Loaded += OnLoaded;
            this.Closing += SplitStigWindow_Closing;

            // Load persisted directory settings
            _lastDestinationDirectory = AppSettings.Instance.LastSplitDestinationDirectory;
        }

        private void OnLoaded(object? sender, RoutedEventArgs e)
        {
            // If already set (e.g., from previous run), keep it.
            if (!string.IsNullOrWhiteSpace(ModulePathTextBox.Text) && File.Exists(ModulePathTextBox.Text))
                return;

            // Try registry (last successful path)
            var last = ReadLastFunctionsPath();
            if (!string.IsNullOrWhiteSpace(last) && File.Exists(last))
            {
                ModulePathTextBox.Text = last;
                return;
            }

            // Discover typical locations
            var discovered = FindPowerStigFunctionsFile();
            if (!string.IsNullOrWhiteSpace(discovered))
            {
                ModulePathTextBox.Text = discovered!;
                WriteLastFunctionsPath(discovered!);
                return;
            }

            // Module not found - expand the Advanced section and show warning
            ModuleExpander.IsExpanded = true;
            ModuleWarningText.Visibility = Visibility.Visible;
            ModulePathTextBox.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 235, 235)); // Light red background

            // Alert user if not found
            ModulePathTextBox.Text = string.Empty;
            System.Windows.MessageBox.Show(
                "Functions.XccdfXml.ps1 file was not found in standard PowerSTIG module locations.\\n\\n" +
                "Please install PowerSTIG (Windows PowerShell 5.x) under:\\n" +
                "  - %ProgramFiles%\\\\WindowsPowerShell\\\\Modules\\\\PowerSTIG\\n" +
                "  - %UserProfile%\\\\Documents\\\\WindowsPowerShell\\\\Modules\\\\PowerSTIG\\n\\n" +
                "Or browse to the Functions.XccdfXml.ps1 file using the Change… button.",
                "PowerSTIG Functions file not found",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private static string? FindPowerStigFunctionsFile()
        {
            const string fileName = "Functions.XccdfXml.ps1";

            // Known repo path under user profile (optional, non-hardcoded)
            var repoCandidate = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "git", "PowerStig", "source", "Module", "Common", fileName);
            if (File.Exists(repoCandidate)) return repoCandidate;

            // Enumerate PSModulePath entries for PowerSTIG
            var psModulePath = Environment.GetEnvironmentVariable("PSModulePath");
            if (!string.IsNullOrWhiteSpace(psModulePath))
            {
                foreach (var dir in psModulePath.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
                {
                    try
                    {
                        // Newer installs: PowerSTIG\\<version>\\...
                        var foundLatest =
                            FindInLatestPowerStigVersion(Path.Combine(dir, "PowerSTIG"), fileName) ??
                            FindInLatestPowerStigVersion(Path.Combine(dir, "PowerStig"), fileName);
                        if (foundLatest is not null) return foundLatest;
                    }
                    catch { /* ignore */ }
                }
            }

            // Typical WindowsPowerShell module locations
            var programFilesModules = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "WindowsPowerShell", "Modules");
            var documentsModules = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WindowsPowerShell", "Modules");

            // Try PowerSTIG latest version under Program Files
            var foundLatestPf =
                FindInLatestPowerStigVersion(Path.Combine(programFilesModules, "PowerSTIG"), fileName) ??
                FindInLatestPowerStigVersion(Path.Combine(programFilesModules, "PowerStig"), fileName);
            if (foundLatestPf is not null) return foundLatestPf;

            // Try PowerSTIG latest version under Documents
            var foundLatestDocs =
                FindInLatestPowerStigVersion(Path.Combine(documentsModules, "PowerSTIG"), fileName) ??
                FindInLatestPowerStigVersion(Path.Combine(documentsModules, "PowerStig"), fileName);
            if (foundLatestDocs is not null) return foundLatestDocs;

            // Recursive search as last resort
            var foundRecursive = FindFileRecursive(Path.Combine(programFilesModules, "PowerSTIG"), fileName) ??
                                 FindFileRecursive(Path.Combine(programFilesModules, "PowerStig"), fileName);
            if (foundRecursive is not null) return foundRecursive;

            return null;
        }

        private static string? FindInLatestPowerStigVersion(string powerStigRoot, string fileName)
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
                    // Common layouts under PowerSTIG\\<version>\\...
                    // 1) ...\\Module\\Common\\Functions.XccdfXml.ps1
                    var underModuleCommon = Path.Combine(vdir, "Module", "Common", fileName);
                    if (File.Exists(underModuleCommon)) return underModuleCommon;

                    // 2) ...\\Common\\Functions.XccdfXml.ps1
                    var underCommon = Path.Combine(vdir, "Common", fileName);
                    if (File.Exists(underCommon)) return underCommon;

                    // 3) Fallback: deep recursive search (handles uncommon layouts)
                    var candidate = FindFileRecursive(vdir, fileName);
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
                if (!Directory.Exists(root))
                    return null;

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

        private static string? ReadLastFunctionsPath()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(LastFunctionsPathRegKey);
                return key?.GetValue(LastFunctionsPathRegValue) as string;
            }
            catch { return null; }
        }

        private static void WriteLastFunctionsPath(string path)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(LastFunctionsPathRegKey);
                key?.SetValue(LastFunctionsPathRegValue, path);
            }
            catch { /* ignore */ }
        }

        private void BrowseModule_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select PowerSTIG Functions.XccdfXml.ps1 file",
                Filter = "PowerShell Scripts (*.ps1)|*.ps1|All files (*.*)|*.*",
                CheckFileExists = true
            };

            // Set initial directory to current module path directory if exists
            if (!string.IsNullOrWhiteSpace(ModulePathTextBox.Text))
            {
                var dir = Path.GetDirectoryName(ModulePathTextBox.Text);
                if (!string.IsNullOrWhiteSpace(dir) && Directory.Exists(dir))
                {
                    dlg.InitialDirectory = dir;
                }
            }

            if (dlg.ShowDialog(this) == true)
            {
                ModulePathTextBox.Text = dlg.FileName;
                WriteLastFunctionsPath(dlg.FileName);

                // Clear warning state
                ModuleWarningText.Visibility = Visibility.Collapsed;
                ModulePathTextBox.Background = System.Windows.Media.Brushes.White;
            }
        }

        private void AppendOutput(string message, System.Windows.Media.Brush? foreground = null, System.Windows.Media.Brush? background = null)
        {
            var para = new Paragraph();
            var run = new Run(message);

            if (foreground is not null) 
                run.Foreground = foreground;
            if (background is not null) 
                para.Background = background;

            para.Inlines.Add(run);
            OutputRichTextBox.Document.Blocks.Add(para);
            OutputRichTextBox.ScrollToEnd();
        }

        private void SplitStigWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            // Save directory settings for next time
            AppSettings.Instance.LastSplitDestinationDirectory = _lastDestinationDirectory;
            AppSettings.Instance.Save();

            // Clean up temp extraction directory
            CleanupTempExtraction();
        }

        private void CleanupTempExtraction()
        {
            if (string.IsNullOrWhiteSpace(_tempExtractPath) || !Directory.Exists(_tempExtractPath))
                return;

            try
            {
                Directory.Delete(_tempExtractPath, true);
                _tempExtractPath = null;
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        private static string? ExtractAndFindXccdfFromZip(string zipPath, out string? tempExtractPath)
        {
            tempExtractPath = null;
            try
            {
                var tempDir = Path.Combine(Path.GetTempPath(), $"PowerStigZip_{Guid.NewGuid():N}");
                Directory.CreateDirectory(tempDir);
                tempExtractPath = tempDir;

                ZipFile.ExtractToDirectory(zipPath, tempDir);

                var xccdfFiles = Directory.GetFiles(tempDir, "*xccdf*.xml", SearchOption.AllDirectories);

                if (xccdfFiles.Length == 0)
                {
                    var allXmlFiles = Directory.GetFiles(tempDir, "*.xml", SearchOption.AllDirectories);
                    xccdfFiles = allXmlFiles.Where(f =>
                        Path.GetFileName(f).Contains("xccdf", StringComparison.OrdinalIgnoreCase) ||
                        Path.GetFileName(f).Contains("Manual", StringComparison.OrdinalIgnoreCase))
                        .ToArray();
                }

                if (xccdfFiles.Length > 0)
                {
                    return xccdfFiles[0];
                }

                return null;
            }
            catch
            {
                if (!string.IsNullOrWhiteSpace(tempExtractPath) && Directory.Exists(tempExtractPath))
                {
                    try { Directory.Delete(tempExtractPath, true); } catch { /* ignore */ }
                }
                tempExtractPath = null;
                return null;
            }
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select DISA Windows OS STIG XML or ZIP File",
                Filter = "STIG files (*.xml;*.zip)|*.xml;*.zip|XML files (*.xml)|*.xml|ZIP files (*.zip)|*.zip|All files (*.*)|*.*",
                CheckFileExists = true
            };

            if (dlg.ShowDialog(this) == true)
            {
                SourcePathTextBox.Text = dlg.FileName;
                ValidateSourceIsWindowsOs();
            }
        }

        private void DestinationBrowse_Click(object sender, RoutedEventArgs e)
        {
            using var fbd = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select destination folder for split STIG files",
                UseDescriptionForTitle = true,
                ShowNewFolderButton = true
            };

            // Only set SelectedPath if we have a saved location
            if (!string.IsNullOrWhiteSpace(_lastDestinationDirectory) && Directory.Exists(_lastDestinationDirectory))
            {
                fbd.SelectedPath = _lastDestinationDirectory;
            }

            var result = fbd.ShowDialog();
            if (result.ToString() == "OK" && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                DestinationPathTextBox.Text = fbd.SelectedPath;
                // Remember the directory for next time
                _lastDestinationDirectory = fbd.SelectedPath;
                AppSettings.Instance.LastSplitDestinationDirectory = _lastDestinationDirectory;
                AppSettings.Instance.Save();
            }
        }

        private void SplitButton_Click(object sender, RoutedEventArgs e)
        {
            // Do not clear; keep history of writes
            StatusText.Text = string.Empty;

            // Clean up any previous temp extraction
            CleanupTempExtraction();
            _actualXccdfPath = null;

            var srcInput = SourcePathTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(srcInput) || !File.Exists(srcInput))
            {
                StatusText.Text = "Invalid source file path.";
                return;
            }

            // Check if input is ZIP and extract if needed
            string? src = srcInput;
            var extension = Path.GetExtension(srcInput);
            bool isZipInput = extension.Equals(".zip", StringComparison.OrdinalIgnoreCase);

            if (isZipInput)
            {
                StatusText.Text = "Extracting ZIP file...";
                AppendOutput($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ZIP file detected, extracting...{Environment.NewLine}", Brushes.DarkBlue);

                src = ExtractAndFindXccdfFromZip(srcInput, out _tempExtractPath);

                if (string.IsNullOrWhiteSpace(src))
                {
                    StatusText.Text = "Could not find XCCDF XML in ZIP archive.";
                    AppendOutput($"Error: Could not find XCCDF XML file in ZIP archive.{Environment.NewLine}{Environment.NewLine}", Brushes.Red);
                    CleanupTempExtraction();
                    return;
                }

                AppendOutput($"Found XCCDF: {Path.GetFileName(src)}{Environment.NewLine}", Brushes.DarkGreen);
            }

            _actualXccdfPath = src;

            if (!IsWindowsOsStig(src))
            {
                StatusText.Text = "Selected file is not a Windows OS STIG. Please choose a Windows Server/Client STIG.";
                AppendOutput($"Error: Not a Windows OS STIG.{Environment.NewLine}{Environment.NewLine}", Brushes.Red);
                CleanupTempExtraction();
                return;
            }

            // For ZIP inputs, determine destination directory
            // Use user-specified destination, or if not set, use the directory where the ZIP file is located
            var destDirText = DestinationPathTextBox.Text?.Trim();
            string destDir;

            if (!string.IsNullOrWhiteSpace(destDirText))
            {
                destDir = destDirText;
            }
            else if (isZipInput)
            {
                // Default to ZIP file's directory for ZIP inputs
                destDir = Path.GetDirectoryName(srcInput)!;
            }
            else
            {
                // Default to XCCDF's directory for direct XCCDF inputs
                destDir = Path.GetDirectoryName(src)!;
            }

            if (!Directory.Exists(destDir))
            {
                StatusText.Text = "Destination folder does not exist.";
                AppendOutput($"Error: Destination folder does not exist: {destDir}{Environment.NewLine}{Environment.NewLine}", Brushes.Red);
                CleanupTempExtraction();
                return;
            }

            try
            {
                string nameNoExt = Path.GetFileNameWithoutExtension(src);
                string baseNameNoEdition = RemoveEditionToken(nameNoExt);

                // Determine output file paths (what PowerSTIG will create)
                var msName = InsertEditionToken(baseNameNoEdition, "MS") + ".xml";
                var dcName = InsertEditionToken(baseNameNoEdition, "DC") + ".xml";
                var msPath = Path.Combine(destDir, msName);
                var dcPath = Path.Combine(destDir, dcName);

                // Check if files exist and handle overwrite logic
                var msExists = File.Exists(msPath);
                var dcExists = File.Exists(dcPath);
                var overwrite = OverwriteCheckBox.IsChecked == true;

                if (msExists || dcExists)
                {
                    if (!overwrite)
                    {
                        var existsMsg =
                            "The following files already exist in the selected destination:\n" +
                            (msExists ? $"  {msPath}\n" : string.Empty) +
                            (dcExists ? $"  {dcPath}\n" : string.Empty) +
                            "\nDo you want to overwrite them?";
                        var overwritePrompt = System.Windows.MessageBox.Show(
                            existsMsg,
                            "Files already exist",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Warning);

                        if (overwritePrompt != MessageBoxResult.Yes)
                        {
                            StatusText.Text = "Split canceled by user.";
                            CleanupTempExtraction();
                            return;
                        }
                    }
                }

                // Header per run with timestamp
                AppendOutput($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Split started using PowerSTIG{Environment.NewLine}", Brushes.DarkBlue);
                AppendOutput($"Source: {(isZipInput ? srcInput : src)}{Environment.NewLine}");
                if (isZipInput)
                {
                    AppendOutput($"Extracted XCCDF: {Path.GetFileName(src)}{Environment.NewLine}", Brushes.DarkCyan);
                }
                AppendOutput($"Destination folder: {destDir}{Environment.NewLine}");

                // Verify module path is set and exists
                var functionsFile = ModulePathTextBox.Text?.Trim();
                if (string.IsNullOrWhiteSpace(functionsFile) || !File.Exists(functionsFile))
                {
                    StatusText.Text = "Functions.XccdfXml.ps1 file not found. Please use the Advanced section to specify the file location.";
                    AppendOutput($"Error: Functions.XccdfXml.ps1 file not found at: {functionsFile ?? "(not set)"}{Environment.NewLine}{Environment.NewLine}", Brushes.Red);
                    CleanupTempExtraction();
                    return;
                }

                AppendOutput($"PowerSTIG Functions file: {functionsFile}{Environment.NewLine}");

                // Create a temporary PowerShell script file
                var tempScript = Path.Combine(Path.GetTempPath(), $"PowerStigSplit_{Guid.NewGuid():N}.ps1");

                var scriptContent = $@"
param(
    [string]$FunctionsFile,
    [string]$XccdfPath,
    [string]$Destination
)
$ErrorActionPreference = 'Stop'
try {{
    # Directly dot-source the Split-StigXccdf function file
    if (-not (Test-Path $FunctionsFile)) {{
        throw ""Functions file not found: $FunctionsFile""
    }}

    . $FunctionsFile
    Write-Output ""Loaded Split-StigXccdf from: $FunctionsFile""

    # Verify function is now available
    if (-not (Get-Command Split-StigXccdf -ErrorAction SilentlyContinue)) {{
        throw ""Split-StigXccdf function was not loaded successfully from $FunctionsFile""
    }}

    Split-StigXccdf -Path $XccdfPath -Destination $Destination
    exit 0
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}}
";
                File.WriteAllText(tempScript, scriptContent, new System.Text.UTF8Encoding(false));

                // Execute PowerShell command
                StatusText.Text = "Splitting STIG using PowerSTIG cmdlet...";

                var psArgs = $"-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File \"{tempScript}\" -FunctionsFile \"{functionsFile}\" -XccdfPath \"{src}\" -Destination \"{destDir}\"";

                // Display the command being executed
                AppendOutput($"{Environment.NewLine}Executing PowerShell command:{Environment.NewLine}", Brushes.DarkBlue);
                AppendOutput($". '{functionsFile}'{Environment.NewLine}", Brushes.Black);
                AppendOutput($"Split-StigXccdf -Path '{src}' -Destination '{destDir}'{Environment.NewLine}{Environment.NewLine}", Brushes.Black);

                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = psArgs,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using var process = System.Diagnostics.Process.Start(psi);
                if (process == null)
                {
                    throw new InvalidOperationException("Failed to start PowerShell process");
                }

                var output = process.StandardOutput.ReadToEnd();
                var error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                // Clean up temp script
                try { File.Delete(tempScript); } catch { /* ignore */ }

                if (process.ExitCode != 0)
                {
                    throw new Exception($"PowerShell execution failed with exit code {process.ExitCode}:\n{error}");
                }

                // Report output if any
                if (!string.IsNullOrWhiteSpace(output))
                {
                    AppendOutput($"{output}{Environment.NewLine}", Brushes.Black);
                }

                // Report success
                AppendOutput($"{(msExists ? "Overwritten" : "Created")}: {msPath}{Environment.NewLine}", msExists ? Brushes.DarkOrange : Brushes.DarkGreen);
                AppendOutput($"{(dcExists ? "Overwritten" : "Created")}: {dcPath}{Environment.NewLine}", dcExists ? Brushes.DarkOrange : Brushes.DarkGreen);

                // Handle log file copying (PowerSTIG doesn't copy log files, so we do it)
                var logSearchPath = isZipInput ? srcInput : src;
                var srcLog = Path.ChangeExtension(logSearchPath, ".log");

                if (File.Exists(srcLog))
                {
                    var msLog = Path.ChangeExtension(msPath, ".log");
                    var dcLog = Path.ChangeExtension(dcPath, ".log");

                    var msLogExisted = File.Exists(msLog);
                    File.Copy(srcLog, msLog, true);
                    AppendOutput($"{(msLogExisted ? "Overwritten" : "Created")} log: {msLog}{Environment.NewLine}", msLogExisted ? Brushes.DarkOrange : Brushes.DarkGreen);

                    var dcLogExisted = File.Exists(dcLog);
                    File.Copy(srcLog, dcLog, true);
                    AppendOutput($"{(dcLogExisted ? "Overwritten" : "Created")} log: {dcLog}{Environment.NewLine}", dcLogExisted ? Brushes.DarkOrange : Brushes.DarkGreen);
                }

                StatusText.Text = "Split completed successfully.";
                AppendOutput($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Split completed{Environment.NewLine}{Environment.NewLine}", Brushes.DarkGreen);
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Split failed: {ex.Message}";
                AppendOutput($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Split failed: {ex.Message}{Environment.NewLine}{Environment.NewLine}", Brushes.Red);
            }
            finally
            {
                // Clean up temp extraction after split completes or fails
                if (isZipInput)
                {
                    CleanupTempExtraction();
                }
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            var main = System.Windows.Application.Current.MainWindow;
            if (main != null)
            {
                if (!main.IsVisible) main.Show();
                main.Activate();
            }
            this.Close();
        }

        private void ClearMessages_Click(object sender, RoutedEventArgs e)
        {
            OutputRichTextBox.Document.Blocks.Clear();
        }

        private void ValidateSourceIsWindowsOs()
        {
            var src = SourcePathTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(src) || !File.Exists(src)) return;

            // Check if it's a ZIP file
            var extension = Path.GetExtension(src);
            if (extension.Equals(".zip", StringComparison.OrdinalIgnoreCase))
            {
                // For ZIP files, we need to extract and check the XCCDF inside
                try
                {
                    var xccdfPath = ExtractAndFindXccdfFromZip(src, out var tempPath);
                    if (string.IsNullOrWhiteSpace(xccdfPath))
                    {
                        System.Windows.MessageBox.Show(
                            "Could not find a valid XCCDF XML file in the ZIP archive.",
                            "Invalid ZIP file",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);

                        // Clean up temp directory
                        if (!string.IsNullOrWhiteSpace(tempPath) && Directory.Exists(tempPath))
                        {
                            try { Directory.Delete(tempPath, true); } catch { /* ignore */ }
                        }
                        return;
                    }

                    // Check if the extracted XCCDF is a Windows OS STIG
                    if (!IsWindowsOsStig(xccdfPath))
                    {
                        System.Windows.MessageBox.Show(
                            "The XCCDF file in the ZIP does not appear to be a Windows OS STIG.\n" +
                            "Expected a name like:\n" +
                            "  U_MS_Windows_Server_2022_STIG_VxRx_Manual-xccdf.xml\n" +
                            "  U_MS_Windows_11_STIG_VxRx_Manual-xccdf.xml",
                            "Invalid STIG selection",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }

                    // Clean up temp directory
                    if (!string.IsNullOrWhiteSpace(tempPath) && Directory.Exists(tempPath))
                    {
                        try { Directory.Delete(tempPath, true); } catch { /* ignore */ }
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(
                        $"Failed to validate ZIP file: {ex.Message}",
                        "Validation Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
            else
            {
                // For direct XCCDF files, check the filename pattern
                if (!IsWindowsOsStig(src))
                {
                    System.Windows.MessageBox.Show(
                        "The selected file does not appear to be a Windows OS STIG.\n" +
                        "Expected a name like:\n" +
                        "  U_MS_Windows_Server_2022_STIG_VxRx_Manual-xccdf.xml\n" +
                        "  U_MS_Windows_11_STIG_VxRx_Manual-xccdf.xml",
                        "Invalid STIG selection",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
        }

        private static bool IsWindowsOsStig(string path)
        {
            var name = Path.GetFileNameWithoutExtension(path);
            if (string.IsNullOrWhiteSpace(name)) return false;

            // Quick filename heuristic
            if (WindowsOsStigName.IsMatch(name))
                return true;

            // Fallback: light XML sniffing (optional, safe)
            try
            {
                using var reader = System.Xml.XmlReader.Create(path, new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });
                while (reader.Read())
                {
                    if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.LocalName.Equals("title", StringComparison.OrdinalIgnoreCase))
                    {
                        var title = reader.ReadElementContentAsString();
                        if (!string.IsNullOrWhiteSpace(title) &&
                            title.IndexOf("Windows", StringComparison.OrdinalIgnoreCase) >= 0 &&
                            (title.IndexOf("Server", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             title.IndexOf("Windows 10", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             title.IndexOf("Windows 11", StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            return true;
                        }
                        break;
                    }
                }
            }
            catch { /* ignore sniff errors */ }
            return false;
        }

        private static string RemoveEditionToken(string nameWithoutExt)
        {
            return Regex.Replace(nameWithoutExt, @"_(MS|DC)_STIG_", "_STIG_", RegexOptions.IgnoreCase);
        }

        private static string InsertEditionToken(string nameWithoutExt, string edition)
        {
            return Regex.Replace(nameWithoutExt, "_STIG_", $"_{edition}_STIG_", RegexOptions.IgnoreCase);
        }
    }
}