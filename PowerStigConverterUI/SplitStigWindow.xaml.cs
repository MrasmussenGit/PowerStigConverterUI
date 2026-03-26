using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using Brushes = System.Windows.Media.Brushes;

namespace PowerStigConverterUI
{
    public partial class SplitStigWindow : Window
    {
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
            this.Closing += SplitStigWindow_Closing;

            // Load persisted directory settings
            _lastDestinationDirectory = AppSettings.Instance.LastSplitDestinationDirectory;
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

            var dir = Path.GetDirectoryName(src)!;
            var nameNoExt = Path.GetFileNameWithoutExtension(src);

            bool hasEditionToken = Regex.IsMatch(nameNoExt, @"_(MS|DC)_STIG_", RegexOptions.IgnoreCase);
            if (hasEditionToken)
            {
                var result = System.Windows.MessageBox.Show(
                    "The selected STIG filename already contains an edition token (MS/DC).\n" +
                    "Proceeding will still produce one MS and one DC file based on the base name.\n\n" +
                    "Do you want to continue?",
                    "Edition token detected",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result != MessageBoxResult.Yes)
                {
                    StatusText.Text = "Split canceled by user.";
                    CleanupTempExtraction();
                    return;
                }
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
                destDir = dir;
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
                string baseNameNoEdition = RemoveEditionToken(nameNoExt);

                var msName = InsertEditionToken(baseNameNoEdition, "MS") + ".xml";
                var dcName = InsertEditionToken(baseNameNoEdition, "DC") + ".xml";

                var msPath = Path.Combine(destDir, msName);
                var dcPath = Path.Combine(destDir, dcName);

                // Warn if target file names collide with source path (only relevant for non-ZIP inputs)
                if (!isZipInput && (string.Equals(src, msPath, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(src, dcPath, StringComparison.OrdinalIgnoreCase)))
                {
                    var overwriteWarn = System.Windows.MessageBox.Show(
                        "One of the output files matches the selected source filename.\n" +
                        "To avoid overwriting, the copy will be skipped for the matching file.\n\n" +
                        "Do you want to continue?",
                        "Potential filename collision",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning);

                    if (overwriteWarn != MessageBoxResult.Yes)
                    {
                        StatusText.Text = "Split canceled by user.";
                        CleanupTempExtraction();
                        return;
                    }
                }

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

                        overwrite = true; // proceed with overwriting after explicit consent
                    }
                }

                // Load source XML once
                var xmlDoc = new System.Xml.XmlDocument();
                xmlDoc.Load(src);

                // Configure XML writer settings for proper formatting
                var writerSettings = new System.Xml.XmlWriterSettings
                {
                    Indent = true,
                    IndentChars = "  ", // 2 spaces
                    NewLineChars = Environment.NewLine,
                    NewLineHandling = System.Xml.NewLineHandling.Replace,
                    Encoding = new System.Text.UTF8Encoding(false) // UTF-8 without BOM
                };

                // Header per run with timestamp
                AppendOutput($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Split started{Environment.NewLine}", Brushes.DarkBlue);
                AppendOutput($"Source: {(isZipInput ? srcInput : src)}{Environment.NewLine}");
                if (isZipInput)
                {
                    AppendOutput($"Extracted XCCDF: {Path.GetFileName(src)}{Environment.NewLine}", Brushes.DarkCyan);
                }
                AppendOutput($"Destination folder: {destDir}{Environment.NewLine}");

                if (!isZipInput && string.Equals(src, msPath, StringComparison.OrdinalIgnoreCase))
                {
                    AppendOutput($"MS target: {msPath} (skipped; same as source){Environment.NewLine}", Brushes.DarkOrange);
                }
                else
                {
                    using (var writer = System.Xml.XmlWriter.Create(msPath, writerSettings))
                    {
                        xmlDoc.Save(writer);
                    }
                    AppendOutput($"{(msExists ? "Overwritten" : "Created")}: {msPath}{Environment.NewLine}", msExists ? Brushes.DarkOrange : Brushes.DarkGreen);
                }

                if (!isZipInput && string.Equals(src, dcPath, StringComparison.OrdinalIgnoreCase))
                {
                    AppendOutput($"DC target: {dcPath} (skipped; same as source){Environment.NewLine}", Brushes.DarkOrange);
                }
                else
                {
                    using (var writer = System.Xml.XmlWriter.Create(dcPath, writerSettings))
                    {
                        xmlDoc.Save(writer);
                    }
                    AppendOutput($"{(dcExists ? "Overwritten" : "Created")}: {dcPath}{Environment.NewLine}", dcExists ? Brushes.DarkOrange : Brushes.DarkGreen);
                }

                // Handle log file copying
                // For ZIP inputs, look for log file in the ZIP's directory (same as ZIP file)
                // For direct XCCDF inputs, look for log file next to the XCCDF
                var logSearchPath = isZipInput ? srcInput : src;
                var srcLog = Path.ChangeExtension(logSearchPath, ".log");

                if (File.Exists(srcLog))
                {
                    var msLog = Path.ChangeExtension(msPath, ".log");
                    var dcLog = Path.ChangeExtension(dcPath, ".log");

                    var msLogExisted = File.Exists(msLog);
                    File.Copy(srcLog, msLog, overwrite);
                    AppendOutput($"{(msLogExisted && overwrite ? "Overwritten" : "Created")} log: {msLog}{Environment.NewLine}", msLogExisted ? Brushes.DarkOrange : Brushes.DarkGreen);

                    var dcLogExisted = File.Exists(dcLog);
                    File.Copy(srcLog, dcLog, overwrite);
                    AppendOutput($"{(dcLogExisted && overwrite ? "Overwritten" : "Created")} log: {dcLog}{Environment.NewLine}", dcLogExisted ? Brushes.DarkOrange : Brushes.DarkGreen);
                }

                StatusText.Text = "Split completed.";
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