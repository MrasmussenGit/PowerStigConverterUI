using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;

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

        public SplitStigWindow()
        {
            InitializeComponent();
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select DISA Windows OS STIG XML",
                Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
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
            var src = SourcePathTextBox.Text?.Trim();
            var initialDir = !string.IsNullOrWhiteSpace(src) && File.Exists(src)
                ? Path.GetDirectoryName(src)!
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            using var fbd = new FolderBrowserDialog
            {
                Description = "Select destination folder for split STIG files",
                UseDescriptionForTitle = true,
                ShowNewFolderButton = true,
                SelectedPath = initialDir
            };

            var result = fbd.ShowDialog();
            //string val = result.ToString();
            //bool val1 = Convert.ToBoolean(val);
            if (result.ToString() == "OK" && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                DestinationPathTextBox.Text = fbd.SelectedPath;
            }
        }

        private void SplitButton_Click(object sender, RoutedEventArgs e)
        {
            // Do not clear; keep history of writes
            StatusText.Text = string.Empty;

            var src = SourcePathTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(src) || !File.Exists(src))
            {
                StatusText.Text = "Invalid source XML path.";
                return;
            }

            if (!IsWindowsOsStig(src))
            {
                StatusText.Text = "Selected file is not a Windows OS STIG. Please choose a Windows Server/Client STIG.";
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
                    return;
                }
            }

            var destDirText = DestinationPathTextBox.Text?.Trim();
            var destDir = string.IsNullOrWhiteSpace(destDirText) ? dir : destDirText;

            if (!Directory.Exists(destDir))
            {
                StatusText.Text = "Destination folder does not exist.";
                return;
            }

            try
            {
                string baseNameNoEdition = RemoveEditionToken(nameNoExt);

                var msName = InsertEditionToken(baseNameNoEdition, "MS") + ".xml";
                var dcName = InsertEditionToken(baseNameNoEdition, "DC") + ".xml";

                var msPath = Path.Combine(destDir, msName);
                var dcPath = Path.Combine(destDir, dcName);

                // Warn if target file names collide with source path
                if (string.Equals(src, msPath, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(src, dcPath, StringComparison.OrdinalIgnoreCase))
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
                            return;
                        }

                        overwrite = true; // proceed with overwriting after explicit consent
                    }
                }

                // Header per run with timestamp
                OutputTextBox.AppendText($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Split started{Environment.NewLine}");
                OutputTextBox.AppendText($"Source: {src}{Environment.NewLine}");
                OutputTextBox.AppendText($"Destination folder: {destDir}{Environment.NewLine}");

                if (!string.Equals(src, msPath, StringComparison.OrdinalIgnoreCase))
                {
                    File.Copy(src, msPath, overwrite);
                    OutputTextBox.AppendText($"{(msExists ? "Overwritten" : "Created")}: {msPath}{Environment.NewLine}");
                }
                else
                {
                    OutputTextBox.AppendText($"MS target: {msPath} (skipped; same as source){Environment.NewLine}");
                }

                if (!string.Equals(src, dcPath, StringComparison.OrdinalIgnoreCase))
                {
                    File.Copy(src, dcPath, overwrite);
                    OutputTextBox.AppendText($"{(dcExists ? "Overwritten" : "Created")}: {dcPath}{Environment.NewLine}");
                }
                else
                {
                    OutputTextBox.AppendText($"DC target: {dcPath} (skipped; same as source){Environment.NewLine}");
                }

                var srcLog = Path.ChangeExtension(src, ".log");
                if (File.Exists(srcLog))
                {
                    var msLog = Path.ChangeExtension(msPath, ".log");
                    var dcLog = Path.ChangeExtension(dcPath, ".log");

                    File.Copy(srcLog, msLog, overwrite);
                    OutputTextBox.AppendText($"{(File.Exists(msLog) && overwrite ? "Overwritten" : "Created")} log: {msLog}{Environment.NewLine}");

                    File.Copy(srcLog, dcLog, overwrite);
                    OutputTextBox.AppendText($"{(File.Exists(dcLog) && overwrite ? "Overwritten" : "Created")} log: {dcLog}{Environment.NewLine}");
                }

                StatusText.Text = "Split completed.";
                OutputTextBox.AppendText($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Split completed{Environment.NewLine}{Environment.NewLine}");
                OutputTextBox.ScrollToEnd();
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Split failed: {ex.Message}";
                OutputTextBox.AppendText($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Split failed: {ex.Message}{Environment.NewLine}{Environment.NewLine}");
                OutputTextBox.ScrollToEnd();
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
            if (string.IsNullOrWhiteSpace(src)) return;

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