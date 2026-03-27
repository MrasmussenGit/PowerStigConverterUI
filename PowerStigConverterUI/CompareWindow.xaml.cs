using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PowerStigConverterUI
{
    public partial class CompareWindow : Window
    {
        private string? _tempExtractPath = null;
        private string? _actualDisaXccdfPath = null; // Track the actual XCCDF path being compared

        public CompareWindow()
        {
            InitializeComponent();
            this.Closing += CompareWindow_Closing;
        }

        private void CompareWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            // Clean up temp extraction directory when window closes
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
                // Create a unique temporary directory for extraction
                var tempDir = Path.Combine(Path.GetTempPath(), $"PowerStigZip_{Guid.NewGuid():N}");
                Directory.CreateDirectory(tempDir);
                tempExtractPath = tempDir;

                // Extract the ZIP file
                ZipFile.ExtractToDirectory(zipPath, tempDir);

                // Search recursively for XCCDF XML files
                var xccdfFiles = Directory.GetFiles(tempDir, "*xccdf*.xml", SearchOption.AllDirectories);

                if (xccdfFiles.Length == 0)
                {
                    // Try broader search for any XML files
                    var allXmlFiles = Directory.GetFiles(tempDir, "*.xml", SearchOption.AllDirectories);
                    xccdfFiles = allXmlFiles.Where(f =>
                        Path.GetFileName(f).Contains("xccdf", StringComparison.OrdinalIgnoreCase) ||
                        Path.GetFileName(f).Contains("Manual", StringComparison.OrdinalIgnoreCase))
                        .ToArray();
                }

                if (xccdfFiles.Length > 0)
                {
                    // Return the first XCCDF file found
                    return xccdfFiles[0];
                }

                return null;
            }
            catch
            {
                // Clean up temp directory if extraction failed
                if (!string.IsNullOrWhiteSpace(tempExtractPath) && Directory.Exists(tempExtractPath))
                {
                    try { Directory.Delete(tempExtractPath, true); } catch { /* ignore cleanup errors */ }
                }
                tempExtractPath = null;
                return null;
            }
        }

        private void BrowseDisa_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select DISA XCCDF XML or STIG ZIP File",
                Filter = "STIG files (*.xml;*.zip)|*.xml;*.zip|XML files (*.xml)|*.xml|ZIP files (*.zip)|*.zip|All files (*.*)|*.*",
                CheckFileExists = true
            };
            if (dlg.ShowDialog(this) == true)
            {
                DisaPathTextBox.Text = dlg.FileName;
            }
        }

        private void BrowsePs_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select Converted PowerSTIG XML",
                Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
                CheckFileExists = true
            };
            if (dlg.ShowDialog(this) == true)
            {
                PsPathTextBox.Text = dlg.FileName;
            }
        }

        private static string? DeriveLogFilePathFromXccdf(string? xccdfPath, string? originalInputPath, bool isZipInput)
        {
            if (string.IsNullOrWhiteSpace(xccdfPath)) return null;

            if (isZipInput && !string.IsNullOrWhiteSpace(originalInputPath))
            {
                // For ZIP files: derive log filename from XCCDF, but place in ZIP's directory
                var zipDirectory = Path.GetDirectoryName(originalInputPath)!;
                var xccdfBaseName = Path.GetFileNameWithoutExtension(xccdfPath);
                return Path.Combine(zipDirectory, xccdfBaseName + ".log");
            }
            else
            {
                // For direct XCCDF files: use XCCDF directory
                var directory = Path.GetDirectoryName(xccdfPath);
                var xccdfBaseName = Path.GetFileNameWithoutExtension(xccdfPath);

                if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(xccdfBaseName))
                    return null;

                return Path.Combine(directory, xccdfBaseName + ".log");
            }
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
                        // Everything else is a Hard Coded rule
                        hardCoded.Add(ruleId);
                    }
                }
            }
            catch { /* ignore parse errors */ }

            return (skips, hardCoded);
        }

        private static string NormalizeVId(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return string.Empty;
            // Normalize "V-123456.a" or "V-123456" to "V-123456"
            var m = System.Text.RegularExpressions.Regex.Match(id, @"^(V-\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value : id;
        }

        private void Compare_Click(object sender, RoutedEventArgs e)
        {
            InfoTextBlock.Text = string.Empty;
            MissingListView.ItemsSource = null;
            AddedListView.ItemsSource = null;
            MatchedListView.ItemsSource = null;
            SkippedListView.ItemsSource = null;
            MissingCountTextBlock.Text = string.Empty;
            AddedCountTextBlock.Text = string.Empty;
            MatchedCountTextBlock.Text = string.Empty;
            SkippedCountTextBlock.Text = string.Empty;

            // Clean up any previous temp extraction
            CleanupTempExtraction();
            _actualDisaXccdfPath = null;

            var disaInputPath = DisaPathTextBox.Text?.Trim();
            var psPath = PsPathTextBox.Text?.Trim();

            if (string.IsNullOrWhiteSpace(disaInputPath) || !File.Exists(disaInputPath))
            {
                InfoTextBlock.Text = "Invalid DISA XCCDF or ZIP path.";
                return;
            }
            if (string.IsNullOrWhiteSpace(psPath) || !File.Exists(psPath))
            {
                InfoTextBlock.Text = "Invalid PowerSTIG XML path.";
                return;
            }

            try
            {
                // Check if DISA input is a ZIP file and extract if needed
                string? disaXccdfPath = disaInputPath;
                var extension = Path.GetExtension(disaInputPath);
                bool isZipInput = extension.Equals(".zip", StringComparison.OrdinalIgnoreCase);

                if (isZipInput)
                {
                    InfoTextBlock.Text = "Extracting ZIP file...";

                    disaXccdfPath = ExtractAndFindXccdfFromZip(disaInputPath, out _tempExtractPath);

                    if (string.IsNullOrWhiteSpace(disaXccdfPath))
                    {
                        InfoTextBlock.Text = "Could not find XCCDF XML file in the ZIP archive.";
                        CleanupTempExtraction();
                        return;
                    }
                }

                // Store the actual XCCDF path for use in double-click handlers
                _actualDisaXccdfPath = disaXccdfPath;

                // Read log file to identify skipped rules
                var logPath = DeriveLogFilePathFromXccdf(disaXccdfPath, disaInputPath, isZipInput);
                var skippedRuleIds = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (!string.IsNullOrWhiteSpace(logPath) && File.Exists(logPath))
                {
                    var (skips, hardCoded) = ParseLogFile(logPath);
                    foreach (var skip in skips)
                        skippedRuleIds.Add(skip);
                    // Note: Hard coded rules are not shown in skipped list, they're still in matched
                }

                InfoTextBlock.Text = "Comparing…";

                var result = RuleIdAnalysis.Compare(disaXccdfPath!, psPath!);

                // Separate skipped rules from matched rules
                var normalizedSkippedIds = new System.Collections.Generic.HashSet<string>(
                    skippedRuleIds.Select(id => NormalizeVId(id)), 
                    StringComparer.OrdinalIgnoreCase);

                var truelyMatched = result.MatchedBaseIds
                    .Where(id => !normalizedSkippedIds.Contains(NormalizeVId(id)))
                    .ToList();

                var skippedMatched = result.MatchedBaseIds
                    .Where(id => normalizedSkippedIds.Contains(NormalizeVId(id)))
                    .ToList();

                MatchedListView.ItemsSource = truelyMatched;
                MatchedCountTextBlock.Text = $"Matched: {truelyMatched.Count}";

                SkippedListView.ItemsSource = skippedMatched;
                SkippedCountTextBlock.Text = $"Skipped: {skippedMatched.Count}";

                MissingListView.ItemsSource = result.MissingBaseIds;
                MissingCountTextBlock.Text = $"Missing: {result.MissingBaseIds.Count}";

                AddedListView.ItemsSource = result.AddedIds;
                AddedCountTextBlock.Text = $"Added: {result.AddedIds.Count}";

                // Calculate coverage: how many DISA base rules are covered
                var totalDisaRules = truelyMatched.Count + skippedMatched.Count + result.MissingBaseIds.Count;
                var coveredRules = truelyMatched.Count + skippedMatched.Count;
                var coveragePercent = totalDisaRules > 0 ? (coveredRules * 100.0 / totalDisaRules) : 0;

                InfoTextBlock.Text = $"Coverage: {coveredRules}/{totalDisaRules} DISA rules ({coveragePercent:F1}%) • Matched: {truelyMatched.Count}, Skipped: {skippedMatched.Count}, Missing: {result.MissingBaseIds.Count}, Added: {result.AddedIds.Count}";
            }
            catch (Exception ex)
            {
                InfoTextBlock.Text = $"Compare failed: {ex.Message}";
                CleanupTempExtraction();
            }
        }

        private void MissingListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (MissingListView.SelectedItem is string id &&
                !string.IsNullOrWhiteSpace(_actualDisaXccdfPath) && File.Exists(_actualDisaXccdfPath) &&
                !string.IsNullOrWhiteSpace(PsPathTextBox.Text) && File.Exists(PsPathTextBox.Text))
            {
                var win = new RuleDetailWindow(id, _actualDisaXccdfPath, PsPathTextBox.Text) { Owner = this };
                win.ShowDialog();
            }
        }

        private void SkippedListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (SkippedListView.SelectedItem is string id &&
                !string.IsNullOrWhiteSpace(_actualDisaXccdfPath) && File.Exists(_actualDisaXccdfPath) &&
                !string.IsNullOrWhiteSpace(PsPathTextBox.Text) && File.Exists(PsPathTextBox.Text))
            {
                var win = new RuleDetailWindow(id, _actualDisaXccdfPath, PsPathTextBox.Text) { Owner = this };
                win.ShowDialog();
            }
        }

        private void AddedListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (AddedListView.SelectedItem is string id &&
                !string.IsNullOrWhiteSpace(_actualDisaXccdfPath) && File.Exists(_actualDisaXccdfPath) &&
                !string.IsNullOrWhiteSpace(PsPathTextBox.Text) && File.Exists(PsPathTextBox.Text))
            {
                var win = new RuleDetailWindow(id, _actualDisaXccdfPath, PsPathTextBox.Text) { Owner = this };
                win.ShowDialog();
            }
        }

        private void MatchedListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (MatchedListView.SelectedItem is string id &&
                !string.IsNullOrWhiteSpace(_actualDisaXccdfPath) && File.Exists(_actualDisaXccdfPath) &&
                !string.IsNullOrWhiteSpace(PsPathTextBox.Text) && File.Exists(PsPathTextBox.Text))
            {
                var win = new RuleDetailWindow(id, _actualDisaXccdfPath, PsPathTextBox.Text) { Owner = this };
                win.ShowDialog();
            }
        }
    }
}