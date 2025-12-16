using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PowerStigConverterUI
{
    public partial class CompareWindow : Window
    {
        public CompareWindow()
        {
            InitializeComponent();
        }

        private void BrowseDisa_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select DISA XCCDF XML",
                Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
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

        private void Compare_Click(object sender, RoutedEventArgs e)
        {
            InfoTextBlock.Text = string.Empty;
            MissingListView.ItemsSource = null;
            AddedListView.ItemsSource = null;
            MatchedListView.ItemsSource = null;
            MissingCountTextBlock.Text = string.Empty;
            AddedCountTextBlock.Text = string.Empty;
            MatchedCountTextBlock.Text = string.Empty;

            var disaPath = DisaPathTextBox.Text?.Trim();
            var psPath = PsPathTextBox.Text?.Trim();

            if (string.IsNullOrWhiteSpace(disaPath) || !File.Exists(disaPath))
            {
                InfoTextBlock.Text = "Invalid DISA XCCDF path.";
                return;
            }
            if (string.IsNullOrWhiteSpace(psPath) || !File.Exists(psPath))
            {
                InfoTextBlock.Text = "Invalid PowerSTIG XML path.";
                return;
            }

            try
            {
                InfoTextBlock.Text = "Comparing…";

                var missing = MainWindow.GetMissingIds(disaPath!, psPath!);
                var added = MainWindow.GetAddedIds(disaPath!, psPath!);

                MissingListView.ItemsSource = missing;
                AddedListView.ItemsSource = added;
                MissingCountTextBlock.Text = $"Missing: {missing.Count}";
                AddedCountTextBlock.Text = $"Added: {added.Count}";

                // Build normalized base IDs to compute matches (V-xxxxxx)
                var disaBase = new HashSet<string>(
                    ExtractDisaRuleIds(disaPath!).Select(NormalizeToBaseV).Where(s => !string.IsNullOrEmpty(s)),
                    StringComparer.OrdinalIgnoreCase);

                var psBase = new HashSet<string>(
                    ExtractPsRuleIds(psPath!).Select(NormalizeToBaseV).Where(s => !string.IsNullOrEmpty(s)),
                    StringComparer.OrdinalIgnoreCase);

                var matched = disaBase.Intersect(psBase, StringComparer.OrdinalIgnoreCase)
                                      .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                                      .ToList();

                MatchedListView.ItemsSource = matched;
                MatchedCountTextBlock.Text = $"Matched: {matched.Count}";

                InfoTextBlock.Text = (missing.Count == 0 && added.Count == 0)
                    ? $"No differences. Matched rules: {matched.Count}"
                    : $"Found {missing.Count} missing and {added.Count} added.";
            }
            catch (Exception ex)
            {
                InfoTextBlock.Text = $"Compare failed: {ex.Message}";
            }
        }

        private static IEnumerable<string> ExtractDisaRuleIds(string disaFile)
        {
            using var reader = System.Xml.XmlReader.Create(
                disaFile,
                new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });

            while (reader.Read())
            {
                if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.LocalName.Equals("Rule", StringComparison.OrdinalIgnoreCase))
                {
                    var id = reader.GetAttribute("id");
                    if (!string.IsNullOrWhiteSpace(id))
                        yield return id.Trim();
                }
            }
        }

        private static IEnumerable<string> ExtractPsRuleIds(string psFile)
        {
            using var reader = System.Xml.XmlReader.Create(
                psFile,
                new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });

            while (reader.Read())
            {
                if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.LocalName.Equals("Rule", StringComparison.OrdinalIgnoreCase))
                {
                    var id = reader.GetAttribute("id");
                    if (!string.IsNullOrWhiteSpace(id))
                        yield return id.Trim();
                }
            }
        }

        private static string NormalizeToBaseV(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            var s = raw.Trim();

            var sv = System.Text.RegularExpressions.Regex.Match(s, @"^SV-(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (sv.Success) return $"V-{sv.Groups[1].Value}";

            var v = System.Text.RegularExpressions.Regex.Match(s, @"^V-(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (v.Success) return $"V-{v.Groups[1].Value}";

            return string.Empty;
        }

        private void MissingListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (MissingListView.SelectedItem is string id &&
                !string.IsNullOrWhiteSpace(DisaPathTextBox.Text) && File.Exists(DisaPathTextBox.Text) &&
                !string.IsNullOrWhiteSpace(PsPathTextBox.Text) && File.Exists(PsPathTextBox.Text))
            {
                var win = new RuleDetailWindow(id, DisaPathTextBox.Text, PsPathTextBox.Text) { Owner = this };
                win.ShowDialog();
            }
        }

        private void AddedListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (AddedListView.SelectedItem is string id &&
                !string.IsNullOrWhiteSpace(DisaPathTextBox.Text) && File.Exists(DisaPathTextBox.Text) &&
                !string.IsNullOrWhiteSpace(PsPathTextBox.Text) && File.Exists(PsPathTextBox.Text))
            {
                var win = new RuleDetailWindow(id, DisaPathTextBox.Text, PsPathTextBox.Text) { Owner = this };
                win.ShowDialog();
            }
        }

        private void MatchedListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (MatchedListView.SelectedItem is string id &&
                !string.IsNullOrWhiteSpace(DisaPathTextBox.Text) && File.Exists(DisaPathTextBox.Text) &&
                !string.IsNullOrWhiteSpace(PsPathTextBox.Text) && File.Exists(PsPathTextBox.Text))
            {
                var win = new RuleDetailWindow(id, DisaPathTextBox.Text, PsPathTextBox.Text) { Owner = this };
                win.ShowDialog();
            }
        }
    }
}