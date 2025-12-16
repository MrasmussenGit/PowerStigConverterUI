using System;
using System.IO;
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

                var result = RuleIdAnalysis.Compare(disaPath!, psPath!);

                MissingListView.ItemsSource = result.MissingBaseIds;
                MissingCountTextBlock.Text = $"Missing: {result.MissingBaseIds.Count}";

                MatchedListView.ItemsSource = result.MatchedBaseIds;
                MatchedCountTextBlock.Text = $"Matched: {result.MatchedBaseIds.Count}";

                AddedListView.ItemsSource = result.AddedIds;
                AddedCountTextBlock.Text = $"Added: {result.AddedIds.Count}";

                InfoTextBlock.Text = (result.MissingBaseIds.Count == 0)
                    ? $"No missing DISA rules. Matched: {result.MatchedBaseIds.Count}, Added: {result.AddedIds.Count}"
                    : $"Missing: {result.MissingBaseIds.Count}, Matched: {result.MatchedBaseIds.Count}, Added: {result.AddedIds.Count}";
            }
            catch (Exception ex)
            {
                InfoTextBlock.Text = $"Compare failed: {ex.Message}";
            }
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