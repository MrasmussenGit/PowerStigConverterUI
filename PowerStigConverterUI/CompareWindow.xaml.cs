using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using PowerStigConverterUI; // Change this to the correct namespace if needed

public partial class CompareWindow : Window
{
    public CompareWindow()
    {
        InitializeComponent();
    }

    void BrowseDisa_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*" };
        if (dlg.ShowDialog() == true) DisaPathTextBox.Text = dlg.FileName;
    }

    void BrowsePs_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*" };
        if (dlg.ShowDialog() == true) PsPathTextBox.Text = dlg.FileName;
    }

    void Compare_Click(object sender, RoutedEventArgs e)
    {
        var disaPath = DisaPathTextBox.Text;
        var psPath = PsPathTextBox.Text;

        if (string.IsNullOrWhiteSpace(disaPath) || string.IsNullOrWhiteSpace(psPath))
        {
            System.Windows.MessageBox.Show("Please select both files before comparing.");
            return;
        }

        try
        {
            var missing = MainWindow.GetMissingIds(disaPath, psPath);
            var added = MainWindow.GetAddedIds(disaPath, psPath);

            MissingCountTextBlock.Text = $"Count: {missing.Count}";
            AddedCountTextBlock.Text = $"Count: {added.Count}";
            MissingListView.ItemsSource = missing;
            AddedListView.ItemsSource = added;

            InfoTextBlock.Text = $"Compared successfully. Missing: {missing.Count}, Added: {added.Count}";
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Compare failed: {ex.Message}");
        }
    }

    private void MissingListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is System.Windows.Controls.ListView lv && lv.SelectedItem is string id && !string.IsNullOrWhiteSpace(id))
        {
            var win = new RuleDetailWindow(id, DisaPathTextBox.Text, PsPathTextBox.Text) { Owner = this };
            win.ShowDialog();
        }
    }

    private void AddedListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is System.Windows.Controls.ListView lv && lv.SelectedItem is string id && !string.IsNullOrWhiteSpace(id))
        {
            var win = new RuleDetailWindow(id, DisaPathTextBox.Text, PsPathTextBox.Text) { Owner = this };
            win.ShowDialog();
        }
    }
}