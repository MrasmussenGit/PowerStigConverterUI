using System.Windows;

namespace PowerStigConverterUI
{
    public partial class RuleInfoWindow : Window
    {
        public RuleInfoWindow()
        {
            InitializeComponent();
        }

        public void SetRuleInfo(RuleInfo info)
        {
            RuleHeader.Text = $"Rule: {info.RuleId}";
            TitleText.Text = string.IsNullOrWhiteSpace(info.Title) ? "" : $"Title: {info.Title}";
            SeverityText.Text = string.IsNullOrWhiteSpace(info.Severity) ? "" : $"Severity: {info.Severity}";
            SvIdText.Text = string.IsNullOrWhiteSpace(info.SvId) ? "" : $"SV Id: {info.SvId}";

            DescriptionText.Text = info.Description ?? "";
            FixTextBlock.Text = info.FixText ?? "";
            RefsTextBox.Text = info.ReferencesXml ?? "";

            ConvertedFileText.Text = string.IsNullOrWhiteSpace(info.ConvertedFile) ? "Converted file: (not found)" : $"Converted file: {info.ConvertedFile}";
            ConvertedSnippetText.Text = info.ConvertedSnippet ?? "";
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    public sealed class RuleInfo
    {
        public string RuleId { get; set; }
        public string? SvId { get; set; }
        public string? Title { get; set; }
        public string? Severity { get; set; }
        public string? Description { get; set; }
        public string? FixText { get; set; }
        public string? ReferencesXml { get; set; }
        public string? ConvertedFile { get; set; }
        public string? ConvertedSnippet { get; set; }
    }
}