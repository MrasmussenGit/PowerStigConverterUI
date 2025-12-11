using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;

namespace PowerStigConverterUI
{
    public partial class RuleDetailWindow : Window
    {
        private readonly string _ruleId;
        private readonly string _disaPath;
        private readonly string _psPath;

        private TextBlock DetailsTextBlock => (TextBlock)FindName("DetailsTextBlock");

        public RuleDetailWindow(string ruleId, string disaPath, string psPath)
        {
            InitializeComponent();
            _ruleId = ruleId;
            _disaPath = disaPath;
            _psPath = psPath;

            Title = $"Rule Details - {_ruleId}";
            RuleIdTextBlock.Text = _ruleId;

            LoadRuleDetails();
        }

        private void LoadRuleDetails()
        {
            var details = TryGetRuleDetails(_psPath) ?? TryGetRuleDetails(_disaPath);

            if (string.IsNullOrWhiteSpace(details))
            {
                DetailsTextBlock.Text = "No details found for this rule in the selected files.";
                return;
            }

            DetailsTextBlock.Text = details;
        }

        private string? TryGetRuleDetails(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;

                var doc = XDocument.Load(path);

                // Try common patterns: attribute 'id' or child element 'id'
                var match = doc
                    .Descendants()
                    .FirstOrDefault(e =>
                        string.Equals((string?)e.Attribute("id"), _ruleId, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals((string?)e.Element("id"), _ruleId, StringComparison.OrdinalIgnoreCase));

                if (match == null) return null;

                // Extract useful fields if present; otherwise show the matched XML fragment text.
                var title = (string?)match.Element("title") ?? (string?)match.Element("Rule_Title");
                var desc = (string?)match.Element("description") ?? (string?)match.Element("Rule_Description");
                var rationale = (string?)match.Element("rationale") ?? (string?)match.Element("Vuln_Discuss");
                var fix = (string?)match.Element("fixtext") ?? (string?)match.Element("Fix_Text");
                var severity = (string?)match.Element("severity") ?? (string?)match.Element("Severity");

                string formatted =
                    $"Title: {title ?? "(unknown)"}{Environment.NewLine}{Environment.NewLine}" +
                    $"Severity: {severity ?? "(unknown)"}{Environment.NewLine}{Environment.NewLine}" +
                    $"Description:{Environment.NewLine}{(desc ?? "(none)")}{Environment.NewLine}{Environment.NewLine}" +
                    $"Rationale:{Environment.NewLine}{(rationale ?? "(none)")}{Environment.NewLine}{Environment.NewLine}" +
                    $"Fix:{Environment.NewLine}{(fix ?? "(none)")}";

                return formatted;
            }
            catch
            {
                return null;
            }
        }
    }
}