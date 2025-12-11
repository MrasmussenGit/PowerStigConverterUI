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

        public RuleDetailWindow(string ruleId, string disaPath, string psPath)
        {
            InitializeComponent();
            _ruleId = ruleId;
            _disaPath = disaPath;
            _psPath = psPath;

            Title = $"Rule Details - {_ruleId}";
            RuleIdTextBlock.Text = _ruleId;

            // Center over owner if set; otherwise center on screen
            Loaded += (_, __) =>
            {
                if (Owner != null)
                {
                    Left = Owner.Left + (Owner.Width - ActualWidth) / 2;
                    Top = Owner.Top + (Owner.Height - ActualHeight) / 2;
                }
                else
                {
                    var wa = SystemParameters.WorkArea;
                    Left = wa.Left + (wa.Width - ActualWidth) / 2;
                    Top = wa.Top + (wa.Height - ActualHeight) / 2;
                }
            };

            LoadRuleDetails();
        }

        private void LoadRuleDetails()
        {
            // Prefer original DISA STIG for rich details, then fall back to PowerSTIG
            var details = TryGetRuleDetails(_disaPath) ?? TryGetRuleDetails(_psPath);

            if (string.IsNullOrWhiteSpace(details))
            {
                DetailsTextBlock.Text = "No details found for this rule in the selected files.";
                return;
            }

            DetailsTextBlock.Text = details;
        }

        private static string NormalizeId(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            var id = raw.Trim();
            var v = System.Text.RegularExpressions.Regex.Match(id, @"\bV-(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (v.Success) return $"V-{v.Groups[1].Value}";
            var sv = System.Text.RegularExpressions.Regex.Match(id, @"\bSV-(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (sv.Success) return $"V-{sv.Groups[1].Value}";
            return id;
        }

        private static string? GetChildValue(XElement e, params string[] names)
        {
            return e.Elements()
                    .FirstOrDefault(x => names.Any(n => string.Equals(x.Name.LocalName, n, StringComparison.OrdinalIgnoreCase)))
                    ?.Value;
        }

        private static XElement? GetChildElement(XElement e, params string[] names)
        {
            return e.Elements()
                    .FirstOrDefault(x => names.Any(n => string.Equals(x.Name.LocalName, n, StringComparison.OrdinalIgnoreCase)));
        }

        private string? TryGetRuleDetails(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;

                var doc = XDocument.Load(path);

                // Try to find a matching Rule by multiple id shapes commonly used in DISA STIGs
                var rule = doc
                    .Descendants()
                    .Where(e => string.Equals(e.Name.LocalName, "Rule", StringComparison.OrdinalIgnoreCase))
                    .FirstOrDefault(e =>
                    {
                        var attrId = (string?)e.Attribute("id") ?? (string?)e.Attribute("Id");
                        var version = GetChildValue(e, "version");
                        var vulnNum = GetChildValue(e, "Vuln_Num");
                        var ruleId2 = GetChildValue(e, "Rule_ID");
                        var candidates = new[] { attrId, version, vulnNum, ruleId2 }
                                         .Where(s => !string.IsNullOrWhiteSpace(s))
                                         .Select(NormalizeId);
                        return candidates.Any(c => string.Equals(c, _ruleId, StringComparison.OrdinalIgnoreCase));
                    });

                // Fallback: any element with id attribute matching after normalization (helps for PowerSTIG structures)
                rule ??= doc.Descendants()
                    .FirstOrDefault(e =>
                    {
                        var attrId = (string?)e.Attribute("id") ?? (string?)e.Attribute("Id");
                        return !string.IsNullOrWhiteSpace(attrId) &&
                               string.Equals(NormalizeId(attrId), _ruleId, StringComparison.OrdinalIgnoreCase);
                    });

                if (rule == null) return null;

                // Extract useful fields (namespace-agnostic)
                var title = GetChildValue(rule, "title", "Rule_Title") ?? "(unknown)";
                var desc = GetChildValue(rule, "description", "Rule_Description") ?? "(none)";
                var rationale = GetChildValue(rule, "rationale", "Vuln_Discuss") ?? "(none)";
                var fix = GetChildValue(rule, "fixtext", "Fix_Text") ?? "(none)";
                var severity = (string?)rule.Attribute("severity") ?? GetChildValue(rule, "Severity") ?? "(unknown)";

                // Extract <check-content> from DISA STIG structure: <check><check-content>...</check-content></check>
                string checkContent = "(none)";
                var check = GetChildElement(rule, "check");
                if (check != null)
                {
                    checkContent = GetChildValue(check, "check-content", "check_content") ?? "(none)";
                }
                else
                {
                    // Some variants may have check-content directly under rule
                    checkContent = GetChildValue(rule, "check-content", "check_content") ?? "(none)";
                }

                string Clean(string s) => string.IsNullOrWhiteSpace(s) ? "(none)" : s.Trim();

                string formatted =
                    $"Title: {Clean(title)}{Environment.NewLine}{Environment.NewLine}" +
                    $"Severity: {Clean(severity)}{Environment.NewLine}{Environment.NewLine}" +
                    $"Description:{Environment.NewLine}{Clean(desc)}{Environment.NewLine}{Environment.NewLine}" +
                    $"Rationale:{Environment.NewLine}{Clean(rationale)}{Environment.NewLine}{Environment.NewLine}" +
                    $"Fix:{Environment.NewLine}{Clean(fix)}{Environment.NewLine}{Environment.NewLine}" +
                    $"Check Content:{Environment.NewLine}{Clean(checkContent)}";

                return formatted;
            }
            catch
            {
                return null;
            }
        }
    }
}