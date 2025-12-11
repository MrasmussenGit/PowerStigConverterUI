using PowerStigConverterUI;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls; // Ensure this is present
using System.Xml;
using System.Xml.Linq;

public partial class MainWindow : Window
{
    public System.Windows.Controls.ListView MissingListView;
    public System.Windows.Controls.ListView AddedListView;
    public TextBlock MissingCountTextBlock;
    public TextBlock AddedCountTextBlock;

    // Wire up double-click handlers (call this after InitializeComponent in your other MainWindow constructor)
    public void InitializeListViewHandlers()
    {
        if (MissingListView != null)
            MissingListView.MouseDoubleClick += ListView_MouseDoubleClick;

        if (AddedListView != null)
            AddedListView.MouseDoubleClick += ListView_MouseDoubleClick;
    }

    private void ListView_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is System.Windows.Controls.ListView lv && lv.SelectedItem is string id && !string.IsNullOrWhiteSpace(id))
        {
            // Prompt for the DISA XML file
            var disaDlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select DISA XML",
                Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
            };
            if (disaDlg.ShowDialog() != true) return;
            string disaPath = disaDlg.FileName;

            // Prompt for the PowerSTIG XML file
            var psDlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select PowerSTIG XML",
                Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
            };
            if (psDlg.ShowDialog() != true) return;
            string psPath = psDlg.FileName;

            var win = new RuleDetailWindow(id, disaPath, psPath) { Owner = this };
            win.ShowDialog();
        }
    }

    static string NormalizeDisaId(string? rawId)
    {
        if (string.IsNullOrWhiteSpace(rawId)) return string.Empty;
        var id = rawId.Trim();
        var vMatch = Regex.Match(id, @"\bV-(\d+)", RegexOptions.IgnoreCase);
        if (vMatch.Success)
            return $"V-{vMatch.Groups[1].Value}";
        var svMatch = Regex.Match(id, @"\bSV-(\d+)", RegexOptions.IgnoreCase);
        if (svMatch.Success)
            return $"V-{svMatch.Groups[1].Value}";
        return id;
    }

    static string NormalizePowerStigId(string? rawId)
    {
        if (string.IsNullOrWhiteSpace(rawId)) return string.Empty;
        var id = rawId.Trim();
        // Normalize "V-123456.a" -> "V-123456"
        var m = Regex.Match(id, @"\b(V-\d+)", RegexOptions.IgnoreCase);
        return m.Success ? m.Groups[1].Value : id;
    }

    public static List<string> GetMissingIds(string disaPath, string psPath)
    {
        var disa = ExtractDisaRuleIds(disaPath);
        var psRaw = ExtractPowerStigRuleIds(psPath);

        // Treat suffix variants in PowerSTIG as matching the base DISA rule
        var psNormalized = new HashSet<string>(psRaw.Select(NormalizePowerStigId), StringComparer.OrdinalIgnoreCase);

        return disa
            .Where(d => !psNormalized.Contains(NormalizePowerStigId(d)))
            .OrderBy(x => x)
            .ToList();
    }

    public static List<string> GetAddedIds(string disaPath, string psPath)
    {
        var disa = ExtractDisaRuleIds(disaPath);
        var psRaw = ExtractPowerStigRuleIds(psPath);

        // If a PowerSTIG id is a suffix variant of a DISA base id, it's not "added"
        var disaNormalized = new HashSet<string>(disa.Select(NormalizePowerStigId), StringComparer.OrdinalIgnoreCase);

        return psRaw
            .Where(p => !disaNormalized.Contains(NormalizePowerStigId(p)))
            .OrderBy(x => x)
            .ToList();
    }

    static HashSet<string> ExtractDisaRuleIds(string path)
    {
        using var reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore });
        var doc = XDocument.Load(reader);
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var groupRuleIds =
            doc.Descendants().Where(x => x.Name.LocalName == "Group")
               .Descendants().Where(x => x.Name.LocalName == "Rule")
               .Select(x => NormalizeDisaId((string?)x.Attribute("id")))
               .Where(id => !string.IsNullOrWhiteSpace(id))!;
        ids.UnionWith(groupRuleIds);

        if (ids.Count == 0)
        {
            var anyRuleIds =
                doc.Descendants().Where(x => x.Name.LocalName == "Rule")
                   .Select(x => NormalizeDisaId((string?)x.Attribute("id")))
                   .Where(id => !string.IsNullOrWhiteSpace(id))!;
            ids.UnionWith(anyRuleIds);
        }
        return ids;
    }

    static HashSet<string> ExtractPowerStigRuleIds(string path)
    {
        using var reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore });
        var doc = XDocument.Load(reader);
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var root = doc.Elements().FirstOrDefault(e => e.Name.LocalName.Equals("DISASTIG", StringComparison.OrdinalIgnoreCase));
        if (root == null) return ids;

        foreach (var ruleType in root.Elements())
            foreach (var rule in ruleType.Elements().Where(e => e.Name.LocalName.Equals("Rule", StringComparison.OrdinalIgnoreCase)))
            {
                var id = rule.Attribute("Id")?.Value ?? rule.Attribute("id")?.Value;
                if (!string.IsNullOrWhiteSpace(id)) ids.Add(id.Trim());
            }
        return ids;
    }

    void CompareButton_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new CompareWindow();
        dlg.Owner = this;
        dlg.ShowDialog();
    }
}