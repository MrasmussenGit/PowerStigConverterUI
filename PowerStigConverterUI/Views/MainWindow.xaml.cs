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
            var win = new RuleDetailWindow(id) { Owner = this };
            win.ShowDialog();
        }
    }

    public static List<string>  GetMissingIds(string disaPath, string psPath)
    {
        var disa = ExtractDisaRuleIds(disaPath);
        var ps = ExtractPowerStigRuleIds(psPath);
        return disa.Except(ps, StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToList();
    }

    public static List<string> GetAddedIds(string disaPath, string psPath)
    {
        var disa = ExtractDisaRuleIds(disaPath);
        var ps = ExtractPowerStigRuleIds(psPath);
        return ps.Except(disa, StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToList();
    }

    void CompareButton_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new CompareWindow();
        dlg.Owner = this;
        dlg.ShowDialog();
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
                if (!string.IsNullOrWhiteSpace(id)) ids.Add(id);
            }
        return ids;
    }
}