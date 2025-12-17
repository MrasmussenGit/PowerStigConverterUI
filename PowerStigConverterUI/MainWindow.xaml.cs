using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Reflection;
using System.Diagnostics;

namespace PowerStigConverterUI
{
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                VersionTextBlock.Text = $"Version: {GetAppVersion()}";
            }
            catch
            {
                // Ignore version display issues; keep default text
            }
        }

        private static string GetAppVersion()
        {
            var asm = Assembly.GetEntryAssembly() ?? typeof(MainWindow).Assembly;

            // Prefer informational version (may include git hash/prerelease metadata)
            var infoAttr = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
            var candidate =
                (!string.IsNullOrWhiteSpace(infoAttr?.InformationalVersion) ? infoAttr!.InformationalVersion :
                (!string.IsNullOrWhiteSpace(FileVersionInfo.GetVersionInfo(asm.Location).ProductVersion) ? FileVersionInfo.GetVersionInfo(asm.Location).ProductVersion :
                asm.GetName().Version?.ToString())) ?? string.Empty;

            // Extract only the numeric dotted version (e.g., 1.2.3.4) and drop prerelease/build metadata
            var m = Regex.Match(candidate, @"\d+(\.\d+){1,3}");
            if (m.Success)
                return m.Value;

            // Fallback: if nothing matched, return the raw candidate or Unknown
            return string.IsNullOrWhiteSpace(candidate) ? "Unknown" : candidate;
        }

        // Normalize to the base numeric V-ID:
        // - "SV-254270r958480_rule" -> "V-254270"
        // - "V-254252.a"            -> "V-254252"
        // - "V-254254.c-extra"      -> "V-254254"
        private static string NormalizeToBaseV(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            var s = raw.Trim();

            // DISA XCCDF formats: SV-<num>r..._rule
            var sv = Regex.Match(s, @"^SV-(\d+)", RegexOptions.IgnoreCase);
            if (sv.Success) return $"V-{sv.Groups[1].Value}";

            // Converted formats: V-<num>... (accept any suffix after the number)
            var v = Regex.Match(s, @"^V-(\d+)", RegexOptions.IgnoreCase);
            if (v.Success) return $"V-{v.Groups[1].Value}";

            return string.Empty;
        }

        // DISA: only read Rule/@id (namespaced safe via LocalName)
        private static IEnumerable<string> ExtractDisaRuleIds(string disaFile)
        {
            using var reader = System.Xml.XmlReader.Create(disaFile, new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });
            while (reader.Read())
            {
                if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.LocalName == "Rule")
                {
                    var id = reader.GetAttribute("id");
                    var norm = NormalizeToBaseV(id);
                    if (!string.IsNullOrEmpty(norm)) yield return norm;
                }
            }
        }

        // Converted: read Rule/@id and normalize
        private static IEnumerable<string> ExtractPsRuleIds(string psFile)
        {
            using var reader = System.Xml.XmlReader.Create(psFile, new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });
            while (reader.Read())
            {
                if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.LocalName == "Rule")
                {
                    var id = reader.GetAttribute("id");
                    var norm = NormalizeToBaseV(id);
                    if (!string.IsNullOrEmpty(norm)) yield return norm;
                }
            }
        }

        public static List<string> GetMissingIdsProper(HashSet<string> disaBase, HashSet<string> psBase)
        {
            List<string> missing = new();
            foreach (var d in disaBase)
            {
                bool found = false;
                foreach (var p in psBase)
                {
                    if (string.Equals(d, p, System.StringComparison.OrdinalIgnoreCase))
                    {
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    missing.Add(d);
                }
            }

            missing.Sort(System.StringComparer.OrdinalIgnoreCase);
            return missing;
        }

        public static List<string> GetMissingIds(string disaFile, string psFile)
        {
            var disaBase = new HashSet<string>(
                ExtractDisaRuleIds(disaFile)
                    .Select(NormalizeToBaseV)
                    .Where(id => !string.IsNullOrEmpty(id)),
                System.StringComparer.OrdinalIgnoreCase);

            var psBase = new HashSet<string>(
                ExtractPsRuleIds(psFile)
                    .Select(NormalizeToBaseV)
                    .Where(id => !string.IsNullOrEmpty(id)),
                System.StringComparer.OrdinalIgnoreCase);

            // FIX: Call the correct overload that accepts HashSet<string> arguments
            List<string> ids = GetMissingIdsProper(disaBase, psBase);

            return disaBase
                .Except(psBase, System.StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();
        }

        public static List<string> GetAddedIds(string disaFile, string psFile)
        {
            var disaBase = new HashSet<string>(
                ExtractDisaRuleIds(disaFile)
                    .Select(NormalizeToBaseV)
                    .Where(id => !string.IsNullOrEmpty(id)),
                System.StringComparer.OrdinalIgnoreCase);

            var psBase = new HashSet<string>(
                ExtractPsRuleIds(psFile)
                    .Select(NormalizeToBaseV)
                    .Where(id => !string.IsNullOrEmpty(id)),
                System.StringComparer.OrdinalIgnoreCase);

            return psBase
                .Except(disaBase, System.StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();
        }

        private void ConvertStigButton_Click(object sender, RoutedEventArgs e)
        {
            var win = new ConvertStigWindow { Owner = this };
            win.ShowDialog();
        }

        private void CompareButton_Click(object sender, RoutedEventArgs e)
        {
            var win = new CompareWindow { Owner = this };
            win.ShowDialog();
        }

        private void SplitOsStigButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var win = new SplitStigWindow { Owner = this };
            win.ShowDialog();
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}