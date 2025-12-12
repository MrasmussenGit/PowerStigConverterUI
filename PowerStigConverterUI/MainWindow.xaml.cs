using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml;
using System.Xml.XPath;

namespace PowerStigConverterUI
{
    public partial class MainWindow
    {
        // Normalize IDs so SV-... and SV-..._rule compare correctly; also support V-... and idref/ruleId forms
        private static string NormalizeId(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            var s = raw.Trim();

            // Remove trailing "_rule"
            s = Regex.Replace(s, "_rule$", "", RegexOptions.IgnoreCase);

            // Some XCCDF ids look like SV-123456r987654; drop revision part
            s = Regex.Replace(s, @"(SV-\d+)(r\d+)?$", "$1", RegexOptions.IgnoreCase);

            // Also normalize V-... forms by upper-casing and trimming
            s = Regex.Replace(s, @"^(sv|v)-", m => m.Value.ToUpperInvariant());

            return s;
        }

        private static List<string> ExtractIdsFromXml(string filePath)
        {
            var ids = new List<string>();
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore };
            using var reader = XmlReader.Create(filePath, settings);
            var doc = new XPathDocument(reader);
            var nav = doc.CreateNavigator();

            // Collect common identifier attributes seen in XCCDF and converted outputs
            var xpaths = new[]
            {
                "//@id",
                "//@idref",
                "//@ruleId",
                "//Rule/@id",
                "//Rule/@Id",
                "//@severityId", // fallback in some converted schemas
            };

            foreach (var xp in xpaths)
            {
                var it = nav.Select(xp);
                while (it.MoveNext())
                {
                    var val = it.Current?.Value;
                    if (!string.IsNullOrWhiteSpace(val))
                        ids.Add(NormalizeId(val));
                }
            }

            // Also check element text containing known IDs (e.g., <reference>V-12345</reference>)
            foreach (var xp in new[] { "//reference/text()", "//ident/@id", "//ident/text()" })
            {
                var it = nav.Select(xp);
                while (it.MoveNext())
                {
                    var val = it.Current?.Value;
                    if (!string.IsNullOrWhiteSpace(val))
                        ids.Add(NormalizeId(val));
                }
            }

            return ids.Where(s => s.Length > 0).Distinct().ToList();
        }

        // DISA XCCDF: Rule id="SV-225238r1069480_rule" -> "V-225238"
        private static IEnumerable<string> ExtractDisaRuleIds(string disaFile)
        {
            using var reader = System.Xml.XmlReader.Create(disaFile, new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });
            while (reader.Read())
            {
                if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.Name.Equals("Rule"))
                {
                    var id = reader.GetAttribute("id");
                    if (string.IsNullOrWhiteSpace(id)) continue;

                    var m = Regex.Match(id, @"^SV-(\d+)", RegexOptions.IgnoreCase);
                    if (m.Success)
                        yield return $"V-{m.Groups[1].Value}";
                }
            }
        }

        // PowerSTIG converted: rule IDs are typically "V-225238"
        private static IEnumerable<string> ExtractPsRuleIds(string psFile)
        {
            using var reader = System.Xml.XmlReader.Create(psFile, new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });
            while (reader.Read())
            {
                if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.Name.Equals("Rule"))
                {
                    var id = reader.GetAttribute("id");
                    if (string.IsNullOrWhiteSpace(id)) continue;

                    if (Regex.IsMatch(id, @"^V-\d+$", RegexOptions.IgnoreCase))
                        yield return id.ToUpperInvariant();
                }
            }
        }

        public static List<string> GetMissingIds(string disaFile, string psFile)
        {
            var disaIds = ExtractDisaRuleIds(disaFile).Distinct().ToList();
            var psIds = ExtractPsRuleIds(psFile).Distinct().ToList();
            return disaIds.Except(psIds).OrderBy(x => x).ToList();
        }

        public static List<string> GetAddedIds(string disaFile, string psFile)
        {
            var disaIds = ExtractDisaRuleIds(disaFile).Distinct().ToList();
            var psIds = ExtractPsRuleIds(psFile).Distinct().ToList();
            return psIds.Except(disaIds).OrderBy(x => x).ToList();
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

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}