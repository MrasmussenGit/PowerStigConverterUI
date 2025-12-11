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

        public static List<string> GetMissingIds(string disaFile, string psFile)
        {
            var disaIds = ExtractIdsFromXml(disaFile); // will include SV- and V- normalized
            var psIds = ExtractIdsFromXml(psFile);   // include whatever the converted file uses
            return disaIds.Except(psIds).ToList();
        }

        public static List<string> GetAddedIds(string disaFile, string psFile)
        {
            var disaIds = ExtractIdsFromXml(disaFile);
            var psIds = ExtractIdsFromXml(psFile);
            return psIds.Except(disaIds).ToList();
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