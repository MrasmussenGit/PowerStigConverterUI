using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PowerStigConverterUI
{
    public static class RuleIdComparer
    {
        private static string NormalizeVId(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return string.Empty;
            var m = Regex.Match(id, @"^(V-\d+)", RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value : id.Trim();
        }

        public static HashSet<string> ExtractRuleIdsFromConverted(string convertedXmlPath)
        {
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var doc = new System.Xml.XmlDocument();
                doc.Load(convertedXmlPath);

                var attrNodes = doc.SelectNodes("//*[@id]");
                if (attrNodes is not null)
                {
                    foreach (System.Xml.XmlNode n in attrNodes)
                    {
                        var id = n.Attributes?["id"]?.Value;
                        if (!string.IsNullOrWhiteSpace(id) && id.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                            ids.Add(NormalizeVId(id));
                    }
                }

                var ruleIdNodes = doc.SelectNodes("//*[local-name()='RuleId' or local-name()='VulnId' or local-name()='BenchmarkId']");
                if (ruleIdNodes is not null)
                {
                    foreach (System.Xml.XmlNode n in ruleIdNodes)
                    {
                        var val = n.InnerText;
                        if (!string.IsNullOrWhiteSpace(val) && val.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                            ids.Add(NormalizeVId(val));
                    }
                }
            }
            catch
            {
                // ignore parse errors
            }
            return ids;
        }

        public static HashSet<string> ExtractRuleIdsFromXccdf(string xccdfPath)
        {
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using var reader = System.Xml.XmlReader.Create(xccdfPath, new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore });
                while (reader.Read())
                {
                    if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                        reader.LocalName.Equals("Rule", StringComparison.OrdinalIgnoreCase))
                    {
                        var id = reader.GetAttribute("id");
                        if (!string.IsNullOrWhiteSpace(id))
                            ids.Add(id.Trim());
                    }
                }
            }
            catch
            {
                foreach (var line in File.ReadLines(xccdfPath))
                {
                    var i = line.IndexOf("id=\"V-", StringComparison.OrdinalIgnoreCase);
                    if (i >= 0)
                    {
                        var start = i + 4;
                        var end = line.IndexOf('"', start);
                        if (end > start)
                        {
                            var id = line.Substring(start, end - start);
                            if (!string.IsNullOrWhiteSpace(id))
                                ids.Add(id.Trim());
                        }
                    }
                }
            }

            var normalized = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var svRegex = new Regex(@"SV-(\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

            foreach (var id in ids)
            {
                if (id.StartsWith("SV-", StringComparison.OrdinalIgnoreCase))
                {
                    var m = svRegex.Match(id);
                    if (m.Success)
                    {
                        normalized.Add($"V-{m.Groups[1].Value}");
                        continue;
                    }
                }
                normalized.Add(NormalizeVId(id));
            }

            return normalized;
        }

        public static List<string> GetMissingRuleIds(string xccdfPath, string convertedXmlPath)
        {
            var xccdfIds = ExtractRuleIdsFromXccdf(xccdfPath);
            var convertedIds = ExtractRuleIdsFromConverted(convertedXmlPath);

            var missing = new HashSet<string>(xccdfIds, StringComparer.OrdinalIgnoreCase);
            missing.ExceptWith(convertedIds);

            return missing
                .Select(id => id.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(id => ExtractNumericKey(id, id.StartsWith("SV-", StringComparison.OrdinalIgnoreCase) ? "SV-" :
                                                    id.StartsWith("V-", StringComparison.OrdinalIgnoreCase) ? "V-" : string.Empty))
                .ThenBy(id => id, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public static int ExtractNumericKey(string id, string prefix)
        {
            if (string.IsNullOrWhiteSpace(id)) return int.MaxValue;
            var s = id.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? id.Substring(prefix.Length) : id;
            int i = 0;
            while (i < s.Length && char.IsDigit(s[i])) i++;
            var digits = (i > 0) ? s.Substring(0, i) : string.Empty;
            return int.TryParse(digits, out var n) ? n : int.MaxValue;
        }
    }
}