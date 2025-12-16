using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace PowerStigConverterUI
{
    public sealed class RuleIdCompareResult
    {
        public IReadOnlyList<string> MissingBaseIds { get; init; } = Array.Empty<string>();
        public IReadOnlyList<string> MatchedBaseIds { get; init; } = Array.Empty<string>();
        public IReadOnlyList<string> AddedIds { get; init; } = Array.Empty<string>();
    }

    public static class RuleIdAnalysis
    {
        private static readonly Regex SvDigits = new(@"^SV-(\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex VDigits = new(@"^V-(\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex VariantRegex = new(@"^V-\d+\.[A-Za-z0-9]+$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static RuleIdCompareResult Compare(string disaXccdfPath, string convertedXmlPath)
        {
            var disaBase = ExtractDisaBaseIds(disaXccdfPath);
            var convertedRaw = ExtractConvertedRawIds(convertedXmlPath);
            var convertedBase = new HashSet<string>(convertedRaw.Select(NormalizeToBaseV).Where(s => !string.IsNullOrWhiteSpace(s)),
                                                    StringComparer.OrdinalIgnoreCase);

            // Missing: in DISA base but not in converted base
            var missing = disaBase
                .Where(d => !convertedBase.Contains(d))
                .OrderBy(x => ExtractNumericKey(x, "V-"))
                .ThenBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Matched: in both bases
            var matched = disaBase
                .Where(d => convertedBase.Contains(d))
                .OrderBy(x => ExtractNumericKey(x, "V-"))
                .ThenBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Added: include variants and true extras (bases not present in DISA)
            var variants = convertedRaw.Where(id => VariantRegex.IsMatch(id));
            var extrasNotInDisa = convertedRaw.Where(id => !disaBase.Contains(NormalizeToBaseV(id)));

            var added = variants
                .Concat(extrasNotInDisa)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => ExtractNumericKey(x, "V-"))
                .ThenBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new RuleIdCompareResult
            {
                MissingBaseIds = missing,
                MatchedBaseIds = matched,
                AddedIds = added
            };
        }

        public static HashSet<string> ExtractDisaBaseIds(string xccdfPath)
        {
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using var reader = XmlReader.Create(xccdfPath, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore });
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName.Equals("Rule", StringComparison.OrdinalIgnoreCase))
                {
                    var id = reader.GetAttribute("id");
                    if (string.IsNullOrWhiteSpace(id)) continue;
                    var baseV = NormalizeToBaseV(id.Trim());
                    if (!string.IsNullOrWhiteSpace(baseV)) ids.Add(baseV);
                }
            }
            return ids;
        }

        public static HashSet<string> ExtractConvertedRawIds(string convertedXmlPath)
        {
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var doc = new System.Xml.XmlDocument();
            doc.Load(convertedXmlPath);

            var attrNodes = doc.SelectNodes("//*[@id]");
            if (attrNodes is not null)
            {
                foreach (System.Xml.XmlNode n in attrNodes)
                {
                    var id = n.Attributes?["id"]?.Value?.Trim();
                    if (!string.IsNullOrWhiteSpace(id) && id.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                        ids.Add(id);
                }
            }

            var ruleIdNodes = doc.SelectNodes("//*[local-name()='RuleId' or local-name()='VulnId' or local-name()='BenchmarkId']");
            if (ruleIdNodes is not null)
            {
                foreach (System.Xml.XmlNode n in ruleIdNodes)
                {
                    var val = n.InnerText?.Trim();
                    if (!string.IsNullOrWhiteSpace(val) && val.StartsWith("V-", StringComparison.OrdinalIgnoreCase))
                        ids.Add(val);
                }
            }

            return ids;
        }

        public static string NormalizeToBaseV(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            var s = raw.Trim();

            var sv = SvDigits.Match(s);
            if (sv.Success) return $"V-{sv.Groups[1].Value}";

            var v = VDigits.Match(s);
            if (v.Success) return $"V-{v.Groups[1].Value}";

            return string.Empty;
        }

        public static int ExtractNumericKey(string id, string prefix)
        {
            if (string.IsNullOrWhiteSpace(id)) return int.MaxValue;
            var s = id.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? id[prefix.Length..] : id;
            int i = 0;
            while (i < s.Length && char.IsDigit(s[i])) i++;
            var digits = (i > 0) ? s[..i] : string.Empty;
            return int.TryParse(digits, out var n) ? n : int.MaxValue;
        }
    }
}