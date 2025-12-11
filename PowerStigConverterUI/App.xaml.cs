using System;
using System.IO;
using System.Linq;
using System.Windows;

namespace PowerStigConverterUI
{
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            // Command line mode: two folder args => compare all matched pairs and print to console, then exit
            if (e.Args.Length == 2 && Directory.Exists(e.Args[0]) && Directory.Exists(e.Args[1]))
            {
                RunFolderCompare(e.Args[0], e.Args[1]);
                Shutdown();
                return;
            }

            // Normal WPF startup
            base.OnStartup(e);
        }

        private static void RunFolderCompare(string disaFolder, string psFolder)
        {
            // Collect candidate files
            var disaFiles = Directory.EnumerateFiles(disaFolder, "*.xml", SearchOption.AllDirectories)
                                     .Where(f => Path.GetFileName(f).Contains("_STIG_", StringComparison.OrdinalIgnoreCase))
                                     .ToList();

            var psFiles = Directory.EnumerateFiles(psFolder, "*.xml", SearchOption.AllDirectories)
                                   .Where(f =>
                                       !Path.GetFileName(f).Contains(".org.default", StringComparison.OrdinalIgnoreCase) && // ignore org.default
                                       !Path.GetFileName(f).EndsWith(".org.default.xml", StringComparison.OrdinalIgnoreCase))
                                   .ToList();

            // Build quick lookup for PowerSTIG by normalized key
            var psByKey = psFiles
                .Select(f => (file: f, key: NormalizePsKey(Path.GetFileNameWithoutExtension(f))))
                .Where(x => !string.IsNullOrEmpty(x.key))
                .GroupBy(x => x.key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.Select(x => x.file).ToList(), StringComparer.OrdinalIgnoreCase);

            int compared = 0;
            foreach (var disa in disaFiles)
            {
                var disaKey = NormalizeDisaKey(Path.GetFileNameWithoutExtension(disa));
                if (string.IsNullOrEmpty(disaKey)) continue;

                if (!psByKey.TryGetValue(disaKey, out var psMatches) || psMatches.Count == 0)
                {
                    Console.WriteLine($"[WARN] No PowerSTIG match for DISA '{Path.GetFileName(disa)}' (key '{disaKey}').");
                    continue;
                }

                // Pick the first match (could extend to compare all)
                var ps = psMatches[0];
                compared++;

                var missing = PowerStigConverterUI.MainWindow.GetMissingIds(disa, ps);
                var added = PowerStigConverterUI.MainWindow.GetAddedIds(disa, ps);

                Console.WriteLine($"=== Compare ===");
                Console.WriteLine($"DISA: {disa}");
                Console.WriteLine($"PowerSTIG: {ps}");
                Console.WriteLine($"Missing ({missing.Count}):");
                foreach (var id in missing) Console.WriteLine($"  {id}");
                Console.WriteLine($"Added ({added.Count}):");
                foreach (var id in added) Console.WriteLine($"  {id}");
                Console.WriteLine();
            }

            Console.WriteLine($"Completed. Compared {compared} DISA file(s).");
        }

        // DISA: U_MS_DotNet_Framework_4-0_STIG_V2R7_Manual-xccdf.xml -> key "dotnetframework-4-2.7"
        private static string NormalizeDisaKey(string name)
        {
            // Extract product and version tokens
            // Product: "DotNet_Framework_4-0" -> "dotnetframework-4-0"
            // Version: "V2R7" -> "2.7"
            try
            {
                var lower = name.ToLowerInvariant();

                // Get product segment between "u_ms_" and "_stig"
                var product = ExtractBetween(lower, "u_ms_", "_stig")?
                              .Replace("_", string.Empty)
                              .Replace("framework", "framework") // keep literal
                              .Replace("-manual-xccdf", string.Empty)
                              .Trim();

                if (string.IsNullOrEmpty(product)) return string.Empty;

                // Ensure product looks like "dotnetframework-4-0" or "iis"
                product = product.Replace("dotnet_framework", "dotnetframework");

                // Version: "v2r7" -> "2.7"
                var vToken = ExtractBetween(lower, "_v", "_manual") ?? ExtractBetween(lower, "_v", "_xccdf");
                var version = ToPowerStigVersion(vToken);

                if (string.IsNullOrEmpty(version)) return string.Empty;

                return $"{product}-{version}";
            }
            catch
            {
                return string.Empty;
            }
        }

        // PowerSTIG: DotNetFramework-4-2.7.xml -> key "dotnetframework-4-2.7"
        private static string NormalizePsKey(string name)
        {
            var lower = name.ToLowerInvariant();
            // Drop trailing environment or profile qualifiers if present (e.g., -windows11, -server)
            // Ignore org.default
            if (lower.Contains(".org.default")) return string.Empty;

            return lower; // already in desired form
        }

        private static string? ExtractBetween(string s, string start, string end)
        {
            var i = s.IndexOf(start, StringComparison.Ordinal);
            if (i < 0) return null;
            i += start.Length;
            var j = s.IndexOf(end, i, StringComparison.Ordinal);
            if (j < 0) return null;
            return s.Substring(i, j - i);
        }

        private static string ToPowerStigVersion(string? vToken)
        {
            if (string.IsNullOrWhiteSpace(vToken)) return string.Empty;
            // v2r7 or 2r7 -> 2.7
            var digits = new string(vToken.Where(char.IsDigit).ToArray());
            if (string.IsNullOrEmpty(digits)) return string.Empty;
            // Expect two numbers: major and minor
            if (digits.Length == 2) return $"{digits[0]}.{digits[1]}";
            // Fallback: if "20" and "7" etc.
            if (digits.Length > 2) return $"{digits[0]}.{digits[^1]}";
            return string.Empty;
        }
    }
}