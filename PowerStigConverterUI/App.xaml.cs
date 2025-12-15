using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;

namespace PowerStigConverterUI
{
    public partial class App : System.Windows.Application
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AllocConsole();

        protected override void OnStartup(StartupEventArgs e)
        {
            // Command line mode: two folder args => compare all matched pairs and print to console, then exit
            if (e.Args.Length == 2 && Directory.Exists(e.Args[0]) && Directory.Exists(e.Args[1]))
            {
                // Ensure a console exists when running as a GUI app
                AllocConsole();
                Console.SetOut(new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true });

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

        // Update NormalizeDisaKey to produce keys that match your PowerSTIG filenames exactly.
        private static string NormalizeDisaKey(string name)
        {
            // Examples:
            // U_MS_DotNet_Framework_4-0_STIG_V2R7_Manual-xccdf.xml -> dotnetframework-4-2.7
            // U_MS_IIS_10-0_Server_STIG_V3R5_Manual-xccdf.xml      -> iisserver-10.0-3.5
            // U_MS_IIS_10-0_Site_STIG_V2R13_Manual-xccdf.xml       -> iissite-10.0-2.3
            try
            {
                var lower = name.ToLowerInvariant();

                var productRaw = ExtractBetween(lower, "u_ms_", "_stig");
                if (string.IsNullOrWhiteSpace(productRaw)) return string.Empty;

                // Normalize product base
                var product = productRaw
                    .Replace("dotnet_framework", "dotnetframework")
                    .Replace("iis_server", "iisserver")
                    .Replace("iis_site", "iissite")
                    .Trim('_');

                // Split remaining tokens and normalize version-like subparts
                var tokens = product.Split('_', StringSplitOptions.RemoveEmptyEntries).ToList();

                // Handle embedded product version tokens:
                // - "4-0" => "4"
                // - "10-0" => "10.0"
                for (int i = 0; i < tokens.Count; i++)
                {
                    if (tokens[i].Equals("4-0", StringComparison.Ordinal)) tokens[i] = "4";
                    else if (tokens[i].Equals("10-0", StringComparison.Ordinal)) tokens[i] = "10.0";
                }

                // Build product segment using '-' as separator
                var productSegment = string.Join("-", tokens);

                // Extract V token (e.g., "v2r7" => "2.7", "v3r5" => "3.5", "v2r13" => "2.13")
                var vToken = ExtractBetween(lower, "_v", "_manual") ?? ExtractBetween(lower, "_v", "_xccdf");
                var version = ToPowerStigVersion(vToken);
                if (string.IsNullOrEmpty(version)) return string.Empty;

                return $"{productSegment}-{version}";
            }
            catch
            {
                return string.Empty;
            }
        }

        // Keep NormalizePsKey simple: filename without extension already contains the desired key.
        // Do not strip by '.' because versions include dots (e.g., "10.0", "2.7").
        private static string NormalizePsKey(string name)
        {
            var lower = name.ToLowerInvariant();
            if (lower.Contains(".org.default")) return string.Empty;

            // Trim trailing qualifiers after version if present, keep first three tokens
            var parts = lower.Split('-', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 3)
                return $"{parts[0]}-{parts[1]}-{parts[2]}";

            return lower;
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