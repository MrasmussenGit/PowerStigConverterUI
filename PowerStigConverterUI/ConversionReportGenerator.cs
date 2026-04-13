using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace PowerStigConverterUI
{
    public static class ConversionReportGenerator
    {
        public static string GenerateHtmlReport(ConversionReportData data)
        {
            var html = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""utf-8"">
    <title>{EscapeHtml(data.StigName)} - Conversion Report</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, sans-serif; margin: 20px; background: #f5f5f5; }}
        .container {{ max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        h1 {{ color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }}
        .timestamp {{ color: #666; font-size: 0.9em; margin-bottom: 20px; }}
        
        /* Summary Cards */
        .summary {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px; margin: 20px 0; }}
        .summary-row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin: 20px 0; }}
        .summary-card {{ padding: 20px; border-radius: 8px; border-left: 4px solid; text-align: center; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; }}
        .summary-card:hover {{ transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,0,0,0.15); }}
        .summary-card[style*=""cursor: default""] {{ cursor: default !important; }}
        .summary-card[style*=""cursor: default""]:hover {{ transform: none; box-shadow: none; }}
        .summary-card-large {{ padding: 30px; border-radius: 8px; border-left: 6px solid; text-align: center; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; }}
        .summary-card-large:hover {{ transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.2); }}
        .summary-card.success {{ background: #e7f3e7; border-left-color: #107c10; }}
        .summary-card.error {{ background: #fce8e8; border-left-color: #d13438; }}
        .summary-card.warning {{ background: #fff4e5; border-left-color: #ff8c00; }}
        .summary-card.info {{ background: #e7f3ff; border-left-color: #0078d4; }}
        .summary-card-large.success {{ background: #e7f3e7; border-left-color: #107c10; }}
        .summary-card-large.error {{ background: #fce8e8; border-left-color: #d13438; }}
        .summary-card-large.warning {{ background: #fff4e5; border-left-color: #ff8c00; }}
        .summary-card-number {{ font-size: 2em; font-weight: bold; margin-bottom: 5px; }}
        .summary-card-large .summary-card-number {{ font-size: 3em; }}
        .summary-card-label {{ font-size: 0.9em; color: #666; }}
        .summary-card-large .summary-card-label {{ font-size: 1.1em; font-weight: 600; }}
        
        /* Collapsible Sections */
        .section {{ margin: 25px 0; border: 1px solid #ddd; border-radius: 8px; overflow: hidden; }}
        .section-header {{ background: #f8f8f8; padding: 15px 20px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; }}
        .section-header:hover {{ background: #efefef; }}
        .section-header.error {{ border-left: 4px solid #d13438; }}
        .section-header.warning {{ border-left: 4px solid #ff8c00; }}
        .section-header.success {{ border-left: 4px solid #107c10; }}
        .section-header.info {{ border-left: 4px solid #0078d4; }}
        .section-title {{ font-weight: 600; font-size: 1.1em; }}
        .section-count {{ background: rgba(0,0,0,0.1); padding: 2px 10px; border-radius: 12px; font-size: 0.9em; }}
        .toggle-icon {{ transition: transform 0.3s; }}
        .section-header.active .toggle-icon {{ transform: rotate(90deg); }}
        .section-content {{ max-height: 0; overflow: hidden; transition: max-height 0.3s ease; }}
        .section-content.active {{ max-height: 10000px; padding: 20px; }}
        
        /* Rule Items */
        .rule-item {{ border-bottom: 1px solid #eee; padding: 12px 0; }}
        .rule-item:last-child {{ border-bottom: none; }}
        .rule-header {{ display: flex; justify-content: space-between; align-items: center; cursor: pointer; padding: 8px 12px; border-radius: 4px; }}
        .rule-header:hover {{ background: #f8f8f8; }}
        .rule-id {{ font-family: 'Consolas', monospace; font-weight: 600; color: #0078d4; }}
        .rule-variant-count {{ font-size: 0.85em; color: #666; margin-left: 8px; }}
        .rule-title {{ color: #666; font-size: 0.9em; font-weight: normal; margin-left: 8px; }}
        .rule-details {{ display: none; padding: 15px; margin-top: 8px; background: #f9f9f9; border-radius: 4px; border-left: 3px solid #0078d4; }}
        .rule-details.active {{ display: block; }}
        .detail-row {{ margin: 8px 0; }}
        .detail-label {{ font-weight: 600; color: #333; margin-right: 8px; }}
        .detail-value {{ color: #666; font-family: 'Consolas', monospace; font-size: 0.9em; }}
        .error-message {{ background: #fff0f0; padding: 10px; border-radius: 4px; margin-bottom: 15px; color: #d13438; font-family: 'Consolas', monospace; font-size: 0.85em; }}
        .variant-list {{ margin-top: 8px; padding-left: 20px; color: #666; font-size: 0.9em; }}
        
        /* Tabs */
        .tabs {{ margin-top: 10px; }}
        .tab-buttons {{ display: flex; gap: 5px; border-bottom: 2px solid #ddd; margin-bottom: 10px; }}
        .tab-button {{ background: none; border: none; padding: 10px 20px; cursor: pointer; font-size: 0.95em; color: #666; border-bottom: 2px solid transparent; margin-bottom: -2px; transition: all 0.2s; }}
        .tab-button:hover {{ background: #f5f5f5; color: #0078d4; }}
        .tab-button.active {{ color: #0078d4; border-bottom-color: #0078d4; font-weight: 600; }}
        .tab-content {{ display: none; padding: 15px 10px; }}
        .tab-content.active {{ display: block; }}
        .detail-text {{ color: #333; margin-top: 5px; }}
        .detail-text-block {{ background: #f9f9f9; padding: 15px; border-radius: 4px; border-left: 3px solid #0078d4; white-space: pre-wrap; word-wrap: break-word; font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 0.9em; margin: 0; line-height: 1.6; }}
        .code-block {{ font-family: 'Consolas', 'Courier New', monospace; background: #1e1e1e; color: #d4d4d4; }}
        
        /* Severity Colors */
        .severity-high {{ color: #d13438; font-weight: 600; }}
        .severity-medium {{ color: #ff8c00; font-weight: 600; }}
        .severity-low {{ color: #107c10; font-weight: 600; }}
        
        .expand-all {{ float: right; margin-bottom: 10px; padding: 8px 16px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer; }}
        .expand-all:hover {{ background: #005a9e; }}
    </style>
    <script>
        function toggleSection(element) {{
            element.classList.toggle('active');
            element.nextElementSibling.classList.toggle('active');
        }}
        
        function toggleRule(element) {{
            const details = element.nextElementSibling;
            details.classList.toggle('active');
        }}
        
        function expandAll(sectionId) {{
            const section = document.getElementById(sectionId);
            const rules = section.querySelectorAll('.rule-details');
            const allExpanded = Array.from(rules).every(r => r.classList.contains('active'));
            
            rules.forEach(r => {{
                if (allExpanded) r.classList.remove('active');
                else r.classList.add('active');
            }});
        }}
        
        function switchTab(event, tabId) {{
            event.stopPropagation();
            const button = event.target;
            const tabButtons = button.parentElement;
            const tabsContainer = tabButtons.parentElement;

            // Deactivate all buttons and contents in this rule's tabs container
            tabButtons.querySelectorAll('.tab-button').forEach(b => b.classList.remove('active'));
            tabsContainer.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

            // Activate clicked button and corresponding content
            button.classList.add('active');
            const tabContent = document.getElementById(tabId);
            if (tabContent) {{
                tabContent.classList.add('active');
            }}
        }}
        
        function jumpToSection(sectionId) {{
            const section = document.getElementById(sectionId);
            if (!section) return;
            
            const header = section.previousElementSibling;
            if (!header) return;
            
            // Check if section is already expanded
            const isExpanded = header.classList.contains('active');
            
            if (isExpanded) {{
                // Collapse if already expanded
                header.classList.remove('active');
                section.classList.remove('active');
            }} else {{
                // Expand and scroll to it
                header.classList.add('active');
                section.classList.add('active');
                
                // Smooth scroll to the section
                header.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
            }}
        }}
    </script>
</head>
<body>
    <div class=""container"">
        <h1>{EscapeHtml(data.StigName)}</h1>
        <p class=""timestamp"">Generated: {data.Timestamp:yyyy-MM-dd HH:mm:ss}</p>

        <!-- Coverage Cards - Two column layout -->
        {(data.TotalDisaRules > 0 ? $@"
        <div class=""summary-row"">
            <div class=""summary-card-large {(data.CoveredDisaRules == data.TotalDisaRules ? "success" : data.CoveredDisaRules >= data.TotalDisaRules * 0.8 ? "warning" : "error")}"" style=""cursor: default;"">
                <div style=""display: inline-block; text-align: left;"">
                    <div style=""font-size: 1.5em; font-weight: 600; color: #333; margin-bottom: 8px;"">
                        {data.TotalDisaRules} Total Rules in Current STIG
                    </div>
                    <div style=""font-size: 1.5em; font-weight: 600; color: #333; margin-bottom: 12px;"">
                        {data.CoveredDisaRules} Successfully Converted ({(data.TotalDisaRules > 0 ? (data.CoveredDisaRules * 100.0 / data.TotalDisaRules).ToString("F1") : "0")}%)
                    </div>
                    <div style=""font-size: 0.9em; color: #666; margin-top: 5px;"">
                        {data.IndividualDISARulesAutomated} Automated - No user intervention
                    </div>
                    <div style=""font-size: 0.9em; color: #666; margin-top: 5px;"">
                        {data.NoDscResourceRules.Count + data.HardCodedRules.Count} Manual - User intervention required
                    </div>
                </div>
            </div>
            <div class=""summary-card-large {(data.FailedCount + data.SkippedRules.Count == 0 ? "success" : "error")}"" onclick=""{(data.FailedCount > 0 ? "jumpToSection('failed-section')" : data.SkippedRules.Count > 0 ? "jumpToSection('skipped-section')" : "")}"" style=""{(data.FailedCount + data.SkippedRules.Count == 0 ? "cursor: default;" : "")}"">
                <div style=""display: inline-block; text-align: left;"">
                    <div style=""font-size: 1.5em; font-weight: 600; color: #333; margin-bottom: 8px;"">{data.FailedCount + data.SkippedRules.Count} Missing DISA Rules ({(data.TotalDisaRules > 0 ? ((data.FailedCount + data.SkippedRules.Count) * 100.0 / data.TotalDisaRules).ToString("F1") : "0")}%)</div>
                    <div style=""font-size: 0.9em; color: #666; margin-top: 10px;"">
                        {(data.FailedCount > 0 && data.SkippedRules.Count > 0 ? $"{data.FailedCount} Failed + {data.SkippedRules.Count} Skipped" : 
                          data.FailedCount > 0 ? $"{data.FailedCount} Failed conversions" :
                          data.SkippedRules.Count > 0 ? $"{data.SkippedRules.Count} Skipped (from log file)" :
                          "All rules converted!")}
                    </div>
                    <div style=""font-size: 0.8em; color: #999; margin-top: 8px; font-style: italic;"">
                        {(data.FailedCount + data.SkippedRules.Count > 0 ? "Click to view details below" : "Not counted in coverage above")}
                    </div>
                </div>
            </div>
        </div>" : "")}

        <div class=""summary"">
            <div class=""summary-card success"" onclick=""jumpToSection('success-section')"">
                <div style=""display: inline-block; text-align: left;"">
                    <div style=""font-size: 1.3em; font-weight: 600; color: #333; margin-bottom: 6px;"">{data.IndividualDISARulesAutomated} Automated rules ({(data.TotalDisaRules > 0 ? (data.IndividualDISARulesAutomated * 100.0 / data.TotalDisaRules).ToString("F1") : "0")}%)</div>
                    <div style=""font-size: 0.85em; color: #666; margin-top: 5px;"">{data.SuccessfulRules.Count} Rules created which includes variants</div>
                    <div style=""font-size: 0.75em; color: #999; margin-top: 5px; font-style: italic;"">Click to view details</div>
                </div>
            </div>
            <div class=""summary-card warning"" onclick=""jumpToSection('nodsc-section')"">
                <div style=""display: inline-block; text-align: left;"">
                    <div style=""font-size: 1.3em; font-weight: 600; color: #333; margin-bottom: 6px;"">{data.NoDscResourceRules.Count + data.HardCodedRules.Count} Manual Intervention Required ({(data.TotalDisaRules > 0 ? ((data.NoDscResourceRules.Count + data.HardCodedRules.Count) * 100.0 / data.TotalDisaRules).ToString("F1") : "0")}%)</div>
                    <div style=""font-size: 0.85em; color: #666; margin-top: 5px;"">
                        {((data.NoDscResourceRules.Count + data.HardCodedRules.Count) > 0 
                            ? (data.NoDscResourceRules.Count > 0 ? $"{data.NoDscResourceRules.Count} No DSC Resource" : "") +
                              (data.HardCodedRules.Count > 0 ? (data.NoDscResourceRules.Count > 0 ? " • " : "") + $"{data.HardCodedRules.Count} Hard Coded" : "")
                            : "All rules automated")}
                    </div>
                    <div style=""font-size: 0.75em; color: #999; margin-top: 5px; font-style: italic;"">{((data.NoDscResourceRules.Count + data.HardCodedRules.Count) > 0 ? "Click to view details" : "No manual intervention needed")}</div>
                </div>
            </div>
        </div>

        <!-- Breakdown cards - REGULAR SIZE -->
        {((data.SuccessfulRules.Count - data.IndividualDISARulesAutomated > 0 || data.SkippedRules.Count > 0 || data.FailedCount > 0) ? $@"
        <div class=""summary"">
            <!-- Non-clickable stat display -->
            <div class=""summary-card info"" style=""cursor: default; opacity: 0.9;"">
                <div style=""display: inline-block; text-align: left;"">
                    <div style=""font-size: 1.3em; font-weight: 600; color: #333; margin-bottom: 6px;"">{data.SuccessfulRules.Count - data.IndividualDISARulesAutomated} Variant Rules Created</div>
                    <div style=""font-size: 0.85em; color: #666; margin-top: 5px;"">From {data.IndividualDISARulesAutomated} base rules</div>
                    <div style=""font-size: 0.75em; color: #999; margin-top: 5px; font-style: italic;"">
                        {data.SuccessfulRules.Count} Total automated rules including variants
                    </div>
                </div>
            </div>
            {(data.SkippedRules.Count > 0 ? $@"
            <div class=""summary-card info"" onclick=""jumpToSection('skipped-section')"">
                <div style=""display: inline-block; text-align: left;"">
                    <div style=""font-size: 1.3em; font-weight: 600; color: #333; margin-bottom: 6px;"">{data.SkippedRules.Count} Rules Skipped</div>
                    <div style=""font-size: 0.85em; color: #666; margin-top: 5px;"">Not Created</div>
                    <div style=""font-size: 0.75em; color: #999; margin-top: 5px; font-style: italic;"">From skip log file • Click to view details</div>
                </div>
            </div>" : "")}
            {(data.FailedCount > 0 ? $@"
            <div class=""summary-card error"" onclick=""jumpToSection('failed-section')"">
                <div style=""display: inline-block; text-align: left;"">
                    <div style=""font-size: 1.3em; font-weight: 600; color: #333; margin-bottom: 6px;"">{data.FailedCount} Failed Conversions</div>
                    <div style=""font-size: 0.85em; color: #666; margin-top: 5px;"">Rerun conversion to skip these failures</div>
                    <div style=""font-size: 0.75em; color: #999; margin-top: 5px; font-style: italic;"">Click to view details</div>
                </div>
            </div>" : "")}
        </div>" : "")}

            {GenerateSection("failed", "Failed Rule Conversions", data.FailedRules, data.FailedRuleDetails, "error")}
            {GenerateSection("nodsc", "Rules with No DSC Resource<br/><span style=\"font-size: 0.75em; font-weight: normal; color: #666;\">Requires Manual Intervention</span>", data.NoDscResourceRules, data.NoDscRuleDetails, "warning")}
            {GenerateSection("skipped", "Skipped Rules<br/><span style=\"font-size: 0.75em; font-weight: normal; color: #666;\">Requires Manual Intervention - Not Included in Converted Output</span>", data.SkippedRules, data.SkippedRuleDetails, "warning")}
            {GenerateSection("hardcoded", "Hard Coded Rules<br/><span style=\"font-size: 0.75em; font-weight: normal; color: #666;\">Requires Manual Intervention</span>", data.HardCodedRules, data.HardCodedRuleDetails, "info")}
            {GenerateSection("success", "Successfully Converted Rules<br/><span style=\"font-size: 0.75em; font-weight: normal; color: #666;\">Automated, No User Intervention</span>", data.SuccessfulRules, data.SuccessfulRuleDetails, "success")}
    </div>
</body>
</html>";
            return html;
        }

        private static string GenerateSection(string id, string title, List<string> rules, Dictionary<string, RuleDetail>? details, string styleClass)
        {
            if (rules.Count == 0) return string.Empty;

            var hasDetails = details != null && details.Count > 0;

            return $@"
        <div class=""section"">
            <div class=""section-header {styleClass}"" onclick=""toggleSection(this)"">
                <div class=""section-title"">{title}</div>
                <div>
                    <span class=""section-count"">{rules.Count}</span>
                    <span class=""toggle-icon"">▶</span>
                </div>
            </div>
            <div class=""section-content"" id=""{id}-section"">
                {(hasDetails ? $@"<button class=""expand-all"" onclick=""expandAll('{id}-section')"">Expand/Collapse All</button>" : "")}
                {string.Join("", rules.Select(ruleId =>
            {
                var detail = details?.GetValueOrDefault(ruleId);
                var hasDetail = detail != null;

                return $@"
                <div class=""rule-item"">
                    <div class=""rule-header"" {(hasDetail ? $@"onclick=""toggleRule(this)""" : "")}>
                        <div>
                            <span class=""rule-id"">{EscapeHtml(ruleId)}</span>
                            {(detail?.VariantCount > 1 ? $@"<span class=""rule-variant-count"">({detail.VariantCount} variants)</span>" : "")}
                            {(!string.IsNullOrWhiteSpace(detail?.Title) ? $@"<span class=""rule-title""> - {EscapeHtml(detail.Title.Length > 80 ? detail.Title.Substring(0, 80) + "..." : detail.Title)}</span>" : "")}
                        </div>
                        {(hasDetail ? @"<span>▼</span>" : "")}
                    </div>
                    {(hasDetail ? GenerateRuleDetails(detail!, ruleId, id) : "")}
                </div>";
            }))}
            </div>
        </div>";
        }

        private static string GenerateRuleDetails(RuleDetail detail, string ruleId, string sectionId)
        {
            // Create unique tab IDs by prefixing with section ID to avoid conflicts between sections
            var tabPrefix = $"{sectionId}-{EscapeHtml(ruleId)}";

            return $@"
                    <div class=""rule-details"">
                        {(detail.ErrorMessage != null ? $@"<div class=""error-message"">{EscapeHtml(detail.ErrorMessage)}</div>" : "")}

                        <div class=""tabs"">
                            <div class=""tab-buttons"">
                                <button class=""tab-button active"" onclick=""switchTab(event, '{tabPrefix}-overview')"">Overview</button>
                                {(!string.IsNullOrWhiteSpace(detail.Description) ? $@"<button class=""tab-button"" onclick=""switchTab(event, '{tabPrefix}-description')"">Description</button>" : "")}
                                {(!string.IsNullOrWhiteSpace(detail.FixText) ? $@"<button class=""tab-button"" onclick=""switchTab(event, '{tabPrefix}-fix')"">Fix</button>" : "")}
                                {(!string.IsNullOrWhiteSpace(detail.CheckText) ? $@"<button class=""tab-button"" onclick=""switchTab(event, '{tabPrefix}-check')"">Check</button>" : "")}
                                {(!string.IsNullOrWhiteSpace(detail.ConvertedSnippet) ? $@"<button class=""tab-button"" onclick=""switchTab(event, '{tabPrefix}-converted')"">Converted</button>" : "")}
                            </div>

                            <div class=""tab-content active"" id=""{tabPrefix}-overview"">
                                {(!string.IsNullOrWhiteSpace(detail.SvId) ? $@"<div class=""detail-row""><span class=""detail-label"">SV ID:</span><span class=""detail-value"">{EscapeHtml(detail.SvId)}</span></div>" : "")}
                                {(!string.IsNullOrWhiteSpace(detail.Severity) ? $@"<div class=""detail-row""><span class=""detail-label"">Severity:</span><span class=""detail-value severity-{detail.Severity?.ToLower()}"">{EscapeHtml(detail.Severity)}</span></div>" : "")}
                                {(!string.IsNullOrWhiteSpace(detail.Title) ? $@"<div class=""detail-row""><span class=""detail-label"">Title:</span><div class=""detail-text"">{EscapeHtml(detail.Title)}</div></div>" : "")}
                                {(detail.VariantCount > 1 ? $@"
                                <div class=""detail-row"">
                                    <span class=""detail-label"">Variants ({detail.VariantCount}):</span>
                                    <div class=""variant-list"">{string.Join(", ", detail.Variants.Select(EscapeHtml))}</div>
                                </div>" : "")}
                                {(!string.IsNullOrWhiteSpace(detail.DscResource) ? $@"<div class=""detail-row""><span class=""detail-label"">DSC Resource:</span><span class=""detail-value"">{EscapeHtml(detail.DscResource)}</span></div>" : "")}
                            </div>

                            {(!string.IsNullOrWhiteSpace(detail.Description) ? $@"
                            <div class=""tab-content"" id=""{tabPrefix}-description"">
                                <pre class=""detail-text-block"">{EscapeHtml(detail.Description)}</pre>
                            </div>" : "")}

                            {(!string.IsNullOrWhiteSpace(detail.FixText) ? $@"
                            <div class=""tab-content"" id=""{tabPrefix}-fix"">
                                <pre class=""detail-text-block"">{EscapeHtml(detail.FixText)}</pre>
                            </div>" : "")}

                            {(!string.IsNullOrWhiteSpace(detail.CheckText) ? $@"
                            <div class=""tab-content"" id=""{tabPrefix}-check"">
                                <pre class=""detail-text-block"">{EscapeHtml(detail.CheckText)}</pre>
                            </div>" : "")}

                            {(!string.IsNullOrWhiteSpace(detail.ConvertedSnippet) ? $@"
                            <div class=""tab-content"" id=""{tabPrefix}-converted"">
                                <pre class=""detail-text-block code-block"">{EscapeHtml(detail.ConvertedSnippet)}</pre>
                            </div>" : "")}
                        </div>
                    </div>";
        }

        private static string EscapeHtml(string? text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }

        public static void SaveReport(string html, string outputPath)
        {
            File.WriteAllText(outputPath, html, System.Text.Encoding.UTF8);
        }
    }

    public class ConversionReportData
    {
        public string StigName { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public int TotalRulesCreated { get; set; }
        public int RulesAutoHandled { get; set; }
        public int ManualHandlingRequired { get; set; }
        public int FailedCount { get; set; }
        public int IndividualDISARulesAutomated { get; set; }
        public int TotalDisaRules { get; set; }  // Total DISA base rules
        public int CoveredDisaRules { get; set; }  // DISA rules covered (automated + manual)
        public string LogFileStatus { get; set; } = string.Empty;
        public List<string> FailedRules { get; set; } = new();
        public List<string> SkippedRules { get; set; } = new();
        public List<string> HardCodedRules { get; set; } = new();
        public List<string> NoDscResourceRules { get; set; } = new();
        public List<string> SuccessfulRules { get; set; } = new();

        // Rich details for expandable items
        public Dictionary<string, RuleDetail> FailedRuleDetails { get; set; } = new();
        public Dictionary<string, RuleDetail> SuccessfulRuleDetails { get; set; } = new();
        public Dictionary<string, RuleDetail> SkippedRuleDetails { get; set; } = new();
        public Dictionary<string, RuleDetail> HardCodedRuleDetails { get; set; } = new();
        public Dictionary<string, RuleDetail> NoDscRuleDetails { get; set; } = new();
    }

    public class RuleDetail
    {
        public string? ErrorMessage { get; set; }
        public int VariantCount { get; set; }
        public List<string> Variants { get; set; } = new();
        public string? DscResource { get; set; }

        // Detailed XCCDF information
        public string? SvId { get; set; }
        public string? Title { get; set; }
        public string? Severity { get; set; }
        public string? Description { get; set; }
        public string? FixText { get; set; }
        public string? CheckText { get; set; }
        public string? ConvertedSnippet { get; set; }
    }
}