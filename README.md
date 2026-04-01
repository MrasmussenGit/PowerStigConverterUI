# PowerSTIG Converter UI

A modern WPF desktop application for converting, comparing, and managing DISA Security Technical Implementation Guides (STIGs) using PowerSTIG automation.

![.NET](https://img.shields.io/badge/.NET-9.0-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)

## 📋 Overview

PowerSTIG Converter UI provides a user-friendly graphical interface for working with DISA STIGs and the PowerSTIG PowerShell module. It streamlines the process of converting STIG compliance requirements into PowerSTIG DSC configurations, comparing conversion results, and splitting Windows OS STIGs into Member Server (MS) and Domain Controller (DC) variants.

## ✨ Features

### 🔄 Convert STIG
- **Automated Conversion**: Convert DISA STIG XCCDF files to PowerSTIG XML format
- **ZIP Support**: Direct import from DISA STIG ZIP packages
- **Smart Module Discovery**: Automatically locates PowerSTIG modules on your system
- **Log File Management**: Tracks failed rules and skips them on subsequent conversions
- **Detailed Reporting**: Generates comprehensive HTML reports with coverage statistics
- **Rule Details**: Double-click any rule ID to view XCCDF details, fix text, and conversion results

### 📊 Compare STIG
- **Side-by-Side Comparison**: Compare DISA base STIG against PowerSTIG converted output
- **Coverage Analysis**: Shows which rules are automated, manual, skipped, or failed
- **Missing Rules Detection**: Identifies DISA rules not present in converted output
- **HTML Reports**: Interactive reports with detailed breakdowns by category
- **Copy-Friendly Output**: All text areas support selection and copying

### 🔀 Split OS STIG
- **Automated Splitting**: Splits Windows OS STIGs into MS and DC variants using PowerSTIG cmdlets
- **Member Server (MS) and Domain Controller (DC)**: Creates separate STIG files for each role
- **ZIP Support**: Extracts and processes STIG files from ZIP archives
- **Command Visibility**: Displays PowerShell commands being executed
- **Log File Handling**: Automatically duplicates log files for both variants

## 🎯 Key Capabilities

- **Dark-Themed Modern UI**: Professional interface matching Visual Studio styling
- **Auto-Discovery**: Finds PowerSTIG modules in standard Windows PowerShell locations
- **Persistent Settings**: Remembers last-used directories and module paths
- **Error Recovery**: Failed rules are logged and skipped on subsequent conversions
- **Comprehensive Statistics**: 
  - Total rules created
  - Automated vs. manual rules
  - Coverage percentages
  - Failed and skipped rules
- **Interactive HTML Reports**: 
  - Expandable/collapsible sections
  - Tabbed rule details (Overview, Description, Fix, Check, Converted)
  - Color-coded severity levels
  - Click-to-jump navigation

## 🖥️ Requirements

### System Requirements
- **Operating System**: Windows 10/11 or Windows Server 2016+
- **.NET Runtime**: .NET 9.0 Desktop Runtime
- **PowerShell**: Windows PowerShell 5.1 (included with Windows)

### Dependencies
- **PowerSTIG Module**: Install from PowerShell Gallery
  ```powershell
  Install-Module -Name PowerSTIG -Scope CurrentUser
  ```
- **PowerSTIG.Convert Module**: Included with PowerSTIG 4.0+
  ```powershell
  Install-Module -Name PowerStig.Convert -Scope CurrentUser
  ```

## 📦 Installation

### Option 1: Download Release (Recommended)
1. Download the latest release from the [Releases](../../releases) page
2. Extract the ZIP file to a folder of your choice
3. Run `PowerStigConverterUI.exe`

### Option 2: Build from Source
1. **Prerequisites**:
   - Visual Studio 2022 or later
   - .NET 9.0 SDK

2. **Clone the Repository**:
   ```bash
   git clone https://github.com/MrasmussenGit/PowerStigConverterUI.git
   cd PowerStigConverterUI
   ```

3. **Build**:
   ```bash
   dotnet build -c Release
   ```

4. **Run**:
   ```bash
   cd PowerStigConverterUI/bin/Release/net9.0-windows
   ./PowerStigConverterUI.exe
   ```

## 🚀 Quick Start

### Converting a STIG

1. **Launch** the application
2. **Click** "Convert STIG"
3. **Select** your DISA STIG XCCDF file or ZIP package
4. **Choose** an output folder
5. **Click** "Convert"
6. **View** the generated HTML report for detailed results

### Comparing STIG Results

1. **Click** "Compare DISA vs PowerSTIG"
2. **Select** the original DISA XCCDF file
3. **Select** the PowerSTIG converted XML file
4. **Click** "Compare"
5. Review coverage statistics and missing rules

### Splitting Windows OS STIG

1. **Click** "Split OS STIG"
2. **Select** a Windows OS STIG XCCDF file (e.g., Windows Server 2022)
3. **Choose** a destination folder (optional)
4. **Click** "Split"
5. Find separate MS and DC STIG files in the destination

## 📖 Usage Details

### Module Path Configuration

The application automatically searches for PowerSTIG modules in:
- `C:\Program Files\WindowsPowerShell\Modules\PowerSTIG\`
- `%UserProfile%\Documents\WindowsPowerShell\Modules\PowerSTIG\`
- Locations in `$env:PSModulePath`

If modules aren't found, expand the **Advanced** section to manually browse for:
- **Convert**: `PowerStig.Convert.psm1`
- **Split OS**: `Functions.XccdfXml.ps1`

### Log Files

PowerSTIG uses log files (`.log`) to track:
- **Skipped Rules**: Rules to skip during conversion (`V-12345::*::.`)
- **Hard-Coded Rules**: Manual overrides for specific configurations

Log files are automatically discovered when placed alongside the XCCDF file with matching names:
- XCCDF: `U_MS_IIS_10-0_Server_STIG_V3R6_Manual-xccdf.xml`
- Log: `U_MS_IIS_10-0_Server_STIG_V3R6_Manual-xccdf.log`

### HTML Reports

Generated reports include:
- **Coverage Summary**: Visual cards showing automation coverage
- **Automated Rules**: Successfully converted rules with DSC resources
- **Manual Rules**: Rules requiring manual intervention (no DSC resource)
- **Hard Coded Rules**: Manually configured rules from log file
- **Skipped Rules**: Rules intentionally skipped
- **Failed Rules**: Rules that failed conversion with error details

Each rule can be expanded to view:
- DISA STIG description
- Fix text
- Check procedures
- PowerSTIG converted DSC configuration

## 🎨 Features in Detail

### Smart ZIP Handling
- Automatically extracts XCCDF files from DISA STIG ZIP packages
- Handles nested directory structures
- Derives log file names from XCCDF files inside ZIPs
- Cleans up temporary extraction directories automatically

### Rule Variant Support
- Handles STIG rules with multiple variants (V-12345.a, V-12345.b)
- Reports total variants created per base rule
- Maintains variant-specific converted snippets in reports

### Edition Token Preservation
- Windows Server STIGs: Preserves MS/DC edition tokens in filenames
- Automatically renames files to maintain PowerSTIG naming conventions
- Example: `WindowsServer-2022-MS-2.6.xml` and `WindowsServer-2022-DC-2.6.xml`

### Coverage Calculation
- **Total DISA Rules**: Unique base rule IDs from DISA STIG
- **Covered Rules**: Automated + Manual (Hard Coded + No DSC Resource)
- **Missing Rules**: Failed + Skipped
- **Coverage %**: (Covered / Total) × 100

## 🛠️ Development

### Project Structure
```
PowerStigConverterUI/
├── MainWindow.xaml              # Main application window
├── ConvertStigWindow.xaml       # STIG conversion interface
├── CompareWindow.xaml           # STIG comparison interface
├── SplitStigWindow.xaml         # OS STIG splitting interface
├── ConversionReportGenerator.cs # HTML report generation
├── RuleInfoWindow.xaml          # Rule detail viewer
├── AppSettings.cs               # Settings persistence
├── Styles/                      # XAML style resources
│   ├── Colors.xaml
│   ├── ButtonStyles.xaml
│   ├── TextStyles.xaml
│   ├── BorderStyles.xaml
│   ├── DataGridStyles.xaml
│   ├── BadgeStyles.xaml
│   └── StyleGuide.xaml
└── RuleIdAnalysis.cs            # STIG comparison logic
```

### Tech Stack
- **Framework**: WPF (Windows Presentation Foundation)
- **.NET**: 9.0
- **Language**: C# 13.0
- **PowerShell Integration**: Windows PowerShell 5.1 via `System.Diagnostics.Process`
- **UI Theme**: Custom dark theme inspired by Visual Studio

### Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 Common STIG Products Supported

The application supports all DISA STIG products that PowerSTIG can convert, including:

- **Windows**: Server 2016/2019/2022, Windows 10/11
- **Office**: 2013/2016, Office 365 ProPlus, Outlook, Excel, Word, etc.
- **Web Servers**: IIS 8.5/10.0 Server and Site
- **Databases**: SQL Server 2012/2016/2019 Instance and Database
- **Browsers**: Microsoft Edge, Internet Explorer 11, Chrome, Firefox
- **Virtualization**: VMware vSphere 6.5/6.7/7.0
- **Linux**: RHEL 7/8, Ubuntu 18.04/20.04, Oracle Linux 7/8
- **Other**: Adobe Acrobat Reader/Pro, .NET Framework, Oracle JRE, McAfee VirusScan

## ⚠️ Known Limitations

- Requires Windows PowerShell 5.1 (PowerShell 7+ not supported for PowerSTIG module execution)
- Some DISA rules cannot be automated and require manual implementation
- ZIP file support assumes standard DISA STIG package structure
- Module auto-discovery searches standard Windows PowerShell locations only

## 🐛 Troubleshooting

### Module Not Found
**Issue**: "PowerStig.Convert module was not found"  
**Solution**: 
1. Install PowerSTIG: `Install-Module PowerSTIG -Scope CurrentUser`
2. Restart the application to trigger auto-discovery
3. Or manually browse to module location via Advanced section

### Failed Rule Conversions
**Issue**: Rules fail during conversion  
**Solution**: 
1. Create or update the `.log` file alongside your XCCDF
2. Add failed rule IDs in format: `V-12345::*::.`
3. Re-run conversion - failed rules will be skipped

### ZIP Extraction Errors
**Issue**: Cannot find XCCDF in ZIP  
**Solution**: 
1. Manually extract the ZIP
2. Use the XCCDF file directly instead

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **PowerSTIG Team**: For the amazing PowerSTIG PowerShell module
- **DISA**: For maintaining and publishing STIG compliance standards
- **Community Contributors**: For feedback and feature requests

## 📞 Support

- **Issues**: [GitHub Issues](../../issues)
- **Discussions**: [GitHub Discussions](../../discussions)
- **PowerSTIG**: [PowerSTIG GitHub](https://github.com/microsoft/PowerStig)

## 🔗 Related Resources

- [PowerSTIG Documentation](https://github.com/microsoft/PowerStig/wiki)
- [DISA STIGs](https://public.cyber.mil/stigs/)
- [PowerShell Gallery - PowerSTIG](https://www.powershellgallery.com/packages/PowerSTIG)

---

**Note**: This is an unofficial community tool and is not affiliated with DISA or Microsoft PowerSTIG team.
