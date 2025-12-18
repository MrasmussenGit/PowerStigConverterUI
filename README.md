# PowerStigConverterUI

PowerStigConverterUI is a Windows desktop (WPF) helper app for working with DISA STIG content:
- Convert DISA XCCDF XML to PowerSTIG XML using the PowerStig.Convert module (Windows PowerShell 5.x).
- Compare DISA vs. converted output to list missing/added rule IDs.
- Split Windows OS STIG XML into separate MS/DC files with safe overwrite and logging.

## Requirements

- Windows 10/11
- .NET 9 SDK
- Visual Studio 2022 (or newer) or `dotnet` CLI
- Windows PowerShell 5.x (`powershell.exe` on Windows)
- PowerSTIG and PowerStig.Convert installed under:
  - %ProgramFiles%\WindowsPowerShell\Modules
  - %UserProfile%\Documents\WindowsPowerShell\Modules

Tip (Windows PowerShell 5.x):
- Install-Module PowerSTIG -Scope AllUsers
- Install-Module PowerStig.Convert -Scope AllUsers

## Build and Run

- Visual Studio:
  - Open the solution and press F5 to run.
- CLI:
  - dotnet build
  - dotnet run --project PowerStigConverterUI

Command-line compare mode:
- PowerStigConverterUI.exe "C:\Path\To\DisaFolder" "C:\Path\To\PowerStigFolder"
- Compares all matched pairs and writes a report to the console, then exits.

## Features

### Convert STIG
- Auto-discovers `PowerStig.Convert.psm1` in common module paths. You can browse to it if needed.
- Select an XCCDF XML and a destination folder, then click Convert.
- Optionally create an organization defaults file (`*.org.default.xml`).
- Edition tokens preserved:
  - Input like `U_MS_Windows_Server_2022_MS_STIG_V2R6_Manual-xccdf.xml`
  - Output renamed to include edition: `WindowsServer-2022-MS-2.6.xml`
  - The `*.org.default.xml` is also renamed to preserve the edition token (e.g., `WindowsServer-2022-MS-2.6.org.default.xml`).
- Comparison:
  - After conversion, the app shows “Comparing…” and the exact converted file path used.
  - Lists “Missing rule IDs” present in XCCDF but not in the converted output.
- Messages panel:
  - Displays conversion progress, errors, success summaries, and comparison results.

### Compare
- Dedicated Compare window (from the main page) to inspect differences between a DISA XCCDF and a PowerSTIG XML.
- Lists missing and added V-IDs using a normalized base ID comparison.

### Split OS STIG
- For Windows OS STIG XML (e.g., `U_MS_Windows_Server_2022_STIG_VxRx_Manual-xccdf.xml`):
  - Creates MS/DC copies in the chosen destination folder:
    - `..._MS_...` → `...-MS.xml`
    - `..._DC_...` → `...-DC.xml`
  - Copies a sibling `.log` file if present.
  - Respects “Overwrite if exist”. When checked, overwrites without prompting; otherwise prompts.
  - Destination-aware: checks and warns only against files in the selected destination.
  - Message area keeps a history for multiple runs and auto-scrolls.

Supported Windows OS STIG name pattern (heuristic):
- `U_MS_Windows_Server_2022_STIG_V2R6_Manual-xccdf.xml`
- `U_MS_Windows_11_STIG_V2R6_Manual-xccdf.xml`
- `U_MS_Windows_10_STIG_V2R6_Manual-xccdf.xml`

## Version Display
- The main window shows “Version: <x.y.z.w>” derived from assembly metadata.

## Troubleshooting

- “PowerStig module not found”:
  - Ensure PowerSTIG and PowerStig.Convert are installed in Windows PowerShell 5.x module paths.
- Conversion fails with missing file under destination:
  - The app now resolves and, when needed, renames outputs to preserve the edition token (MS/DC) for both the main XML and `*.org.default.xml`.
- Split shows no prompts on repeated clicks:
  - Use the Overwrite checkbox to control prompting; with Overwrite checked, the app overwrites without prompts and logs each run.

## Contributing

- Issues and PRs are welcome. Please:
  - Keep WPF UI changes minimal and consistent.
  - Preserve existing code comments when refactoring.
  - Target .NET 9 and C# 13 unless otherwise discussed.

## License

- See the repository’s license file if present.
