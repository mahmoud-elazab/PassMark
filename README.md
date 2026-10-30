# PassMark Tool

A Python desktop application for processing voltage logs from a PassMark USB Power Delivery Tester and exporting formatted Excel summaries.

Select log files, choose operating modes, review average voltages and tolerance results, and export an Initial or Final report. An optional reference workbook can supply loaded and unloaded current values.

## Features

- Desktop interface built with CustomTkinter.
- Extracts voltage readings from diagnostic log lines.
- Groups logs by device, operating mode, and port.
- Pairs measurements taken at 0° and 180°.
- Checks average voltages against a ±0.25 V tolerance.
- Displays warnings for missing angle files, duplicates, and unselected modes.
- Exports `.xlsx` reports with mode-specific columns, highlighted out-of-tolerance values, and black cells for unused ports.
- Includes an interactive generator for synthetic example logs.

## Requirements and setup

Use Python 3.10 or newer with Tkinter available. The folder contains Python 3.12 bytecode from a previous build; compatibility across Python versions and operating systems has not been fully verified.

From the project directory, create a virtual environment:

```sh
python -m venv .venv
```

Activate it on Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Or on macOS/Linux:

```sh
source .venv/bin/activate
```

Install the runtime dependencies and launch the app:

```sh
python -m pip install customtkinter openpyxl
python main.py
```

If your system uses `python3` instead of `python`, use that command to create the environment. On some systems, Tkinter must be installed separately through the Python installer or operating system package manager. You can check its availability with `python -m tkinter`.

Dependencies are currently installed directly; the project does not yet include a pinned dependency file.

## Input format

Name each log using four numeric fields:

```text
<device>-<mode>-<port>-<angle>.txt
```

For example, `10001-5-1-1.txt` identifies device `10001`, mode `5`, port `1`, at 0°.

| Field | Meaning |
| --- | --- |
| Device | Numeric device identifier |
| Mode | Operating mode, from 2 to 7 |
| Port | 1 or 2 |
| Angle | 1 = 0°; 2 = 180° |

Logs must be UTF-8 text. The parser reads lines containing `Diagnostics-` and extracts positive decimal voltage values in the form `V=15.02V`:

```text
Diagnostics- V=15.02V, I Tester=2000mA, I Loopback=0mA, I Total~=2000mA
```

Other lines are ignored. Current values in the text logs are not processed.

## Operating modes

| Mode | Port 1 / Port 2 power | Port 1 nominal voltage | Port 2 nominal voltage |
| --- | --- | --- | --- |
| 2 | 45 W / 15 W | 15 V | 5 V |
| 3 | 15 W / 45 W | 5 V | 15 V |
| 4 | 15 W / 15 W | 5 V | 5 V |
| 5 | 30 W / 30 W | 15 V | 15 V |
| 6 | 0 W / 60 W | Unused | 20 V |
| 7 | 60 W / 0 W | 20 V | Unused |

Modes 5, 6, and 7 are selected by default. Voltage limits include both endpoints of the nominal voltage ±0.25 V range. The app does not calculate a separate mode 1 voltage result; the report includes an unloaded-current column.

## Usage

1. Run `python main.py`.
2. Select the operating modes to process.
3. Choose **Initial** or **Final** for the report heading.
4. Click **Select Files** and choose the `.txt` logs.
5. Optionally click **Load Reference Excel (optional)** to import current values.
6. Click **Process** and review results, errors, and coverage warnings.
7. Click **Export to Excel** and choose where to save the workbook.

The report is created in code; no bundled report template is required.

### How averages are calculated

The app first calculates the mean of the voltage readings in each file. When both angles are available for a device/mode/port combination, it takes the mean of those two file averages. Each angle receives equal weight, even when the files contain different numbers of readings. Results are rounded to two decimal places before tolerance checks.

When only one angle is available, the app uses that file's average and adds a warning. For duplicate angle files, the last file encountered in the selected-file list takes precedence. Files belonging to unselected modes are ignored with a warning.

### Optional reference workbook

The reader prefers a worksheet named `FT Data Summary`; otherwise it reads the first worksheet. It looks for `Sample #`, `Unloaded Current`, and `Loaded Current` headers, with power labels such as `30W/30W` above the loaded-current columns.

**Current values are copied by row position, not matched by device identifier.** Put reference rows in the same order as the exported device rows, which are sorted by the device identifier as text. Extra reference rows are ignored, and missing values remain blank. Formula cells need cached results saved by a spreadsheet application because the reader does not calculate formulas.

## Generate example logs

```sh
python create_test_files.py
```

The generator asks for device identifiers, operating modes, and the number of readings. It writes synthetic logs for both angles of each active port into `test_files/` relative to the current directory. Use fictional numeric identifiers. Existing files with matching names are overwritten after confirming generation.

This is a sample-data generator; the project does not currently contain an automated test suite.

## Project structure

```text
main.py                         Application entry point
gui/app_window.py               Desktop interface and workflow
core/filename_parser.py         Filename metadata parsing
core/file_parser.py             Voltage extraction and per-file averages
core/calculator.py              Pairing, mode definitions, and tolerance checks
core/reference_excel_reader.py  Optional current-value import
output/excel_writer.py          Excel report creation
create_test_files.py            Synthetic log generator
PassMark Tool*.spec             Existing PyInstaller build configurations
```

`output/` contains application source code and must be included in the repository. Generated reports are saved to the location chosen in the app.

## Windows packaging

The project includes PyInstaller specifications for earlier Windows builds. To attempt a V3 build, run the following on Windows from the project directory:

```sh
python -m pip install pyinstaller
python -m PyInstaller "PassMark Tool V3.spec"
```

Build artifacts are written to `build/` and `dist/`. The existing specifications have not been revalidated in a clean environment; verify the packaged app and its CustomTkinter resources before distributing it.

## Current limitations

- The report's **Reverse polarity** value is derived from the available voltage tolerance results. It does not independently detect polarity or prove that all required measurements are present.
- The **LED On?** column is always filled with `On`; LED state is not read from the logs.
- Report headings contain fixed `13.5 V Input` and `25° C` labels. These conditions are not measured or validated by the app.
- Files without matching voltage readings are recorded as `0.0 V` and listed in the processing errors.
- Coverage warnings group files by mode and port across the selection, so they do not establish completeness for every device. They also expect both ports for modes with an unused port.
- Unusual filename fields produce warnings but are not consistently rejected. Use the documented mode, port, and angle values.
- Changing inputs or modes after processing requires clicking **Process** again before export.

Review the warnings and report values before relying on the results.
