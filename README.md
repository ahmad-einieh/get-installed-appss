# get-installed-appss

Python/Tkinter utilities that export software registered on a Windows computer to an Excel workbook.

## Requirements

- Windows and Python 3 with Tkinter.
- `windows-tools` for installed-software enumeration.
- `openpyxl` for the newer script, or `XlsxWriter` for the original version.

The enumeration library is Windows-specific; these tools are not Linux/macOS application scanners and may not include every portable or unregistered app.

## Run the newer version

In Windows Command Prompt, from the repository directory:

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install windows-tools openpyxl
python "installed v2 python3.py"
```

Click **Save**, then choose an existing, writable output folder. The script writes `apps.xlsx` with software name, version, and publisher in three columns, without a header row.

## Original version

```bat
python -m pip install windows-tools XlsxWriter
python installedApps.py
```

Click **get my apps** and choose an output folder. This version also includes support/GitHub buttons and a status label.

## Notes

- Both versions use the fixed filename `apps.xlsx`; an existing file at that path can be overwritten. Back it up first and close it in Excel before exporting.
- Choose a folder rather than cancelling the dialog; cancellation is not handled explicitly.
- The newer script does not catch export errors; the original reports only a generic error.
- The workbook can reveal details about your computer's software. Review it before sharing.

## Support

[Support via PayPal](https://paypal.me/progahmad?locale.x=ar_EG), as linked in the original app.
