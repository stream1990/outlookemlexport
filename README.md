# outlookemlexport

Export locally cached Outlook mails to `.eml` (Python CLI + .NET Framework WinForms).

## Download

Prebuilt exe: [OutlookEmlExport/bin/Release/OutlookEmlExport.exe](OutlookEmlExport/bin/Release/OutlookEmlExport.exe)

## .NET Framework GUI

1. Run `OutlookEmlExport.exe` (requires desktop Outlook)
2. Click **列出账号**
3. Uncheck accounts you do not want
4. Click **导出**

Build from source:

```
cd OutlookEmlExport
build.bat
```

## Python CLI

```
pip install -r requirements.txt
python export_outlook_eml.py --list-stores
python export_outlook_eml.py -o ./out
```