# Excel2Openness-Installer
This tool provides a streamlined, engineering-centred interface between Excel and ABT Site. Excel2Openness is an independent tool that enables Excel-based interaction with ABT Openness via a gRPC interface.

---

## 🚀 Download

👉 [Download Latest Version](../../releases/latest)

---

## 📦 Installation

1. Download the latest `.msi` installer
2. Run the installer
3. Open the Excel template from the added desktop icon 'Excel2Openness'
4. The Excel2Openness worksheet and add-in will load automatically


## 🛠 Post-install add-in registration (From V61.13.0.0 onwards)

In some environments, Excel may not automatically register the add-in after installation.

If that happens, run the following PowerShell script manually:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\Program Files\Excel2Openness\InstallerScripts\RegisterExcelAddin.ps1"
```

To remove the add-in registration manually, use:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\Program Files\Excel2Openness\InstallerScripts\UnregisterExcelAddin.ps1"
```
These helper scripts are installed here:
```
C:\Program Files\Excel2Openness\InstallerScripts\
```
After running the registration script, close and reopen Excel.
---

## 🔔 Stay Updated

To receive notifications for new versions:

1. Click **Watch** (top right of this repository)
2. Select **Custom**
3. Enable **Releases**

---

## ⚠️ Disclaimer

This is an independent, unofficial tool.  
It is not affiliated with, endorsed by, or supported by any specific company.

This tool is provided without warranty and used at your own risk.

---

## 🍺 Support

If you find this tool useful, consider supporting the project:

👉 https://buymeabeer.com/timpendlebury
