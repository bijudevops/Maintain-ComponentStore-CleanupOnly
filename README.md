# Maintain-ComponentStore-CleanupOnly.ps1

A safe PowerShell script to **analyze and reduce the size of the Windows WinSxS component store** (`C:\Windows\WinSxS`) on Windows Server.  

This script uses **only Microsoft-supported cleanup methods** and runs cleanup **only when DISM recommends it**.  
It does **not** use `/ResetBase` (which blocks uninstalling older updates).

---

## ✨ Features
- Runs `DISM /AnalyzeComponentStore` to check for recommended cleanup.
- Executes `DISM /StartComponentCleanup` only if cleanup is recommended.
- Optionally runs:
  - `/SPSuperseded` (for legacy service pack backups on older Windows Server versions).
  - The scheduled servicing task `\Microsoft\Windows\Servicing\StartComponentCleanup`.
- Creates detailed logs in `C:\Windows\Logs\ComponentCleanup`.
- Detects reboot-pending state before and after cleanup.
- Provides a summary of actions taken.

---

## ⚠️ Requirements
- Windows Server 2012 R2 and newer (tested up to Server 2022).
- Must be run from an **elevated PowerShell session** (`Run as Administrator`).
- PowerShell 5.1 or later.

---

## 🚀 Usage

### Default (analyze + cleanup if recommended)
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Maintain-ComponentStore-CleanupOnly.ps1 -Verbose
```

### Analyze only (no cleanup)
```powershell
.\Maintain-ComponentStore-CleanupOnly.ps1 -AnalyzeOnly
```

### Include legacy service pack cleanup
```powershell
.\Maintain-ComponentStore-CleanupOnly.ps1 -IncludeSPSuperseded
```

### Trigger the Windows servicing scheduled task
```powershell
.\Maintain-ComponentStore-CleanupOnly.ps1 -TriggerScheduledTask
```

---

## 📂 Logging
- All output is saved in:
  - `C:\Windows\Logs\ComponentCleanup\ComponentCleanup_<timestamp>.log`
  - Transcript files alongside the logs.
- The summary is also printed to console.

---

## ✅ Example Output
```text
== Summary ==
Analyzed                 : True
CleanupRecommended       : True
CleanupPerformed         : True
SPSupersededRun          : False
ScheduledTaskTriggered   : False
RebootPendingBefore      : False
RebootPendingAfter       : True

Detailed log:
  C:\Windows\Logs\ComponentCleanup\ComponentCleanup_20250820_103000.log
```

---

## 📌 Notes
- The script **never uses `/ResetBase`**. This keeps rollback/uninstall of older updates possible.
- If DISM reports cleanup is **not recommended**, nothing is changed.
- If a reboot is pending before cleanup, the script will notify you.

---

## 🔒 Safety
- Only uses **Microsoft-supported DISM commands**.
- Built-in safeguards (admin check, log capture, no destructive actions).
- Suitable for scheduled maintenance tasks or manual health checks.

---

## 📝 License
MIT License – feel free to use and modify.
