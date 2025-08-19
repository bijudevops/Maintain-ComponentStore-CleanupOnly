<# 
.SYNOPSIS
  Safe WinSxS / Component Store maintenance (analyze + recommended cleanup only).

.DESCRIPTION
  - Runs DISM /AnalyzeComponentStore.
  - Parses whether "Component Store Cleanup Recommended : Yes".
  - If recommended, runs DISM /StartComponentCleanup (safe).
  - Optional: run /SPSuperseded (legacy OS) and trigger the servicing scheduled task.
  - Logs to a timestamped file and prints a summary.
  - Never uses /ResetBase.

.PARAMETER AnalyzeOnly
  Only analyze; do not perform cleanup.

.PARAMETER TriggerScheduledTask
  Trigger "\Microsoft\Windows\Servicing\StartComponentCleanup" scheduled task.

.PARAMETER IncludeSPSuperseded
  Attempt DISM /SPSuperseded (useful only on older Server versions).

.PARAMETER LogPath
  Folder for logs (default C:\Windows\Logs\ComponentCleanup).

.EXAMPLE
  .\Maintain-ComponentStore-CleanupOnly.ps1 -Verbose
#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
  [switch] $AnalyzeOnly,
  [switch] $TriggerScheduledTask,
  [switch] $IncludeSPSuperseded,
  [string] $LogPath = "C:\Windows\Logs\ComponentCleanup"
)

#region Helpers
function Test-Administrator {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-RebootPending {
  $paths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",
    "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations"
  )
  foreach ($p in $paths) { if (Test-Path $p) { return $true } }
  return $false
}

function Invoke-Dism {
  param(
    [Parameter(Mandatory=$true)][string[]] $Arguments,
    [Parameter(Mandatory=$true)][string]   $LogFile
  )
  $disp = "dism.exe " + ($Arguments -join ' ')
  if ($PSCmdlet.ShouldProcess($disp, "Run DISM")) {
    Write-Verbose "Running: $disp"
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "dism.exe"
    $psi.Arguments = ($Arguments -join ' ')
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    $null = $p.Start()
    $out = $p.StandardOutput.ReadToEnd()
    $err = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    Add-Content -Path $LogFile -Value ("`n=== {0} ===`n{1}`n" -f (Get-Date), $out)
    if ($err) { Add-Content -Path $LogFile -Value ("`n[STDERR]`n{0}`n" -f $err) }

    if ($p.ExitCode -ne 0) { throw "DISM exited with $($p.ExitCode). See $LogFile" }
    return $out
  }
}

function Parse-CleanupRecommended {
  param([string]$AnalysisText)
  # Looks for: "Component Store Cleanup Recommended : Yes"
  $m = ($AnalysisText -split "`r?`n") |
       Where-Object { $_ -match 'Component Store Cleanup Recommended\s*:\s*(Yes|No)' } |
       Select-Object -First 1
  if (-not $m) { return $null }
  return ($m -match 'Yes')
}
#endregion Helpers

#region Preconditions
if (-not (Test-Administrator)) {
  Write-Error "Run this script from an elevated PowerShell session."
  exit 1
}

try {
  if (-not (Test-Path -LiteralPath $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
  }
} catch {
  Write-Error "Failed to access/create log folder '$LogPath': $_"
  exit 1
}

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $LogPath "ComponentCleanup_$ts.log"

try { Start-Transcript -Path (Join-Path $LogPath "Transcript_$ts.txt") -ErrorAction Stop | Out-Null } catch { Write-Warning "Could not start transcript: $_" }
#endregion Preconditions

Write-Host "== Component Store Maintenance (Cleanup-Only) ==" -ForegroundColor Cyan
Write-Host "Log file: $LogFile`n"

$summary = [ordered]@{
  Analyzed                 = $false
  CleanupRecommended       = $null
  CleanupPerformed         = $false
  SPSupersededRun          = $false
  ScheduledTaskTriggered   = $false
  RebootPendingBefore      = (Test-RebootPending)
  RebootPendingAfter       = $null
}

try {
  # 1) Analyze
  Write-Host "Step 1/3: Analyzing component store..." -ForegroundColor Yellow
  $analysis = Invoke-Dism -Arguments @("/Online","/Cleanup-Image","/AnalyzeComponentStore") -LogFile $LogFile
  $summary.Analyzed = $true

  $recommended = Parse-CleanupRecommended -AnalysisText $analysis
  if ($null -eq $recommended) {
    Write-Warning "Could not determine if cleanup is recommended from DISM output. Skipping cleanup."
  } else {
    $summary.CleanupRecommended = $recommended
  }

  if ($AnalyzeOnly) {
    Write-Host "`nAnalyze-only mode: skipping cleanup." -ForegroundColor Yellow
  } elseif ($recommended -eq $true) {
    # 2) Safe cleanup (only when recommended)
    Write-Host "Step 2/3: Running StartComponentCleanup (recommended)..." -ForegroundColor Yellow
    Invoke-Dism -Arguments @("/Online","/Cleanup-Image","/StartComponentCleanup") -LogFile $LogFile
    $summary.CleanupPerformed = $true
  } else {
    Write-Host "Cleanup not recommended by DISM. Nothing to do." -ForegroundColor Green
  }

  # Optional legacy SP backup removal
  if ($IncludeSPSuperseded) {
    Write-Host "Attempting SPSuperseded (legacy systems only)..." -ForegroundColor Yellow
    try {
      Invoke-Dism -Arguments @("/Online","/Cleanup-Image","/SPSuperseded") -LogFile $LogFile
      $summary.SPSupersededRun = $true
    } catch {
      Write-Warning "SPSuperseded failed or not supported: $_"
    }
  }

  # Optional: trigger scheduled cleanup task
  if ($TriggerScheduledTask) {
    Write-Host "Triggering scheduled task: \Microsoft\Windows\Servicing\StartComponentCleanup" -ForegroundColor Yellow
    try {
      schtasks.exe /Run /TN "\Microsoft\Windows\Servicing\StartComponentCleanup" | Out-Null
      $summary.ScheduledTaskTriggered = $true
      Add-Content -Path $LogFile -Value ("`n[{0}] Triggered scheduled task StartComponentCleanup.`n" -f (Get-Date))
    } catch {
      Write-Warning "Could not trigger scheduled task: $_"
    }
  }

} catch {
  Write-Error $_
} finally {
  $summary.RebootPendingAfter = Test-RebootPending
  try { Stop-Transcript | Out-Null } catch {}
}

# Summary
Write-Host "`n== Summary ==" -ForegroundColor Cyan
$summary.GetEnumerator() | Sort-Object Name | ForEach-Object { "{0,-24} : {1}" -f $_.Key, $_.Value }

Write-Host "`nDetailed log:" -ForegroundColor Cyan
Write-Host "  $LogFile"
Write-Host "  (Transcript saved alongside the log.)"

# Exit codes
if ($summary.RebootPendingAfter) {
  Write-Host "`nA system reboot is pending. Consider rebooting to complete servicing." -ForegroundColor Yellow
  exit 3010
} else {
  exit 0
}
