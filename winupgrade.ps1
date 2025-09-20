<#
.SYNOPSIS
Windows Upgrade Orchestration Script

.DESCRIPTION
Performs in-place Windows upgrades remotely.

.AUTHOR
John C.

.VERSION
1.1

.CREATED
2025-09-19

.COMPONENTS
- winupgrade.ps1: Launches upgrade
- upgr-progressmon.ps1: Monitors upgrade progress
#>
param (
    [string]$ComputerName
)

# --------------------------- CONFIG ---------------------------
$SourcePath     = 'C:\winsetup'
$TaskName       = 'InPlaceUpgrade'
$TimeoutMinutes = 180
# --------------------------------------------------------------

Write-Host "🚀 Starting remote upgrade orchestration for $ComputerName..."

# Connectivity check
try {
    Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction Stop | Out-Null
    Test-WSMan -ComputerName $ComputerName -ErrorAction Stop | Out-Null
} catch {
    throw "Cannot reach $ComputerName. Check ping and WinRM availability."
}

# Remote check: source path and disk space
$remoteCheck = {
    param($SourcePath)
    $result = [ordered]@{ Exists=$false; FreeGB=0; Edition=(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').EditionID }
    if (Test-Path -LiteralPath $SourcePath) {
        $result.Exists = $true
        $drive = Get-Item -LiteralPath $SourcePath
        $root  = (Get-PSDrive -Name $drive.PSDrive.Name)
        $result.FreeGB = [math]::Round($root.Free/1GB,2)
    }
    return $result
}
$pre = Invoke-Command -ComputerName $ComputerName -ScriptBlock $remoteCheck -ArgumentList $SourcePath
if (-not $pre.Exists) {
    throw ("Source path not found on {0}: {1}" -f $ComputerName, $SourcePath)
}
Write-Host "✅ Remote source located. Free space: $($pre.FreeGB) GB. Edition: $($pre.Edition)"

# Create and launch scheduled task remotely
$remoteSetup = {
    param($TaskName, $SourcePath)
    $setupExe = Join-Path $SourcePath 'setup.exe'
    if (-not (Test-Path -LiteralPath $setupExe)) {
        throw "setup.exe not found at $setupExe"
    }

    $arguments = @(
        '/auto', 'upgrade',
        '/quiet',
        '/eula', 'accept',
        '/DynamicUpdate', 'Enable',
        '/Compat', 'IgnoreWarning',
        '/Telemetry', 'Disable'
    )

    $action  = New-ScheduledTaskAction -Execute $setupExe -Argument ($arguments -join ' ')
    $trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(1))
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew

    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings | Out-Null
    Start-ScheduledTask -TaskName $TaskName

    # Wait for setup to start
    $started = $false
    for ($i=0; $i -lt 120; $i++) {
        $proc = Get-Process -Name 'SetupHost','setup' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($proc) { $started = $true; break }
        Start-Sleep -Seconds 5
    }
    if ($started) {
        Write-Host "🛠️ Setup is running on $env:COMPUTERNAME."
    } else {
        Write-Host "⚠️ Setup did not start within expected time."
    }
}
Write-Host "Launching silent setup task..."
Invoke-Command -ComputerName $ComputerName -ScriptBlock $remoteSetup -ArgumentList $TaskName,$SourcePath

# Monitor WinRM availability
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
$phase = 'pre-reboot'
$wasDown = $false

while ((Get-Date) -lt $deadline) {
    $up = $false
    try { Test-WSMan -ComputerName $ComputerName -ErrorAction Stop | Out-Null; $up = $true } catch { $up = $false }

    if ($phase -eq 'pre-reboot') {
        if (-not $up) {
            Write-Host "🔄 WinRM went down — likely rebooting into setup."
            $phase = 'rebooting'
            $wasDown = $true
        } else {
            Write-Host "⏳ Setup staging; waiting for reboot..."
        }
    }
    elseif ($phase -eq 'rebooting') {
        if ($up) {
            Write-Host "✅ WinRM is back — post-upgrade phase detected."
            break
        }
    }

    Start-Sleep -Seconds 15
}

if (-not $wasDown) {
    Write-Host "⚠️ WinRM never dropped. Setup may still be staging."
}

# Final version check
$verify = {
    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    [pscustomobject]@{
        ProductName   = $cv.ProductName
        EditionID     = $cv.EditionID
        DisplayVersion= $cv.DisplayVersion
        CurrentBuild  = $cv.CurrentBuild
        UBR           = $cv.UBR
    }
}
try {
    $ver = Invoke-Command -ComputerName $ComputerName -ScriptBlock $verify
    Write-Host "🧾 Post-upgrade version: $($ver.ProductName), Edition: $($ver.EditionID), DisplayVersion: $($ver.DisplayVersion), Build $($ver.CurrentBuild).$($ver.UBR)"

    $isWin11 = ($ver.ProductName -like '*Windows 11*')
    $isEdu   = ($ver.ProductName -like '*Education*' -or $ver.EditionID -like '*Education*')
    $isTargetVersion = ($ver.DisplayVersion -in @('23H2','24H2'))

    if ($isWin11 -and $isEdu -and $isTargetVersion) {
        Write-Host "✅ Upgrade completed successfully to Windows 11 Education $($ver.DisplayVersion)."
    } else {
        Write-Host "⚠️ Upgrade status uncertain. Version does not match expected 23H2 or 24H2 Education."
    }
} catch {
    Write-Host "❌ Could not query post-upgrade version. The device may still be finalizing."
}

Write-Host "🎯 Done."
