<#
.SYNOPSIS
Windows Upgrade Orchestration Script

.DESCRIPTION
Monitors in-place Windows upgrades remotely using WinRM and scheduled tasks.

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

function Show-ProgressBar {
    param($Percent, $Message)
    $bar = '=' * ($Percent / 2) + '>' + ' ' * (50 - ($Percent / 2))
    Write-Host ("[{0}] {1}% - {2}" -f $bar, $Percent, $Message)
}

Write-Host "`n🚀 Starting remote upgrade orchestration for $ComputerName..."

try {
    Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction Stop | Out-Null
    Test-WSMan -ComputerName $ComputerName -ErrorAction Stop | Out-Null
} catch {
    throw "❌ Cannot reach $ComputerName. Check ping and WinRM availability."
}

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
    throw ("❌ Source path not found on {0}: {1}" -f $ComputerName, $SourcePath)
}
Write-Host "Remote source located. Free space: $($pre.FreeGB) GB. Edition: $($pre.Edition)"

$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
$phase = 'pre-reboot'
$wasDown = $false
$startTime = Get-Date

while ((Get-Date) -lt $deadline) {
    $elapsed = (Get-Date) - $startTime
    $percent = [math]::Min([math]::Round(($elapsed.TotalMinutes / $TimeoutMinutes) * 100), 100)

    $up = $false
    try { Test-WSMan -ComputerName $ComputerName -ErrorAction Stop | Out-Null; $up = $true } catch { $up = $false }

    if ($phase -eq 'pre-reboot') {
        if (-not $up) {
            Write-Host "`n🔄 WinRM went down — likely rebooting into setup."
            $phase = 'rebooting'
            $wasDown = $true
        } else {
            Show-ProgressBar -Percent $percent -Message "Setup staging on $ComputerName..."
        }
    }
    elseif ($phase -eq 'rebooting') {
        if ($up) {
            Write-Host "`nWinRM is back — post-upgrade phase detected."
            break
        } else {
            Show-ProgressBar -Percent $percent -Message "Rebooting into setup..."
        }
    }

    Start-Sleep -Seconds 15
}

if (-not $wasDown) {
    Write-Host "`nWinRM never dropped. Setup may still be staging."
}

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
    Write-Host "`n Post-upgrade version: $($ver.ProductName), Edition: $($ver.EditionID), DisplayVersion: $($ver.DisplayVersion), Build $($ver.CurrentBuild).$($ver.UBR)"

    $isWin11 = ($ver.ProductName -like '*Windows 11*')
    $isEdu   = ($ver.ProductName -like '*Education*' -or $ver.EditionID -like '*Education*')
    $isTargetVersion = ($ver.DisplayVersion -in @('23H2','24H2'))

    if ($isWin11 -and $isEdu -and $isTargetVersion) {
        Write-Host "`n Upgrade completed successfully to Windows 11 Education $($ver.DisplayVersion)."
    } else {
        Write-Host "`n Upgrade status uncertain. Version does not match expected 23H2 or 24H2 Education."
    }
} catch {
    Write-Host "`n Could not query post-upgrade version. The device may still be finalizing."
}

Write-Host "`n Done."
