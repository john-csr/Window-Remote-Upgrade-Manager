# Windows Upgrade Orchestration Suite

**Author:** John C.  
**Version:** 1.0  
**Date:** September 2025  

---

## Overview

This solution provides a streamlined, remote-controlled method for performing in-place Windows upgrades across multiple devices in an enterprise or educational environment. It consists of two PowerShell scripts and an HTA (HTML Application) interface that allows IT administrators to launch upgrades, monitor progress, and verify completion — all from a central console.

---

## Components

### 1. `winupgrade.ps1`

**Purpose:**  
Remotely initiates a silent in-place upgrade on a target machine using `setup.exe`.

**Functions:**
- Validates remote connectivity and disk space
- Confirms edition compatibility
- Locates upgrade source files
- Creates and launches a scheduled task to run `setup.exe /auto upgrade /quiet`
- Monitors task status and confirms launch

**Use Case:**  
An IT admin needs to upgrade a fleet of student laptops to Windows 11 23H2 and 24H2 without user interaction. This script ensures the upgrade is launched silently and reliably, even if the user is logged in.

---

### 2. `upgr-progressmon.ps1`

**Purpose:**  
Monitors the upgrade progress of a remote machine in real time.

**Functions:**
- Connects to the target machine via UNC path
- Checks for presence of setup logs (`setupact.log`)
- Displays a dynamic progress bar and status messages
- Tracks upgrade phases (pre-reboot, setup environment, post-upgrade)

**Use Case:**  
While upgrades are running silently, the admin wants visual feedback on progress. This script provides a live console view of each machine’s upgrade status.

---

### 3. HTA Interface

**Purpose:**  
Provides a user-friendly GUI for launching and managing upgrades.

**Functions:**
- `CopySetupFiles`: Copies setup files from a network share to the target machine’s `C:\WinSetup` folder using `robocopy`
- `VerifyFiles`: Uses WMI to confirm presence of setup files on the remote machine
- `LaunchUpgrade`: Executes `winupgrade.ps1` with the specified computer name
- `LaunchMonitor`: Executes `upgr-progressmon.ps1` to track upgrade progress

---

## Requirements

- Admin rights on target machines  
- WinRM enabled and accessible  
- Network share containing setup files (e.g., `\\server1\share1\24h2`)  
- PowerShell execution policy set to allow script execution  

**Use Case:**  
An IT technician uses the HTA to select a machine, copy setup files, verify readiness, launch the upgrade, and monitor progress — all without touching the remote device.

---

## Deployment Notes

- Scripts should reside in `H:\Scripts` (can be adjusted in the HTA code as needed)  
- Setup files must be staged in a UNC-accessible folder  
- HTA should be run from a technician’s workstation with network access to all targets  

---

## Benefits

- Centralized control of upgrades  
- Silent, user-transparent execution  
- Real-time monitoring  
- Minimal disruption to end users  
- Scalable across dozens or hundreds of machines  
