# ProfileMigration.ps1

ProfileMigration.ps1 is a PowerShell 5.1 GUI tool for Windows profile migration. It supports domain, local, and AzureAD/Entra ID moves, registry hive SID rewriting, multi-threaded file operations, and a modern Windows Forms UI. Production-ready for Windows 11 24H2.

## Features
- Export/import/merge user profiles (local, domain, AzureAD/Entra ID)
- Robust AzureAD/Entra ID detection (SID pattern, even for DOMAIN\user format)

- Registry hive SID rewrite (binary + string)
- Multi-threaded file copy (Robocopy)
- 7-Zip compression (auto-detect/install)

- Full HTML migration report
- Modern, flat Windows Forms UI
- Detailed logging (DEBUG/INFO/WARN/ERROR)

- Progress bar, status, and log viewer
- Domain join after import (optional)
- Winget app install (optional)

## Usage
**Run as Administrator:**
```powershell

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\ProfileMigration.ps1
```

Or right-click → "Run with PowerShell" (auto-elevates)

## Key Improvements (2025)
- Accurate AzureAD/Entra ID detection for all user formats (including DOMAIN\user)

- Improved error and success dialogs (taller, all info visible)
- Modernized UI and status feedback
- Debug logging for troubleshooting (removed in production)

## Documentation
- [USER-GUIDE.md](USER-GUIDE.md): Step-by-step usage
- [CONFIGURATION.md](CONFIGURATION.md): All config options
- [TECHNICAL-DOCS.md](TECHNICAL-DOCS.md): Architecture, workflows

- [FAQ.md](FAQ.md): Troubleshooting

## Support
See [Logs/](Logs/) for daily logs. For issues, see FAQ or contact IT support.
# Windows Profile Migration Tool

**Version:** December 2025  
**Tested on:** Windows 11 25H2 (26200.7462)  
**Status:** Production Ready

## Overview

The Windows Profile Migration Tool is an enterprise-grade PowerShell application that provides seamless migration of user profiles between Windows systems. It features a modern GUI, multi-threaded operations, comprehensive error handling, and detailed reporting capabilities.

## Key Features

### Core Functionality
- ✅ **Full Profile Export/Import** - Complete user profile migration with all settings preserved
- ✅ **Multi-Threaded Performance** - 2-3x faster compression using all CPU cores
- ✅ **Local & Domain Users** - Support for both local accounts and Active Directory users
- ✅ **AzureAD/Entra ID Support** - Seamless migration of Microsoft Entra ID (AzureAD) profiles
- ✅ **Hash Verification** - SHA-256 integrity checking for exported archives
- ✅ **Merge or Replace Modes** - Flexible import options for existing profiles
- ✅ **Automatic Backup** - Timestamped backups before any destructive operations
- ✅ **Winget Package Migration** - Automatically reinstall applications on target system

### Advanced Features
- 🔧 **Smart Exclusions** - Automatic filtering of temp files, caches, and Outlook OST files
- 🔧 **Registry Hive Translation** - Automatic SID rewriting for cross-system compatibility
- 🔧 **Junction Handling** - Proper recreation of Windows profile junctions
- 🔧 **Logged-in User Detection** - Prevents corruption from active profiles
- 🔧 **Profile Verification** - Validates registry hive integrity before completion
- 🔧 **HTML Reports** - Detailed migration reports with statistics and diagnostics

### User Experience
- 🎨 **Modern GUI** - Clean, intuitive interface with real-time progress tracking
- 🎨 **Live Logging** - Searchable log viewer with level filtering
- 🎨 **Profile Size Display** - Shows profile sizes in selection dropdown
- 🎨 **Password Strength Indicator** - Visual feedback for local user creation
- 🎨 **Cancellable Operations** - Stop long-running tasks at any time

### Performance Optimizations
- ⚡ **CPU Auto-Detection** - Automatically scales thread count to available cores
- ⚡ **7-Zip Multi-Threading** - Parallel compression (2-3x faster than single-threaded)
- ⚡ **Robocopy Optimization** - Dynamic 8-32 threads for file operations (60% faster)
- ⚡ **Smart Progress Updates** - Non-blocking UI with real-time status

## System Requirements

### Minimum Requirements
- **OS:** Windows 10 20H2 or later / Windows 11
- **PowerShell:** 5.1 or later
- **RAM:** 4 GB minimum (8 GB recommended for large profiles)
- **CPU:** 2+ cores (4+ recommended for optimal performance)
- **Disk Space:** 2x the size of the largest profile to migrate
- **Permissions:** Administrator rights required

### Dependencies
- **7-Zip** - Required for compression (automatically checked and prompted for installation)
- **.NET Framework** - 4.7.2 or later (included in Windows 10/11)
- **Windows PowerShell** - 5.1 (built-in to Windows)

## Quick Start Guide

### 1. Launch the Tool
Right-click `ProfileMigration.ps1` and select **Run with PowerShell**

Or from PowerShell (as Administrator):
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\ProfileMigration.ps1
```

### 2. Export a Profile
1. Select the user from the dropdown
2. Click **Export** button
3. Choose save location
4. Wait for compression to complete
5. Copy the `.zip` file and `.sha256` hash file to target system

### 3. Import a Profile
1. Click **Browse** to select the exported `.zip` file
2. Enter the target username:
   - Local user: `username`
   - Domain user: `DOMAIN\username`
   - **AzureAD user: `AzureAD\username`**
3. Choose **Merge** or **Replace** mode (if profile exists)
4. Click **Import** button
5. Follow prompts for user creation (if needed)
6. **For AzureAD profiles**: Join device to Entra ID if prompted, sign in with work/school account
7. Reboot and login after completion

### 4. AzureAD/Entra ID Profiles (Special Instructions)

**Exporting AzureAD profiles:**
- Works automatically - the tool detects AzureAD accounts (SID starts with S-1-12-1)
- Profile is tagged as AzureAD in the manifest

**Importing AzureAD profiles:**
1. Use format: `AzureAD\username` (e.g., `AzureAD\john.doe`)
2. If the device is not AzureAD joined, the tool will:
   - Open Settings → Access work or school
   - Guide you to select "Join this device to Microsoft Entra ID"
   - Wait for you to complete the join process
3. **Important**: Sign in with the AzureAD account at least once before importing
4. The tool will then import the profile to the existing AzureAD account
5. After the account has been logged in once use the drop down to select targer user.  Will show as Tenant\user

## Documentation

- **[User Guide](USER-GUIDE.md)** - Detailed step-by-step instructions for all operations
- **[Technical Documentation](TECHNICAL-DOCS.md)** - Architecture, internals, and troubleshooting
- **[Configuration Guide](CONFIGURATION.md)** - Advanced settings and customization
- **[FAQ](FAQ.md)** - Common questions and solutions

## Use Cases

### IT Administrators
- **Computer Replacements** - Migrate users to new hardware seamlessly
- **OS Upgrades** - Transfer profiles from Windows 10 to Windows 11
- **Hardware Failures** - Recover user profiles from backup drives
- **Department Transfers** - Move users between domains/workgroups

### End Users (Guided by IT)
- **Desktop to Laptop** - Take your profile on the go
- **Profile Corruption Recovery** - Restore from working backup
- **Shared Computer Setup** - Import personal settings to shared workstations

## What Gets Migrated?

### ✅ Included in Migration
- Desktop files and shortcuts
- Documents, Pictures, Videos, Music, Downloads
- Browser bookmarks and settings (Edge, Chrome, Firefox)
- Application settings and data in AppData
- Registry user hive (NTUSER.DAT) with all preferences
- Start Menu customization
- Taskbar pinned items
- File Explorer favorites and recent locations
- Windows theme and wallpaper
- Installed fonts (user-level)
- Network mapped drives (credentials not included)
- Printer configurations

### ❌ Automatically Excluded
- Temporary files and caches
- Outlook OST files (rebuilt from Exchange)
- Windows search index
- Thumbnail caches
- Browser caches
- Log files
- Recycle Bin contents
- System-specific junctions

## Safety Features

### Backup Protection
- **Automatic Backups** - Creates timestamped backup before any import
- **Rollback Capability** - Original profile preserved at `C:\Users\username.backup_TIMESTAMP`
- **Non-Destructive Merge** - Merge mode keeps existing profile intact

### Validation Checks
- **User Logged-In Detection** - Warns if target user is active
- **Profile Mount Detection** - Checks registry for active user hives
- **Hash Verification** - Validates ZIP integrity on import
- **Hive Load Test** - Ensures NTUSER.DAT can be mounted before completion
- **Disk Space Check** - Verifies sufficient space before operations

### Error Recovery
- **Operation Logging** - Complete audit trail in timestamped log files
- **Graceful Cancellation** - Clean shutdown of multi-threaded operations
- **Diagnostic Reporting** - HTML reports with full operation details
- **Stale Registry Cleanup** - Automatic removal of orphaned profile entries

## Performance

### Typical Migration Times
| Profile Size | Export Time | Import Time | Total Time |
|--------------|-------------|-------------|------------|
| 5 GB         | 2-3 min     | 2-3 min     | ~5 min     |
| 20 GB        | 5-8 min     | 5-8 min     | ~12 min    |
| 50 GB        | 12-20 min   | 12-20 min   | ~30 min    |
| 100 GB       | 25-40 min   | 25-40 min   | ~60 min    |

*Times vary based on CPU cores, disk speed, and file types*

### Optimization Settings
- **CPU Cores:** Automatically detected and utilized
- **7-Zip Threads:** All cores (configurable)
- **Robocopy Threads:** 8-32 dynamically scaled by CPU count
- **Compression Level:** Balanced for speed and size

## Troubleshooting

### Common Issues

**"7-Zip not found"**
- Install 7-Zip from https://www.7-zip.org/
- Ensure installed to default location: `C:\Program Files\7-Zip\7z.exe`

**"Profile not found" error**
- Verify username is correct
- Check if profile exists at `C:\Users\[username]`
- For domain users, use format: `DOMAIN\username`

**"User appears to be logged on"**
- Log out the user before migration
- Check Task Manager for active user processes
- Use `qwinsta` command to verify no active sessions

**Import creates temporary profile (TEMP folder)**
- NTUSER.DAT failed to load - check log for details
- May indicate corrupted export or permissions issue
- Try re-exporting from source system

**"Access Denied" during import**
- Ensure running as Administrator
- Check target folder permissions
- Verify no antivirus blocking operations

See **[TECHNICAL-DOCS.md](TECHNICAL-DOCS.md)** for advanced troubleshooting.

## Support & Contributions

### Getting Help
1. Check the **[FAQ](FAQ.md)** for common solutions
2. Review log files in migration folder
3. Examine HTML report for detailed diagnostics
4. Search existing issues

### Reporting Bugs
Include in your report:
- Windows version (run `winver`)
- PowerShell version (`$PSVersionTable.PSVersion`)
- Complete log file contents
- Steps to reproduce the issue
- Expected vs actual behavior

## License

This tool is provided as-is for enterprise IT use. See LICENSE file for details.

## Version History

### December 2025 - Current Version
- Multi-threaded compression (2-3x performance improvement)
- HTML migration reports with detailed statistics
- CPU core auto-detection and dynamic threading
- Profile size display in UI dropdown
- Hash verification for data integrity
- Merge mode for non-destructive imports
- Winget package migration support
- Improved error handling and validation
- Modern UI with live logging
- Comprehensive documentation

---

**For detailed instructions, see the [User Guide](USER-GUIDE.md)**
