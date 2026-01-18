# Windows Profile Migration Tool v2.12.25

**Version:** v2.12.25 (January 2026)  
**Tested on:** Windows 11 25H2 (26200.7623)  
**Status:** Production Ready

## Overview

The Windows Profile Migration Tool is an enterprise-grade PowerShell application that provides seamless migration of user profiles between Windows systems. It features a modern GUI, multi-threaded operations, comprehensive error handling, and detailed reporting capabilities.

ProfileMigration.ps1 is a PowerShell 5.1 GUI tool for Windows profile migration supporting domain, local, and AzureAD/Entra ID accounts with **in-place profile conversion**, **automatic unjoin/join for domain↔AzureAD transitions**, registry hive SID rewriting, multi-threaded file operations, and a modern Windows Forms UI.

---

## Key Features

### Core Functionality
- ✅ **Full Profile Export/Import** - Complete user profile migration with all settings preserved
- ✅ **Profile Conversion** - Convert profiles between Local ↔ Domain ↔ AzureAD without re-export
- ✅ **Automatic Unjoin/Join** - Seamless domain↔AzureAD transitions with automatic device management
- ✅ **Multi-Threaded Performance** - 2-3x faster compression using all CPU cores
- ✅ **Hash Verification** - SHA-256 integrity checking for exported archives
- ✅ **Merge or Replace Modes** - Flexible import options for existing profiles
- ✅ **Automatic Backup** - Timestamped backups before any destructive operations
- ✅ **Winget Package Migration** - Automatically reinstall applications on target system

### Advanced Features
- 🔧 **Microsoft Graph Integration** - Resolve AzureAD SIDs without requiring prior user login
- 🔧 **Smart Exclusions** - Automatic filtering of temp files, caches, and Outlook OST files
- 🔧 **Registry Hive Translation** - Automatic SID rewriting for cross-system compatibility
- 🔧 **Enhanced Profile Cleanup Wizard** - Remove large items and deduplicate files before export
- 🔧 **Junction Handling** - Proper recreation of Windows profile junctions
- 🔧 **Logged-in User Detection** - Prevents corruption from active profiles
- 🔧 **Profile Verification** - Validates registry hive integrity before completion
- 🔧 **HTML Reports** - Detailed migration reports with statistics and diagnostics

### User Experience
- 🎨 **Modern GUI** - Clean, intuitive interface with real-time progress tracking
- 🎨 **Light/Dark Theme Support** - Automatic Windows theme detection with manual toggle
- 🎨 **Live Logging** - Searchable log viewer with level filtering (DEBUG/INFO/WARN/ERROR)
- 🎨 **Profile Size Display** - Shows profile sizes in selection dropdown
- 🎨 **Password Strength Indicator** - Visual feedback for local user creation
- 🎨 **Debug Mode** - Verbose export analysis and artifact preservation
- 🎨 **Cancellable Operations** - Stop long-running tasks at any time

### Performance Optimizations
- ⚡ **CPU Auto-Detection** - Automatically scales thread count to available cores
- ⚡ **7-Zip Multi-Threading** - Parallel compression (2-3x faster than single-threaded)
- ⚡ **Robocopy Optimization** - Dynamic 8-32 threads for file operations (60% faster)
- ⚡ **Smart Progress Updates** - Non-blocking UI with real-time status
- ⚡ **Helper Function Library** - 81 reusable functions for improved code maintainability

---

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
- **Microsoft Graph PowerShell SDK** - Auto-installed when needed for AzureAD conversions

---

## Quick Start Guide

### 1. Launch the Tool

**Method 1: PowerShell Command Line (Recommended)**

Open PowerShell as Administrator and run:
```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\ProfileMigration.ps1"
```

### 2. Export a Profile
1. Select the user from the dropdown
2. Click **Export** button
3. Choose save location
4. Wait for compression to complete

### 3. Import a Profile
1. Click **Browse** to select the exported `.zip` file
2. Enter the target username:
   - Local user: `username`
   - Domain user: `DOMAIN\username`
   - **AzureAD user: `AzureAD\username` OR `user@domain.com` (UPN format)**
3. Click **Import** button
4. Choose **Merge** or **Replace** mode (if profile exists)
5. Follow prompts for user creation (if needed)
6. **For AzureAD profiles**: Join device to Entra ID if prompted, sign in with work/school account
7. Reboot and login after completion

### 4. Profile Conversion

**Convert existing profiles without re-export:**
1. Click **Convert Profile** button
2. Select source profile type and username
3. Select target profile type and username
4. Supported conversions:
   - Local ↔ AzureAD
   - Domain ↔ AzureAD
   - Local ↔ Domain
5. The tool will:
   - Copy profile to new location (if needed)
   - Update registry entries
   - Rewrite SIDs in NTUSER.DAT
   - Apply correct permissions
   - Handle domain/AzureAD unjoin/join automatically

**Universal Repair:**
- **In-Place Repair** - Select the **Source** and **Target** as the SAME user to trigger repair mode
- **Zero Data Loss** - No files are moved or copied
- **Fixes Profile Issues** - Resets permissions, rewrites NTUSER.DAT SIDs, and re-registers AppX apps
- **Solves "Temporary Profile"** - Corrects registry pointers and ACLs enabling successful login

**AzureAD Conversion Features:**
- **No prior login required** - Uses Microsoft Graph API to resolve AzureAD SIDs
- **Email format supported** - Enter AzureAD username as UPN (e.g., `user@domain.com`)
- **Auto-join guidance** - Tool provides step-by-step instructions if device isn't AzureAD-joined
- **One-time setup** - Microsoft Graph module auto-installs when first needed

---

## Documentation

- **[User Guide](USER-GUIDE.md)** - Detailed step-by-step instructions for all operations
- **[Technical Documentation](TECHNICAL-DOCS.md)** - Architecture, internals, and troubleshooting
- **[Configuration Guide](CONFIGURATION.md)** - Advanced settings and customization
- **[FAQ](FAQ.md)** - Common questions and solutions
- **[Functions Reference](FUNCTIONS.md)** - Complete documentation of all 81 functions

---

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

---

## Performance

### Typical Migration Times
| Profile Size | Export Time | Import Time | Total Time |
|--------------|-------------|-------------|------------|
| 5 GB         | 2-3 min     | 2-3 min     | ~5 min     |
| 20 GB        | 5-8 min     | 5-8 min     | ~12 min    |
| 50 GB        | 12-20 min   | 12-20 min   | ~30 min    |
| 100 GB       | 25-40 min   | 25-40 min   | ~60 min    |

*Times vary based on CPU cores, disk speed, and file types*

---

## Restore from Backup
There are two ways to restore a profile from a backup created by this tool:

### Method 1: Automated Restore (Recommended)
Use the **Import Profile** tab within the tool.
1. Select "Import Profile" tab.
2. Choose **Source Type**: "ZIP Backup".
3. Browse to your backup ZIP file (`.zip`).
4. Select the target user you want to restore to.
5. Click **Start Import**.
   - *This handles file extraction, registry updates, and permission fixes automatically.*

### Method 2: Manual / Emergency Restore
If the tool is unavailable, you can perform a manual restore:
1. **Registry:** Locate the `.reg` file created alongside your backup ZIP (e.g., `ProfileBackup_User_Date.zip.reg`). Double-click it to import the profile registry key.
2. **Files:** Extract the contents of the `.zip` file to the profile path specified in the registry (usually `C:\Users\username`).
3. **Permissions:** Ensure the user has Full Control over their profile folder.

## Version History

### v2.12.25 (Current - January 2026)
- **Refactor**: Standardized Profile Cleanup Wizard logic:
    - Integrated `New-CleanupItem` helper for all 6 cleanup categories.
    - Added support for Duplicate File groups in cleanup helper.
    - Improved code maintainability by removing manual object construction.
- **Reliability**: Resolved syntax errors and improved brace matching in wizard functions.

### v2.12.21 (January 2026)
- **Refactor**: Modernized codebase by implementing helper functions:
    - `Confirm-DomainUnjoin` for consistent UI.
    - `Test-ValidProfilePath` for robust profile validation (checking `NTUSER.DAT`).
    - `Mount-RegistryHive` for safer hive operations.
- **Restoration**: Restored utility functions `Get-FolderSize`, `Test-PathWithRetry`, and `Convert-SIDToAccountName` for future use.

### v2.12.19 (January 2026)
**Safety: Conversion Cleanup**
- **Fixed:** Added robust cleanup and cancellation logic for all profile conversion types (Local/Domain/AzureAD)
- **Benefit:** Prevents data corruption and partial folders if conversion is cancelled or fails during file transfer

### v2.12.18 (January 2026)
**Fix: Import Cleanup**
- **Fixed:** Cancelling a Merge Mode import now correctly removes temporary extraction folders
- **Benefit:** Prevents disk clutter from failed/cancelled imports

### v2.12.17 (January 2026)
**Fix: Cancel Cleanup**
- **Fixed:** Cancelling an export now correctly deletes the partial/corrupted ZIP file
- **Benefit:** Prevents confusion from leaving 0KB or corrupted backups on disk

### v2.12.16 (January 2026)
**Fix: Skip Cleanup Behavior**
- **Fixed:** Clicking "Skip Cleanup" in the export wizard no longer cancels the operation
- **Benefit:** Users can now bypass the cleanup step and proceed directly to export as intended

### v2.12.15 (January 2026)
**Code Refactor: UI Consistency**
- **Refactor:** Replaced legacy custom warning dialog in Import workflow with standard `Show-ModernDialog`
- **Cleanup:** Removed ~65 lines of redundant UI code

### v2.12.14 (January 2026)
**Bug Fix: AppX Repair Concurrency**
- **Fixed:** Prevented "cross-fire" issue where any user login would trigger and consume pending repair scripts for other users
- **Method:** Added smart username validation to generated repair script; it now exits gracefully (without self-deletion) if the current user doesn't match the target

### v2.12.13 (January 2026)
**Feature Enhancement: Backup Location Prompt**
- **New:** Users are now prompted to select a destination for profile backups during conversion
- **Benefit:** Allows saving backups to external drives or custom locations directly
- **Fallback:** Logic handles cancellations gracefully (skip or default to `C:\Users\ProfileBackups`)

### v2.12.12 (January 2026)
**Bug Fixes: UI & Reporting**
- **UI Fix:** Standardized newline characters in dialog messages for consistent formatting
- **Reporting Fix:** Corrected an issue where "AzureAD status: Not Changed" was reported even after a successful unjoin.

### v2.12.11 (January 2026)
**Feature Enhancement: Simplified AppX Repair**
- **Enhanced:** AppX repair now uses logged-in user context naturally
- **Fixed:** Eliminates username mismatch errors completely by removing fragile validation logic
- **Simplify:** Removed complex username string parsing/extraction
- **Robustness:** Ensures reliable AppX repair for all user types (Domain/Local/AzureAD)

### v2.12.10 (January 2026)
**Bug Fixes & Refinements**
- **AzureAD Fix:** Fixed AzureAD username extraction in AppX repair script (added support for UPN/email formats).
- **Local Repair Fix:** Refined logic to only trigger "Repair Mode" when Source User == Target User.
- **User Deletion:** Added option to delete source local user after successful conversion (with safety sanitization).

### v2.12.9 (January 2026)
**Complete Dotted Username Logic Removal:**
- **Legacy Code Cleanup:** Removed ~198 lines of obsolete dotted username workaround code
- **Functions Modified:** Updated 6 functions (`New-ConversionReport`, `Convert-LocalToDomain`, `Convert-DomainToLocal`, `Convert-AzureADToLocal`, `Rewrite-HiveSID`)
- **Cleaner Code:** Eliminated disabled code blocks, conditional checks, and legacy comments

### v2.12.8 (January 2026)
**Feature: Universal Profile Repair**
- **Domain & AzureAD Repair:** Enabled in-place repair for Domain and AzureAD profiles (Source = Target).
- **Performance:** Optimized ACL application (replaced `takeown` with `icacls`, fixing hangs).
- **Fixes:** Resolves "Temporary Profile" issues via recursive permissions and log cleanup.

### v2.12.7 (January 2026)
**Critical Fix: Temporary Profile Resolution**
- **Recursive ACLs:** Updated `Set-ProfileFolderAcls` to grant user permissions recursively (/T).
- **Hive Cleanup:** Added aggressive deletion of `NTUSER.DAT` transaction logs (.LOG*, .blf) to prevent "dirty hive" load failures.

### v2.12.6 (January 2026)
**Hotfix / Feature Release:**
- **Local-to-Local Repair**: Enabled in-place conversion for same-user scenarios.
- **ACL Ownership**: Fixed `NTUSER.DAT` and profile folder ownership.
- **Registry Safety**: Added safeguard to prevent deleting the registry key during self-repair operations.


### v2.12.11 (January 2026)
- **Improvement**: Simplified AppX repair logic by removing unnecessary username validation
- **Fix**: Resolves all potential username mismatch issues by relying on naturally correct user context

### v2.12.10 (January 2026)
- **Bug Fix**: Fixed AzureAD username extraction in AppX repair script
- **Issue**: AppX repair was failing for AzureAD conversions due to username format mismatch
- **Solution**: Added proper extraction logic for Domain, AzureAD UPN, and email formats

### v2.12.9 (January 2026)
- **Code Cleanup**: Removed ~198 lines of legacy dotted username workaround code
- **Functions Modified**: 6 functions updated to remove dotted username logic
- **AppX Repair**: Now relies entirely on AppX repair mechanism
- **Function Count**: 81 functions documented

### v2.10.109 (January 2026)
- **Code Refactoring**: Added 8 new helper functions
- **Documentation**: Created comprehensive testing plan and implementation plan
- **Function Count**: Increased from 70 to 79 functions
- **Code Quality**: Established foundation for reducing code duplication

### v2.10.108 (January 2026)
- **Code Cleanup**: Removed duplicate `Get-ProfileType` function definition

### v2.10.107 (January 2026)
- **Code Cleanup**: Removed duplicate `Get-LocalProfiles` function definition

### v2.10.104-106 (January 2026)
- **Code Cleanup**: Removed 3 unused functions
- **Reduced code by ~178 lines**

### v2.10.100-103 (January 2026)
- **Bug Fix**: Fixed report auto-open scope
- **Bug Fix**: Fixed log viewer filter to use "WARN"
- **Bug Fix**: Fixed log viewer filter pattern matching

### v2.10.96-99 (January 2026)
- **Critical Fix**: Resolved black screen issue for dotted usernames
- **Root Cause**: Profile path incorrectly included domain prefix
- **Solution**: Modified profile path logic

### v2.10.73-95 (January 2026)
- **Critical Fixes**: Registry handle leak fixes
- **Hive Unload**: Improved reliability with retry logic
- **Transaction Logs**: Enhanced cleanup for NTUSER.DAT and UsrClass.dat
- **Binary SID Replacement**: Refined with safety mechanisms

### v2.7.2-2.10.72 (December 2025 - January 2026)
- **Automatic Domain↔AzureAD Unjoin/Join**: Seamless transitions
- **Profile Conversion**: Convert between Local/Domain/AzureAD
- **Microsoft Graph Integration**: Resolve SIDs without prior login
- **Enhanced UI**: Light/Dark theme support
- **Profile Cleanup Wizard**: Identify and remove bloated files
- **Debug Mode**: Verbose logging and artifact preservation

### v1.0.0 (December 2025)
- **Initial Release**: Core export/import functionality
- **Multi-threaded Performance**: 2-3x faster compression
- **HTML Reports**: Detailed migration statistics
- **Modern UI**: Windows Forms interface

---

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

See **[FAQ.md](FAQ.md)** and **[TECHNICAL-DOCS.md](TECHNICAL-DOCS.md)** for advanced troubleshooting.

---

## Support & Contributions

### Getting Help
1. Check the **[FAQ](FAQ.md)** for common solutions
2. Review log files in `Logs/` directory
3. Examine HTML report for detailed diagnostics
4. Contact IT support with log files

### Reporting Issues
Include in your report:
- Windows version (run `winver`)
- PowerShell version (`$PSVersionTable.PSVersion`)
- Complete log file contents
- Steps to reproduce

---

**For detailed instructions, see the [User Guide](USER-GUIDE.md)**

**Last Updated:** January 17, 2026
