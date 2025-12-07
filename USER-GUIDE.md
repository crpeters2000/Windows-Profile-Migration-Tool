# User Guide - Windows Profile Migration Tool

## Table of Contents
1. [Getting Started](#getting-started)
2. [Exporting a Profile](#exporting-a-profile)
3. [Importing a Profile](#importing-a-profile)
4. [Domain Operations](#domain-operations)
5. [Advanced Features](#advanced-features)
6. [Best Practices](#best-practices)

---

## Getting Started

### Prerequisites Check
Before starting, ensure you have:
- ✅ Administrator rights on both source and target computers
- ✅ 7-Zip installed (tool will prompt if missing)
- ✅ Sufficient disk space (2x profile size recommended)
- ✅ Network/USB drive for transferring exported profiles

### Launching the Tool

**Method 1: Right-Click (Recommended)**
1. Right-click `ProfileMigration.ps1`
2. Select **Run with PowerShell**
3. If prompted, click **Yes** to allow administrator elevation

**Method 2: PowerShell Console**
1. Open PowerShell as Administrator
2. Navigate to tool directory:
   ```powershell
   cd "C:\Path\To\ProfileMigration"
   ```
3. Set execution policy (if needed):
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
4. Run the script:
   ```powershell
   .\ProfileMigration.ps1
   ```

### First Launch

When the tool starts, you'll see:
- **User Selection Dropdown** - Shows all local user profiles with sizes
- **Export/Import Buttons** - Main operation controls
- **Log Viewer** - Real-time operation logging
- **Progress Bar** - Visual progress indicator

---

## Exporting a Profile

### Step-by-Step Export Process

#### 1. Select Source User
1. Click the **user dropdown** at the top
2. Select the profile to export
3. Profile size is displayed: `username - [12.5 GB]`

**Notes:**
- System accounts (Administrator, Guest) are hidden by default
- Sizes are calculated quickly (may be approximate for very large profiles)
- Currently logged-in users will show a warning if selected

#### 2. Start Export
1. Click the **Export** button
2. Choose save location in file dialog
3. Suggested filename: `username-Export-YYYYMMDD_HHMMSS.zip`

#### 3. Monitor Progress
Watch the progress bar and status messages:
- `Checking prerequisites...` - Validating 7-Zip, disk space
- `Calculating profile size...` - Determining archive size
- `Compressing profile... XX%` - Multi-threaded compression in progress
- `Generating hash file...` - Creating SHA-256 checksum
- `Creating migration report...` - Building HTML report

#### 4. Export Complete
When finished, you'll have:
- **username-Export-YYYYMMDD_HHMMSS.zip** - Compressed profile
- **username-Export-YYYYMMDD_HHMMSS.zip.sha256** - Hash file for verification
- **Migration_Report_Export_YYYYMMDD_HHMMSS.html** - Detailed report
- **Export-username-YYYYMMDD_HHMMSS.log** - Operation log

**✅ Success Checklist:**
- ZIP file size seems reasonable (check against original profile)
- SHA256 hash file exists alongside ZIP
- HTML report opens and shows success
- No errors in log viewer

---

## Importing a Profile

### Preparation

**Before Import:**
1. ✅ Copy exported ZIP and SHA256 files to target computer
2. ✅ Ensure target user is NOT logged in
3. ✅ Have user password ready (if creating new local account)
4. ✅ Verify sufficient disk space on C: drive

### Step-by-Step Import Process

#### 1. Select ZIP File
1. Click **Browse** button
2. Navigate to exported `.zip` file
3. Click **Open**
4. ZIP path shows in text field

**Auto-Detection:**
- Tool extracts username from filename: `user-Export-20251207_143022.zip` → `user`
- Username field is auto-populated
- Edit username if importing to different account

#### 2. Enter Target Username
Enter username in one of these formats:

**For Local Users:**
```
username
```
or
```
COMPUTERNAME\username
```

**For Domain Users:**
```
DOMAIN\username
```

#### 3. Start Import
Click the **Import** button

The tool will check:
- ✅ ZIP file exists and is readable
- ✅ Hash verification (if .sha256 file present)
- ✅ Target user logged-in status
- ✅ Existing profile detection

### Handling Different Scenarios

#### Scenario A: New User (Doesn't Exist)
**For Local Users:**
1. Tool detects user doesn't exist
2. Password creation dialog appears:
   - Enter password (twice for confirmation)
   - Shows strength indicator (weak/moderate/strong)
   - Optional: Check "Add to Administrators group"
3. Click **Create User**
4. Import proceeds to extraction

**For Domain Users:**
1. Domain credentials prompt appears
2. Enter domain admin username and password
3. Tool queries Active Directory for user
4. If user exists in domain, import proceeds
5. If not found, operation stops with error

#### Scenario B: Existing Profile - Replace Mode
If profile exists at `C:\Users\username`:

1. **Profile Exists** dialog appears with two options:
   
   **REPLACE** (Recommended for clean migration):
   - Backs up existing profile to `C:\Users\username.backup_TIMESTAMP`
   - Deletes existing profile
   - Extracts imported profile to clean directory
   - Rewrites registry SIDs
   - **Use when:** Migrating from old computer, profile corrupted, want fresh start
   
   **MERGE** (Advanced - preserves existing settings):
   - Backs up existing profile (safety)
   - Extracts to temporary location
   - Copies files into existing profile (doesn't overwrite newer files)
   - Keeps existing NTUSER.DAT (preserves current registry settings)
   - **Use when:** Adding files from backup, supplementing existing profile, keeping current settings

2. Choose mode and click button

#### Scenario C: Domain User First Login
1. Profile extracted to `C:\Users\username`
2. Registry permissions configured for domain
3. SID translation performed
4. User can login on next boot

### Post-Import Steps

#### 4. Monitor Import Progress
Progress indicators show:
- `Verifying ZIP integrity...` - Hash check (if enabled)
- `Checking if user is logged on...` - Safety validation
- `Backing up existing profile...` - Creating safety backup (if exists)
- `Extracting ZIP archive... XX%` - Multi-threaded extraction
- `Merging profile files...` - File copy (merge mode only)
- `Applying permissions...` - ACL and SID rewriting
- `Registering profile in Windows...` - Registry configuration
- `Recreating profile junctions...` - Fixing special folders

#### 5. Completion
When import finishes:
- **Success Dialog** appears with instructions
- **HTML Report** generated and auto-opened (if enabled)
- **Log File** saved with complete details

**✅ Post-Import Checklist:**
- Read success message carefully
- Note the reboot requirement
- Review HTML report for any warnings
- Keep backup folder until user confirms everything works

---

## Domain Operations

### Domain Join Functionality

The tool includes integrated domain join capability:

#### Prerequisites
- Target computer not currently domain-joined
- Domain admin credentials available
- Network connectivity to domain controller
- DNS configured correctly

#### Steps to Join Domain
1. Navigate to **Domain Join** tab in tool
2. Enter domain name: `CONTOSO.COM` or `CONTOSO`
3. Click **Test Domain Connectivity**
   - Green: Domain reachable, proceed
   - Red: Check network/DNS, cannot join
4. Enter computer name (optional, or keep current)
5. Click **Join Domain** button
6. Enter domain admin credentials when prompted
7. Choose restart option:
   - Restart Now
   - Restart after delay (10-60 seconds)
   - Don't restart (manual restart later)

#### Post-Join
- Computer restarts and joins domain
- Domain users can now login
- Import domain profiles using `DOMAIN\username` format

### Domain User Migration
When importing for domain users:
1. Ensure computer is domain-joined first
2. Use format: `DOMAIN\username` in import
3. Provide domain admin credentials when prompted
4. User account is queried from Active Directory
5. Profile configured with correct domain SID
6. User can login immediately after reboot

---

## Advanced Features

### Hash Verification
**What it does:** Validates ZIP integrity using SHA-256 checksum

**Usage:**
- Automatic if `.sha256` file exists alongside ZIP
- Import shows: `Verifying ZIP integrity...`
- If hash mismatch:
  - Warning dialog appears
  - Option to continue anyway or cancel
  - Logged as security warning

**When to use:**
- Transferring over network (detect corruption)
- Long-term archive storage (detect bit rot)
- Security compliance requirements

### Winget Package Migration
**What it does:** Detects and reinstalls user-installed applications

**Automatic Detection:**
- Export scans Winget package list
- Saves to `Winget-Packages.json`
- Import detects file and prompts

**Import Dialog:**
- Shows list of detected applications
- Check/uncheck packages to install
- Click Install to download and install selected apps
- Progress shown in log viewer

**Limitations:**
- Requires Windows 10 1809+ or Windows 11
- Requires internet connection
- Only reinstalls apps available in Winget catalog
- Settings/data may need reconfiguration

### Log Filtering
**Built-in log viewer features:**

1. **Level Filtering Dropdown:**
   - ALL - Shows everything
   - INFO - Normal operations (default)
   - WARN - Warnings and errors
   - ERROR - Only critical failures

2. **Search Box:**
   - Type keyword to filter log entries
   - Real-time filtering
   - Case-insensitive
   - Combines with level filter

3. **Auto-Scroll:**
   - Automatically follows new log entries
   - Temporarily pauses when scrolling up
   - Resumes when scrolled to bottom

### Operation Cancellation
**How to cancel:**
1. Click **Cancel** button (appears during operations)
2. Tool safely stops:
   - Terminates compression/extraction
   - Cleans up temporary files
   - Unmounts any loaded registry hives
   - Logs cancellation

**Safe cancellation points:**
- During compression
- During extraction
- During file copy operations
- During Winget installs

**⚠️ Warning:** Cancelling mid-operation may leave incomplete files

---

## Best Practices

### Before Export
1. ✅ **Close all user applications** - Ensure files aren't locked
2. ✅ **Run Disk Cleanup** - Remove temporary files first
3. ✅ **Check profile size** - Plan for adequate storage/transfer time
4. ✅ **Test 7-Zip** - Export a small test profile first
5. ✅ **Document settings** - Screenshot important configurations

### During Export
1. ✅ **Don't interrupt** - Let process complete fully
2. ✅ **Monitor progress** - Watch for errors in log
3. ✅ **Verify completion** - Check for success message
4. ✅ **Review HTML report** - Look for any warnings
5. ✅ **Test ZIP file** - Try opening to verify not corrupted

### Before Import
1. ✅ **Verify ZIP and hash** - Ensure files copied completely
2. ✅ **Check disk space** - Need 2x profile size
3. ✅ **Logout target user** - Prevent profile corruption
4. ✅ **Close all apps** - Especially Outlook, OneDrive, browsers
5. ✅ **Disable antivirus** - Temporarily (may block file operations)

### During Import
1. ✅ **Don't interrupt** - Critical operation
2. ✅ **Monitor logs** - Watch for errors
3. ✅ **Read prompts carefully** - Merge vs Replace is important
4. ✅ **Keep backup** - Don't delete .backup folder until verified
5. ✅ **Note warnings** - May need manual fixes

### After Import
1. ✅ **REBOOT IMMEDIATELY** - Required for registry changes
2. ✅ **Login as target user** - Verify profile loads correctly
3. ✅ **Check Desktop/Documents** - Ensure files present
4. ✅ **Test applications** - Outlook, browsers, etc.
5. ✅ **Verify network drives** - May need to re-enter credentials
6. ✅ **Check printers** - Test printing
7. ✅ **Reconnect OneDrive** - Re-sync if needed
8. ✅ **Configure Outlook** - Rebuild OST (10-30 min)
9. ✅ **Keep backup 30 days** - Until fully verified
10. ✅ **Save HTML report** - For documentation

### Troubleshooting Tips
1. **Enable DEBUG logging** - Use dropdown to show all messages
2. **Save log files** - Keep for at least 30 days
3. **Check Windows Event Viewer** - System and Application logs
4. **Test with small profile first** - Verify process works
5. **Contact IT if stuck** - Don't force operations

---

## Common Workflows

### Workflow 1: Computer Replacement
**Scenario:** Upgrading user from old PC to new PC

**Source Computer (Old PC):**
1. Login as user one last time
2. Save any open work
3. Logout user
4. Login as admin
5. Run ProfileMigration tool
6. Export user profile
7. Copy ZIP + SHA256 to USB drive
8. Keep old computer running until verified

**Target Computer (New PC):**
1. Complete Windows setup
2. Create temporary admin account
3. Copy ZIP + SHA256 from USB
4. Run ProfileMigration tool
5. Import profile (Replace mode)
6. Choose "Create User" when prompted
7. Set password for user
8. Reboot computer
9. Login as migrated user
10. Verify everything works
11. After 30 days, decommission old PC

### Workflow 2: Domain Migration
**Scenario:** Moving user from workgroup to domain

**Source Computer (Workgroup):**
1. Export local user profile
2. Note current computer name

**Target Computer (Domain):**
1. Join computer to domain first
2. Import profile using local username format
3. User login with local account
4. Export profile again (now as local user)
5. Import profile using `DOMAIN\username` format
6. Provide domain credentials when prompted
7. Reboot
8. Login with domain account

### Workflow 3: Profile Corruption Recovery
**Scenario:** User profile corrupted, using last good backup

1. Login as admin (not affected user)
2. Rename corrupted profile:
   ```
   C:\Users\john → C:\Users\john.corrupt
   ```
3. Import last good backup (Replace mode)
4. Reboot
5. Test user login
6. If working, delete john.corrupt folder
7. If issues, rename back and investigate

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+L` | Clear log viewer |
| `Ctrl+F` | Focus search box |
| `Ctrl+C` | Copy selected log text |
| `Escape` | Cancel current operation |
| `F5` | Refresh user list |

---

## Next Steps

- **Configuration:** See [CONFIGURATION.md](CONFIGURATION.md) for advanced settings
- **Troubleshooting:** See [TECHNICAL-DOCS.md](TECHNICAL-DOCS.md) for error resolution
- **FAQ:** See [FAQ.md](FAQ.md) for common questions

---

**Need Help?** Check the FAQ or contact your IT administrator.
