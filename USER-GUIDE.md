# ProfileMigration.ps1 – User Guide

## Getting Started
1. Run as Administrator:
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\ProfileMigration.ps1"
   ```
2. Select user profile (local, domain, or AzureAD/Entra ID)
3. Click "Set Target User" (robust detection for AzureAD/Entra ID, even if selected as DOMAIN\user)
4. Use Export or Import as needed
## Exporting a Profile
1. Click "Set Target User" and confirm user type
2. Click "Export"
## Importing a Profile
1. Click "Set Target User" and confirm user type
2. Click "Browse" select .zip
2. Click "Import"
## AzureAD/Entra ID Users
- Tool detects AzureAD/Entra ID users by SID pattern (S-1-12-1-...)
- Handles DOMAIN\username and AzureAD\username formats
- Device must be AzureAD joined
## Dialogs & UI
- Modern flat design, color-coded status
- Custom dialogs for errors/success (taller for full info display)
- Improved dialog sizing for all feedback (2025)
## Troubleshooting
- See [FAQ.md](FAQ.md) and [Logs/](Logs/)
# User Profile Migration Tool - User Guide (v2.12.25)

## Table of Contents
1. [Getting Started](#getting-started)
2. [Exporting a Profile](#exporting-a-profile)
3. [Importing a Profile](#importing-a-profile)
4. [Restoring from Backup](#restoring-from-backup)
5. [Domain Operations](#domain-operations)
5. [Repairing a Profile](#repairing-a-profile)
6. [Advanced Features](#advanced-features)
7. [Best Practices](#best-practices)

---

## Getting Started

### Prerequisites Check
Before starting, ensure you have:
1. Ensure you have **Administrator** privileges.
2. The tool version is displayed in the window title: **Profile Migration Tool v2.12.25**.
3. The tool automatically detects your Windows theme (Light/Dark).
4. **Hardware/Software Checks:**
   - ✅ Administrator rights on both source and target computers
   - ✅ 7-Zip installed (tool will prompt if missing)
   - ✅ Sufficient disk space (2x profile size recommended)
   - ✅ Network/USB drive for transferring exported profiles

### Launching the Tool

Open PowerShell as Administrator and run:
```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\ProfileMigration.ps1"
```

### First Launch

When the tool starts, you'll see:
- **User Selection Dropdown** - To see sizes click the "R" button
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
1. ✅ Ensure exported ZIP and SHA256 files are accessible (Flash drive, Network Share, or Local Disk)
2. ✅ Ensure target user is NOT logged in
3. ✅ Have user password ready (if creating new local account)
4. ✅ Verify sufficient disk space on C: drive

### Step-by-Step Import Process

#### 1. Select ZIP File
1. Click **Browse** button
2. Navigate to exported `.zip` file
3. Click **Open**
4. ZIP path shows in text field

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

**For AzureAD/Entra ID Users:**
```
AzureAD\username
```
**Example:** `AzureAD\john.doe`

**OR use UPN format (NEW in v2.8.6):**
```
user@domain.com
```
**Example:** `john.doe@company.com`

**UPN Format Benefits:**
- Automatically detected as AzureAD user (no `AzureAD\` prefix needed)
- Simpler to enter (just copy the email address)
- Works in both Set Target User and Import fields
- Tool automatically extracts username for folder creation

**Important for AzureAD:**
- Device must be Entra ID joined (Settings → Access work or school)
- **No prior login required** - Tool uses Microsoft Graph to resolve SID
- Tool validates join status and provides guided setup if needed
- If user has already logged in, you can also select from the dropdown

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
5. If not found, the tool displays an "Import failed" error dialog stating the user was not found and stops the operation

#### Scenario B: Existing Profile - Replace Mode
If profile exists at `C:\Users\username`:
1. Tool detects existing folder
2. User selects **Replace** mode
3. Existing folder is backed up:
   - **Fast Method:** Attempts to rename folder to `username.backup_timestamp` (Instant)
   - **Fallback:** If rename fails (files locked), performs a backup copy and then deletes the original
4. New profile extracted to clean folder
5. Registry SIDs rewritten for the new user
6. **Use when:** Migrating from old computer, profile corrupted, want fresh start
   
   
   **MERGE** (Advanced - preserves existing settings):
   - Backs up existing profile (safety)
   - Extracts to temporary location
   - Copies files into existing profile (doesn't overwrite newer files)
   - Keeps existing NTUSER.DAT (preserves current registry settings)
   - **Use when:** Adding files from backup, supplementing existing profile, keeping current settings
   
2. Choose mode and click button

The tool supports **Light** and **Dark** modes. It defaults to your Windows system setting, but you can override it by clicking the **L** or **D** button in the top-right corner.

## Settings Configuration
Click the **...** button in the header to open the Settings dialog. From here, you can:
- **Set 7-Zip Path**: Manually locate `7z.exe` if the tool cannot find it automatically.

## Profile Cleanup Wizard
Before exporting, you can use the **Profile Cleanup Wizard** to reduce the size of the migration and remove unnecessary files.

### Features
The wizard automatically scans for and categorizes:
1.  **Browser Caches:** Temporary data from Chrome, Edge, and Firefox.
2.  **Temporary Files:** User TEMP folder contents and other safe-to-delete items.
3.  **Large Files:** Files larger than 100MB (likely videos or installers).
4.  **Duplicate Files:** Identical files found in multiple locations.
5.  **Recycle Bin:** Forgotten trash items.
6.  **Large Downloads:** Old installers or ISOs in the Downloads folder.

### Using the Duplicate File Finder
The duplicate file feature is particularly powerful:
- **Detection:** It finds files that match in both **size** AND **hash** (content), so it is 100% safe.
- **Grouping:** Duplicates are shown in groups.
- **Selection:** By default, the *newest* file in each group is kept, and older duplicates are selected for deletion. You can expand any group to change your selection.
- **Safety:** By default, it preserves the newest copy of the file. You can override this to exclude all copies, but the default selection ensures one is kept.

### How to use:
1. Start the tool and select a user.
2. Click **Export**.
3. The **Profile Cleanup Wizard** will launch automatically.
4. Wait for the scan to complete.
5. Review the categories. Expand "Duplicate Files" or "Large Files" to see details.
6. Check the items you want to remove.
7. Click **Clean & Export** to proceed.

## Troubleshooting with Debug Mode
If you encounter errors during the **Export Profile** phase (especially 7-Zip errors), you can enable **Debug Mode**.

### How to use:
1. Locate the **Debug Mode** checkbox next to the **Export** button.
2. Check the box before clicking **Export**.
3. The Activity Log will now show every single file 7-Zip processes.
4. If the export fails, the tool will **preserve** all temporary files in your `TEMP` folder for IT analysis.
5. Exclusion list artifacts will also be saved to the `Logs/` directory.

#### Scenario C: User is Logged In
If the target user is currently logged in (or has a mounted registry hive):

1. **User Logged In** warning appears.
2. You will be given the option to **Force Logoff** the user.
3. If you choose **Yes**, the tool will attempt to forcefully sign out the user to release file locks.
4. If you choose **No**, the operation will be cancelled to prevent data corruption.

#### Scenario D: Corrupted ZIP File
If the imported ZIP file is corrupted or incomplete:

1. **Zip Integrity Check** runs automatically before any file operations.
2. If the archive is invalid (fails `7z t` test), the import will be blocked immediately.
3. An error message will display: "The ZIP archive is corrupted or invalid."
4. This prevents partial restores or data loss from bad archives.

#### Scenario E: Domain User First Login
1. Profile extracted to `C:\Users\username`
2. Registry permissions configured for domain
3. SID translation performed
4. User can login on next boot

### Post-Import Steps

#### 4. Monitor Import Progress
Progress indicators show:
- `Verifying archive integrity...` - 7-Zip internal consistency check
- `Verifying ZIP hash...` - SHA-256 hash check (if .sha256 file present)
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

## Restoring from Backup

There are two primary ways to restore a profile from a backup created by this tool.

### Method 1: Automated Restore (Recommended)
This uses the standard **Import Profile** feature of the tool.
1. Select the **Import Profile** tab.
2. Choose **Source Type**: "ZIP Backup".
3. Browse to your backup ZIP file.
4. Select the target user.
5. Click **Start Import**.
   - *The tool automatically handles file extraction, registry updates, and permission fixes.*

### Method 2: Manual / Emergency Restore
If the tool is unavailable or you need to recover data manually:

**1. Restore Files:**
- Extract the contents of the backup `.zip` file to the user's profile directory (e.g., `C:\Users\username`).

**2. Restore Registry:**
- Locate the `.reg` file created alongside your backup ZIP (e.g., `ProfileBackup_User_Date.zip.reg`).
- Double-click it to import the profile registry key into Windows.
- *Note: You may need to restart the computer for registry changes to take effect.*

**3. Verify Permissions:**
- Ensure the user account has **Full Control** over their restored profile folder.

---

## Profile Conversion (NEW!)

### Overview

The Profile Conversion feature allows you to convert existing user profiles between different account types **without needing to export and re-import**. This is significantly faster and more efficient than the traditional export/import workflow.

### Supported Conversion Paths

| From | To | Status |
|------|-----|--------|
| Local | AzureAD | ✅ Supported |
| AzureAD | Local | ✅ Supported |
| Domain | AzureAD | ✅ Supported |
| AzureAD | Domain | ✅ Supported |
| Local | Domain | ✅ Supported |
| Domain | Local | ✅ Supported |

### When to Use Profile Conversion

**Use Profile Conversion when:**
- ✅ Migrating from local account to AzureAD/Entra ID
- ✅ Converting AzureAD account back to local (e.g., leaving organization)
- ✅ Moving between domain and AzureAD
- ✅ Source and target are on the same computer
- ✅ You want to preserve all settings and files in-place

**Use Export/Import when:**
- ❌ Moving to a different computer
- ❌ Creating a backup archive
- ❌ Transferring over network
- ❌ Need hash verification

### Step-by-Step: Converting to AzureAD

#### Prerequisites
1. ✅ Device must be AzureAD/Entra ID joined
   - Go to Settings → Accounts → Access work or school
   - Should show "Connected to [Your Organization]'s Azure AD"
2. ✅ Source user must be logged out
3. ✅ Administrator rights required
4. ✅ AzureAD user email address (UPN) ready

#### Conversion Process

**1. Launch Profile Conversion**
- Click the **Convert Profile** button in the main window

**2. Select Source Profile**
- **Profile Type:** Choose source type (Local, Domain, or AzureAD)
- **Username:** Select from dropdown or enter username
  - Local: `username`
  - Domain: `DOMAIN\username`
  - AzureAD: Select from dropdown (shows as `TENANT\username`)

**3. Select Target Profile Type**
- **Profile Type:** Choose "AzureAD Profile"
- **Username:** Enter AzureAD User Principal Name (UPN)
  - **Format:** `user@domain.com` (email format required)
  - **Example:** `john.doe@company.com`
  - **Note:** A hint appears below the input box reminding you of the format

**4. Microsoft Graph Authentication (First Time Only)**
- If this is your first AzureAD conversion, you'll see:
  - **Module Installation Dialog:** Informs you that Microsoft Graph PowerShell module needs to be installed
  - **NuGet Provider Prompt:** You may see a PowerShell prompt asking to install NuGet provider
    - Press **Y** (Yes) when prompted
    - This is a one-time installation
  - **Microsoft Graph Sign-In:** Browser window opens for authentication
    - Sign in with an account that has permissions to read user information
    - Typically an admin account or the user's own account
    - Grant requested permissions

**5. Conversion Progress**
The tool will:
1. Verify AzureAD join status (guides you if not joined)
2. Retrieve AzureAD user SID via Microsoft Graph API
3. Moves (or copies) profile data to new location (e.g., `C:\Users\john.doe`)
4. Update Windows registry (ProfileList)
5. Rewrite SIDs in NTUSER.DAT
6. Apply correct permissions with AzureAD SID
7. Clean up old profile registry association

**6. Completion**
- Success dialog appears
- **Reboot required** for changes to take effect
- User can now login with AzureAD credentials

### Step-by-Step: Converting from AzureAD

#### Prerequisites
1. ✅ AzureAD user must be logged out
2. ✅ Administrator rights required
3. ✅ Target username decided (for local) or domain credentials ready

#### Conversion Process

**1. Launch Profile Conversion**
- Click the **Convert Profile** button

**2. Select Source Profile**
- **Profile Type:** Choose "AzureAD Profile"
- **Username:** Select from dropdown (shows as `TENANT\username`)

**3. Select Target Profile Type**

**For Local Account:**
- **Profile Type:** Choose "Local Profile"
- **Username:** Enter desired local username (e.g., `john.doe`)
- **Create User Dialog:** Appears if user doesn't exist
  - Enter password (twice for confirmation)
  - Shows strength indicator
  - Optional: Check "Add to Administrators group"

**For Domain Account:**
- **Profile Type:** Choose "Domain Profile"
- **Username:** Enter as `DOMAIN\username`
- **Domain Credentials:** Provide domain admin credentials when prompted

**4. Conversion Progress**
The tool will:
1. Create target user (if needed)
2. Copy profile to new location
3. Update Windows registry
4. Rewrite SIDs in NTUSER.DAT
5. Apply correct permissions
6. Clean up old AzureAD profile

**5. Completion**
- Success dialog appears
- Reboot required
- User can now login with local/domain credentials

### AzureAD-Specific Features

#### Microsoft Graph Integration
- **No Prior Login Required:** Unlike traditional methods, you don't need the AzureAD user to have logged in before
- **Automatic SID Resolution:** Tool queries Microsoft Graph API to get the user's SID
- **ObjectId to SID Conversion:** Converts Entra ID ObjectId to Windows SID format

#### Smart Profile Naming
- **Username-Only Folders:** AzureAD profiles are named with just the username part
  - Input: `john.doe@company.com`
  - Folder: `C:\Users\john.doe` (not `C:\Users\john.doe@company.com`)

#### Auto-Join Guidance
If device is not AzureAD-joined when attempting conversion:
1. Helpful dialog appears with step-by-step instructions
2. Option to open Windows Settings directly
3. Guided through:
   - Settings → Accounts → Access work or school
   - Click "Connect"
   - Select "Join this device to Azure Active Directory"
   - Sign in with work/school account
   - Complete setup wizard
4. Return to tool and retry conversion

### Troubleshooting Profile Conversion

**"Device is not AzureAD joined"**
- Follow the auto-join guidance dialog
- Verify in Settings → Accounts → Access work or school
- Must show "Connected to [Organization]'s Azure AD"

**"AzureAD username must be in email format"**
- Use full UPN: `user@domain.com`
- Not just username: ~~`user`~~
- Not domain format: ~~`DOMAIN\user`~~

**"Failed to retrieve SID from Microsoft Graph"**
- Ensure you signed in during authentication
- Check you have permissions to read user information
- Verify the email address is correct
- User must exist in your AzureAD tenant

**"Microsoft Graph module installation failed"**
- Check internet connection
- Ensure you pressed Y when prompted for NuGet
- Try running PowerShell as Administrator
- Manually install: `Install-Module Microsoft.Graph.Users -Scope CurrentUser`

**"Profile conversion failed: A parameter cannot be found"**
- This was a bug in earlier versions
- Ensure you're running v2.4.7 or later
- Check the version in the window title

### Best Practices for Profile Conversion

**Before Conversion:**
1. ✅ **Backup important data** - Just in case
2. ✅ **Close all applications** - Ensure no files are locked
3. ✅ **Logout source user** - Critical for success
4. ✅ **Verify account exists** - For AzureAD, confirm user in tenant
5. ✅ **Check disk space** - Need space for profile copy

**During Conversion:**
1. ✅ **Don't interrupt** - Let process complete
2. ✅ **Monitor progress bar** - Watch for errors
3. ✅ **Read dialogs carefully** - Important information provided
4. ✅ **Authenticate when prompted** - For Microsoft Graph

**After Conversion:**
1. ✅ **Reboot immediately** - Required for registry changes
2. ✅ **Test login** - Verify user can sign in
3. ✅ **Check files** - Ensure Desktop/Documents present
4. ✅ **Verify applications** - Test key apps work
5. ✅ **Monitor for 24-48 hours** - Watch for any issues

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

### Easy "Join Now" Button (v2.10.x)
New in v2.10, a **"Join Now"** button is available directly on the main form (top right corner).

**How to use:**
1. Enter the target domain name in the input field.
2. Click **Join Now**.
3. You will be prompted for domain credentials.
4. The tool performs checks, joins the domain, and notifies you of the status instantly.

---

### Domain User Migration
When importing for domain users:
1. Use format: `DOMAIN\username` in import
2. Provide domain admin credentials when prompted
3. User account is queried from Active Directory
4. Profile configured with correct domain SID
5. User can login immediately after reboot

---

## Repairing a Profile

### Overview
If a local user profile becomes corrupted (e.g., Start Menu broken, "Temporary Profile" errors, or permissions issues), you can use the **Local-to-Local conversion** feature to repair it in place.

### How to Repair
1. Click the **Convert Profile** button.
2. **Select Source:** Choose the corrupted user (e.g., `LocalUser1`).
3. **Select Target:** Choose the **SAME** user (`LocalUser1`) as the target.
   - You can copy/paste the username or type it manually.
4. The tool will detect this is a **Repair Operation**.
5. Click **Convert**.

### What Happens
The tool performs a non-destructive repair:
- **Registry Refresh:** Safely deletes and recreates the `ProfileList` registry entry to fix temporary profile states.
- **ACL Reset:** Scans the entire profile folder (`C:\Users\LocalUser1`) and forces ownership/permissions to be correct.
- **AppX Repair:** Schedules a full re-registration of specific Windows apps on the next login.
- **Data Preservation:** Your files (Desktop, Documents, etc.) are **NOT** copied or moved; they stay exactly where they are.

### After Repair
1. Reboot the computer.
2. Log in as the user.
3. Allow 1-2 minutes for the AppX repair script to complete (before the desktop fully loads).
4. Verify the issue is resolved.

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

### Log Filtering & Search
**Built-in log viewer features:**

1.  **Search Box:**
    - Type text to instantly filter log lines (e.g., "error", "registry", "user").
    - Clears automatically when text is removed.

2.  **Level Filtering Dropdown:**
   - **DEBUG** - Shows everything
   - **INFO** - Normal operations (default)
   - **WARN** - Warnings and errors
   - **ERROR** - Only critical failures

3.  **Refresh Button:**
    - Manually reloads the log file if external updates occur.

4.  **Auto-Scroll:**
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
1. Import profile using DOMAIN\username format
2. Provide domain credentials when prompted
3. Reboot
4. Login with domain account

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

## Next Steps

- **Configuration:** See [CONFIGURATION.md](CONFIGURATION.md) for advanced settings
- **Troubleshooting:** See [TECHNICAL-DOCS.md](TECHNICAL-DOCS.md) for error resolution
- **FAQ:** See [FAQ.md](FAQ.md) for common questions

---

**Need Help?** Check the FAQ or contact your IT administrator.
### Optional AzureAD Unjoin Feature

When converting FROM an AzureAD profile to Local or Domain, you can optionally unjoin the device from AzureAD.

**How to use:**
1. In the Profile Conversion dialog, select an AzureAD source profile
2. A checkbox appears to the right of the target username: "Unjoin from AzureAD after conversion"
3. Check the box if you want to unjoin the device
4. Complete the conversion
5. A confirmation dialog will appear explaining the consequences
6. Confirm to proceed with unjoin

**What happens when you unjoin:**
- Device is removed from your organization's management
- Conditional access policies are disabled
- SSO capabilities are removed
- Requires a reboot to complete

**When to use:**
- Converting personal device from work to personal use
- Leaving organization and removing corporate management
- Device repurposing

**When NOT to use:**
- Keeping device in organization
- Testing conversions (want to preserve join status)
- Hybrid scenarios (AzureAD + Domain)

---

## Domain  AzureAD Conversions (v2.7.2 - Automatic Unjoin/Join)

### Overview

**NEW in v2.7.2:** The tool now automatically handles domain and AzureAD unjoin/join operations during profile conversions. No manual intervention required!

### What's Automatic

When converting between domain and AzureAD profiles, the tool automatically:
-  Detects current join state (domain or AzureAD)
-  Prompts for mandatory unjoin (if needed)
-  Executes unjoin safely
-  Prompts for join credentials/setup
-  Joins target environment
-  Converts profile
-  Updates registry
-  Prompts for reboot

**You don't need to manually unjoin or join anymore!**

### AzureAD  Domain Conversion

#### Prerequisites
1.  Device is currently AzureAD joined
2.  Source AzureAD user is logged out
3.  Domain admin credentials ready
4.  Network connectivity to domain controller

#### Step-by-Step Process

**1. Launch Profile Conversion**
- Click **Convert Profile** button

**2. Select Source Profile**
- **Profile Type:** AzureAD Profile
- **Username:** Select from dropdown (shows as TENANT\username)

**3. Select Target Profile**
- **Profile Type:** Domain Profile
- **Username:** Enter as DOMAIN\username

**4. Automatic Unjoin from AzureAD**
The tool will:
1. Detect device is AzureAD joined
2. Show confirmation dialog explaining:
   - Device will be unjoined from AzureAD
   - Management policies will be removed
   - After unjoin, domain join will proceed
3. If you click **Yes**:
   - Tool executes dsregcmd /leave
   - Verifies unjoin succeeded
   - Proceeds to next step
4. If you click **No**:
   - Conversion is cancelled
   - No changes made

**5. Automatic Domain Join**
The tool will:
1. Prompt for domain credentials (Windows credential dialog)
2. Enter domain admin username and password
3. Tool executes Add-Computer with credentials
4. Verifies domain join succeeded
5. Proceeds to profile conversion

**6. Profile Conversion**
The tool will:
1. Convert profile using pre-resolved AzureAD SID
2. Update registry (domain SID  profile folder)
3. Apply correct permissions
4. Clean up old profile entries

**7. Completion**
- Success dialog appears
- **Reboot required** to complete domain join
- User can now login with domain credentials

#### Special Case: Same Username

If domain and AzureAD tenant have the same name (e.g., "cpmethod"):
- Source: AzureAD user cpeters@cpmethod.com  C:\Users\cpeters
- Target: Domain user cpmethod\cpeters  C:\Users\cpeters

**What happens:**
- Tool detects paths are identical
- Skips all file operations
- Only updates registry entries
- **Instant conversion** (< 1 minute)
- All data preserved in place

### Domain  AzureAD Conversion

#### Prerequisites
1.  Device is currently domain joined
2.  Source domain user is logged out
3.  AzureAD tenant credentials ready
4.  Network connectivity to internet

#### Step-by-Step Process

**1. Launch Profile Conversion**
- Click **Convert Profile** button

**2. Select Source Profile**
- **Profile Type:** Domain Profile
- **Username:** Select from dropdown (shows as DOMAIN\username)

**3. Select Target Profile**
- **Profile Type:** AzureAD Profile
- **Username:** Enter as user@domain.com (UPN format)

**4. Automatic Unjoin from Domain**
The tool will:
1. Detect device is domain joined
2. Show confirmation dialog explaining:
   - Device will be unjoined from domain
   - Group policies will be removed
   - After unjoin, AzureAD join will proceed
3. If you click **Yes**:
   - Tool executes Remove-Computer
   - Verifies unjoin succeeded
   - Proceeds to next step
4. If you click **No**:
   - Conversion is cancelled
   - No changes made

**5. Manual AzureAD Join**
The tool will:
1. Show instructions for AzureAD join
2. Launch Windows Settings (ms-settings:workplace)
3. You complete the join:
   - Click "Connect"
   - Select "Join this device to Azure Active Directory"
   - Sign in with work/school account
   - Complete setup wizard
4. Tool verifies join succeeded
5. Proceeds to profile conversion

**6. Profile Conversion**
The tool will:
1. Convert profile to AzureAD format
2. Update registry (AzureAD SID  profile folder)
3. Apply correct permissions
4. Clean up old profile entries

**7. Completion**
- Success dialog appears
- **Reboot required** to complete AzureAD join
- User can now login with AzureAD credentials

### Troubleshooting

**"Unjoin failed" error**
- Check you have admin rights
- For domain unjoin: Ensure network connectivity
- For AzureAD unjoin: Device must be AzureAD joined
- Review error message for specific cause

**"Join failed" error**
- For domain join:
  - Verify domain credentials are correct
  - Check network connectivity to domain controller
  - Ensure DNS is configured correctly
- For AzureAD join:
  - Complete the join process in Settings
  - Sign in with correct work/school account
  - Wait for join to complete before clicking OK

**"Cannot find profile" after conversion**
- This shouldn't happen with v2.7.2
- If it does, check logs for SID resolution errors
- Profile may be at unexpected location

### Best Practices

**Before Conversion:**
1.  Backup important data
2.  Logout source user completely
3.  Close all applications
4.  Have credentials ready (domain admin or AzureAD account)
5.  Verify network connectivity

**During Conversion:**
1.  Read all dialogs carefully
2.  Don't interrupt the process
3.  Provide correct credentials when prompted
4.  Complete AzureAD join fully (if prompted)

**After Conversion:**
1.  Reboot immediately
2.  Login with new credentials
3.  Verify Desktop/Documents are intact
4.  Test applications
5.  Monitor for 24-48 hours

