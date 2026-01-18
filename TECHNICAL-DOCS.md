# USER PROFILE TRANSFER TOOL - v2.12.25 (JANUARY 2026)
**Tested on:** Windows 11 25H2 (26200.7623)

## Architecture
- **Centralized Versioning:** Versioning is controlled by a single `Version` key in the `$Config` hashtable.
- **Dynamic Theming:** A UI-wide theme system using a `$Themes` dictionary.
- **AppX Re-registration:** Uses **Active Setup** (`HKLM\SOFTWARE\Microsoft\Active Setup\Installed Components`) to trigger a self-destructing PowerShell script on the user's next login, ensuring apps register *before* the desktop loads.
- **Local Repair:** In-place conversion logic that detects when Source==Target, skipping file operations but refreshing Registry and ACLs.
- **Debug Mode Logic:** Checkbox-driven toggle that increases 7-Zip verbosity (`-bb3`), preserves temporary listfiles in `Logs\`, and disables automatic cleanup of the `$tmp` directory on error or completion.
- **Tool Automation:** Auto-detection of 7-Zip (`7z.exe`) with built-in download/install fallback.
- **Security:** Elevation check (`Is-Admin`) and privilege bypass logic for hive mounting.
- Monolithic script: All logic in ProfileMigration.ps1
- Global state: All persistent state and UI controls use $global:
- External tools: Robocopy, 7-Zip, reg.exe
- Registry hive SID rewrite: Binary and UTF-16 string replacement
- AzureAD/Entra ID: Detect SIDs with ^S-1-12-1-
- Privileges: SeBackupPrivilege, SeRestorePrivilege, SeTakeOwnershipPrivilege
- Threading: CPU core count for Robocopy/7-Zip
- Progress/UI: [System.Windows.Forms.Application]::DoEvents() after UI/progress updates
- Logging: DEBUG, INFO, WARN, ERROR

## Key Workflows
- Export: Detect profile, copy NTUSER.DAT, Robocopy, manifest, compress, hash, HTML report
- Import: Verify ZIP/hash, extract, read manifest, create user, rewrite hive SID, copy files, set ACLs, update registry, Winget app install, HTML report, prompt reboot
- Conversion: Host-side migration (Local/Domain/AzureAD) → Auto-detect join state, unjoin/join if needed, copy/merge files, rewrite hive SID, update ProfileList, set ACLs
- Repair: Local-to-Local self-conversion → Detect identity, Skip file copy, Refresh Registry ProfileList, Reset ACLs, Trigger AppX repair

## AzureAD/Entra ID Detection
- Checks all Win32_UserProfile entries for username
- If any SID matches ^S-1-12-1-, treats as AzureAD/Entra ID
- Handles DOMAIN\username and AzureAD\username formats
- Robust detection logic for all user selection scenarios

## GUI Usability
- Modern flat design, color-coded status
- Custom dialogs for errors/success (taller for full info display)
- Progress bar, log viewer, tooltips
- Improved dialog sizing for all feedback (2025)

## Logging
- Four levels: DEBUG, INFO, WARN, ERROR
- Console, GUI, and file output
- Set via $Config.LogLevel

## Error Handling
- try/catch, log user # Technical Documentation (v2.10.109)ls
- Rollback on import failure if backup exists

## Gotchas
- Always unload registry hives with retry loop
- 7-Zip: check/install if missing
- Exports fail if profile is active
- Outlook OST files excluded
- Robocopy: /XJ excludes junctions
- UI freezes if DoEvents() omitted

## Files
- ProfileMigration.ps1 – all logic
- README.md, TECHNICAL-DOCS.md, CONFIGURATION.md, USER-GUIDE.md
---

## AzureAD Profile Conversion Architecture (v2.4.7)

### Overview
The AzureAD profile conversion system enables seamless migration between Local, Domain, and AzureAD account types without requiring export/import workflows.

### Core Components

#### Helper Functions

**Test-IsAzureADSID**
- **Purpose:** Validates if SID belongs to AzureAD user
- **Pattern:** ^S-1-12-1- (AzureAD SID prefix)
- **Returns:** Boolean
- **Usage:** Profile type detection, validation

**Test-IsAzureADJoined**
- **Purpose:** Checks if computer is AzureAD-joined
- **Method:** WMI query to Win32_ComputerSystem
- **Returns:** Boolean

**Get-ProfileType**
- **Purpose:** Determines profile type (Local/Domain/AzureAD)
- **Returns:** String ('Local', 'Domain', or 'AzureAD')

#### Microsoft Graph Integration

**Convert-EntraObjectIdToSid**
- **Purpose:** Converts Azure AD ObjectId (GUID) to Windows SID
- **Algorithm:** Parse GUID  byte array  construct SID with prefix S-1-12-1-
- **Input:** ObjectId (e.g., 2e3ff6dd-cdde-43ed-b31e-5ddc65316615)
- **Output:** SID (e.g., S-1-12-1-775943901-1139658206-3697090227-359018853)

**Get-AzureADUserSID**
- **Purpose:** Retrieves AzureAD user SID via Microsoft Graph API
- **Parameters:** UserPrincipalName (email format)
- **Process:** Install module  Connect to Graph  Query user  Convert ObjectId  Return SID

**Integration in Get-LocalUserSID**
- **Fallback Logic:** If local SID lookup fails for AzureAD user
**Validation:** Checks username contains @ (UPN format)
- **Automatic:** Transparent to caller

**Get-DomainCredential**
- **Purpose:** Standardized, theme-aware GUI credential prompt.
- **Features:**
  - Modern Windows Forms UI matching app theme.
  - Prevents "credential fatigue" by checking `$global:DomainCredential` first (optional).
  - Validates non-empty input.
- **Used By:** `Get-DomainAdminCredential` (Join/Import), Convert-Profile (Domain).

**Get-DomainAdminCredential**
- **Purpose:** Centralized helper for obtaining and validating Domain Admin credentials.
- **Features:** 
  - Standardizes UI by calling `Get-DomainCredential`.
  - Implements **robust retry logic** for failed authentication or permissions.
  - Inlines permission checks (Domain Admin group membership).
  - Handles specific error codes (`INVALID_CREDENTIALS`, `DC_UNREACHABLE`) with user-friendly retry prompts.
- **Used By:** `Join-Domain-Enhanced` (Join Now button), `Import-Profile` (Domain user verification).

**Invoke-ProactiveUserCheck**
- **Purpose:** Checks if a user is logged in (mounted registry) and handles force logoff.
- **Features:**
  - Uses `Test-ProfileMounted` to detect `HKU` hive.
  - Prompts user with "Force Logoff" dialog if detected.
  - Calls `Invoke-ForceUserLogoff` (logoff.exe) if confirmed.
- **Used By:** `cmbSourceProfile.SelectedIndexChanged`, `btnClean`, `btnSetTargetUser`.

**Test-ZipIntegrity**
- **Purpose:** Verifies ZIP archive integrity before extraction.
- **Features:** 
  - Runs `7z.exe t -bsp1` (test) command.
  - Returns boolean success/failure.
- **Used By:** `Import-UserProfile` (Pre-flight check).

### Conversion Functions

#### Convert-LocalToAzureAD
**Process:** Validate join  Get SIDs  Create profile folder  Copy  Update registry  Rewrite NTUSER.DAT  Apply permissions  Cleanup

**Parameters:**
- LocalUsername - Source local username
- AzureADUsername - Target UPN (email format)

#### Convert-AzureADToLocal
**Process:** Get SIDs  Create local user  Create profile folder  Copy  Update registry  Rewrite NTUSER.DAT  Apply permissions  Cleanup

**Parameters:**
- AzureADUsername - Source AzureAD UPN
- LocalUsername - Target local username
- LocalPassword - Password for new local user
- MakeAdmin - Add to Administrators group

### Registry Operations

**ProfileList Updates:**
- Key: HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\{SID}
- Properties: ProfileImagePath, Sid
- Old key deleted after successful conversion

**NTUSER.DAT Rewriting:**
- Function: Rewrite-HiveSID
- Mounts hive  Rewrites SIDs  Unmounts  Validates

### Permission Management

**Set-ProfileFolderAcls:**
- Takes ownership  Grants Administrators access  Resets ACLs  Applies new ACLs  Removes Administrators  Transfers ownership

**Parameters:**
- ProfilePath, UserName, UserSID

### Security Considerations

**Microsoft Graph Authentication:**
- OAuth 2.0 flow via browser
- Requires User.Read.All permission
- Credentials never stored

**SID Handling:**
- Pattern validation prevents injection
- Binary SID conversion verified

### Invoke-AzureADUnjoin Function

**Location:** Line 1362  
**Purpose:** Unjoins the computer from AzureAD/Entra ID using dsregcmd /leave

**Returns:** Hashtable with:
- Success (bool) - Whether unjoin succeeded
- Message (string) - Success or error message

**Process:**
1. Verifies device is AzureAD joined using Test-IsAzureADJoined
2. Executes dsregcmd /leave
3. Checks exit code
4. Returns result

**Error Handling:**
- Returns failure if device not joined
- Captures stderr output on failure
- Logs all attempts

**Usage:**
`powershell
 = Invoke-AzureADUnjoin
if (.Success) {
    Write-Host "Successfully unjoined"
} else {
    Write-Host "Failed: "
}
`

**Integration:**
- Called by Convert-AzureADToLocal when -UnjoinAzureAD True
- Confirmation dialog shown before execution
- Conversion succeeds even if unjoin fails

---

## Domain  AzureAD Conversion Architecture (v2.7.2)

### Overview
Automatic domain/AzureAD unjoin and join handling for seamless profile conversions between domain and AzureAD environments.

### SID Resolution Strategy

**Challenge:** After unjoining from AzureAD, source SID cannot be resolved via name lookup.

**Solution:** Pre-resolve source SID from registry before unjoin.

**Registry-Based Lookup:**
- Search ProfileList for AzureAD SIDs (S-1-12-1-...)
- Match by profile folder name
- Return SID before unjoin occurs
- Pass to conversion function via -SourceSID parameter

### AzureAD  Domain Flow
1. Resolve source AzureAD SID from registry
2. Detect AzureAD join, prompt for unjoin
3. Execute Invoke-AzureADUnjoin
4. Prompt for domain credentials
5. Join domain with Add-Computer
6. Convert profile using pre-resolved SID
7. Update registry (domain SID  profile folder)

### Domain  AzureAD Flow
1. Detect domain join, prompt for unjoin
2. Execute Invoke-DomainUnjoin
3. Prompt for AzureAD join
4. Launch ms-settings:workplace
5. Wait for user to complete join
6. Verify AzureAD join succeeded
7. Convert profile normally

### Same-Path Optimization
When source and target paths are identical:
- Skip all file operations
- Registry-only update
- Instant conversion
- Zero data movement

### Precondition Checks
Domain unjoin logic added in THREE places:
1. Precondition checks (runs first!)
2. Convert-LocalToAzureAD function
3. Profile Conversion UI flow

---

## Complete API Reference (Internal Functions)

The following is a comprehensive list of internal functions orchestrated by `ProfileMigration.ps1`.

### Core Orchestrators
- **Export-UserProfile** - Main export logic: validation, compression (7z), hashing, reporting.
- **Import-UserProfile** - Main import logic: validation, backup, extraction, permissioning, registry update.
- **Convert-Profile** - Main conversion logic: handles Local/Domain/AzureAD conversions in-place.
- **Show-ProfileCleanupWizard** - UI wizard for analyzing and deleting large/duplicate files before export.
- **Show-ProfileConversionDialog** - UI dialog for selecting conversion source/target and initiating the process.

### UI & Theming
- **Show-ModernDialog** - Unified dialog box for Success, Error, Warning, Question (theme-aware).
- **Show-FileDetailsDialog** - Scrollable text dialog for license files, logs, or lists.
- **Apply-Theme** - Recursively applies Light/Dark theme colors to all Form controls.
- **Init-Theme** - Detects system theme (Registry) or user preference and sets global theme.
- **Update-Status** - Helper to update the main status label and force UI refresh (`DoEvents`).

### AzureAD & Graph Integration
- **Test-IsAzureADSID** - Regex check (`^S-1-12-1-`) to identify AzureAD users.
- **Test-IsAzureADJoined** - WMI check (`Win32_ComputerSystem.PartOfDomain`) for AzureAD join status.
- **Get-AzureADUserSID** - Uses Microsoft Graph API to resolve UPN (`user@domain.com`) to SID.
- **Convert-EntraObjectIdToSid** - Helper to convert Graph GUID to Windows SID format.
- **Invoke-AzureADUnjoin** - Wrapper for `dsregcmd /leave` with error handling.

### Domain Handling
- **Get-DomainCredential** - Standardized UI prompt for domain user credentials.
- **Get-DomainAdminCredential** - Wrapper for `Get-DomainCredential` with retry logic for Admin operations.
- **Join-Domain-Enhanced** - Logic behind "Join Now" button: connectivity check, join, reboot prompt.
- **Invoke-DomainUnjoin** - Wrapper for `Remove-Computer` to unjoin domain.
- **Test-DomainReachability** - LDAP/DNS check to verify domain controller availability.

### Registry & Filesystem
- **Rewrite-HiveSID** - Critical function to binary-patch `NTUSER.DAT` with new user SID.
- **Set-ProfileFolderAcls** - Resets and applies correct ACLs (Full Control) for new owner.
- **Get-RobocopyExclusions** - Returns list of `/XF` (files) and `/XD` (dirs) for Robocopy.
- **Test-ProfileMounted** - Checks `HKU` to see if a user's hive is currently loaded (logged in).
- **Invoke-ProactiveUserCheck** - Checks if user is logged in and offers "Force Logoff" (quser/logoff).

### Utilities
- **Test-IsAdmin** - specific check for elevated privileges.
- **Test-InternetConnectivity** - connectivity check (Microsoft endpoints) for Graph/AzureAD.
- **Test-ZipIntegrity** - Runs `7z t` to verify archive structure before import.
- **Log-Message** - Central logging function (Console + File + UI).
- **Log-Debug** / **Log-Error** - Wrappers for `Log-Message` with specific levels.
- **New-CleanupItem** - Standardized object creator for Cleanup Wizard items, enforcing strong typing and duplicate group support.
- **Test-PathWithRetry** - Robust file existence check that retries on access denied or network glitch errors.

