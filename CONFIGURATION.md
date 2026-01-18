# ProfileMigration.ps1 – Configuration Guide (v2.12.25)

## Config Object ($Config)
- Version: Current script version (e.g., '2.12.22')
- LogLevel: DEBUG, INFO, WARN, ERROR (Minimum level for display/recording)
- RobocopyThreads: 8–32 (auto-detects CPU cores)
- SevenZipPath: Path to 7z.exe (auto-detects/installs if missing)
- ExcludeOST: Exclude Outlook OST files (default: true)
- ExcludeTemp: Exclude temp folders (default: true)
- ExcludeHives: Exclude registry hives (default: true)
- ExportReport: Generate HTML report (default: true)
- ImportReport: Generate HTML report (default: true)
- WingetApps: Install apps from manifest (default: false)
- GenerateHTMLReports: Globally enable/disable HTML reports
- AutoOpenReports: Automatically open reports in default browser
- **Debug Mode**: Toggleable via UI for verbose 7-Zip logging and artifact preservation (overrides default log cleanup)
## AzureAD/Entra ID Detection
- If any profile SID matches ^S-1-12-1-, treats as AzureAD/Entra ID
- Handles DOMAIN\username and AzureAD\username formats
- Robust detection logic for all user selection scenarios
## Dialogs & UI
- Modern flat design, color-coded status
- Custom dialogs for errors/success (taller for full info display)
- Improved dialog sizing for all feedback (2025)
## Logging
- Four levels: DEBUG, INFO, WARN, ERROR
- Console, GUI, and file output
## Example
```powershell
$Config = @{
    Version                     = '1.0.0'
    LogLevel                    = 'INFO'
    RobocopyThreads             = 16
    SevenZipThreads             = 8
    GenerateHTMLReports         = $true
    AutoOpenReports             = $true
    # ... other internal tunables
}
```

## Theme System
The tool features a custom theme system that supports **Light** and **Dark** modes.

### Automatic Detection
On startup, the script detects the Windows "AppsUseLightTheme" registry setting and applies the matching theme.

### Manual Toggle
Users can toggle between Light (L) and Dark (D) modes using the button in the top-right corner of the header.

## Debug Mode Operations
Debug Mode can be enabled via the checkbox next to the **Export** button.

### Functional Impact:
1.  **7-Zip Verbosity**: Adds `-bb3` to 7-Zip arguments, logging every file processed.
2.  **Exclusion Artifacts**: Copies temporary exclusion listfiles to the `Logs/` directory for audit.
3.  **Cleanup Bypass**: Prevents the script from deleting the temporary scratch folder (`$env:TEMP\ProfileMigration_...`) on completion, allowing for manual inspection of artifacts like `manifest.json` and `Winget-Packages.json`.

### Theme Definitions
Themes are defined in the `$Themes` hashtable, allowing for easy customization of colors for panels, buttons, text, and logs.
# Configuration Guide - Windows Profile Migration Tool

## Overview

This guide covers all configurable options in the ProfileMigration tool. Most users can use default settings, but advanced administrators may want to customize behavior for specific environments.

---

## Configuration Location

All settings are defined in the `$Config` hashtable near the top of `ProfileMigration.ps1` (around line 40):

```powershell
$Config = @{
    Version                     = '2.12.22'     # Current script version
    DomainReachabilityTimeout   = 3000      # milliseconds for LDAP/DNS checks
    DomainJoinCountdown         = 10        # seconds for restart warning
    HiveUnloadMaxAttempts       = 3         # retry attempts for hive cleanup
    HiveUnloadWaitMs            = 500       # milliseconds between retry attempts
    MountPointMaxAttempts       = 5         # max collision detection attempts
    ProgressUpdateIntervalMs    = 1000      # milliseconds between progress display updates
    ExportProgressCheckMs       = 500       # milliseconds between export file count checks
    RobocopyThreads             = [Math]::Min(32, [Math]::Max(8, $cpuCores))
    RobocopyRetryCount          = 1         # /R parameter
    RobocopyRetryWaitSec        = 1         # /W parameter
    SevenZipThreads             = $cpuCores # Use all available CPU cores
    ProfileValidationTimeoutSec = 10      # max time to validate profile path writeability
    SizeEstimationDepth         = 3           # depth for fast size estimation
    HashVerificationEnabled     = $true       # verify ZIP integrity
    LogLevel                    = 'INFO'      # Minimum log level: DEBUG, INFO, WARN, ERROR
    GenerateHTMLReports         = $true       # Generate HTML migration reports
    AutoOpenReports             = $true       # Automatically open HTML reports after generation
}
```

---

## Settings Reference

### Performance Settings

#### SevenZipThreads
**Purpose:** Number of CPU threads used for compression/extraction  
**Type:** Integer  
**Default:** `$cpuCores` (all available cores)  
**Range:** 1 to number of CPU cores

**Effect:**
- Higher = faster compression/extraction
- Uses more CPU resources
- Diminishing returns beyond 8-12 cores for most profiles

**When to change:**
```powershell
# Limit CPU usage on production systems
$Config.SevenZipThreads = 4  # Use only 4 cores

# Maximize speed on dedicated migration workstation
$Config.SevenZipThreads = 16  # Use all 16 cores
```

**Performance impact:**
- 1 thread: Baseline
- 4 threads: ~3.5x faster
- 8 threads: ~6x faster
- 12 threads: ~8x faster
- 16+ threads: ~10x faster (limited by I/O)

#### RobocopyThreads
**Purpose:** Number of parallel file copy threads  
**Type:** Integer  
**Default:** `[Math]::Min(32, [Math]::Max(8, $cpuCores))`  
**Range:** 1-128 (recommended: 8-32)

**Default behavior:**
- Minimum: 8 threads (even on dual-core)
- Maximum: 32 threads (robocopy limit for optimal performance)
- Scales between based on CPU count

**When to change:**
```powershell
# Slow HDD - reduce threads to avoid thrashing
$Config.RobocopyThreads = 8

# Fast NVMe SSD - maximize parallelism
$Config.RobocopyThreads = 32

# Network share - optimize for network bandwidth
$Config.RobocopyThreads = 16
```

**Guidelines:**
| Storage Type | Recommended Threads |
|--------------|---------------------|
| HDD (5400 RPM) | 4-8 |
| HDD (7200 RPM) | 8-12 |
| SATA SSD | 16-24 |
| NVMe SSD | 24-32 |
| Network (1 Gbps) | 8-12 |
| Network (10 Gbps) | 16-24 |

#### RobocopyRetryCount
**Purpose:** Number of retry attempts for failed file copies  
**Type:** Integer  
**Default:** `1`  
**Range:** 0-1000000

**Effect:**
- Higher = more resilient to transient errors
- Higher = slower on permanently locked files
- 0 = skip failed files immediately

**When to change:**
```powershell
# Unreliable storage or network
$Config.RobocopyRetryCount = 5

# Fast failure for troubleshooting
$Config.RobocopyRetryCount = 0

# Very patient operation (locked file scenarios)
$Config.RobocopyRetryCount = 10
```

#### RobocopyRetryWaitSec
**Purpose:** Seconds to wait between retry attempts  
**Type:** Integer  
**Default:** `1`  
**Range:** 0-3600

**Effect:**
- Higher = gives locked files time to release
- Higher = slower overall operation
- Lower = faster failure detection

**Typical scenarios:**
```powershell
# Quick retries for network glitches
$Config.RobocopyRetryWaitSec = 1

# Wait for locked files (e.g., Outlook PST)
$Config.RobocopyRetryWaitSec = 5

# Very patient (applications closing)
$Config.RobocopyRetryWaitSec = 10
```

---

### UI & Progress Settings

#### ProgressUpdateIntervalMs
**Purpose:** Milliseconds between progress bar updates  
**Type:** Integer  
**Default:** `1000` (1 second)  
**Range:** 100-5000

**Effect:**
- Lower = smoother progress bar
- Lower = more CPU for UI updates
- Higher = less responsive but more efficient

**When to change:**
```powershell
# Smooth visual updates (modern fast PC)
$Config.ProgressUpdateIntervalMs = 250

# Reduce overhead (slow PC or headless operation)
$Config.ProgressUpdateIntervalMs = 2000
```

#### ExportProgressCheckMs
**Purpose:** Milliseconds between checking export progress  
**Type:** Integer  
**Default:** `500`  
**Range:** 100-5000

**Effect:**
- Lower = more accurate progress reporting
- Lower = more disk I/O for file counting
- Higher = less overhead

**Recommendation:** Keep at 500ms for most scenarios

---

### Reliability Settings

#### HiveUnloadMaxAttempts
**Purpose:** Retry attempts for unloading registry hives  
**Type:** Integer  
**Default:** `3`  
**Range:** 1-10

**Effect:**
- Higher = more resilient to locked hives
- Registry unload can fail if processes hold handles

**When to change:**
```powershell
# Very aggressive cleanup (may have locked hives)
$Config.HiveUnloadMaxAttempts = 5

# Fast failure for troubleshooting
$Config.HiveUnloadMaxAttempts = 1
```

#### HiveUnloadWaitMs
**Purpose:** Milliseconds between hive unload retry attempts  
**Type:** Integer  
**Default:** `500`  
**Range:** 100-5000

**Effect:**
- Higher = gives system time to release handles
- Higher = slower cleanup on failure

**Typical values:**
```powershell
# Quick retries
$Config.HiveUnloadWaitMs = 200

# Patient cleanup
$Config.HiveUnloadWaitMs = 1000
```

#### MountPointMaxAttempts
**Purpose:** Maximum attempts to find unique registry mount point  
**Type:** Integer  
**Default:** `5`  
**Range:** 1-100

**Effect:**
- Used when loading registry hives to avoid collisions
- Prevents conflicts with other operations

**Rarely needs changing**

---

### Domain Settings

#### DomainReachabilityTimeout
**Purpose:** Milliseconds to wait for domain connectivity test  
**Type:** Integer  
**Default:** `3000` (3 seconds)  
**Range:** 1000-30000

**Effect:**
- Higher = more patient for slow networks
- Lower = faster failure detection

**When to change:**
```powershell
# Fast local network
$Config.DomainReachabilityTimeout = 1000

# Slow VPN or WAN link
$Config.DomainReachabilityTimeout = 10000

# Very slow satellite link
$Config.DomainReachabilityTimeout = 30000
```

#### DomainJoinCountdown
**Purpose:** Seconds countdown before reboot after domain join  
**Type:** Integer  
**Default:** `10`  
**Range:** 0-300

**Effect:**
- Gives user time to save work before reboot
- 0 = immediate reboot (use with caution!)

**When to change:**
```powershell
# Automated scripts (no user interaction)
$Config.DomainJoinCountdown = 0

# Give users more time
$Config.DomainJoinCountdown = 60

# Maximum warning time
$Config.DomainJoinCountdown = 300
```

---

### Validation Settings

#### ProfileValidationTimeoutSec
**Purpose:** Maximum seconds to test profile path writeability  
**Type:** Integer  
**Default:** `10`  
**Range:** 1-60

**Effect:**
- Used during pre-flight validation
- Tests if profile folder is accessible

**Rarely needs changing**

#### SizeEstimationDepth
**Purpose:** Directory depth for fast size calculation  
**Type:** Integer  
**Default:** `3`  
**Range:** 1-10

**Effect:**
- Lower = faster but less accurate
- Higher = slower but more accurate
- Used for initial profile size display

**When to change:**
```powershell
# Very fast estimation (may be inaccurate)
$Config.SizeEstimationDepth = 1

# Accurate size (slower on large profiles)
$Config.SizeEstimationDepth = 10
```

---

### Security Settings

#### HashVerificationEnabled
**Purpose:** Enable/disable SHA-256 hash verification on import  
**Type:** Boolean  
**Default:** `$true`  
**Values:** `$true` or `$false`

**Effect:**
- `$true` = Verifies ZIP integrity using .sha256 file
- `$false` = Skips verification (faster import)

**When to disable:**
```powershell
# Trusted local transfers (no network)
$Config.HashVerificationEnabled = $false

# Speed over security
$Config.HashVerificationEnabled = $false
```

**Security recommendation:** Keep enabled for:
- Network transfers
- Long-term archives
- Compliance requirements
- Untrusted sources

---

### Logging Settings

#### LogLevel
**Purpose:** Minimum log level to display/record  
**Type:** String  
**Default:** `'INFO'`  
**Values:** `'DEBUG'`, `'INFO'`, `'WARN'`, `'ERROR'`

**Levels explained:**
```powershell
# DEBUG - Very verbose, shows all operations
#   - File paths
#   - Every registry change
#   - Thread operations
#   - Performance metrics
$Config.LogLevel = 'DEBUG'

# INFO - Normal operations (default)
#   - Major steps
#   - Success messages
#   - Progress updates
$Config.LogLevel = 'INFO'

# WARN - Only warnings and errors
#   - Non-critical issues
#   - Skipped operations
#   - Performance warnings
$Config.LogLevel = 'WARN'

# ERROR - Only critical errors
#   - Operation failures
#   - Data corruption
#   - Access denied
$Config.LogLevel = 'ERROR'
```

**When to change:**
```powershell
# Troubleshooting issues
$Config.LogLevel = 'DEBUG'

# Production (reduce log file size)
$Config.LogLevel = 'INFO'

# Quiet operation (errors only)
$Config.LogLevel = 'ERROR'
```

**Performance impact:**
- DEBUG: ~10% slower (lots of string formatting/I/O)
- INFO: Minimal impact
- WARN/ERROR: Negligible impact

---

### Reporting Settings

#### GenerateHTMLReports
**Purpose:** Enable/disable HTML report generation  
**Type:** Boolean  
**Default:** `$true`  
**Values:** `$true` or `$false`

**Effect:**
- `$true` = Creates detailed HTML report after each operation
- `$false` = No report generated (slightly faster)

**When to disable:**
```powershell
# Automated bulk operations
$Config.GenerateHTMLReports = $false

# Minimize disk writes
$Config.GenerateHTMLReports = $false

# Headless servers
$Config.GenerateHTMLReports = $false
```

**When to keep enabled:**
- Manual migrations (good for documentation)
- Compliance/audit requirements
- Troubleshooting (provides diagnostics)

#### AutoOpenReports
**Purpose:** Automatically open HTML reports in browser  
**Type:** Boolean  
**Default:** `$true`  
**Values:** `$true` or `$false`

**Effect:**
- `$true` = Opens report in default browser when complete
- `$false` = Report saved but not opened

**When to disable:**
```powershell
# Automated scripts
$Config.AutoOpenReports = $false

# Multiple sequential operations
$Config.AutoOpenReports = $false

# Remote/headless systems
$Config.AutoOpenReports = $false
```

---

## Advanced Customization

### Compression Level
**Location:** `Export-UserProfile` function, ~line 3378

**Current:**
```powershell
$args = @('a', '-t7z', '-m0=LZMA2', '-mx=5', "-mmt=$threadCount", ...)
```

**Options:**
```powershell
# Fastest (largest files)
'-mx=1'  # ~70% of original, 3x faster

# Fast (good compression)
'-mx=3'  # ~50% of original, 2x faster

# Balanced (default)
'-mx=5'  # ~40% of original, baseline

# Good compression
'-mx=7'  # ~35% of original, 1.5x slower

# Maximum compression
'-mx=9'  # ~30% of original, 3x slower
```

**Example modification:**
```powershell
# Change line ~3378 from:
$args = @('a', '-t7z', '-m0=LZMA2', '-mx=5', "-mmt=$threadCount", ...)

# To (for maximum compression):
$args = @('a', '-t7z', '-m0=LZMA2', '-mx=9', "-mmt=$threadCount", ...)
```

### Exclusion Patterns
**Location:** `Get-RobocopyExclusions` function, ~line 1200

**Add custom exclusions:**
```powershell
function Get-RobocopyExclusions {
    param([ValidateSet('Export','Import')][string]$Mode)
    
    if ($Mode -eq 'Export') {
        return @(
            '/XF',  # Exclude Files
            'ntuser.dat.LOG1', 'ntuser.dat.LOG2',
            '*.tmp', '*.temp',
            
            # === ADD YOUR CUSTOM FILE EXCLUSIONS HERE ===
            '*.iso', '*.vmdk',           # Virtual machine images
            '*.mp4', '*.avi', '*.mkv',   # Video files
            '*.pst',                     # Outlook archives (optional)
            
            '/XD',  # Exclude Directories
            'AppData\Local\Temp',
            'AppData\LocalLow',
            
            # === ADD YOUR CUSTOM FOLDER EXCLUSIONS HERE ===
            'Videos',                    # Skip Videos folder entirely
            'VirtualBox VMs',            # Skip VMs
            'Downloads',                 # Skip Downloads
            'OneDrive',                  # Skip OneDrive (if synced)
            'Dropbox'                    # Skip Dropbox (if synced)
        )
    }
}
```

### UI Customization
**Location:** Main form initialization, ~line 5000-5200

**Change window title:**
```powershell
# Line ~5100
$global:Form.Text = "Contoso Corp - Profile Migration Tool"
```

**Change colors:**
```powershell
# Header color (line ~5120)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(42, 87, 141)  # Corporate blue

# Success color (line ~4775)
$global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)  # Dark green

# Button colors (lines ~5800-5900)
$btnExport.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)  # Windows blue
```

**Change fonts:**
```powershell
# Main form font (line ~5105)
$global:Form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Header font (line ~5125)
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
```

---

## Configuration Profiles

### Profile 1: Maximum Speed
**Use case:** Dedicated migration workstation, fast hardware

```powershell
$Config = @{
    SevenZipThreads           = 16
    RobocopyThreads           = 32
    RobocopyRetryCount        = 0
    RobocopyRetryWaitSec      = 0
    ProgressUpdateIntervalMs  = 2000
    ExportProgressCheckMs     = 1000
    HashVerificationEnabled   = $false
    GenerateHTMLReports       = $false
    AutoOpenReports           = $false
    LogLevel                  = 'WARN'
}

# Also change compression level to -mx=3 (fast)
```

**Expected impact:** 30-40% faster migrations, less validation

### Profile 2: Maximum Reliability
**Use case:** Critical migrations, compliance requirements

```powershell
$Config = @{
    SevenZipThreads           = 8
    RobocopyThreads           = 16
    RobocopyRetryCount        = 5
    RobocopyRetryWaitSec      = 3
    HiveUnloadMaxAttempts     = 5
    HiveUnloadWaitMs          = 1000
    HashVerificationEnabled   = $true
    GenerateHTMLReports       = $true
    AutoOpenReports           = $true
    LogLevel                  = 'DEBUG'
}

# Also change compression level to -mx=7 (better)
```

**Expected impact:** Slower but maximum data integrity

### Profile 3: Low Resource Usage
**Use case:** Shared systems, background operations

```powershell
$Config = @{
    SevenZipThreads           = 2
    RobocopyThreads           = 4
    RobocopyRetryCount        = 1
    RobocopyRetryWaitSec      = 1
    ProgressUpdateIntervalMs  = 5000
    ExportProgressCheckMs     = 2000
    HashVerificationEnabled   = $true
    GenerateHTMLReports       = $false
    AutoOpenReports           = $false
    LogLevel                  = 'INFO'
}
```

**Expected impact:** Minimal system impact, slower migrations

### Profile 4: Network/Remote
**Use case:** Migrations over slow networks

```powershell
$Config = @{
    SevenZipThreads           = $cpuCores
    RobocopyThreads           = 8
    RobocopyRetryCount        = 10
    RobocopyRetryWaitSec      = 5
    DomainReachabilityTimeout = 10000
    HashVerificationEnabled   = $true
    GenerateHTMLReports       = $true
    LogLevel                  = 'INFO'
}

# Also change compression level to -mx=9 (smallest files)
```

**Expected impact:** Smaller transfer sizes, more resilient

---

## Environment Variables

Some behaviors can be overridden with environment variables (for automation):

```powershell
# Override log level
$env:PROFILE_MIGRATION_LOG_LEVEL = 'DEBUG'

# Override report generation
$env:PROFILE_MIGRATION_NO_REPORTS = '1'

# Override auto-open
$env:PROFILE_MIGRATION_NO_AUTOOPEN = '1'

# Custom 7-Zip path
$env:SEVEN_ZIP_PATH = 'D:\Tools\7-Zip\7z.exe'
```

**Note:** These are not implemented in current version but could be added.

---

## Testing Your Configuration

After making changes:

1. **Syntax check:**
   ```powershell
   powershell -NoProfile -Command {
       $null = [System.Management.Automation.PSParser]::Tokenize(
           (Get-Content 'ProfileMigration.ps1' -Raw), [ref]$null
       )
       Write-Host 'Syntax OK' -ForegroundColor Green
   }
   ```

2. **Test with small profile:**
   - Create test user with minimal data
   - Export with new settings
   - Verify completion
   - Check logs for expected behavior

3. **Benchmark performance:**
   - Export same profile with different settings
   - Compare operation times
   - Monitor resource usage (Task Manager)

4. **Validate results:**
   - Import test export
   - Verify data integrity
   - Check log quality

---

## Troubleshooting Configuration

### Settings not taking effect
**Check:**
1. Edited correct file (not a copy)
2. Restarted script after changes
3. No syntax errors (prevented loading)
4. Variable scope correct (`$Config` is global)

### Performance worse after changes
**Common mistakes:**
```powershell
# Too many threads for storage
$Config.RobocopyThreads = 128  # Way too high!
# Fix: Use 8-32 range

# Thread count exceeds CPU cores
$Config.SevenZipThreads = 64  # But only 8 cores available
# Fix: Use $cpuCores or lower

# Too frequent updates
$Config.ProgressUpdateIntervalMs = 10  # Too fast!
# Fix: Use 250-1000 range
```

### Errors after customization
**Rollback:**
1. Keep original file as `.bak`
2. Copy original over modified version
3. Retest with defaults
4. Make smaller incremental changes

---

## Best Practices

1. **Document changes:**
   - Add comments above modified settings
   - Note why changed and expected benefit
   - Include change date and author

2. **Test incrementally:**
   - Change one setting at a time
   - Verify impact before next change
   - Keep log of what worked/didn't work

3. **Use version control:**
   - Keep original pristine copy
   - Save customized versions with descriptive names
   - Consider Git for change tracking

4. **Profile-specific configs:**
   - Create variants for different scenarios
   - Name files: `ProfileMigration-Fast.ps1`, `ProfileMigration-Reliable.ps1`
   - Document which to use when

5. **Validate regularly:**
   - After Windows updates
   - After PowerShell updates
   - Monthly test migrations

---

**Document Version:** 2.12.8
**Last Updated:** December 2025
---

### AzureAD & Microsoft Graph Settings

#### Microsoft Graph PowerShell Module
**Purpose:** Required for AzureAD profile conversions  
**Type:** Auto-installed dependency  
**Installation:** Automatic on first AzureAD conversion

**Module details:**
- **Name:** Microsoft.Graph.Users
- **Scope:** CurrentUser (no admin rights needed for installation)
- **Permissions:** User.Read.All (granted during first sign-in)
- **Installation trigger:** First time converting to/from AzureAD profile

**First-time setup:**
1. Tool detects module not installed
2. Dialog informs user about installation
3. NuGet provider prompt appears (press Y)
4. Module installs automatically
5. Browser opens for Microsoft Graph authentication
6. Sign in with account that has user read permissions
7. Grant requested permissions
8. Conversion proceeds

**Manual installation (if needed):**
```powershell
Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser -Force
```

**Manual uninstallation:**
```powershell
Uninstall-Module -Name Microsoft.Graph.Users
```

#### AzureAD Username Format
**Requirement:** User Principal Name (UPN) in email format  
**Format:** user@domain.com  
**Examples:**
-  john.doe@company.com
-  jane.smith@contoso.onmicrosoft.com
-  john.doe (missing domain)
-  DOMAIN\john.doe (wrong format)

**UI Guidance:**
- Hint label appears when AzureAD radio button selected
- Validation occurs before Microsoft Graph lookup
- Clear error messages if format incorrect

#### AzureAD Profile Folder Naming
**Behavior:** Extracts username part from UPN  
**Example:**
- Input: john.doe@company.com
- Folder: C:\Users\john.doe
- Not: C:\Users\john.doe@company.com

**Implementation:** Regex pattern ^([^@]+)@ extracts username before @

#### AzureAD Join Requirements
**For conversions to AzureAD:**
- Device must be AzureAD/Entra ID joined
- Verified via WMI query to Win32_ComputerSystem
- Tool provides guided setup if not joined
- Opens Settings  Accounts  Access work or school

**Validation function:** Test-IsAzureADJoined

#### AzureAD SID Format
**Pattern:** S-1-12-1-*  
**Example:** S-1-12-1-775943901-1139658206-3697090227-359018853  
**Validation:** Test-IsAzureADSID function checks pattern match

**Conversion process:**
1. Query Microsoft Graph for user ObjectId
2. Convert ObjectId to Windows SID format
3. Validate SID matches AzureAD pattern
4. Use SID for profile conversion
