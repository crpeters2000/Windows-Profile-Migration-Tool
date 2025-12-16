# Frequently Asked Questions (FAQ)

## General Questions

### Q: What is this tool used for?
**A:** The Windows Profile Migration Tool transfers complete user profiles between Windows computers. It's designed for IT administrators performing hardware replacements, OS upgrades, or profile recovery operations.

### Q: Is this tool safe to use?
**A:** Yes, when used correctly:
- ✅ Creates automatic backups before any destructive operations
- ✅ Validates data integrity with SHA-256 hashing
- ✅ Tested on Windows 10/11 in production environments
- ✅ Non-destructive merge mode available
- ⚠️ **Always test on non-critical profiles first**
- ⚠️ **Ensure you have backups before migration**

### Q: Do I need special permissions?
**A:** Yes, you must:
- Run PowerShell as **Administrator**
- Have local admin rights on both source and target computers
- For domain users, have domain admin credentials available

### Q: How long does migration take?
**A:** Typical times:
- 5 GB profile: ~5 minutes (2-3 min export + 2-3 min import)
- 20 GB profile: ~12 minutes
- 50 GB profile: ~30 minutes
- 100 GB profile: ~60 minutes

*Performance varies based on CPU cores, disk speed, and file types*

---

## Pre-Migration

### Q: What gets migrated?
**A:** Included:
- ✅ Desktop files and shortcuts
- ✅ Documents, Pictures, Videos, Music, Downloads
- ✅ Browser bookmarks and history (Edge, Chrome, Firefox)
- ✅ Application settings (AppData folder)
- ✅ Windows preferences and registry settings
- ✅ Start Menu layout and taskbar pins
- ✅ File Explorer favorites
- ✅ Mapped network drives (configurations, not credentials)
- ✅ Printer configurations
- ✅ Installed fonts (user-level)

**Not included:**
- ❌ Passwords (Windows DPAPI-encrypted, may not work on different computer)
- ❌ Outlook OST files (excluded intentionally, rebuilt from Exchange)
- ❌ Temporary files and caches
- ❌ Windows Search index
- ❌ Browser caches
- ❌ System files

### Q: Can I migrate from Windows 10 to Windows 11?
**A:** Yes, fully supported. The tool handles OS version differences automatically.

### Q: Can I migrate between 32-bit and 64-bit Windows?
**A:** Generally yes, but some applications may have issues:
- User files and documents: ✅ Full compatibility
- Registry settings: ✅ Compatible
- 32-bit applications: ⚠️ May need reinstallation on 64-bit
- 64-bit applications: ❌ Won't work on 32-bit target

**Recommendation:** Stick to same architecture when possible.

### Q: What about antivirus software?
**A:** During migration:
1. Disable real-time scanning temporarily (can block file operations)
2. Run migration
3. Re-enable after completion

Most antivirus software can be migrated, but license reactivation may be required.

### Q: Can I migrate a profile while the user is logged in?
**A:** **No, strongly discouraged:**
- ❌ Risk of file corruption
- ❌ Incomplete data capture
- ❌ Registry inconsistencies
- ✅ Tool will warn you if user is logged in
- ✅ Best practice: Log out user first

**Exception:** Export can work if user is logged in elsewhere (different session), but not recommended.

---

## Export Questions

### Q: Where should I save the exported ZIP file?
**A:** Options:
- ✅ **USB drive** - Direct physical transfer (fastest, most secure)
- ✅ **Network share** - Convenient for multiple migrations
- ✅ **External hard drive** - Good for large profiles
- ⚠️ **OneDrive/Dropbox** - Works but slow for large files
- ❌ **Email** - Too large, security risk

**Recommendation:** USB 3.0 drive or SMB network share

### Q: The export is taking forever. Is it stuck?
**A:** Check these:
1. Look at progress percentage in status bar
2. Check log viewer for activity
3. Open Task Manager:
   - CPU at 100%? Normal during compression
   - Disk at 100%? Slow disk, but working
   - Both low? May be stuck on locked file

**If truly stuck:**
- Click Cancel button
- Close any open applications
- Retry export

### Q: How much disk space do I need for export?
**A:** Formula:
```
Required Space = (Profile Size × 1.5) + Compressed ZIP

Example:
Profile: 20 GB
Temporary: 30 GB (during compression)
Final ZIP: ~8 GB (40% compression ratio)
Total needed: ~38 GB
```

**Tip:** Check with `Get-PSDrive C:`

### Q: Can I compress to a different drive?
**A:** Yes! When the save dialog appears:
- Navigate to D:\ or any other drive
- Select location
- Click Save

### Q: What if 7-Zip is not installed?
**A:** The tool will:
1. Detect missing 7-Zip
2. Show error with download link
3. You install 7-Zip from https://www.7-zip.org/
4. Restart tool
5. Proceed with export

**Installation:**
- Use default location: `C:\Program Files\7-Zip\`
- Install 64-bit version (if on 64-bit Windows)

### Q: Can I export multiple profiles at once?
**A:** No, one profile at a time. For batch operations:
1. Export first profile
2. Wait for completion
3. Select next profile
4. Repeat

**Tip:** Exports can run sequentially without closing the tool.

### Q: What's the .sha256 file for?
**A:** It's a checksum file that:
- Verifies ZIP wasn't corrupted during transfer
- Detects tampering or incomplete copies
- Automatically validated on import (if present)

**Keep it with the ZIP file!**

---

## Import Questions

### Q: Do I need to create the user account first?
**A:** Depends:

**Local User:**
- Tool will prompt to create if doesn't exist
- You set password during import
- Optional admin rights checkbox

**Domain User:**
- Account must exist in Active Directory
- Computer can be joined to domain using built-in Domain Join feature (if not already joined)
- You provide domain credentials during import

### Q: What's the difference between Merge and Replace mode?
**A:** 

**Replace Mode (Recommended):**
- Backs up existing profile
- Deletes old profile
- Imports fresh profile
- **Use when:** Clean migration, old profile corrupted

**Merge Mode (Advanced):**
- Keeps existing profile
- Copies imported files alongside existing
- Doesn't replace newer files
- Preserves current NTUSER.DAT (registry settings)
- **Use when:** Adding files from backup, supplementing profile

**Most users should choose Replace.**

### Q: The import failed. What do I do?
**A:** Check the error message and logs:

**Common errors:**

| Error | Meaning | Fix |
|-------|---------|-----|
| "Profile not found" | Username doesn't match | Check spelling, use dropdown |
| "ZIP not found" | Wrong path | Click Browse, select correct ZIP |
| "Hash verification failed" | Corrupted transfer | Re-copy ZIP from source |
| "User logged on" | Target user active | Log out user, retry |
| "Access denied" | Not administrator | Right-click > Run as Administrator |
| "Disk full" | No space | Free up space on C: drive |

**General recovery:**
1. Review log file: `Import-username-timestamp.log`
2. Check HTML report for details
3. Fix issue listed
4. Retry import

### Q: Can I import to a different username?
**A:** Yes!
1. Click Browse and select ZIP
2. Tool auto-fills username from ZIP filename
3. **Edit the username field** to desired target name
4. Click Import
5. Profile is imported to new username

**Example:**
- Export: `john-Export-20251207.zip`
- Import as: `jane` (just edit the text field)
- Result: Profile imported to `C:\Users\jane`

### Q: Import succeeded but user gets temporary profile. Why?
**A:** This means NTUSER.DAT failed to load. Causes:

1. **Corrupted hive during transfer**
   - Re-export from source
   - Verify hash matches

2. **Permissions issue**
   - Check logs for "Access Denied"
   - Re-run import as Administrator

3. **Antivirus blocking**
   - Temporarily disable AV
   - Retry import

4. **SID rewriting failed**
   - Check logs for SID errors
   - May need domain admin credentials

**Quick fix:**
```powershell
# Test if hive can load
reg load "HKU\TEST" "C:\Users\john\NTUSER.DAT"
reg unload "HKU\TEST"

# If error, restore from backup
Copy-Item "C:\Users\john.backup_TIMESTAMP\NTUSER.DAT" "C:\Users\john\" -Force
```

### Q: After import, applications don't work. Why?
**A:** Common issues:

**Outlook:**
- OST file was excluded (intentionally)
- Wait 10-30 minutes for Exchange to rebuild cache
- Check status bar: "Updating folders..."

**OneDrive:**
- Needs to re-link to account
- Settings > Unlink this PC
- Sign in again
- Wait for re-sync

**Chrome/Edge:**
- Bookmarks should migrate
- If missing, check `AppData\Local\Microsoft\Edge\User Data`
- Restore from import backup if needed

**VPN/Network Drives:**
- Credentials don't migrate (Windows DPAPI limitation)
- Re-enter passwords manually
- Mapped drives configurations migrate, but not passwords

### Q: How do I rollback a failed import?
**A:** Easy! Automatic backup created:

1. Boot to Safe Mode (if profile won't load normally)
2. Login as Administrator
3. Delete corrupted profile:
   ```
   Remove-Item "C:\Users\john" -Recurse -Force
   ```
4. Rename backup folder:
   ```
   Rename-Item "C:\Users\john.backup_20251207_143022" "C:\Users\john"
   ```
5. Reboot normally
6. User logs in with original profile restored

**Backup location:** `C:\Users\username.backup_TIMESTAMP`

### Q: Can I delete the backup folder?
**A:** Yes, but wait:
- ✅ After 30 days of confirmed working profile
- ✅ After user verifies all files and apps work
- ✅ After you have other backups

**Don't delete:**
- ❌ Immediately after import
- ❌ If user hasn't tested everything
- ❌ If you have no other backup

---

## Domain & Network

### Q: How do I migrate domain users?
**A:** Process:

**Source Computer:**
1. Export profile normally
2. Copy ZIP to target

**Target Computer:**
1. Join computer to domain if needed (use built-in Domain Join tab, or pre-join manually)
2. Click Browse, select ZIP
3. Enter username as: `DOMAIN\username`
4. Click Import
5. Provide domain admin credentials when prompted
6. Wait for completion
7. Reboot
8. User logs in with domain credentials

### Q: Can I join the computer to domain using this tool?
**A:** Yes! Built-in domain join:

1. Switch to "Domain Join" tab
2. Enter domain: `CONTOSO.COM`
3. Click "Test Domain Connectivity"
4. If green, click "Join Domain"
5. Enter domain admin credentials
6. Choose restart option
7. Computer joins and reboots

**Requirements:**
- Network connectivity to domain controller
- DNS configured to find domain
- Domain admin credentials

### Q: Can I migrate AzureAD/Entra ID profiles?
**A:** Yes! Full support for Microsoft Entra ID (formerly AzureAD):

**Export (automatic):**
- Tool detects AzureAD profiles by SID pattern (S-1-12-1)
- Profile is tagged as AzureAD in manifest
- Export works normally

**Import (requires setup):**
1. Target computer **must be Entra ID joined**
2. User **must sign in** with work/school account at least once
3. Enter username as: `AzureAD\username` (e.g., `AzureAD\john.doe`)
4. Tool validates device join status
5. If not joined, tool opens Settings → Access work or school
6. Complete join, sign in, then retry import

**Key differences from domain migration:**
- Username format: `AzureAD\username` (NOT email address)
- No domain admin credentials needed
- Device must be joined via Settings app
- Tool provides guided setup with `ms-settings:workplace` link

### Q: Domain join fails. What's wrong?
**A:** Check these:

1. **Network connectivity**
   ```powershell
   Test-NetConnection -ComputerName CONTOSO.COM -Port 389  # LDAP
   ```

2. **DNS configuration**
   ```powershell
   nslookup CONTOSO.COM
   # Should return domain controller IPs
   ```

3. **Time sync**
   ```powershell
   w32tm /query /status
   # Should show synced with domain
   ```

4. **Firewall**
   - Allow ports: 88 (Kerberos), 389 (LDAP), 445 (SMB), 53 (DNS)

5. **Credentials**
   - Use format: `DOMAIN\admin` or `admin@DOMAIN.COM`
   - Verify password is correct

### Q: Can I migrate from domain to local user?
**A:** Yes, but limitations:

**What works:**
- ✅ Files and folders migrate
- ✅ Desktop, Documents, etc.
- ✅ Most application settings

**What needs reconfiguration:**
- ⚠️ Domain-specific apps may break
- ⚠️ Group Policy settings lost
- ⚠️ Domain credentials don't work
- ⚠️ Network resources need local credentials

**Process:**
1. Export domain profile: `DOMAIN\john`
2. Import as local user: `john`
3. User loses domain features
4. Works as standard local account

---

## Performance & Optimization

### Q: Can I make it faster?
**A:** Yes, several ways:

**1. Use faster storage:**
- NVMe SSD >> SATA SSD >> HDD
- USB 3.0 >> USB 2.0
- Local drives >> Network shares

**2. Increase thread count (advanced):**
Edit `$Config` at top of script:
```powershell
$Config = @{
    RobocopyThreads = 32    # Increase from default
    SevenZipThreads = 16    # Use more CPU cores
}
```

**3. Reduce compression level (advanced):**
Edit Export-UserProfile function:
```powershell
# Change -mx=5 to -mx=3
$args = @('a', '-t7z', '-m0=LZMA2', '-mx=3', ...)
```

**4. Exclude large unnecessary folders:**
Edit Get-RobocopyExclusions:
```powershell
'/XD',
'Videos',      # Skip videos
'Downloads',   # Skip downloads
'VirtualBox VMs'  # Skip VM files
```

### Q: Why is my export only using 25% CPU?
**A:** Depends on bottleneck:

**CPU at 25% on 4-core system = 100% on 1 core:**
- Compression maxing out single thread
- Normal, increase `SevenZipThreads`

**CPU low, Disk at 100%:**
- Disk is bottleneck, not CPU
- Upgrade to SSD
- Reduce thread count (disk can't keep up)

**Both low:**
- Waiting on I/O
- Check for antivirus scanning
- Close other applications

### Q: How do I reduce the ZIP file size?
**A:** Methods:

**1. Run Disk Cleanup first:**
```powershell
cleanmgr /d C:
# Select user profile drive
# Check all boxes
```

**2. Clear browser caches manually:**
```
Edge: edge://settings/clearBrowserData
Chrome: chrome://settings/clearBrowserData
Firefox: about:preferences#privacy > Clear Data
```

**3. Exclude large media files:**
Edit exclusions (advanced):
```powershell
'/XF', '*.mp4', '*.avi', '*.mkv', '*.iso'
```

**4. Increase compression level:**
```powershell
-mx=9  # Best compression (slower)
```

**5. Delete Downloads folder before export:**
- Manually empty Downloads
- Or exclude in script

**Typical compression ratios:**
- Documents: 50-70% reduction
- Pictures: 5-10% reduction (already compressed)
- Videos: 1-5% reduction (already compressed)
- Overall: 40-60% reduction

---

## Troubleshooting

### Q: I get "Access Denied" errors. What do I do?
**A:** Checklist:

1. ✅ **Run as Administrator**
   - Right-click PowerShell
   - Select "Run as administrator"
   - Retry

2. ✅ **Check file permissions**
   ```powershell
   Get-Acl "C:\Users\john" | Format-List
   ```

3. ✅ **Disable antivirus temporarily**
   - Antivirus may block operations
   - Disable, retry, re-enable

4. ✅ **Check if files are locked**
   - Close all user applications
   - Use Process Explorer to find locks

5. ✅ **Take ownership (last resort)**
   ```powershell
   takeown /F "C:\Users\john" /R /D Y
   icacls "C:\Users\john" /grant Administrators:F /T
   ```

### Q: Tool crashes or freezes. What's wrong?
**A:** Common causes:

**1. Out of memory:**
- Close other applications
- Upgrade RAM (8GB minimum recommended)

**2. Corrupted profile:**
- Try exporting different profile (test)
- Check source disk for errors: `chkdsk /F`

**3. PowerShell version:**
```powershell
$PSVersionTable.PSVersion
# Should be 5.1 or higher
```

**4. .NET Framework:**
- Ensure .NET 4.7.2+ installed
- Windows Update usually handles this

**5. Antivirus interference:**
- Add PowerShell to exclusions
- Temporarily disable

### Q: "Hash verification failed" - should I continue?
**A:** Depends on reason:

**If hash file is outdated:**
- ✅ Safe to continue
- Generate new hash: `Get-FileHash -Path file.zip -Algorithm SHA256`

**If file was modified:**
- ❌ Don't continue
- Re-copy from source
- Verify copy completed successfully

**If transferred over network:**
- ⚠️ May be corruption
- Compare file sizes (source vs transferred)
- If same, probably safe
- If different, re-copy

**If from old backup:**
- ⚠️ Possible bit rot
- Try different backup
- Or re-export from source if possible

### Q: User can't login after import. What do I do?
**A:** Diagnosis:

**1. Check account exists:**
```powershell
Get-LocalUser -Name john
# or for domain:
Get-ADUser -Identity john
```

**2. Verify password:**
- Reset password if needed
- Ensure account is enabled

**3. Check profile path:**
```powershell
$sid = "S-1-5-21-..."  # User SID
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
```

**4. Look at login error:**
- "User profile service failed logon" = NTUSER.DAT issue
- "Trust relationship failed" = Domain connectivity
- "Account is disabled" = Re-enable account

**5. Test with different account:**
- If other users can login → Profile specific issue
- If no one can login → System issue

### Q: Some files are missing after import. Why?
**A:** Check these:

**1. Intentionally excluded:**
- Temp files, caches (normal)
- Outlook OST (normal, rebuilds)
- Browser cache (normal, rebuilds)

**2. Hidden files:**
```powershell
# Show hidden files in Explorer
# View > Show > Hidden items
```

**3. File permissions:**
```powershell
# Check if you have access
Get-ChildItem "C:\Users\john\Desktop" -Force
```

**4. Still in backup:**
```powershell
# Compare backup to current
Compare-Object (Get-ChildItem "C:\Users\john\Desktop") `
               (Get-ChildItem "C:\Users\john.backup_TIMESTAMP\Desktop")
```

**5. Never in export:**
- Check export log file
- Look for "excluding" messages
- May have been in excluded folder

### Q: How do I view detailed logs?
**A:** Multiple ways:

**1. Built-in log viewer:**
- Bottom panel of tool UI
- Use search box to filter
- Change level dropdown (DEBUG for all messages)

**2. Log files on disk:**
```
Export: Export-username-YYYYMMDD_HHMMSS.log
Import: Import-username-YYYYMMDD_HHMMSS.log
Location: Same folder as ZIP file
```

**3. Open in Notepad:**
```powershell
notepad "Export-john-20251207_143022.log"
```

**4. Search logs:**
```powershell
Select-String -Path "Export-*.log" -Pattern "ERROR"
```

**5. Enable DEBUG level:**
- UI: Dropdown at bottom > Select "DEBUG"
- Script: `$Config.LogLevel = 'DEBUG'`

---

## Advanced Topics

### Q: Can I script/automate migrations?
**A:** Yes, the tool can be called programmatically:

```powershell
# Source the script
. .\ProfileMigration.ps1

# Export profile
Export-UserProfile -Username "john" -OutputPath "C:\Exports\john.7z"

# Import profile  
Import-UserProfile -Username "jane" -ZipPath "C:\Exports\john.7z"
```

**For bulk operations:**
```powershell
$users = @('john', 'jane', 'bob')
foreach ($user in $users) {
    Export-UserProfile -Username $user -OutputPath "C:\Exports\$user.7z"
}
```

**Caution:** Testing recommended before automation.

### Q: Can I customize what gets migrated?
**A:** Yes, edit the exclusion function:

```powershell
function Get-RobocopyExclusions {
    param([string]$Mode)
    
    if ($Mode -eq 'Export') {
        return @(
            '/XF',
            # Add files to exclude
            '*.tmp', '*.log', 'desktop.ini',
            
            '/XD',
            # Add folders to exclude
            'Downloads', 'Videos', 'Temp'
        )
    }
}
```

### Q: Can I migrate to a network drive?
**A:** For export, yes:
```powershell
# Export directly to network share
# When save dialog appears, navigate to:
\\server\share\ProfileBackups\
```

For import, not recommended:
- Import needs to extract to `C:\Users\`
- Can import FROM network ZIP
- But destination must be local drive

### Q: What about Server Core or no GUI?
**A:** PowerShell cmdlets work without GUI:

```powershell
# Suppress GUI, use cmdlets directly
$global:Form = $null  # Disable UI

# Call functions directly
Export-UserProfile -Username "john" -OutputPath "C:\export.7z"
Import-UserProfile -Username "john" -ZipPath "C:\export.7z"
```

**Or write wrapper script:**
```powershell
# HeadlessMigration.ps1
. .\ProfileMigration.ps1

param($Action, $User, $Path)

if ($Action -eq 'Export') {
    Export-UserProfile -Username $User -OutputPath $Path
} elseif ($Action -eq 'Import') {
    Import-UserProfile -Username $User -ZipPath $Path
}
```

Usage:
```powershell
.\HeadlessMigration.ps1 -Action Export -User john -Path C:\export.7z
.\HeadlessMigration.ps1 -Action Import -User john -Path C:\export.7z
```

---

## Best Practices

### Q: What's the recommended migration workflow?
**A:** Step-by-step:

**Preparation:**
1. ✅ Document current setup (screenshots, notes)
2. ✅ Verify backup strategy exists
3. ✅ Test on non-critical profile first
4. ✅ Schedule during off-hours/downtime
5. ✅ Notify affected users

**Source Computer:**
1. ✅ Update all applications
2. ✅ Run Windows Update
3. ✅ Run Disk Cleanup
4. ✅ Log out user
5. ✅ Login as admin
6. ✅ Run export
7. ✅ Verify ZIP created successfully
8. ✅ Copy ZIP + SHA256 to secure location

**Target Computer:**
1. ✅ Complete Windows setup/updates
2. ✅ Join domain (if applicable)
3. ✅ Install required applications
4. ✅ Copy ZIP + SHA256 to local drive
5. ✅ Run import
6. ✅ Verify success message
7. ✅ Review HTML report
8. ✅ Reboot immediately

**Verification:**
1. ✅ Login as migrated user
2. ✅ Verify Desktop/Documents present
3. ✅ Test applications (Outlook, browsers)
4. ✅ Check network drives
5. ✅ Test printers
6. ✅ Verify all data accessible
7. ✅ Keep backup 30 days

### Q: How often should I test the tool?
**A:** Recommended schedule:

- **Before production use:** Test on 3-5 different profiles
- **Monthly:** Quick test migration to verify still working
- **After Windows updates:** Verify compatibility
- **Before mass migrations:** Test again on sample set

### Q: Should I keep export ZIPs as backups?
**A:** Good idea! Benefits:

- ✅ Point-in-time profile snapshot
- ✅ Quick recovery if profile corrupts
- ✅ User can revert to previous state
- ✅ Compliance/audit trail

**Storage strategy:**
```
\\Server\ProfileBackups\
├── john\
│   ├── john-20251201_weekly.7z
│   ├── john-20251207_pre-upgrade.7z
│   └── john-20251215_monthly.7z
```

**Retention policy example:**
- Daily backups: Keep 7 days
- Weekly backups: Keep 4 weeks  
- Monthly backups: Keep 12 months
- Major events: Keep indefinitely

---

## Getting Help

### Q: Where can I get more help?
**A:** Resources:

1. **Documentation:**
   - README.md - Overview and features
   - USER-GUIDE.md - Step-by-step instructions
   - TECHNICAL-DOCS.md - Advanced troubleshooting
   - This FAQ

2. **Diagnostic info:**
   - Log files (Export/Import-*.log)
   - HTML reports
   - Windows Event Viewer

3. **IT Support:**
   - Contact your IT administrator
   - Provide log files and HTML report

4. **Community:**
   - Search for similar issues
   - Check Windows admin forums

### Q: How do I report a bug?
**A:** Include this information:

1. **System details:**
   ```powershell
   winver  # Windows version
   $PSVersionTable.PSVersion  # PowerShell version
   Get-ComputerInfo | Select-Object WindowsProductName, OsArchitecture
   ```

2. **Profile details:**
   - Profile size
   - Export or import?
   - Local or domain user?

3. **Error details:**
   - Complete error message
   - Full log file contents
   - HTML report (if generated)

4. **Steps to reproduce:**
   - What you did
   - What you expected
   - What actually happened

5. **Environment:**
   - Antivirus software
   - Domain vs workgroup
   - Network or local storage

---

## Miscellaneous

### Q: Is there a GUI-only version?
**A:** The current version IS GUI-based! It has:
- Modern Windows Forms interface
- Point-and-click operation
- Progress bars and visual feedback
- Built-in log viewer

**No command-line required** (except launching the script).

### Q: Can I use this for server profiles?
**A:** Yes, but considerations:

**Works fine for:**
- ✅ Terminal Server user profiles
- ✅ RDS user profiles
- ✅ Standard user accounts on servers

**May have issues with:**
- ⚠️ Service accounts (complex permissions)
- ⚠️ System accounts
- ⚠️ IIS application pool identities

**Recommendation:** Test thoroughly on server platforms.

### Q: What languages is this available in?
**A:** Currently English only. The tool displays:
- English UI labels
- English log messages
- English error messages

**Customization:** Advanced users can edit UI strings in the script.

### Q: Can I rebrand/customize the UI?
**A:** Yes! Edit these sections:

**Window title (line ~5100):**
```powershell
$global:Form.Text = "Your Company - Profile Migration"
```

**Colors (lines ~5000-5200):**
```powershell
$headerColor = [System.Drawing.Color]::FromArgb(0, 120, 212)  # Blue
$successColor = [System.Drawing.Color]::FromArgb(16, 124, 16)  # Green
```

**Logo (advanced):**
Add company logo image to form:
```powershell
$logo = New-Object System.Windows.Forms.PictureBox
$logo.Image = [System.Drawing.Image]::FromFile("C:\Path\Logo.png")
$logo.SizeMode = 'Zoom'
$global:Form.Controls.Add($logo)
```

### Q: Is this tool open source?
**A:** Check the LICENSE file in the distribution for terms of use and redistribution rights.

---

**Last Updated:** December 2025  
**Version:** 1.0  

**Didn't find your answer?** Check the [Technical Documentation](TECHNICAL-DOCS.md) or contact IT support.
