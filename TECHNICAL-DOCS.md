# ProfileMigration.ps1 – Technical Documentation

## Overview
ProfileMigration.ps1 is a single-file PowerShell 5.1 GUI tool for Windows profile migration. It supports domain, local, and AzureAD/Entra ID moves, registry hive SID rewriting, multi-threaded file operations, and a modern Windows Forms UI. Production-ready for Windows 11 24H2.

## Architecture
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
- try/catch, log user + technical details
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
