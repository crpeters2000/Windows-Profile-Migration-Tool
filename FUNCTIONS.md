# Profile Migration Tool - Function Reference

This document provides a comprehensive reference of all functions in the Profile Migration Tool.

**Total Functions:** 82
- **ProfileMigration.ps1:** 17 functions (Business Logic)
- **Functions.ps1:** 65 functions (Helper/Utility)

---

# ProfileMigration.ps1 Functions

**Business Logic Functions (17 total)**

These functions contain the core business logic, UI workflows, and main operations.

---

## Confirm-DomainUnjoin

**File:** ProfileMigration.ps1

**Defined at:** Line 46

**Description:** Function: Confirm-DomainUnjoin

**Called at lines:** *(Not called directly or only called dynamically)*

---

## Generate-MigrationReport

**File:** ProfileMigration.ps1

**Defined at:** Line 71

**Description:** Function: Generate-MigrationReport

**Called at lines:** 5909, 5957, 7853, 8012

**Total calls:** 4

---

## New-ConversionReport

**File:** ProfileMigration.ps1

**Defined at:** Line 770

**Description:** Function: New-ConversionReport

**Called at lines:** 3854, 3893

**Total calls:** 2

---

## Start-OperationLog

**File:** ProfileMigration.ps1

**Defined at:** Line 1024

**Description:** Function: Start-OperationLog

**Called at lines:** 5152, 5987

**Total calls:** 2

---

## Stop-OperationLog

**File:** ProfileMigration.ps1

**Defined at:** Line 1072

**Description:** Function: Stop-OperationLog

**Called at lines:** 5970, 8025

**Total calls:** 2

---

## Repair-UserProfile

**File:** ProfileMigration.ps1

**Defined at:** Line 1084

**Description:** Enhanced profile detection with size estimates

**Called at lines:** 3446, 3481, 3543

**Total calls:** 3

---

## Convert-LocalToDomain

**File:** ProfileMigration.ps1

**Defined at:** Line 1150

**Description:** Convert a local user profile to a domain user profile

**Called at lines:** 3456, 3531

**Total calls:** 2

---

## Convert-DomainToLocal

**File:** ProfileMigration.ps1

**Defined at:** Line 1459

**Description:** Convert a domain user profile to a local user profile

**Called at lines:** 3464, 3496

**Total calls:** 2

---

## Convert-AzureADToLocal

**File:** ProfileMigration.ps1

**Defined at:** Line 1702

**Description:** Function: Convert-AzureADToLocal

**Called at lines:** 3503

**Total calls:** 1

---

## Convert-LocalToAzureAD

**File:** ProfileMigration.ps1

**Defined at:** Line 1887

**Description:** Function: Convert-LocalToAzureAD

**Called at lines:** 3512, 3520

**Total calls:** 2

---

## Show-ProfileConversionDialog

**File:** ProfileMigration.ps1

**Defined at:** Line 2157

**Description:** Show Profile Conversion Dialog

**Called at lines:** 8239

**Total calls:** 1

---

## Handle-Restart

**File:** ProfileMigration.ps1

**Defined at:** Line 3935

**Description:** Function: Handle-Restart

**Called at lines:** 4045, 4072, 4186

**Total calls:** 3

---

## Join-Domain-Enhanced

**File:** ProfileMigration.ps1

**Defined at:** Line 4015

**Description:** Function: Join-Domain-Enhanced

**Called at lines:** 2787, 6708, 9071

**Total calls:** 3

---

## Show-ProfileCleanupWizard

**File:** ProfileMigration.ps1

**Defined at:** Line 4208

**Description:** Function: Show-ProfileCleanupWizard

**Called at lines:** 5172

**Total calls:** 1

---

## Export-UserProfile

**File:** ProfileMigration.ps1

**Defined at:** Line 5141

**Description:** Function: Export-UserProfile

**Called at lines:** 3066, 8852

**Total calls:** 2

---

## Import-UserProfile

**File:** ProfileMigration.ps1

**Defined at:** Line 5979

**Description:** Function: Import-UserProfile

**Called at lines:** 8902

**Total calls:** 1

---

## Test-AndFix-ProfileHive

**File:** ProfileMigration.ps1

**Defined at:** Line 7483

**Description:** Function: Test-AndFix-ProfileHive

**Called at lines:** 7534

**Total calls:** 1

---


# Functions.ps1 Module

**Helper/Utility Functions (65 total)**

These reusable helper functions are organized into 6 categories:
- Logging & UI (15 functions)
- Profile Management (20 functions)
- Registry & File System (12 functions)
- Domain & Azure AD (10 functions)
- Applications & Packages (5 functions)
- Archive & Validation (3 functions)

---

## Log-Message

**File:** Functions.ps1

**Defined at:** Line 5

**Description:** Function: Log-Message

**Called at lines (within Functions.ps1):** 69, 72, 75, 78, 789, 795, 1015, 1024, 1046, 1066, 1071, 1091, 1103, 1111, 1367, 1371, 1386, 1393, 1394, 1408, 1410, 1413, 1421, 1490, 1510, 1757, 1763, 1791, 1802, 1805, 1810, 1811, 1817, 1823, 1824, 1826, 1833, 1834, 1840, 1843, 1865, 1871, 1879, 1889, 1894, 1900, 1903, 1907, 1922, 1929, 1936, 1943, 1954, 1962, 1969, 1976, 1983, 1987, 1991, 2006, 2009, 2014, 2017, 2024, 2029, 2039, 2044, 2051, 2056, 2062, 2069, 2076, 2081, 2097, 2100, 2111, 2116, 2123, 2127, 2130, 2137, 2141, 2144, 2164, 2172, 2177, 2208, 2212, 2218, 2222, 2225, 2253, 2260, 2264, 2292, 2302, 2306, 2313, 2333, 2334, 2335, 2336, 2337, 2365, 2368, 2371, 2376, 2382, 2385, 2390, 2397, 2398, 2410, 2419, 2426, 2430, 2439, 2470, 2475, 2478, 2485, 2488, 2493, 2494, 2497, 2501, 2502, 2504, 2509, 2518, 2524, 2531, 2533, 2541, 2551, 2556, 2573, 2590, 2595, 2636, 2641, 2663, 2676, 2684, 2692, 2704, 2737, 2740, 2751, 2764, 2767, 2779, 2782, 2789, 2793, 2802, 2808, 2813, 2820, 2823, 2824, 2835, 2842, 2845, 2850, 2863, 2866, 2867, 2878, 2884, 2891, 2895, 2942, 3178, 3196, 3209, 3382, 3406, 3424, 3449, 3481, 3485, 3491, 3496, 3504, 3508, 3512, 3516, 3522, 3697, 3704, 3993, 4017, 4021, 4035, 4042, 4047, 4060, 4065, 4237, 4252, 4265, 4277, 4292, 4296, 4301, 4306, 4312

**Total calls:** 208

---

## Log-Debug

**File:** Functions.ps1

**Defined at:** Line 69

**Description:** Function: Log-Debug

**Called at lines (within Functions.ps1):** 1097, 1137, 1158, 1165, 1189, 1200, 1204, 1686

**Total calls:** 8

---

## Log-Info

**File:** Functions.ps1

**Defined at:** Line 72

**Description:** Function: Log-Info

**Called at lines (within Functions.ps1):** 1238, 1244, 1251, 1274, 1282, 1290, 1295, 1300, 1305, 1325, 1522, 1534, 1541, 1545, 1551, 1639, 1650, 1662, 1666, 1681, 1694, 1698, 1704, 1706, 1717, 1719, 1722, 1726, 3745, 3756, 3761, 3800, 3812, 3817, 3822, 3853, 3883, 3887, 3894, 3914, 3916, 4329, 4521, 4538, 4539, 4747, 4768

**Total calls:** 47

---

## Log-Warning

**File:** Functions.ps1

**Defined at:** Line 75

**Description:** Function: Log-Warning

**Called at lines (within Functions.ps1):** 936, 942, 975, 1208, 1657, 1690, 1709, 3025, 3906, 3922, 3929, 3933, 4038, 4346, 4545

**Total calls:** 15

---

## Log-Error

**File:** Functions.ps1

**Defined at:** Line 78

**Description:** Function: Log-Error

**Called at lines (within Functions.ps1):** 980, 1342, 1571, 1735, 3769, 3777, 3830, 3838, 3938, 4750, 4772

**Total calls:** 11

---

## Refresh-LogDisplay

**File:** Functions.ps1

**Defined at:** Line 81

**Description:** Function: Refresh-LogDisplay

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Show-ModernDialog

**File:** Functions.ps1

**Defined at:** Line 101

**Description:** Function: Show-ModernDialog

**Called at lines (within Functions.ps1):** 1530, 1658, 1714, 3386, 3395, 3425, 3428, 3453, 3462, 3705, 4068, 4318, 4677, 4680, 4806, 4812, 4819

**Total calls:** 17

---

## Show-InputDialog

**File:** Functions.ps1

**Defined at:** Line 249

**Description:** Function: Show-InputDialog

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Show-LogViewer

**File:** Functions.ps1

**Defined at:** Line 380

**Description:** Function: Show-LogViewer

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Test-InternetConnectivity

**File:** Functions.ps1

**Defined at:** Line 579

**Description:** Function: Test-InternetConnectivity

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Get-WindowsTheme

**File:** Functions.ps1

**Defined at:** Line 606

**Description:** Function: Get-WindowsTheme

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Apply-Theme

**File:** Functions.ps1

**Defined at:** Line 623

**Description:** Function: Apply-Theme

**Called at lines (within Functions.ps1):** 755

**Total calls:** 1

---

## Toggle-Theme

**File:** Functions.ps1

**Defined at:** Line 753

**Description:** Function: Toggle-Theme

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Test-PathWithRetry

**File:** Functions.ps1

**Defined at:** Line 759

**Description:** Function: Test-PathWithRetry

**Called at lines (within Functions.ps1):** 1012, 1021

**Total calls:** 2

---

## Get-ProfileInfo

**File:** Functions.ps1

**Defined at:** Line 807

**Description:** Function: Get-ProfileInfo

**Called at lines (within Functions.ps1):** 865

**Total calls:** 1

---

## Get-ProfileDisplayEntries

**File:** Functions.ps1

**Defined at:** Line 856

**Description:** Function: Get-ProfileDisplayEntries

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Get-ProfileType

**File:** Functions.ps1

**Defined at:** Line 923

**Description:** Function: Get-ProfileType

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Test-ValidProfilePath

**File:** Functions.ps1

**Defined at:** Line 986

**Description:** Function: Test-ValidProfilePath

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Test-ProfilePathWriteable

**File:** Functions.ps1

**Defined at:** Line 1033

**Description:** Function: Test-ProfilePathWriteable

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Test-ProfileMounted

**File:** Functions.ps1

**Defined at:** Line 1081

**Description:** Function: Test-ProfileMounted

**Called at lines (within Functions.ps1):** 4805, 4811

**Total calls:** 2

---

## Test-UserLoggedOut

**File:** Functions.ps1

**Defined at:** Line 1117

**Description:** Function: Test-UserLoggedOut

**Called at lines (within Functions.ps1):** 1239

**Total calls:** 1

---

## Test-ProfileConversionPreconditions

**File:** Functions.ps1

**Defined at:** Line 1215

**Description:** Function: Test-ProfileConversionPreconditions

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Get-LocalProfiles

**File:** Functions.ps1

**Defined at:** Line 1352

**Description:** Function: Get-LocalProfiles

**Called at lines (within Functions.ps1):** 861

**Total calls:** 1

---

## Get-LocalUserSID

**File:** Functions.ps1

**Defined at:** Line 1361

**Description:** Function: Get-LocalUserSID

**Called at lines (within Functions.ps1):** 933, 1229, 1756

**Total calls:** 3

---

## Convert-SIDToAccountName

**File:** Functions.ps1

**Defined at:** Line 1442

**Description:** Function: Convert-SIDToAccountName

**Called at lines (within Functions.ps1):** 875, 1384

**Total calls:** 2

---

## Test-IsAzureADSID

**File:** Functions.ps1

**Defined at:** Line 1496

**Description:** Function: Test-IsAzureADSID

**Called at lines (within Functions.ps1):** 947, 1383, 1409, 2088, 2203

**Total calls:** 5

---

## Test-IsAzureADJoined

**File:** Functions.ps1

**Defined at:** Line 1503

**Description:** Function: Test-IsAzureADJoined

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Get-AzureADUserSID

**File:** Functions.ps1

**Defined at:** Line 1516

**Description:** Function: Get-AzureADUserSID

**Called at lines (within Functions.ps1):** 1402

**Total calls:** 1

---

## Convert-EntraObjectIdToSid

**File:** Functions.ps1

**Defined at:** Line 1578

**Description:** Function: Convert-EntraObjectIdToSid

**Called at lines (within Functions.ps1):** 1550

**Total calls:** 1

---

## Update-ConversionProgress

**File:** Functions.ps1

**Defined at:** Line 1594

**Description:** Function: Update-ConversionProgress

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Update-ProfileListRegistry

**File:** Functions.ps1

**Defined at:** Line 1631

**Description:** Function: Update-ProfileListRegistry

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Set-ProfileAcls

**File:** Functions.ps1

**Defined at:** Line 1744

**Description:** Function: Set-ProfileAcls

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Set-ProfileFolderAcls

**File:** Functions.ps1

**Defined at:** Line 2033

**Description:** Function: Set-ProfileFolderAcls

**Called at lines (within Functions.ps1):** 1815

**Total calls:** 1

---

## Set-ProfileHiveAcl

**File:** Functions.ps1

**Defined at:** Line 2150

**Description:** Function: Set-ProfileHiveAcl

**Called at lines (within Functions.ps1):** 1825

**Total calls:** 1

---

## Mount-RegistryHive

**File:** Functions.ps1

**Defined at:** Line 2231

**Description:** Function: Mount-RegistryHive

**Called at lines (within Functions.ps1):** 2423

**Total calls:** 1

---

## Dismount-RegistryHive

**File:** Functions.ps1

**Defined at:** Line 2270

**Description:** Function: Dismount-RegistryHive

**Called at lines (within Functions.ps1):** 2425

**Total calls:** 1

---

## Rewrite-HiveSID

**File:** Functions.ps1

**Defined at:** Line 2324

**Description:** Function: Rewrite-HiveSID

**Called at lines (within Functions.ps1):** 2013, 2333, 2850

**Total calls:** 3

---

## Update-RegistryStringValues

**File:** Functions.ps1

**Defined at:** Line 2561

**Description:** SAFER APPROACH: Use native PowerShell object iteration This avoids "reg export" corruption and handles special characters correctly

**Called at lines (within Functions.ps1):** 2652

**Total calls:** 1

---

## Report-RewriteSummary

**File:** Functions.ps1

**Defined at:** Line 2696

**Description:** Diagnostic: summarize rewrite status across critical keys

**Called at lines (within Functions.ps1):** 2757

**Total calls:** 1

---

## Test-AndFix-ProfileHive

**File:** Functions.ps1

**Defined at:** Line 2854

**Description:** Function: Test-AndFix-ProfileHive

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Remove-FolderRobust

**File:** Functions.ps1

**Defined at:** Line 2902

**Description:** Function: Remove-FolderRobust

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Get-FolderSize

**File:** Functions.ps1

**Defined at:** Line 2917

**Description:** Function: Get-FolderSize

**Called at lines (within Functions.ps1):** 829

**Total calls:** 1

---

## Get-RobocopyExclusions

**File:** Functions.ps1

**Defined at:** Line 2948

**Description:** Function: Get-RobocopyExclusions

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Enable-Privilege

**File:** Functions.ps1

**Defined at:** Line 3015

**Description:** Function: Enable-Privilege

**Called at lines (within Functions.ps1):** 1779, 1780, 1781, 2340, 2341, 2342

**Total calls:** 6

---

## New-CleanupItem

**File:** Functions.ps1

**Defined at:** Line 3034

**Description:** Function: New-CleanupItem

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Get-FriendlyAppName

**File:** Functions.ps1

**Defined at:** Line 3085

**Description:** Function: Get-FriendlyAppName

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Get-DomainFQDN

**File:** Functions.ps1

**Defined at:** Line 3150

**Description:** Function: Get-DomainFQDN

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Test-DomainReachability

**File:** Functions.ps1

**Defined at:** Line 3175

**Description:** Function: Test-DomainReachability

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Get-DomainCredential

**File:** Functions.ps1

**Defined at:** Line 3222

**Description:** Function: Get-DomainCredential

**Called at lines (within Functions.ps1):** 3372

**Total calls:** 1

---

## Get-DomainAdminCredential

**File:** Functions.ps1

**Defined at:** Line 3359

**Description:** Function: Get-DomainAdminCredential

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Test-DomainCredentials

**File:** Functions.ps1

**Defined at:** Line 3478

**Description:** Function: Test-DomainCredentials

**Called at lines (within Functions.ps1):** 3380

**Total calls:** 1

---

## Get-DomainJoinErrorDetails

**File:** Functions.ps1

**Defined at:** Line 3528

**Description:** Function: Get-DomainJoinErrorDetails

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Show-AzureADJoinDialog

**File:** Functions.ps1

**Defined at:** Line 3588

**Description:** Function: Show-AzureADJoinDialog

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Invoke-AzureADUnjoin

**File:** Functions.ps1

**Defined at:** Line 3731

**Description:** Function: Invoke-AzureADUnjoin

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Invoke-DomainUnjoin

**File:** Functions.ps1

**Defined at:** Line 3786

**Description:** Function: Invoke-DomainUnjoin

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Invoke-ForceUserLogoff

**File:** Functions.ps1

**Defined at:** Line 3847

**Description:** Function: Invoke-ForceUserLogoff

**Called at lines (within Functions.ps1):** 4809

**Total calls:** 1

---

## Test-WingetFunctionality

**File:** Functions.ps1

**Defined at:** Line 3944

**Description:** Function: Test-WingetFunctionality

**Called at lines (within Functions.ps1):** 4036

**Total calls:** 1

---

## Repair-WingetSources

**File:** Functions.ps1

**Defined at:** Line 3992

**Description:** Function: Repair-WingetSources

**Called at lines (within Functions.ps1):** 4039

**Total calls:** 1

---

## Install-WingetAppsFromExport

**File:** Functions.ps1

**Defined at:** Line 4027

**Description:** Function: Install-WingetAppsFromExport

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Add-AppxReregistrationActiveSetup

**File:** Functions.ps1

**Defined at:** Line 4322

**Description:** Function: Add-AppxReregistrationActiveSetup

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Write-Log

**File:** Functions.ps1

**Defined at:** Line 4381

**Description:** Function: Write-Log

**Called at lines (within Functions.ps1):** 4387, 4388, 4389, 4390, 4397, 4398, 4404, 4408, 4468, 4485, 4488, 4493, 4505, 4506, 4507, 4514

**Total calls:** 16

---

## Show-SevenZipRecoveryDialog

**File:** Functions.ps1

**Defined at:** Line 4551

**Description:** Function: Show-SevenZipRecoveryDialog

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Get-ZipUncompressedSize

**File:** Functions.ps1

**Defined at:** Line 4709

**Description:** Function: Get-ZipUncompressedSize

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Test-ZipIntegrity

**File:** Functions.ps1

**Defined at:** Line 4741

**Description:** Function: Test-ZipIntegrity

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---

## Invoke-ProactiveUserCheck

**File:** Functions.ps1

**Defined at:** Line 4778

**Description:** Function: Invoke-ProactiveUserCheck

**Called at lines:** *(Not called within Functions.ps1 - called from ProfileMigration.ps1)*

---


