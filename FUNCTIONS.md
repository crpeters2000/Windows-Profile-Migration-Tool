# ProfileMigration.ps1 - Function Reference

This document provides a comprehensive reference of all functions in the ProfileMigration.ps1 script.

**Total Functions:** 81

---

## Show-ModernDialog

**Defined at:** Line 33

**Description:** Function: Show-ModernDialog

**Called at lines:** 623, 1084, 1087, 1185, 1207, 1210, 1214, 1218, 1669, 3259, 3265, 3272, 3643, 5354, 5410, 5547, 5560, 5598, 5645, 5649, 6013, 6190, 6201, 6273, 6285, 6312, 6321, 6338, 6345, 6350, 6358, 6407, 6438, 6937, 6942, 6972, 6977, 7057, 7067, 7108, 7112, 7131, 7158, 7164, 7188, 7197, 7222, 7231, 7250, 7258, 7263, 7272, 7277, 7301, 7311, 7408, 7416, 7466, 7474, 7526, 7560, 7615, 7625, 7632, 7664, 7689, 7698, 7717, 7731, 7759, 7769, 7776, 7834, 7905, 7919, 7966, 7979, 7985, 8000, 8228, 8263, 8271, 8302, 8308, 8330, 8334, 8340, 8355, 8701, 8710, 8740, 8743, 8768, 8777, 8876, 8899, 8909, 8938, 8984, 9013, 9040, 9087, 9337, 10549, 11053, 11102, 11156, 11178, 11399, 11408, 11743, 11847, 11903, 12280, 13011, 13157, 13304, 13308, 13378, 13471, 13477, 13483, 13796, 13863, 13867, 13943, 13984, 13991, 14019, 14025, 14034, 14041, 14065, 14174, 14188, 14216, 14245, 14255, 14260

**Total calls:** 139

---

## Test-InternetConnectivity

**Defined at:** Line 182

**Description:** Function: Test-InternetConnectivity

**Called at lines:** *(Not called directly or only called dynamically)*

---

## Show-InputDialog

**Defined at:** Line 209

**Description:** Function: Show-InputDialog

**Called at lines:** 7054, 11356

**Total calls:** 2

---

## Get-WindowsTheme

**Defined at:** Line 434

**Description:** Function to detect Windows theme

**Called at lines:** 14428

**Total calls:** 1

---

## Apply-Theme

**Defined at:** Line 451

**Description:** Function to apply theme to all controls

**Called at lines:** 583, 14429

**Total calls:** 2

---

## Toggle-Theme

**Defined at:** Line 581

**Description:** Function to toggle theme

**Called at lines:** 13351

**Total calls:** 1

---

## Convert-EntraObjectIdToSid

**Defined at:** Line 593

**Description:** Function to detect profile type (Local, Domain, or AzureAD) Function to convert Entra/AzureAD ObjectId to Windows SID

**Called at lines:** 643

**Total calls:** 1

---

## Get-AzureADUserSID

**Defined at:** Line 609

**Description:** Function to get AzureAD user SID using Microsoft Graph

**Called at lines:** 3709, 11378

**Total calls:** 2

---

## Log-Message

**Defined at:** Line 730

**Description:** Function: Log-Message

**Called at lines:** 793, 794, 795, 796, 866, 890, 894, 1320, 1355, 1361, 1420, 1454, 1463, 1530, 1537, 1541, 1568, 1578, 1582, 1589, 3033, 3038, 3041, 3046, 3064, 3084, 3089, 3110, 3122, 3130, 3520, 3635, 3642, 3674, 3678, 3693, 3700, 3701, 3715, 3717, 3720, 3728, 3779, 3784, 3791, 3796, 3802, 3809, 3816, 3821, 3837, 3840, 3851, 3856, 3863, 3867, 3870, 3877, 3881, 3884, 3904, 3912, 3917, 3948, 3952, 3958, 3962, 3965, 3986, 3992, 4020, 4031, 4034, 4039, 4040, 4046, 4052, 4053, 4055, 4062, 4063, 4069, 4072, 4094, 4100, 4108, 4118, 4123, 4129, 4132, 4136, 4151, 4158, 4165, 4172, 4183, 4191, 4198, 4205, 4212, 4216, 4220, 4235, 4238, 4243, 4246, 4253, 4258, 4273, 4274, 4275, 4276, 4277, 4305, 4308, 4311, 4316, 4322, 4325, 4330, 4337, 4338, 4350, 4359, 4366, 4370, 4379, 4410, 4415, 4418, 4425, 4428, 4433, 4434, 4437, 4441, 4442, 4444, 4449, 4458, 4464, 4471, 4473, 4481, 4491, 4496, 4513, 4530, 4535, 4576, 4581, 4603, 4616, 4624, 4632, 4644, 4677, 4680, 4691, 4704, 4707, 4719, 4722, 4729, 4733, 4742, 4748, 4753, 4760, 4763, 4764, 4775, 4782, 4785, 4790, 5769, 5993, 6170, 6474, 8301, 8303, 8307, 8311, 8316, 8323, 8328, 8333, 8337, 8350, 8354, 8375, 8393, 8406, 8557, 8561, 8567, 8572, 8580, 8584, 8588, 8592, 8598, 8697, 8721, 8739, 8764, 8865, 8874, 8880, 8887, 8893, 8901, 8906, 8914, 8922, 8924, 8941, 8946, 8949, 8955, 8963, 8968, 8972, 8983, 8987, 8994, 9001, 9007, 9011, 9015, 9020, 9023, 9028, 9029, 9038, 9054, 9061, 9066, 9079, 9084, 9256, 9271, 9284, 9296, 9311, 9315, 9320, 9325, 9331, 9346, 9347, 10237, 10243, 10245, 10257, 10258, 10291, 10292, 10310, 10320, 10321, 10324, 10358, 10362, 10384, 10389, 10441, 10445, 10450, 10510, 10514, 10521, 10522, 10523, 10527, 10529, 10530, 10531, 10535, 10548, 10576, 10582, 10601, 10606, 10610, 10613, 10618, 10625, 10631, 10642, 10649, 10655, 10662, 10668, 10684, 10695, 10757, 10772, 10780, 10802, 10806, 10850, 10859, 10863, 10866, 10869, 10873, 10883, 10898, 10901, 10905, 10910, 10915, 10927, 10930, 10942, 10957, 11007, 11010, 11014, 11019, 11060, 11063, 11067, 11070, 11094, 11099, 11128, 11130, 11142, 11148, 11153, 11154, 11155, 11170, 11176, 11182, 11226, 11234, 11244, 11259, 11260, 11266, 11273, 11279, 11294, 11295, 11299, 11307, 11310, 11320, 11322, 11327, 11335, 11337, 11340, 11345, 11350, 11355, 11375, 11384, 11388, 11396, 11397, 11407, 11412, 11417, 11609, 11612, 11632, 11718, 11722, 11742, 11746, 11796, 11800, 11811, 11828, 11833, 11856, 11861, 11902, 11909, 11914, 11920, 11923, 11935, 11943, 11953, 11957, 11974, 11998, 12002, 12008, 12012, 12015, 12019, 12025, 12035, 12044, 12047, 12053, 12126, 12128, 12133, 12143, 12145, 12180, 12186, 12189, 12196, 12202, 12204, 12207, 12211, 12227, 12232, 12248, 12256, 12262, 12265, 12270, 12275, 12276, 12287, 12290, 12294, 12300, 12310, 12312, 12315, 12331, 12336, 12337, 12342, 12346, 12354, 12355, 12363, 12370, 12378, 12416, 12423, 12426, 12428, 12440, 12443, 12448, 12459, 12465, 12468, 12472, 12476, 12485, 12490, 12496, 12499, 12507, 12508, 12520, 12523, 12527, 12535, 12537, 12542, 12547, 12551, 12555, 12558, 12560, 12563, 12565, 12570, 12578, 12580, 12583, 12594, 12601, 12606, 12607, 12608, 12611, 12627, 12630, 12631, 12642, 12648, 12655, 12659, 12671, 12672, 12673, 12674, 12675, 12676, 12677, 12682, 12683, 12684, 12687, 12694, 12703, 12714, 12715, 12716, 12726, 12737, 12748, 12751, 12759, 12762, 12766, 12778, 12783, 12791, 12795, 12800, 12807, 12814, 12818, 12820, 12823, 12826, 12829, 12833, 12837, 12840, 12843, 12844, 12845, 12846, 12847, 12848, 12855, 12858, 12879, 12893, 12942, 12945, 12949, 12954, 12961, 13015, 13016, 13017, 13021, 13029, 13032, 13035, 13036, 13041, 13044, 13052, 13062, 13064, 13076, 13081, 13084, 13089, 13096, 13098, 13103, 13107, 13108, 13112, 13118, 13149, 13154, 14068, 14335, 14358, 14367, 14373, 14381, 14387, 14423, 14424, 14425

**Total calls:** 578

---

## Log-Debug

**Defined at:** Line 793

**Description:** Function: Log-Debug

**Called at lines:** 1809, 2475, 3116, 4881, 4902, 4909, 4933, 4944, 4948, 5382, 9529, 9672, 9788, 10021, 10036

**Total calls:** 15

---

## Log-Info

**Defined at:** Line 794

**Description:** Function: Log-Info

**Called at lines:** 615, 627, 634, 638, 644, 1688, 1699, 1704, 1742, 1754, 1759, 1764, 1824, 2471, 2780, 3141, 3171, 3175, 3182, 3202, 3204, 4982, 4988, 4995, 5018, 5026, 5034, 5039, 5044, 5049, 5069, 5105, 5297, 5314, 5315, 5335, 5346, 5358, 5362, 5377, 5390, 5394, 5400, 5402, 5413, 5415, 5418, 5422, 5448, 5449, 5452, 5455, 5464, 5471, 5480, 5490, 5521, 5522, 5523, 5533, 5536, 5553, 5565, 5571, 5575, 5579, 5584, 5588, 5592, 5623, 5629, 5632, 5633, 5657, 5666, 5687, 5691, 5697, 5701, 5711, 5722, 5745, 5750, 5755, 5763, 5773, 5790, 5792, 5795, 5800, 5835, 5836, 5837, 5838, 5841, 5843, 5850, 5853, 5856, 5860, 5862, 5867, 5875, 5877, 5887, 5888, 5897, 5918, 5922, 5927, 5931, 5938, 5949, 5972, 5979, 5987, 5996, 6001, 6008, 6017, 6037, 6039, 6041, 6044, 6078, 6079, 6080, 6083, 6091, 6096, 6099, 6101, 6105, 6113, 6125, 6134, 6144, 6154, 6175, 6179, 6196, 6205, 6225, 6227, 6229, 6232, 6254, 6255, 6256, 6276, 6280, 6290, 6326, 6330, 6337, 6339, 6342, 6344, 6353, 6357, 6366, 6378, 6382, 6386, 6389, 6397, 6424, 6428, 6452, 6462, 6479, 6494, 6496, 6498, 6501, 7076, 7091, 7111, 7116, 7138, 7149, 7191, 7201, 7204, 7234, 7238, 7242, 7249, 7251, 7254, 7256, 7266, 7270, 7282, 7321, 7322, 7323, 7324, 7325, 7328, 7329, 7330, 7331, 7332, 7333, 7334, 7337, 7341, 7344, 7348, 7352, 7357, 7360, 7362, 7415, 7435, 7486, 7487, 7492, 7512, 7513, 7534, 7535, 7536, 7539, 7542, 7545, 7559, 7561, 7564, 7570, 7573, 7579, 7588, 7591, 7595, 7597, 7631, 7638, 7642, 7646, 7649, 7663, 7665, 7668, 7674, 7677, 7683, 7701, 7705, 7709, 7716, 7718, 7723, 7734, 7738, 7748, 7775, 7782, 7786, 7788, 7802, 7837, 7843, 7899, 7938, 7939, 7941, 7969, 7978, 8131, 8139, 8153, 8158, 8215, 8220, 8245, 8246, 8256, 8641, 8662, 10343, 11046, 11885, 12079, 12990, 13510, 13522, 13533, 13539, 13545, 13563, 13573, 13579, 13585, 13614, 13629, 13639, 13645, 13651, 13727, 13899

**Total calls:** 302

---

## Log-Warning

**Defined at:** Line 795

**Description:** Function: Log-Warning

**Called at lines:** 947, 1194, 3194, 3210, 3217, 3221, 4811, 4817, 4850, 4952, 5122, 5321, 5353, 5386, 5405, 5595, 5660, 5682, 5704, 5805, 5913, 5934, 6049, 6146, 6237, 6262, 6297, 6320, 6349, 6454, 6506, 7095, 7171, 7177, 7230, 7262, 7410, 7438, 7697, 7730, 8126, 8164, 8223, 8259, 9057, 10352, 11050, 11303, 11892, 11897, 11947, 12096, 12101, 12994, 13047

**Total calls:** 55

---

## Log-Error

**Defined at:** Line 796

**Description:** Function: Log-Error

**Called at lines:** 664, 1712, 1720, 1772, 1780, 2483, 2736, 3226, 4855, 5086, 5431, 5498, 5786, 6033, 6221, 6361, 6372, 6490, 6587, 7163, 7275, 7288, 7525, 7611, 7621, 7743, 7755, 7765, 7984, 8145, 8270, 8644, 8666, 9536, 11846, 13377, 13990, 14040, 14213

**Total calls:** 39

---

## Refresh-LogDisplay

**Defined at:** Line 798

**Description:** Function: Refresh-LogDisplay

**Called at lines:** 13897

**Total calls:** 1

---

## Test-WingetFunctionality

**Defined at:** Line 818

**Description:** Function: Test-WingetFunctionality

**Called at lines:** 1130, 1192, 9055

**Total calls:** 3

---

## Repair-WingetSources

**Defined at:** Line 865

**Description:** Function: Repair-WingetSources

**Called at lines:** 1133, 1195, 9058

**Total calls:** 3

---

## Enable-Privilege

**Defined at:** Line 937

**Description:** Function: Enable-Privilege

**Called at lines:** 4008, 4009, 4010, 4280, 4281, 4282

**Total calls:** 6

---

## Show-SevenZipRecoveryDialog

**Defined at:** Line 958

**Description:** Function to show 7-Zip recovery dialog

**Called at lines:** 1171

**Total calls:** 1

---

## Get-FriendlyAppName

**Defined at:** Line 1231

**Description:** Function: Get-FriendlyAppName

**Called at lines:** 2377

**Total calls:** 1

---

## Get-FolderSize

**Defined at:** Line 1295

**Description:** Function: Get-FolderSize

**Called at lines:** 2824

**Total calls:** 1

---

## Test-PathWithRetry

**Defined at:** Line 1325

**Description:** Function: Test-PathWithRetry

**Called at lines:** 1451, 1460

**Total calls:** 2

---

## Convert-SIDToAccountName

**Defined at:** Line 1372

**Description:** Function: Convert-SIDToAccountName

**Called at lines:** 2870, 3691

**Total calls:** 2

---

## Test-ValidProfilePath

**Defined at:** Line 1425

**Description:** Function: Test-ValidProfilePath

**Called at lines:** 5636, 5641, 5891, 6128

**Total calls:** 4

---

## Update-ConversionProgress

**Defined at:** Line 1471

**Description:** Function: Update-ConversionProgress

**Called at lines:** 5667, 5700, 5746, 5764, 5774, 5898, 5930, 5973, 5988, 5997, 6115, 6136, 6155, 6165, 6173

**Total calls:** 15

---

## Mount-RegistryHive

**Defined at:** Line 1508

**Description:** Function: Mount-RegistryHive

**Called at lines:** 4363

**Total calls:** 1

---

## Dismount-RegistryHive

**Defined at:** Line 1546

**Description:** Function: Dismount-RegistryHive

**Called at lines:** 4365

**Total calls:** 1

---

## New-CleanupItem

**Defined at:** Line 1599

**Description:** Function: New-CleanupItem

**Called at lines:** 9462, 9499, 9551, 9646, 9686, 9722

**Total calls:** 6

---

## Confirm-DomainUnjoin

**Defined at:** Line 1649

**Description:** Function: Confirm-DomainUnjoin

**Called at lines:** *(Not called directly or only called dynamically)*

---

## Invoke-AzureADUnjoin

**Defined at:** Line 1674

**Description:** Function: Invoke-AzureADUnjoin

**Called at lines:** 5550, 6193, 7569, 8945

**Total calls:** 4

---

## Invoke-DomainUnjoin

**Defined at:** Line 1728

**Description:** Function: Invoke-DomainUnjoin

**Called at lines:** 6005, 6277, 7192, 7673

**Total calls:** 4

---

## Generate-MigrationReport

**Defined at:** Line 1789

**Description:** Function: Generate-MigrationReport

**Called at lines:** 11044, 11092, 12988, 13147

**Total calls:** 4

---

## New-ConversionReport

**Defined at:** Line 2488

**Description:** Function: New-ConversionReport

**Called at lines:** 8211, 8250

**Total calls:** 2

---

## Start-OperationLog

**Defined at:** Line 2742

**Description:** Function: Start-OperationLog

**Called at lines:** 10287, 11122

**Total calls:** 2

---

## Stop-OperationLog

**Defined at:** Line 2790

**Description:** Function: Stop-OperationLog

**Called at lines:** 11105, 13160

**Total calls:** 2

---

## Get-ProfileInfo

**Defined at:** Line 2802

**Description:** Enhanced profile detection with size estimates

**Called at lines:** 2860

**Total calls:** 1

---

## Get-ProfileDisplayEntries

**Defined at:** Line 2851

**Description:** Build display names as DOMAIN\\username or COMPUTERNAME\\username for the dropdown

**Called at lines:** 6581, 13413, 13770

**Total calls:** 3

---

## Get-RobocopyExclusions

**Defined at:** Line 2918

**Description:** Function: Get-RobocopyExclusions

**Called at lines:** 2986, 10367, 12105

**Total calls:** 3

---

## Test-ProfilePathWriteable

**Defined at:** Line 3051

**Description:** Function: Test-ProfilePathWriteable

**Called at lines:** 11857

**Total calls:** 1

---

## Test-ProfileMounted

**Defined at:** Line 3100

**Description:** Function: Test-ProfileMounted

**Called at lines:** 3258, 3264, 10547, 11901, 13470, 13476

**Total calls:** 6

---

## Invoke-ForceUserLogoff

**Defined at:** Line 3135

**Description:** Function: Invoke-ForceUserLogoff

**Called at lines:** 3262, 13474

**Total calls:** 2

---

## Invoke-ProactiveUserCheck

**Defined at:** Line 3231

**Description:** Function: Invoke-ProactiveUserCheck

**Called at lines:** 6829, 6953

**Total calls:** 2

---

## Get-LocalProfiles

**Defined at:** Line 3285

**Description:** Function: Get-LocalProfiles

**Called at lines:** 2856, 10297, 13766

**Total calls:** 3

---

## Remove-FolderRobust

**Defined at:** Line 3293

**Description:** Function: Remove-FolderRobust

**Called at lines:** 5801, 6045, 6233, 6502, 10916, 12011, 12427, 12486, 12819, 13022, 13033, 13097

**Total calls:** 12

---

## Show-LogViewer

**Defined at:** Line 3307

**Description:** Function: Show-LogViewer

**Called at lines:** 13850, 13860

**Total calls:** 2

---

## Test-IsAzureADSID

**Defined at:** Line 3506

**Description:** Check if a SID is an AzureAD/Entra ID account

**Called at lines:** 3690, 3716, 3828, 3943, 4822, 6087, 6393, 10564, 11255, 11331, 11803, 13609

**Total calls:** 12

---

## Test-IsAzureADJoined

**Defined at:** Line 3513

**Description:** Check if system is AzureAD/Entra ID joined

**Called at lines:** 7541, 7542

**Total calls:** 2

---

## Show-AzureADJoinDialog

**Defined at:** Line 3526

**Description:** Show AzureAD join guidance dialog

**Called at lines:** 11312, 12302, 13514, 13556, 13622

**Total calls:** 5

---

## Get-LocalUserSID

**Defined at:** Line 3668

**Description:** Function: Get-LocalUserSID

**Called at lines:** 3985, 4808, 4973, 5454, 5583, 5591, 5842, 5876, 6084, 6112, 6385, 6391, 7374, 9665, 11298

**Total calls:** 15

---

## Get-DomainFQDN

**Defined at:** Line 3749

**Description:** Resolve NetBIOS domain name to FQDN

**Called at lines:** 7043, 11265, 11272, 13675, 13708

**Total calls:** 5

---

## Set-ProfileFolderAcls

**Defined at:** Line 3773

**Description:** Function: Set-ProfileFolderAcls

**Called at lines:** 4044

**Total calls:** 1

---

## Set-ProfileHiveAcl

**Defined at:** Line 3890

**Description:** Function: Set-ProfileHiveAcl

**Called at lines:** 4054

**Total calls:** 1

---

## Set-ProfileAcls

**Defined at:** Line 3973

**Description:** Function: Set-ProfileAcls

**Called at lines:** 5488, 5770, 5994, 6171, 6475, 12528, 12536

**Total calls:** 7

---

## Rewrite-HiveSID

**Defined at:** Line 4264

**Description:** Function: Rewrite-HiveSID

**Called at lines:** 4242, 4273, 4790

**Total calls:** 3

---

## Update-RegistryStringValues

**Defined at:** Line 4501

**Description:** SAFER APPROACH: Use native PowerShell object iteration This avoids "reg export" corruption and handles special characters correctly

**Called at lines:** 4592

**Total calls:** 1

---

## Report-RewriteSummary

**Defined at:** Line 4636

**Description:** Diagnostic: summarize rewrite status across critical keys

**Called at lines:** 4697

**Total calls:** 1

---

## Get-ProfileType

**Defined at:** Line 4798

**Description:** Get the type of a user profile (Local, Domain, or AzureAD)

**Called at lines:** 6838, 6983

**Total calls:** 2

---

## Test-UserLoggedOut

**Defined at:** Line 4861

**Description:** Check if a user is currently logged in

**Called at lines:** 4983

**Total calls:** 1

---

## Test-ProfileConversionPreconditions

**Defined at:** Line 4959

**Description:** Test all preconditions for profile conversion

**Called at lines:** 7029

**Total calls:** 1

---

## Add-AppxReregistrationActiveSetup

**Defined at:** Line 5098

**Description:** Add Active Setup registry key to re-register AppX packages on first login This runs BEFORE the desktop loads, fixing file locking issues with Search/Start Menu

**Called at lines:** 8207, 12862

**Total calls:** 2

---

## Write-Log

**Defined at:** Line 5157

**Description:** Function: Write-Log

**Called at lines:** 5163, 5164, 5165, 5166, 5173, 5174, 5180, 5184, 5244, 5261, 5264, 5269, 5281, 5282, 5283, 5290

**Total calls:** 16

---

## Update-ProfileListRegistry

**Defined at:** Line 5327

**Description:** Update ProfileList registry entry for profile conversion

**Called at lines:** 5473, 5756, 5980, 6161, 6469

**Total calls:** 5

---

## Repair-UserProfile

**Defined at:** Line 5441

**Description:** Repair an existing user profile (permissions and registry)

**Called at lines:** 7803, 7838, 7900

**Total calls:** 3

---

## Convert-LocalToDomain

**Defined at:** Line 5507

**Description:** Convert a local user profile to a domain user profile

**Called at lines:** 7813, 7888

**Total calls:** 2

---

## Convert-DomainToLocal

**Defined at:** Line 5816

**Description:** Convert a domain user profile to a local user profile

**Called at lines:** 7821, 7853

**Total calls:** 2

---

## Convert-AzureADToLocal

**Defined at:** Line 6059

**Description:** Function: Convert-AzureADToLocal

**Called at lines:** 7860

**Total calls:** 1

---

## Convert-LocalToAzureAD

**Defined at:** Line 6244

**Description:** Function: Convert-LocalToAzureAD

**Called at lines:** 7869, 7877

**Total calls:** 2

---

## Show-ProfileConversionDialog

**Defined at:** Line 6514

**Description:** Show Profile Conversion Dialog

**Called at lines:** 13374

**Total calls:** 1

---

## Handle-Restart

**Defined at:** Line 8292

**Description:** Function: Handle-Restart

**Called at lines:** 8883, 8910, 9024

**Total calls:** 3

---

## Test-DomainReachability

**Defined at:** Line 8372

**Description:** Function: Test-DomainReachability

**Called at lines:** 8978, 11737, 14251

**Total calls:** 3

---

## Get-DomainCredential

**Defined at:** Line 8418

**Description:** Function: Get-DomainCredential

**Called at lines:** 7064, 8687, 14191

**Total calls:** 3

---

## Test-DomainCredentials

**Defined at:** Line 8554

**Description:** Function: Test-DomainCredentials

**Called at lines:** 8695

**Total calls:** 1

---

## Get-ZipUncompressedSize

**Defined at:** Line 8604

**Description:** Helper to get uncompressed size of a ZIP file using 7-Zip listing

**Called at lines:** 11870, 12067

**Total calls:** 2

---

## Test-ZipIntegrity

**Defined at:** Line 8635

**Description:** Function: Test-ZipIntegrity

**Called at lines:** 12058

**Total calls:** 1

---

## Get-DomainAdminCredential

**Defined at:** Line 8674

**Description:** Function: Get-DomainAdminCredential

**Called at lines:** 8996, 11761

**Total calls:** 2

---

## Get-DomainJoinErrorDetails

**Defined at:** Line 8791

**Description:** Function: Get-DomainJoinErrorDetails

**Called at lines:** 8886, 8913, 9027

**Total calls:** 3

---

## Join-Domain-Enhanced

**Defined at:** Line 8853

**Description:** Function: Join-Domain-Enhanced

**Called at lines:** 7144, 11843, 14206

**Total calls:** 3

---

## Install-WingetAppsFromExport

**Defined at:** Line 9046

**Description:** Function: Install-WingetAppsFromExport

**Called at lines:** 12323, 12327

**Total calls:** 2

---

## Show-ProfileCleanupWizard

**Defined at:** Line 9343

**Description:** Function: Show-ProfileCleanupWizard

**Called at lines:** 10307

**Total calls:** 1

---

## Export-UserProfile

**Defined at:** Line 10276

**Description:** Function: Export-UserProfile

**Called at lines:** 7423, 13987

**Total calls:** 2

---

## Import-UserProfile

**Defined at:** Line 11114

**Description:** Function: Import-UserProfile

**Called at lines:** 14037

**Total calls:** 1

---

## Test-AndFix-ProfileHive

**Defined at:** Line 12618

**Description:** Function: Test-AndFix-ProfileHive

**Called at lines:** 12669

**Total calls:** 1

---

