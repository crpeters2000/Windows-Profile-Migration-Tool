# === LOAD REQUIRED ASSEMBLIES ===
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# === CONFIGURATION (MUST BE DEFINED BEFORE MODULE) ===
$cpuCores = [Environment]::ProcessorCount
$Config = @{
    Version                     = 'v2.13.0' 
    DomainReachabilityTimeout   = 3000
    DomainJoinCountdown         = 10
    HiveUnloadMaxAttempts       = 3
    HiveUnloadWaitMs            = 500
    MountPointMaxAttempts       = 5
    ProgressUpdateIntervalMs    = 1000
    ExportProgressCheckMs       = 500
    RobocopyThreads             = [Math]::Min(32, [Math]::Max(8, $cpuCores))
    RobocopyRetryCount          = 1
    RobocopyRetryWaitSec        = 1
    SevenZipThreads             = $cpuCores
    ProfileValidationTimeoutSec = 10
    SizeEstimationDepth         = 3
    HashVerificationEnabled     = $true
    LogLevel                    = 'INFO'
    GenerateHTMLReports         = $true
    AutoOpenReports             = $true
    LogPath                     = "C:\Logs\ProfileMigration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
}

# === LOG LEVELS ===
$global:LogLevels = @{
    DEBUG = 0
    INFO  = 1
    WARN  = 2
    ERROR = 3
}

# === THEME DEFINITIONS ===
$global:Themes = @{
    'Dark'  = @{
        FormBackColor                 = [System.Drawing.Color]::FromArgb(32, 32, 32)
        HeaderBackColor               = [System.Drawing.Color]::FromArgb(45, 45, 48)
        PanelBackColor                = [System.Drawing.Color]::FromArgb(45, 45, 48)
        BorderColor                   = [System.Drawing.Color]::FromArgb(60, 60, 60)
        LabelTextColor                = [System.Drawing.Color]::FromArgb(240, 240, 240)
        HeaderTextColor               = [System.Drawing.Color]::FromArgb(0, 120, 215)
        SubHeaderTextColor            = [System.Drawing.Color]::FromArgb(180, 180, 180)
        TextBoxBackColor              = [System.Drawing.Color]::FromArgb(30, 30, 30)
        TextBoxForeColor              = [System.Drawing.Color]::White
        LogBoxBackColor               = [System.Drawing.Color]::Black
        LogBoxForeColor               = [System.Drawing.Color]::FromArgb(200, 200, 200)
        ProgressBarForeColor          = [System.Drawing.Color]::FromArgb(0, 120, 215)
        ButtonPrimaryBackColor        = [System.Drawing.Color]::FromArgb(0, 120, 215)
        ButtonPrimaryForeColor        = [System.Drawing.Color]::White
        ButtonPrimaryHoverBackColor   = [System.Drawing.Color]::FromArgb(0, 100, 195)
        ButtonSecondaryBackColor      = [System.Drawing.Color]::FromArgb(60, 60, 60)
        ButtonSecondaryForeColor      = [System.Drawing.Color]::FromArgb(240, 240, 240)
        ButtonSecondaryHoverBackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
        ButtonSuccessBackColor        = [System.Drawing.Color]::FromArgb(16, 124, 16)
        ButtonSuccessForeColor        = [System.Drawing.Color]::White
        ButtonDangerBackColor         = [System.Drawing.Color]::FromArgb(232, 17, 35)
        ButtonDangerForeColor         = [System.Drawing.Color]::White
    }
    'Light' = @{
        FormBackColor                 = [System.Drawing.Color]::White
        HeaderBackColor               = [System.Drawing.Color]::FromArgb(240, 240, 240)
        PanelBackColor                = [System.Drawing.Color]::FromArgb(255, 255, 255)
        BorderColor                   = [System.Drawing.Color]::FromArgb(200, 200, 200)
        LabelTextColor                = [System.Drawing.Color]::Black
        HeaderTextColor               = [System.Drawing.Color]::FromArgb(0, 120, 215)
        SubHeaderTextColor            = [System.Drawing.Color]::FromArgb(100, 100, 100)
        TextBoxBackColor              = [System.Drawing.Color]::White
        TextBoxForeColor              = [System.Drawing.Color]::Black
        LogBoxBackColor               = [System.Drawing.Color]::FromArgb(250, 250, 250)
        LogBoxForeColor               = [System.Drawing.Color]::Black
        ProgressBarForeColor          = [System.Drawing.Color]::FromArgb(0, 120, 215)
        ButtonPrimaryBackColor        = [System.Drawing.Color]::FromArgb(0, 120, 215)
        ButtonPrimaryForeColor        = [System.Drawing.Color]::White
        ButtonPrimaryHoverBackColor   = [System.Drawing.Color]::FromArgb(0, 100, 195)
        ButtonSecondaryBackColor      = [System.Drawing.Color]::FromArgb(230, 230, 230)
        ButtonSecondaryForeColor      = [System.Drawing.Color]::Black
        ButtonSecondaryHoverBackColor = [System.Drawing.Color]::FromArgb(210, 210, 210)
        ButtonSuccessBackColor        = [System.Drawing.Color]::FromArgb(16, 124, 16)
        ButtonSuccessForeColor        = [System.Drawing.Color]::White
        ButtonDangerBackColor         = [System.Drawing.Color]::FromArgb(232, 17, 35)
        ButtonDangerForeColor         = [System.Drawing.Color]::White
    }
}
$global:CurrentTheme = "Dark"  # Default to Dark theme

# === LOAD HELPER FUNCTIONS MODULE (AFTER CONFIG) ===
$FunctionsPath = Join-Path $PSScriptRoot "Functions.ps1"
if (Test-Path $FunctionsPath) {
    . $FunctionsPath
    Write-Host "Loaded helper functions module from: $FunctionsPath" -ForegroundColor Green
}
else {
    Write-Error "Required module not found: $FunctionsPath"
    Write-Error "Please ensure Functions.ps1 exists in the same directory as this script."
    Read-Host "Press Enter to exit"
    exit 1
}

# === MODERN DIALOG FUNCTION ===
function Confirm-DomainUnjoin {
    <#
    .SYNOPSIS
    Shows domain unjoin confirmation dialog
    
    .DESCRIPTION
    Displays standard domain unjoin warning and confirmation
    
    .OUTPUTS
    Returns $true if user confirms, $false otherwise
    #>
    
    $message = "You are about to unjoin this device from the domain.`r`n`r`n"
    $message += "This will:`r`n"
    $message += "- Remove device from domain management`r`n"
    $message += "- Disable domain group policies`r`n"
    $message += "- Remove domain authentication`r`n`r`n"
    $message += "This action requires a reboot to complete.`r`n`r`n"
    $message += "Continue with unjoin?"
    
    $response = Show-ModernDialog -Message $message -Title "Confirm Domain Unjoin" -Type Warning -Buttons YesNo
    
    return ($response -eq 'Yes')
}

function Generate-MigrationReport {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Export', 'Import')]
        [string]$OperationType,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$ReportData,
        
        [Parameter(Mandatory = $false)]
        [bool]$Success = $true,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "",
        
        [Parameter(Mandatory = $false)]
        [string]$RollbackStatus = ""
    )
    
    if (-not $Config.GenerateHTMLReports) {
        Log-Debug "Report generation disabled in configuration"
        return $null
    }
    
    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $reportFileName = "Migration_Report_${OperationType}_${timestamp}.html"
        
        # Save report to the Logs directory
        $logsDir = Join-Path $PSScriptRoot "Logs"
        if (-not (Test-Path $logsDir)) {
            New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
        }
        $reportPath = Join-Path $logsDir $reportFileName
        
        Log-Info "Generating $OperationType migration report..."
        
        # Build HTML report
        $currentDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Profile Migration Report</title>
    <style>
        :root {
            --primary-color: #0078d4;
            --success-color: #107c10;
            --error-color: #d13438;
            --bg-color: #f3f2f1;
            --card-bg: #ffffff;
            --text-primary: #323130;
            --text-secondary: #605e5c;
            --border-radius: 8px;
            --shadow: 0 1.6px 3.6px 0 rgba(0,0,0,0.13), 0 0.3px 0.9px 0 rgba(0,0,0,0.11);
        }
        body {
            font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif;
            background-color: var(--bg-color);
            color: var(--text-primary);
            margin: 0;
            padding: 40px 20px;
            line-height: 1.5;
        }
        .container {
            max-width: 900px;
            margin: 0 auto;
        }
        .header {
            background-color: var(--success-color);
            color: white;
            padding: 30px;
            border-radius: var(--border-radius);
            box-shadow: var(--shadow);
            margin-bottom: 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .header.error {
            background-color: var(--error-color);
        }
        .header-content {
            color: white;
        }
        .header-content h1 { margin: 0; font-size: 24px; font-weight: 600; }
        .header-content p { margin: 5px 0 0; opacity: 0.9; }
        .status-badge {
            background-color: rgba(255,255,255,0.2);
            padding: 8px 16px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 24px;
            margin-bottom: 24px;
        }
        .card {
            background-color: var(--card-bg);
            border-radius: var(--border-radius);
            padding: 24px;
            box-shadow: var(--shadow);
        }
        .section {
            background-color: var(--card-bg);
            border-radius: var(--border-radius);
            padding: 24px;
            box-shadow: var(--shadow);
            margin-bottom: 24px;
        }
        .section-title {
            font-size: 18px;
            font-weight: 600;
            color: var(--text-primary);
            margin-bottom: 16px;
            padding-bottom: 8px;
            border-bottom: 2px solid var(--primary-color);
        }
        .info-grid {
            display: grid;
            grid-template-columns: auto 1fr;
            gap: 12px 24px;
            align-items: start;
        }
        .info-label {
            font-weight: 600;
            color: var(--text-secondary);
            font-size: 13px;
        }
        .info-value {
            color: var(--text-primary);
            font-size: 13px;
            word-break: break-word;
        }
        .card-title {
            font-size: 16px;
            font-weight: 600;
            color: var(--text-secondary);
            margin-bottom: 16px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            border-bottom: 1px solid #e1dfdd;
            padding-bottom: 8px;
        }
        .stat-group { margin-bottom: 16px; }
        .stat-label { font-size: 12px; color: var(--text-secondary); font-weight: 600; }
        .stat-value { font-size: 16px; font-weight: 400; color: var(--text-primary); word-break: break-all; }
        .stat-value.highlight { font-weight: 600; color: var(--primary-color); }
        .success { color: var(--success-color); font-weight: 600; }
        .warning { color: var(--warning-color); font-weight: 600; }
        .error { color: var(--error-color); font-weight: 600; }
        table { width: 100%; border-collapse: separate; border-spacing: 0; margin: 16px 0; overflow: hidden; border-radius: var(--border-radius); }
        table th { background-color: var(--primary-color); color: white; padding: 12px; text-align: left; font-size: 13px; font-weight: 600; }
        table th:first-child { border-top-left-radius: var(--border-radius); }
        table th:last-child { border-top-right-radius: var(--border-radius); }
        table td { padding: 10px 12px; border-bottom: 1px solid #e1dfdd; font-size: 13px; background-color: white; }
        table tr:hover td { background-color: #f9f9f9; }
        table tr:last-child td:first-child { border-bottom-left-radius: var(--border-radius); }
        table tr:last-child td:last-child { border-bottom-right-radius: var(--border-radius); }
        .footer {
            text-align: center;
            color: var(--text-secondary);
            font-size: 12px;
            margin-top: 40px;
        }
        code {
            background-color: #f3f2f1;
            padding: 2px 4px;
            border-radius: 4px;
            font-family: Consolas, monospace;
            font-size: 0.9em;
        }
    </style>
</head>
<body>
    <div class="container">
"@
        
        # Determine header class and status text based on success/failure
        $headerClass = if ($Success) { "header" } else { "header error" }
        $statusText = if ($Success) { "Success" } else { "Failed" }
        $statusIcon = if ($Success) { [char]0x2713 } else { [char]0x2717 }
        
        # Add header with conditional styling
        $html += @"
        <div class="$headerClass">
            <div class="header-content">
                <h1>Profile Migration Report</h1>
                <p>$OperationType operation summary</p>
            </div>
            <div class="status-badge">
                <span>$statusIcon</span>
                <span>$statusText</span>
            </div>
        </div>
"@
        
        # Add error summary section if operation failed
        if (-not $Success) {
            $html += @"
        <div class="section">
            <div class="section-title">Error Summary</div>
            <div class="info-grid">
                <div class="info-label">Status:</div>
                <div class="info-value error">$statusIcon Operation Failed</div>
                <div class="info-label">Target User:</div>
                <div class="info-value">$($ReportData.Username)</div>
                <div class="info-label">Error Message:</div>
                <div class="info-value error">$ErrorMessage</div>
"@
            if ($RollbackStatus) {
                $html += @"
                <div class="info-label">Rollback Status:</div>
                <div class="info-value">$RollbackStatus</div>
"@
            }
            $html += @"
                <div class="info-label">Elapsed Time:</div>
                <div class="info-value">$($ReportData.ElapsedMinutes) minutes ($($ReportData.ElapsedSeconds) seconds)</div>
            </div>
        </div>
"@
            
            # Add Operation Details section for failure reports
            $html += @"
        <div class="section">
            <div class="section-title">Operation Details</div>
            <div class="info-grid">
                <div class="info-label">User Type:</div>
                <div class="info-value">$($ReportData.UserType)</div>
                <div class="info-label">Profile Path:</div>
                <div class="info-value">$($ReportData.TargetPath)</div>
"@
            if ($OperationType -eq 'Import') {
                $html += @"
                <div class="info-label">Source Zip:</div>
                <div class="info-value">$($ReportData.ZipPath)</div>
                <div class="info-label">Archive Size:</div>
                <div class="info-value">$($ReportData.ZipSizeMB) MB</div>
                <div class="info-label">Import Mode:</div>
                <div class="info-value">$($ReportData.MergeMode)</div>
"@
            }
            $html += @"
                <div class="info-label">Operation Time:</div>
                <div class="info-value">$($ReportData.ElapsedMinutes) minutes ($($ReportData.ElapsedSeconds) seconds)</div>
                <div class="info-label">Timestamp:</div>
                <div class="info-value">$($ReportData.Timestamp)</div>
            </div>
        </div>
"@
            
            # Add Profile Statistics if available
            if ($ReportData.ProfileStats) {
                $html += @"
        <div class="section">
            <div class="section-title">Profile Statistics</div>
            <div class="info-grid">
                <div class="info-label">Total Files:</div>
                <div class="info-value">$($ReportData.ProfileStats.TotalFiles)</div>
                <div class="info-label">Total Folders:</div>
                <div class="info-value">$($ReportData.ProfileStats.TotalFolders)</div>
                <div class="info-label">Total Size:</div>
                <div class="info-value">$($ReportData.ProfileStats.TotalSizeMB) MB</div>
            </div>
        </div>
"@
            }
            
            # Add Applications Installed if available
            if ($ReportData.Applications -and $ReportData.Applications.Count -gt 0) {
                $html += @"
        <div class="section">
            <div class="section-title">Applications Installed ($($ReportData.Applications.Count))</div>
            <table>
                <tr>
                    <th>Application Name</th>
                    <th>Version</th>
                    <th>Publisher</th>
                </tr>
"@
                foreach ($app in $ReportData.Applications) {
                    $html += @"
                <tr>
                    <td>$($app.DisplayName)</td>
                    <td>$($app.DisplayVersion)</td>
                    <td>$($app.Publisher)</td>
                </tr>
"@
                }
                $html += @"
            </table>
        </div>
"@
            }
            
            # Add Integrity Verification if available
            if ($ReportData.IntegrityCheck) {
                $html += @"
        <div class="section">
            <div class="section-title">Integrity Verification</div>
            <div class="info-grid">
                <div class="info-label">Registry Hive:</div>
                <div class="info-value">$($ReportData.IntegrityCheck.RegistryHive)</div>
                <div class="info-label">Critical Folders:</div>
                <div class="info-value">$($ReportData.IntegrityCheck.CriticalFolders)</div>
                <div class="info-label">User Permissions:</div>
                <div class="info-value">$($ReportData.IntegrityCheck.Permissions)</div>
            </div>
        </div>
"@
            }
            
            # Add Next Steps section for failure reports
            $html += @"
        <div class="section">
            <div class="section-title">Next Steps</div>
            <p><strong>Recommended Actions:</strong></p>
            <ul>
                <li>Review the error message above to identify the root cause</li>
                <li>Check the activity log for detailed error information</li>
                <li>If rollback was successful, the original profile has been restored</li>
                <li>Verify the source archive is not corrupted by testing extraction manually</li>
                <li>Ensure sufficient disk space is available on the target drive</li>
                <li>Check that no other processes are locking the profile folder</li>
                <li>If the target user exists, try logging in to verify the profile is not corrupted</li>
                <li>Review Windows Event Viewer for additional system-level errors</li>
                <li>Consider running the operation again after addressing the error</li>
            </ul>
        </div>
"@
        }
        
        # === EXPORT REPORT (SUCCESS ONLY) ===
        if ($OperationType -eq 'Export' -and $Success) {
            $checkmark = [char]0x2713
            $elapsedMins = $ReportData.ElapsedMinutes
            $elapsedSecs = $ReportData.ElapsedSeconds
            $html += @"
        <div class="section">
            <div class="section-title">Export Summary</div>
            <div class="info-grid">
                <div class="info-label">Status:</div>
                <div class="info-value success">$checkmark Export Completed Successfully</div>
                <div class="info-label">Source User:</div>
                <div class="info-value">$($ReportData.Username)</div>
                <div class="info-label">Source SID:</div>
                <div class="info-value">$($ReportData.SourceSID)</div>
                <div class="info-label">Source Path:</div>
                <div class="info-value">$($ReportData.SourcePath)</div>
                <div class="info-label">Export File:</div>
                <div class="info-value">$($ReportData.ZipPath)</div>
                <div class="info-label">Archive Size:</div>
                <div class="info-value">$($ReportData.ZipSizeMB) MB</div>
                <div class="info-label">Operation Time:</div>
                <div class="info-value">$elapsedMins minutes ($elapsedSecs seconds)</div>
                <div class="info-label">Timestamp:</div>
                <div class="info-value">$($ReportData.Timestamp)</div>
            </div>
        </div>
        
        <div class="section">
            <div class="section-title">Profile Statistics</div>
            <div class="stat-box">
                <h3>Files &amp; Folders</h3>
                <p>Total Files: $($ReportData.FileCount)</p>
                <p>Total Folders: $($ReportData.FolderCount)</p>
                <p>Uncompressed Size: $($ReportData.UncompressedSizeMB) MB</p>
                <p>Compression Ratio: $($ReportData.CompressionRatio)</p>
            </div>
        </div>
"@
            
            if ($ReportData.Exclusions -and $ReportData.Exclusions.Count -gt 0) {
                $html += @"
        <div class="section">
            <div class="section-title">Exclusions Applied</div>
            <ul>
"@
                foreach ($exclusion in $ReportData.Exclusions) {
                    $html += "                <li>$exclusion</li>`n"
                }
                $html += @"
            </ul>
        </div>
"@
            }
            
            # Add cleanup optimization section
            if ($ReportData.CleanupCategories -and $ReportData.CleanupCategories.Count -gt 0) {
                $html += @"
        <div class="section">
            <div class="section-title">Profile Cleanup Optimization</div>
            <div class="info-grid">
                <div class="info-label">Categories Cleaned:</div>
                <div class="info-value">$($ReportData.CleanupCategories.Count)</div>
                <div class="info-label">Space Saved:</div>
                <div class="info-value success">$($ReportData.CleanupSavingsMB) MB</div>
            </div>
"@
                
                # Detailed reporting for specific categories
                $detailCategories = @('Large Files', 'Duplicate Files', 'Large Downloads')
                foreach ($cat in $detailCategories) {
                    if ($ReportData.CategoryExclusions -and $ReportData.CategoryExclusions.ContainsKey($cat)) {
                        $paths = $ReportData.CategoryExclusions[$cat]
                        if ($paths.Count -gt 0) {
                            $html += "            <p><strong>Excluded $($cat):</strong></p>`n"
                            $html += "            <div class='paths-list'><ul>`n"
                            foreach ($p in $paths) {
                                # Use just the filename if it's too long, or relative path
                                $displayPath = if ($p.Length -gt 80) { "...$(Split-Path $p -Leaf)" } else { $p }
                                $html += "                <li>$displayPath</li>`n"
                            }
                            $html += "            </ul></div>`n"
                        }
                    }
                }

                $html += @"
            <p><strong>Other Excluded Categories:</strong></p>
            <ul>
"@
                foreach ($category in $ReportData.CleanupCategories) {
                    if ($detailCategories -notcontains $category) {
                        $html += "                <li>$category</li>`n"
                    }
                }
                $html += @"
            </ul>
        </div>
"@
            }

            # Add Cloud Exclusions section
            if ($ReportData.CloudExclusions -and $ReportData.CloudExclusions.Count -gt 0) {
                $html += @"
        <div class="section">
            <div class="section-title">Cloud Storage folders (Excluded)</div>
            <p>The following top-level cloud sync folders were found and automatically excluded to prevent errors with offline files:</p>
            <div class='paths-list'><ul>
"@
                foreach ($cloudPath in $ReportData.CloudExclusions) {
                    $html += "                <li>$cloudPath</li>`n"
                }
                $html += @"
            </ul></div>
            <p class="warning">Note: These folders are usually synced to OneDrive or SharePoint and must be resynced on the target computer.</p>
        </div>
"@
            }
            
            $html += "        <p class=`"success`">Significant items were excluded to reduce export size and improve transfer speed.</p>"
            
            if ($ReportData.HashEnabled) {
                $html += @"
        <div class="section">
            <div class="section-title">Integrity Verification</div>
            <div class="info-grid">
                <div class="info-label">Hash Algorithm:</div>
                <div class="info-value">SHA-256</div>
                <div class="info-label">Hash File:</div>
                <div class="info-value">$($ReportData.ZipPath).sha256</div>
                <div class="info-label">Status:</div>
                <div class="info-value success">&#10003; Hash generated</div>
            </div>
        </div>
"@
            }
            
            # Add installed programs list
            if ($ReportData.InstalledPrograms -and $ReportData.InstalledPrograms.Count -gt 0) {
                $html += @"
        <div class="section">
            <div class="section-title">Installed Programs on Source Computer</div>
            <p>Total Programs: $($ReportData.InstalledPrograms.Count)</p>
            <table>
                <thead>
                    <tr>
                        <th>Program Name</th>
                        <th>Version</th>
                        <th>Source</th>
                    </tr>
                </thead>
                <tbody>
"@
                foreach ($program in $ReportData.InstalledPrograms) {
                    $progName = if ($program.Name) { $program.Name } else { "Unknown" }
                    $progVersion = if ($program.Version) { $program.Version } else { "N/A" }
                    $progSource = if ($program.Source) { $program.Source } else { "N/A" }
                    $html += "                    <tr><td>$progName</td><td>$progVersion</td><td>$progSource</td></tr>`n"
                }
                $html += @"
                </tbody>
            </table>
        </div>
"@
            }
        }
        
        
        # === IMPORT REPORT (SUCCESS ONLY) ===
        if ($OperationType -eq 'Import' -and $Success) {
            $checkmark = [char]0x2713
            $elapsedMins = $ReportData.ElapsedMinutes
            $elapsedSecs = $ReportData.ElapsedSeconds
            $html += @"
        <div class="section">
            <div class="section-title">Import Summary</div>
            <div class="info-grid">
                <div class="info-label">Status:</div>
                <div class="info-value success">$checkmark Import Completed Successfully</div>
                <div class="info-label">Target User:</div>
                <div class="info-value">$($ReportData.Username)</div>
                <div class="info-label">User Type:</div>
                <div class="info-value">$($ReportData.UserType)</div>
                <div class="info-label">Target SID:</div>
                <div class="info-value">$($ReportData.TargetSID)</div>
                <div class="info-label">Source SID:</div>
                <div class="info-value">$($ReportData.SourceSID)</div>
                <div class="info-label">Profile Path:</div>
                <div class="info-value">$($ReportData.ProfilePath)</div>
                <div class="info-label">Source ZIP:</div>
                <div class="info-value">$($ReportData.ZipPath)</div>
                <div class="info-label">Archive Size:</div>
                <div class="info-value">$($ReportData.ZipSizeMB) MB</div>
                <div class="info-label">Import Mode:</div>
                <div class="info-value">$($ReportData.ImportMode)</div>
                <div class="info-label">Operation Time:</div>
                <div class="info-value">$elapsedMins minutes ($elapsedSecs seconds)</div>
                <div class="info-label">Timestamp:</div>
                <div class="info-value">$($ReportData.Timestamp)</div>
            </div>
        </div>
        
        <div class="section">
            <div class="section-title">Profile Statistics</div>
            <div class="stat-box">
                <h3>Files &amp; Folders</h3>
                <p>Total Files: $($ReportData.FileCount)</p>
                <p>Total Folders: $($ReportData.FolderCount)</p>
                <p>Uncompressed Size: $($ReportData.UncompressedSizeMB) MB</p>
                <p>Compression Ratio: $($ReportData.CompressionRatio)</p>
            </div>
        </div>
"@
            
            if ($ReportData.BackupPath) {
                $html += @"
        <div class="section">
            <div class="section-title">Backup Information</div>
            <div class="info-grid">
                <div class="info-label">Backup Created:</div>
                <div class="info-value success">&#10003; Yes</div>
                <div class="info-label">Backup Location:</div>
                <div class="info-value">$($ReportData.BackupPath)</div>
                <div class="info-label">Note:</div>
                <div class="info-value">Backup will be automatically cleaned up after 2 newest backups</div>
            </div>
        </div>
"@
            }
            
            # Add installed apps section
            if ($ReportData.InstalledApps -and $ReportData.InstalledApps.Count -gt 0) {
                $html += @"
        <div class="section">
            <div class="section-title">Applications Installed by Migration Tool</div>
            <p>Total Apps Installed: $($ReportData.InstalledApps.Count)</p>
            <table>
                <thead>
                    <tr>
                        <th>Application Name</th>
                        <th>Package ID</th>
                        <th>Version</th>
                    </tr>
                </thead>
                <tbody>
"@
                foreach ($app in $ReportData.InstalledApps) {
                    $appName = Get-FriendlyAppName -PackageId ($app.PackageIdentifier)
                    $appId = if ($app.PackageIdentifier) { $app.PackageIdentifier } else { "N/A" }
                    $appVersion = if ($app.Version) { $app.Version } else { "Latest" }
                    $html += "                    <tr><td>$appName</td><td>$appId</td><td>$appVersion</td></tr>`n"
                }
                $html += @"
                </tbody>
            </table>
        </div>
"@
            }
            
            if ($ReportData.HashVerified) {
                # Determine status color and icon based on verification result
                $statusClass = "success"
                $statusIcon = "&#10003;"
                if ($ReportData.HashVerificationStatus -like "*Failed*" -or $ReportData.HashVerificationStatus -like "*Bypassed*") {
                    $statusClass = "warning"
                    $statusIcon = "&#9888;"
                }
                elseif ($ReportData.HashVerificationStatus -like "*Skipped*" -or $ReportData.HashVerificationStatus -like "*Error*") {
                    $statusClass = "warning"
                    $statusIcon = "&#9888;"
                }
                
                $html += @"
        <div class="section">
            <div class="section-title">Integrity Verification</div>
            <div class="info-grid">
                <div class="info-label">Hash Verification:</div>
                <div class="info-value $statusClass">$statusIcon $($ReportData.HashVerificationStatus)</div>
                <div class="info-label">Details:</div>
                <div class="info-value">$($ReportData.HashVerificationDetails)</div>
            </div>
        </div>
"@
            }
            
            $html += @"
        <div class="section">
            <div class="section-title">Next Steps</div>
            <div class="stat-box">
                <h3>Post-Migration Actions Required</h3>
                <ul>
                    <li><strong>REBOOT NOW</strong> - Restart the computer before logging in</li>
                    <li>Log in as <strong>$($ReportData.Username)</strong></li>
"@
            
            # Only show Outlook-related items in REPLACE mode
            # In MERGE mode, we don't touch NTUSER.DAT or email settings
            if ($ReportData.ImportMode -eq 'Replace') {
                $html += @"
                    <li>Outlook will rebuild mailbox cache (10-30 minutes)</li>
                    <li>Re-enter email passwords if prompted</li>
                    <li>Reconnect to Exchange/Microsoft 365 if needed</li>
                    <li>Copy PST archives manually if stored outside profile</li>
"@
            }
            
            $html += @"
                </ul>
            </div>
        </div>
"@
        }
        
        # === WARNINGS/ERRORS (if any) ===
        if ($ReportData.Warnings -and $ReportData.Warnings.Count -gt 0) {
            $html += @"
        <div class="section">
            <div class="section-title">Warnings</div>
            <ul>
"@
            foreach ($warning in $ReportData.Warnings) {
                $html += "                <li class=`"warning`">WARNING: $warning</li>`n"
            }
            $html += @"
            </ul>
        </div>
"@
        }
        
        # Footer
        $html += @"
        <div class="footer">
            <p>Profile Migration Tool 2.13.0 | Generated on $(hostname) | PowerShell $($PSVersionTable.PSVersion)</p>
        </div>
    </div>
</body>
</html>
"@
        
        # Save HTML report
        $html | Out-File -FilePath $reportPath -Encoding UTF8 -Force
        Log-Info "HTML report saved: $reportPath"
        
        # Auto-open report if configured
        if ($global:Config.AutoOpenReports) {
            Log-Debug "Opening HTML report..."
            Start-Process $reportPath
        }
        
        return $reportPath
        
    }
    catch {
        Log-Error "Failed to generate migration report: $_"
        return $null
    }
}

function New-ConversionReport {
    param(
        [string]$SourceUser,
        [string]$TargetUser,
        [string]$SourceType,
        [string]$TargetType,
        [string]$Status,
        [DateTime]$StartTime,
        [DateTime]$EndTime,
        [string]$BackupPath,
        [string]$BackupSizeMB,
        [string]$LogPath,
        [string]$JoinedDomain = "Not Changed",
        [string]$AzureStatus = "Not Changed"
    )

    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        # Sanitize usernames for filename
        $sourceUserSafe = ($SourceUser -split '\\')[-1] -replace '[\\/:*?"<>|]', '_'
        $targetUserSafe = ($TargetUser -split '\\')[-1] -replace '[\\/:*?"<>|]', '_'
        $reportFilename = "ConversionReport_${sourceUserSafe}_to_${targetUserSafe}_${timestamp}.html"
        
        $logsDir = Split-Path -Parent $LogPath
        if (-not (Test-Path $logsDir)) { New-Item -ItemType Directory -Path $logsDir -Force | Out-Null }
        $reportPath = Join-Path $logsDir $reportFilename

        $duration = $EndTime - $StartTime
        $durationStr = "{0:hh\:mm\:ss}" -f $duration
        
        # Colors
        $isSuccess = $Status -eq "Success"
        $statusColor = if ($isSuccess) { "#107c10" } else { "#d13438" } # Green/Red
        $statusIcon = if ($isSuccess) { "&#10004;" } else { "&#10060;" }

        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Profile Conversion Report</title>
    <style>
        :root {
            --primary-color: #0078d4;
            --success-color: #107c10;
            --error-color: #d13438;
            --bg-color: #f3f2f1;
            --card-bg: #ffffff;
            --text-primary: #323130;
            --text-secondary: #605e5c;
            --border-radius: 8px;
            --shadow: 0 1.6px 3.6px 0 rgba(0,0,0,0.13), 0 0.3px 0.9px 0 rgba(0,0,0,0.11);
        }
        body {
            font-family: 'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif;
            background-color: var(--bg-color);
            color: var(--text-primary);
            margin: 0;
            padding: 40px 20px;
            line-height: 1.5;
        }
        .container {
            max-width: 900px;
            margin: 0 auto;
        }
        .header {
            background-color: $statusColor;
            color: white;
            padding: 30px;
            border-radius: var(--border-radius);
            box-shadow: var(--shadow);
            margin-bottom: 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .header-content h1 { margin: 0; font-size: 24px; font-weight: 600; }
        .header-content p { margin: 5px 0 0; opacity: 0.9; }
        .status-badge {
            background-color: rgba(255,255,255,0.2);
            padding: 8px 16px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 24px;
            margin-bottom: 24px;
        }
        .card {
            background-color: var(--card-bg);
            border-radius: var(--border-radius);
            padding: 24px;
            box-shadow: var(--shadow);
        }
        .card-title {
            font-size: 16px;
            font-weight: 600;
            color: var(--text-secondary);
            margin-bottom: 16px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            border-bottom: 1px solid #e1dfdd;
            padding-bottom: 8px;
        }
        .stat-group { margin-bottom: 16px; }
        .stat-label { font-size: 12px; color: var(--text-secondary); font-weight: 600; }
        .stat-value { font-size: 16px; font-weight: 400; color: var(--text-primary); word-break: break-all; }
        .stat-value.highlight { font-weight: 600; color: var(--primary-color); }
        
        .footer {
            text-align: center;
            color: var(--text-secondary);
            font-size: 12px;
            margin-top: 40px;
        }
        code {
            background-color: #f3f2f1;
            padding: 2px 4px;
            border-radius: 4px;
            font-family: Consolas, monospace;
            font-size: 0.9em;
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Hero Header -->
        <div class="header">
            <div class="header-content">
                <h1>Profile Conversion Report</h1>
                <p>Migration execution summary</p>
            </div>
            <div class="status-badge">
                <span>$statusIcon</span>
                <span>$Status</span>
            </div>
        </div>

        <!-- Quick Stats Row -->
        <div class="grid" style="grid-template-columns: repeat(3, 1fr);">
            <div class="card" style="text-align: center; padding: 16px;">
                <div class="stat-label">DURATION</div>
                <div class="stat-value" style="font-size: 24px;">$durationStr</div>
            </div>
            <div class="card" style="text-align: center; padding: 16px;">
                <div class="stat-label">DATE</div>
                <div class="stat-value" style="font-size: 18px;">$(Get-Date -Format "yyyy-MM-dd")</div>
            </div>
            <div class="card" style="text-align: center; padding: 16px;">
                <div class="stat-label">COMPUTER</div>
                <div class="stat-value" style="font-size: 18px;">$env:COMPUTERNAME</div>
            </div>
        </div>

        <!-- Main Details Grid -->
        <div class="grid">
            <!-- Identity Card -->
            <div class="card">
                <div class="card-title">Identity Migration</div>
                <div class="stat-group">
                    <div class="stat-label">SOURCE PROFILE</div>
                    <div class="stat-value highlight">$SourceUser</div>
                    <div class="stat-value" style="font-size: 12px; color: #605e5c;">Type: $SourceType</div>
                </div>
                <div style="text-align: center; margin: 10px 0; color: #ccc;">&darr;</div>
                <div class="stat-group">
                    <div class="stat-label">TARGET PROFILE</div>
                    <div class="stat-value highlight">$TargetUser</div>
                    <div class="stat-value" style="font-size: 12px; color: #605e5c;">Type: $TargetType</div>
                </div>
            </div>

            <!-- System & Network Card -->
            <div class="card">
                <div class="card-title">System Status</div>
                <div class="stat-group">
                    <div class="stat-label">DOMAIN JOIN STATUS</div>
                    <div class="stat-value">$JoinedDomain</div>
                </div>
                <div class="stat-group">
                    <div class="stat-label">AZURE AD STATUS</div>
                    <div class="stat-value">$AzureStatus</div>
                </div>
                <div class="stat-group">
                    <div class="stat-label">LOG FILE PATH</div>
                    <div class="stat-value"><code>$LogPath</code></div>
                </div>
            </div>
        </div>

        <!-- Storage Card -->
        <div class="card">
            <div class="card-title">Storage & Backup</div>
            <div class="grid" style="grid-template-columns: 1fr 1fr; gap: 16px; margin: 0; box-shadow: none; padding: 0;">
                <div class="stat-group">
                    <div class="stat-label">BACKUP STATUS</div>
                    <div class="stat-value">$(if ($BackupPath) { "Created" } else { "Skipped" })</div>
                </div>
                <div class="stat-group">
                    <div class="stat-label">BACKUP SIZE</div>
                    <div class="stat-value">$(if ($BackupSizeMB) { "$BackupSizeMB MB" } else { "-" })</div>
                </div>
            </div>
            <div class="stat-group">
                <div class="stat-label">BACKUP PATH</div>
                <div class="stat-value"><code>$(if ($BackupPath) { $BackupPath } else { "N/A" })</code></div>
            </div>
        </div>


        <!-- Next Steps Section -->
        <div class="card" style="margin-top: 24px;">
            <div class="card-title">Next Steps</div>
"@

        $html += @"
            <div class="success" style="background: #dff6dd; border-left: 4px solid #107c10; border-radius: 4px; padding: 16px; margin: 16px 0;">
                <div style="font-weight: bold; margin-bottom: 8px; color: #0b5a08;">&#10003; Standard Login Procedure</div>
                <ol style="margin: 8px 0; padding-left: 20px;">
                    <li>Restart the computer</li>
                    <li>Log in as <strong>$TargetUser</strong></li>
                    <li>Desktop will load normally</li>
                </ol>
            </div>
"@

        $html += @"
        </div>

        <div class="footer">
            Generated by Profile Migration Tool v$($global:Config.Version) &bull; $(Get-Date -Format "F")
        </div>
    </div>
</body>
</html>
"@

        $html | Out-File -FilePath $reportPath -Encoding UTF8
        return $reportPath
    }
    catch {
        Log-Error "Failed to generate HTML report: $_"
        return $null
    }
}


function Start-OperationLog {
    param(
        [Parameter(Mandatory = $true)][string]$Operation,
        [string]$Username
    )
    
    # Create logs directory if it doesn't exist
    $logsDir = Join-Path $PSScriptRoot "Logs"
    if (-not (Test-Path $logsDir)) {
        New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    }
    
    # Create timestamped log file
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    # Strip size suffix if present (e.g., "username - [756.6 MB]" -> "username")
    $cleanUsername = $Username
    if ($cleanUsername -match '^(.+?)\s+-\s+\[.+\]$') {
        $cleanUsername = $matches[1].Trim()
    }
    $sanitizedUser = if ($cleanUsername) { $cleanUsername -replace '[\\/:*?"<>|]', '_' } else { "Unknown" }
    $logFileName = "${Operation}_${sanitizedUser}_${timestamp}.log"
    $global:CurrentLogFile = Join-Path $logsDir $logFileName
    
    # Write header to log file
    $header = @"
================================================================================
Profile Migration Tool - $Operation Log
Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
User: $sanitizedUser
Computer: $env:COMPUTERNAME
PowerShell: $($PSVersionTable.PSVersion)
OS: $([Environment]::OSVersion.VersionString)
================================================================================

"@
    
    try {
        Set-Content -Path $global:CurrentLogFile -Value $header -ErrorAction Stop
        Log-Info "Log file created: $global:CurrentLogFile"
        return $global:CurrentLogFile
    }
    catch {
        Write-Host "WARNING: Could not create log file: $_" -ForegroundColor Yellow
        $global:CurrentLogFile = $null
        return $null
    }
}

function Stop-OperationLog {
    if ($global:CurrentLogFile) {
        $footer = "`n" + "=" * 80 + "`nOperation completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" + "=" * 80
        try {
            Add-Content -Path $global:CurrentLogFile -Value $footer -ErrorAction SilentlyContinue
        }
        catch { }
        $global:CurrentLogFile = $null
    }
}

# Enhanced profile detection with size estimates
function Repair-UserProfile {
    param(
        [Parameter(Mandatory = $true)][string]$Username,
        [Parameter(Mandatory = $true)][string]$UserType
    )

    try {
        Log-Info "=== Starting Profile Repair (Universal) ==="
        Log-Info "User: $Username ($UserType)"

        # Step 1: Get SID
        Log-Info "Resolving user SID..."
        # Get-LocalUserSID handles Local, Domain and AzureAD (if graphed/cached)
        $sid = Get-LocalUserSID -UserName $Username
        Log-Info "SID: $sid"

        if (-not $sid) { throw "Could not resolve SID for user: $Username" }

        # Step 2: Get Profile Path
        $profileListBase = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
        try {
            $profileKey = Get-ItemProperty -Path "$profileListBase\$sid" -ErrorAction Stop
            $profilePath = $profileKey.ProfileImagePath
            Log-Info "Profile path: $profilePath"
        }
        catch {
            throw "Profile entry not found in registry for SID: $sid"
        }

        # Step 3: Update Registry (Reset State)
        Log-Info "Resetting profile state in registry..."
        # We pass the same SID and Path. Update-ProfileListRegistry handles the rest (State=0)
        $regResult = Update-ProfileListRegistry -OldSID $sid -NewSID $sid -NewProfilePath $profilePath

        if (-not $regResult.Success) {
            throw "Registry update failed: $($regResult.Error)"
        }

        # Step 4: Apply ACLs
        Log-Info "Applying permissions repair (ACLs & Ownership)..."
        if ($global:ConversionProgressBar) { 
            $global:ConversionProgressBar.Value = 50 
            if ($global:ConversionStatusLabel) { $global:ConversionStatusLabel.Text = "Repairing permissions..." }
            [System.Windows.Forms.Application]::DoEvents() 
        }

        # Set-ProfileAcls handles file permissions, hive permissions, and ownership
        Set-ProfileAcls -ProfileFolder $profilePath -UserSID $sid -UserName $Username -SourceSID $sid -OldProfilePath $profilePath -NewProfilePath $profilePath

        Log-Info "=== Profile Repair Completed Successfully ==="
        return @{
            Success = $true
            SID     = $sid
            Path    = $profilePath
        }
    }
    catch {
        Log-Error "Profile repair failed: $_"
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}

# Convert a local user profile to a domain user profile
function Convert-LocalToDomain {
    param(
        [Parameter(Mandatory = $true)][string]$LocalUsername,
        [Parameter(Mandatory = $true)][string]$DomainUsername,
        [PSCredential]$DomainCredential,
        
        [Parameter(Mandatory = $false)]
        [bool]$UnjoinAzureAD = $false,
        
        [Parameter(Mandatory = $false)]
        [string]$SourceSID = $null  # Pre-resolved source SID (for AzureAD→Domain after unjoin)
    )
    
    try {
        Log-Info "=== Starting Local to Domain Profile Conversion ==="
        Log-Info "Source: $LocalUsername (Local)"
        Log-Info "Target: $DomainUsername (Domain)"
        
        $azureAdStatus = "Not Changed"
        
        # CRITICAL: Check if device is AzureAD joined
        # Windows does not allow domain join while AzureAD joined - unjoin is REQUIRED
        # EXCEPTION: If device is already domain joined (Hybrid), we don't need to unjoin
        
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        if ($computerSystem.PartOfDomain) {
            Log-Info "Device is already domain joined (Hybrid state possible) - skipping AzureAD unjoin check"
        }
        elseif (Test-IsAzureADJoined) {
            Log-Info "Device is AzureAD joined - unjoin is REQUIRED before domain join"
            
            $message = "This device is currently joined to AzureAD.`r`n`r`n"
            $message += "Windows REQUIRES unjoining from AzureAD before joining a domain.`r`n`r`n"
            $message += "This will:`r`n"
            $message += "- Remove device from your organization's management`r`n"
            $message += "- Disable conditional access policies`r`n"
            $message += "- Remove SSO capabilities`r`n`r`n"
            $message += "After unjoin, the conversion will proceed with domain join.`r`n`r`n"
            $message += "Continue with required AzureAD unjoin?"
            
            $response = Show-ModernDialog -Message $message -Title "AzureAD Unjoin Required" -Type Warning -Buttons YesNo
            
            if ($response -eq 'Yes') {
                $unjoinResult = Invoke-AzureADUnjoin
                
                if ($unjoinResult.Success) {
                    Log-Info "AzureAD unjoin successful - can now proceed with domain join"
                    $azureAdStatus = "Unjoined"
                }
                else {
                    $errorMsg = "AzureAD unjoin failed: $($unjoinResult.Message)`r`n`r`n"
                    $errorMsg += "Cannot proceed with domain join while AzureAD joined.`r`n`r`n"
                    $errorMsg += "You can manually unjoin using: dsregcmd /leave"
                    Show-ModernDialog -Message $errorMsg -Title "Unjoin Failed" -Type Error -Buttons OK
                    throw "AzureAD unjoin failed - cannot join domain"
                }
            }
            else {
                Log-Info "User cancelled required AzureAD unjoin"
                throw "User cancelled AzureAD unjoin - cannot join domain while AzureAD joined"
            }
        }
        elseif ($UnjoinAzureAD) {
            # User requested unjoin but device is not AzureAD joined
            Log-Info "User requested AzureAD unjoin but device is not AzureAD joined - skipping"
        }
        
        # Step 1: Get source SID
        Log-Info "Resolving source user SID..."
        if ($SourceSID) {
            # Use pre-resolved SID (for AzureAD→Domain after unjoin)
            $sourceSID = $SourceSID
            Log-Info "Using pre-resolved source SID: $sourceSID"
        }
        else {
            # Resolve SID normally
            $sourceSID = Get-LocalUserSID -UserName $LocalUsername
            Log-Info "Source SID: $sourceSID"
        }
        
        # Step 2: Resolve domain user SID from Active Directory
        Log-Info "Resolving domain user SID from Active Directory..."
        try {
            # Get-LocalUserSID uses NTAccount.Translate() which queries AD (no local login required)
            $targetSID = Get-LocalUserSID -UserName $DomainUsername
            Log-Info "Domain user SID: $targetSID"
        }
        catch {
            Log-Warning "Could not resolve domain user SID from Active Directory: $_"
            
            # Prompt user to verify the account exists in AD
            $response = Show-ModernDialog -Message "The domain user '$DomainUsername' could not be found in Active Directory.`r`n`r`nThis usually means:`r`n`r`n1. The username is incorrect (check spelling)`r`n2. The user account doesn't exist in Active Directory`r`n3. The domain is not reachable from this computer`r`n`r`nPlease verify:`r`n- The username is spelled correctly (format: DOMAIN\username)`r`n- The user account exists in Active Directory`r`n- This computer can reach the domain controller`r`n`r`nCorrect the username and retry the conversion." -Title "Domain User Not Found" -Type Warning -Buttons OK
            throw "Domain user '$DomainUsername' not found in Active Directory - verify username is correct"
        }
        
        if ($global:ConversionProgressBar) { 
            $global:ConversionProgressBar.Value = 35
            if ($global:ConversionStatusLabel) { $global:ConversionStatusLabel.Text = "Resolving user SIDs..." }
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        # Step 3: Get profile paths
        $profileListBase = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
        $sourceProfileKey = Get-ItemProperty -Path "$profileListBase\$sourceSID" -ErrorAction Stop
        $sourceProfilePath = $sourceProfileKey.ProfileImagePath
        
        $targetProfileKey = Get-ItemProperty -Path "$profileListBase\$targetSID" -ErrorAction SilentlyContinue
        
        # CRITICAL: Handle dotted usernames to prevent black screen
        # Windows creates profile folders as C:\Users\username (NO domain prefix by default)
        # The EXCEPTION is dotted usernames which can cause issues
        # If ProfileImagePath doesn't match the actual folder Windows creates, user gets black screen
        
        if ($targetProfileKey) {
            # Profile already exists - use existing path
            $targetProfilePath = $targetProfileKey.ProfileImagePath
            Log-Info "Target profile already exists, using existing path"
        }
        else {
            # New profile - Windows creates folder as C:\Users\username (no domain prefix)
            $usernameOnly = $DomainUsername.Split('\')[1]
            $targetProfilePath = "C:\Users\$usernameOnly"
            Log-Info "Target profile path will be: $targetProfilePath"
        }
        
        Log-Info "Source profile path: $sourceProfilePath"
        Log-Info "Target profile path: $targetProfilePath"
        
        # Verify source profile integrity
        if (-not (Test-ValidProfilePath -Path $sourceProfilePath -RequireNTUSER)) {
            throw "Source profile corrupted or missing (NTUSER.DAT not found): $sourceProfilePath"
        }
        
        # Step 4: Check if target profile already has data
        if (Test-ValidProfilePath -Path $targetProfilePath -RequireNTUSER) {
            # Check if source and target paths are the same
            if ($sourceProfilePath -eq $targetProfilePath) {
                # Same path - registry-only conversion
                $response = Show-ModernDialog -Message "The profile folder is already in use by the AzureAD user.`r`n`r`nThis conversion will UPDATE the registry to associate this profile with the domain user instead.`r`n`r`nNo files will be moved or deleted - only registry SID associations will change.`r`n`r`nContinue with conversion?" -Title "Profile Conversion" -Type Info -Buttons YesNo
            }
            else {
                # Different paths - file copy required
                $response = Show-ModernDialog -Message "The domain user already has a profile on this computer.`r`n`r`nDo you want to OVERWRITE it with the AzureAD user's profile?`r`n`r`nWARNING: This will delete the existing domain profile!" -Title "Overwrite Existing Profile" -Type Warning -Buttons YesNo
            }
                
            if ($response -ne "Yes") {
                throw "User cancelled: Target profile already exists"
            }
                
            if ($sourceProfilePath -eq $targetProfilePath) {
                Log-Info "User approved registry-only conversion (same profile path)"
            }
            else {
                Log-Warning "User chose to overwrite existing domain profile"
            }
        }
        
        # Step 5: Copy OR Rename profile data
        if ($global:CancelRequested) { throw "Operation cancelled by user" }
        Log-Info "Transferring profile data from local to domain user..."
        Update-ConversionProgress -PercentComplete 40
        
        # If paths are different, check if we can RENAME (Move) instead of Copy
        if ($sourceProfilePath -ne $targetProfilePath) {
            
            # --- RENAME STRATEGY CHECK ---
            $canRename = $false
            try {
                $srcDrive = [System.IO.Path]::GetPathRoot($sourceProfilePath)
                $dstDrive = [System.IO.Path]::GetPathRoot($targetProfilePath)
                if ($srcDrive -eq $dstDrive) {
                    $canRename = $true
                }
            }
            catch {
                Log-Warning "Could not determine drive volumes: $_"
            }

            $renameSuccess = $false
            if ($canRename) {
                Log-Info "Source and Target are on the same volume ($srcDrive). Attempting fast RENAME..."
                try {
                    # If target exists (and we are here, meaning overwrite was approved), we must delete it first
                    if (Test-Path $targetProfilePath) {
                        Log-Info "Removing existing target profile folder to allow rename..."
                        # Use robust deletion (sometimes simple Remove-Item fails on profiles)
                        Remove-Item -Path $targetProfilePath -Recurse -Force -ErrorAction Stop
                    }
                    
                    # Perform the Rename
                    Log-Info "Renaming '$sourceProfilePath' -> '$targetProfilePath'"
                    Rename-Item -Path $sourceProfilePath -NewName $targetProfilePath -ErrorAction Stop
                    $renameSuccess = $true
                    Update-ConversionProgress -PercentComplete 70 -StatusMessage "Copying profile files..."
                    Log-Info "Fast Rename successful!"
                }
                catch {
                    Log-Warning "Rename failed: $_. Falling back to Robocopy strategy."
                    # If rename failed, source might still be there, or partially moved? 
                    # Rename is usually atomic-ish, but if it failed, we assume Source is still Source.
                    # If Target was deleted, Robocopy will recreate it.
                }
            }
            else {
                Log-Info "Source and Target are on different volumes. Rename not possible."
            }
            
            # --- ROBOCOPY FALLBACK ---
            if (-not $renameSuccess) {
                # Create target directory if it doesn't exist
                if (-not (Test-Path $targetProfilePath)) {
                    New-Item -ItemType Directory -Path $targetProfilePath -Force | Out-Null
                }
                
                # Use robocopy to copy all files
                Log-Info "Using robocopy to copy profile files..."
                $robocopyArgs = @(
                    "`"$sourceProfilePath`"",
                    "`"$targetProfilePath`"",
                    "/E",           # Copy subdirectories including empty
                    "/COPYALL",     # Copy all file info
                    "/R:1",         # Retry once
                    "/W:1",         # Wait 1 second between retries
                    "/MT:$($Config.RobocopyThreads)",  # Multi-threaded
                    "/NFL",         # No file list
                    "/NDL",         # No directory list
                    "/NJH",         # No job header
                    "/NJS",         # No job summary
                    "/DCOPY:DAT"    # Copy directory timestamps
                )
                
                $robocopyProcess = Start-Process -FilePath "robocopy.exe" -ArgumentList $robocopyArgs -NoNewWindow -Wait -PassThru
                
                # Robocopy exit codes: 0-7 are success, 8+ are errors
                if ($robocopyProcess.ExitCode -ge 8) {
                    throw "Robocopy failed with exit code: $($robocopyProcess.ExitCode)"
                }
                
                Log-Info "Profile files copied successfully"
                Update-ConversionProgress -PercentComplete 70
            }
        }
        else {
            Log-Info "Source and target paths are the same, skipping file transfer"
        }
        
        # Step 6: Update registry
        if ($global:CancelRequested) { throw "Operation cancelled by user" }
        Log-Info "Updating ProfileList registry..."
        $regResult = Update-ProfileListRegistry -OldSID $sourceSID -NewSID $targetSID -NewProfilePath $targetProfilePath
        
        if (-not $regResult.Success) {
            throw "Registry update failed: $($regResult.Error)"
        }
        
        # Step 7: Apply ACLs
        Log-Info "Applying folder ACLs for domain user..."
        Update-ConversionProgress -PercentComplete 90 -StatusMessage "Applying folder permissions..."
        # Redundant call removed: Set-ProfileFolderAcls is called internally by Set-ProfileAcls
        
        # Step 8 & 9: Apply hive ACLs AND Rewrite SIDs (Unified via Set-ProfileAcls)
        # This wrapper handles UsrClass.dat reset and other critical fixes
        Log-Message "Applying hive ACLs and rewriting SIDs..."
        Set-ProfileAcls -ProfileFolder $targetProfilePath -UserSID $targetSID -UserName $DomainUsername -SourceSID $sourceSID -OldProfilePath $sourceProfilePath -NewProfilePath $targetProfilePath
        
        
        Log-Info "=== Local to Domain Conversion Completed Successfully ==="
        Update-ConversionProgress -PercentComplete 100 -StatusMessage "Conversion completed successfully!"
        
        
        return @{
            Success           = $true
            SourceSID         = $sourceSID
            TargetSID         = $targetSID
            SourceProfilePath = $sourceProfilePath
            TargetProfilePath = $targetProfilePath
        }
    }
    catch {
        Log-Error "Local to Domain conversion failed: $_"
        
        # ROLLBACK LOGIC
        try {
            Log-Info "Attempting rollback..."
            if ($renameSuccess) {
                Log-Info "Rolling back RENAME operation..."
                # Target became Source. Rename it back.
                Rename-Item -Path $targetProfilePath -NewName $sourceProfilePath -Force -ErrorAction Stop
                Log-Info "Rollback successful: Restored $sourceProfilePath"
            }
            elseif ($sourceProfilePath -ne $targetProfilePath -and (Test-Path $targetProfilePath)) {
                # Only delete if we were copying (not same path) and target exists
                # If usage was 'Registry Only' (Same Path), we do NOTHING.
                Log-Info "Rolling back COPY operation: Removing target $targetProfilePath"
                Remove-FolderRobust -Path $targetProfilePath -ErrorAction SilentlyContinue
            }
        }
        catch {
            Log-Warning "Rollback failed: $_"
        }

        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}

# Convert a domain user profile to a local user profile
function Convert-DomainToLocal {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DomainUsername,
        
        [Parameter(Mandatory = $true)]
        [string]$LocalUsername,
        
        [Parameter(Mandatory = $true)]
        [SecureString]$LocalPassword,
        
        [Parameter(Mandatory = $false)]
        [bool]$MakeAdmin = $false,
        
        [Parameter(Mandatory = $false)]
        [bool]$UnjoinDomain = $false
    )
    
    try {
        Log-Info "Starting Domain to Local profile conversion..."
        Log-Info "Domain user: $DomainUsername"
        Log-Info "Target local user: $LocalUsername"
        Log-Info "Make administrator: $MakeAdmin"
        
        # Step 1: Get source SID
        Log-Info "Resolving source domain user SID..."
        $sourceSID = Get-LocalUserSID -UserName $DomainUsername
        Log-Info "Source SID: $sourceSID"
        
        # Step 2: Create local user if it doesn't exist
        $localUserExists = $false
        try {
            $existingUser = Get-LocalUser -Name $LocalUsername -ErrorAction Stop
            $localUserExists = $true
            Log-Info "Local user '$LocalUsername' already exists"
        }
        catch {
            Log-Info "Creating new local user '$LocalUsername'..."
            try {
                New-LocalUser -Name $LocalUsername -Password $LocalPassword -FullName $LocalUsername -Description "Converted from domain profile" -ErrorAction Stop
                Log-Info "Local user created successfully"
                
                # Add to Administrators group if requested
                if ($MakeAdmin) {
                    Log-Info "Adding user to Administrators group..."
                    Add-LocalGroupMember -Group "Administrators" -Member $LocalUsername -ErrorAction Stop
                    Log-Info "User added to Administrators group"
                }
                
                # Add to Users group
                Add-LocalGroupMember -Group "Users" -Member $LocalUsername -ErrorAction SilentlyContinue
                Log-Info "Added to Users group"
            }
            catch {
                throw "Failed to create local user: $_"
            }
        }
        
        # Step 3: Get target SID
        Log-Info "Resolving target local user SID..."
        $targetSID = Get-LocalUserSID -UserName $LocalUsername
        Log-Info "Target SID: $targetSID"
        
        # Step 4: Get profile paths
        if ($global:ConversionProgressBar) { $global:ConversionProgressBar.Value = 35; [System.Windows.Forms.Application]::DoEvents() }
        $profileListBase = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
        $sourceProfileKey = Get-ItemProperty -Path "$profileListBase\$sourceSID" -ErrorAction Stop
        $sourceProfilePath = $sourceProfileKey.ProfileImagePath
        
        $targetProfilePath = "C:\Users\$LocalUsername"
        
        Log-Info "Source profile path: $sourceProfilePath"
        Log-Info "Target profile path: $targetProfilePath"
        
        # Verify source profile integrity
        if (-not (Test-ValidProfilePath -Path $sourceProfilePath -RequireNTUSER)) {
            throw "Source profile corrupted or missing (NTUSER.DAT not found): $sourceProfilePath"
        }
        
        # Step 5: Copy OR Rename profile data
        if ($global:CancelRequested) { throw "Operation cancelled by user" }
        Log-Info "Transferring profile data from domain to local user..."
        Update-ConversionProgress -PercentComplete 40
        
        # If paths are different, check if we can RENAME (Move) instead of Copy
        if ($sourceProfilePath -ne $targetProfilePath) {
            
            # --- RENAME STRATEGY CHECK ---
            $canRename = $false
            try {
                $srcDrive = [System.IO.Path]::GetPathRoot($sourceProfilePath)
                $dstDrive = [System.IO.Path]::GetPathRoot($targetProfilePath)
                if ($srcDrive -eq $dstDrive) {
                    $canRename = $true
                }
            }
            catch {
                Log-Warning "Could not determine drive volumes: $_"
            }

            $renameSuccess = $false
            if ($canRename) {
                Log-Info "Source and Target are on the same volume ($srcDrive). Attempting fast RENAME..."
                try {
                    # If target exists, we must delete it first (overwrite scenario)
                    if (Test-Path $targetProfilePath) {
                        Log-Info "Removing existing target profile folder to allow rename..."
                        Remove-Item -Path $targetProfilePath -Recurse -Force -ErrorAction Stop
                    }
                    
                    # Perform the Rename
                    Log-Info "Renaming '$sourceProfilePath' -> '$targetProfilePath'"
                    Rename-Item -Path $sourceProfilePath -NewName $targetProfilePath -ErrorAction Stop
                    $renameSuccess = $true
                    Update-ConversionProgress -PercentComplete 70
                    Log-Info "Fast Rename successful!"
                }
                catch {
                    Log-Warning "Rename failed: $_. Falling back to Robocopy strategy."
                }
            }
            else {
                Log-Info "Source and Target are on different volumes. Rename not possible."
            }

            # --- ROBOCOPY FALLBACK ---
            if (-not $renameSuccess) {
                # Create target directory if it doesn't exist
                if (-not (Test-Path $targetProfilePath)) {
                    New-Item -ItemType Directory -Path $targetProfilePath -Force | Out-Null
                }
                
                # Use robocopy to copy all files
                Log-Info "Using robocopy to copy profile files..."
                $robocopyArgs = @(
                    "`"$sourceProfilePath`"",
                    "`"$targetProfilePath`"",
                    "/E",           # Copy subdirectories including empty
                    "/COPYALL",     # Copy all file info
                    "/R:1",         # Retry once
                    "/W:1",         # Wait 1 second between retries
                    "/MT:$($Config.RobocopyThreads)",  # Multi-threaded
                    "/NFL",         # No file list
                    "/NDL",         # No directory list
                    "/NJH",         # No job header
                    "/NJS",         # No job summary
                    "/DCOPY:DAT"    # Copy directory timestamps
                )
                
                $robocopyProcess = Start-Process -FilePath "robocopy.exe" -ArgumentList $robocopyArgs -NoNewWindow -Wait -PassThru
                
                # Robocopy exit codes: 0-7 are success, 8+ are errors
                if ($robocopyProcess.ExitCode -ge 8) {
                    throw "Robocopy failed with exit code: $($robocopyProcess.ExitCode)"
                }
                
                Log-Info "Profile files copied successfully"
                Update-ConversionProgress -PercentComplete 70
            }
        }
        
        # Step 6: Update registry
        if ($global:CancelRequested) { throw "Operation cancelled by user" }
        Log-Info "Updating ProfileList registry..."
        $regResult = Update-ProfileListRegistry -OldSID $sourceSID -NewSID $targetSID -NewProfilePath $targetProfilePath
        
        if (-not $regResult.Success) {
            throw "Registry update failed: $($regResult.Error)"
        }
        
        # Step 7: Apply ACLs
        Log-Info "Applying folder ACLs for local user..."
        Update-ConversionProgress -PercentComplete 90
        # Redundant call removed: Set-ProfileFolderAcls is called internally by Set-ProfileAcls
        
        # Step 8 & 9: Apply hive ACLs AND Rewrite SIDs (Unified via Set-ProfileAcls)
        # This wrapper handles UsrClass.dat reset and other critical fixes
        Log-Message "Applying hive ACLs and rewriting SIDs..."
        Set-ProfileAcls -ProfileFolder $targetProfilePath -UserSID $targetSID -UserName $LocalUsername -SourceSID $sourceSID -OldProfilePath $sourceProfilePath -NewProfilePath $targetProfilePath
        
        Log-Info "=== Domain to Local Conversion Completed Successfully ==="
        Update-ConversionProgress -PercentComplete 100 -StatusMessage "Domain to Local Conversion Completed Successfully"
        
        # Handle optional domain unjoin
        if ($UnjoinDomain) {
            Log-Info "User requested domain unjoin..."
            
            # Show warning dialog
            if (Confirm-DomainUnjoin) {
                $unjoinResult = Invoke-DomainUnjoin
                
                if ($unjoinResult.Success) {
                    Log-Info "Domain unjoin successful"
                }
                else {
                    Log-Warn "Domain unjoin failed: $($unjoinResult.Message)"
                    # Don't fail the entire conversion, just warn
                    Show-ModernDialog -Message "Profile conversion succeeded, but domain unjoin failed:`r`n`r`n$($unjoinResult.Message)`r`n`r`nYou can manually unjoin via System Properties." -Title "Unjoin Warning" -Type Warning -Buttons OK
                }
            }
            else {
                Log-Info "User cancelled domain unjoin"
            }
        }
        
        
        # AppX re-registration handled in main Convert-UserProfile function now
        
        return @{
            Success           = $true
            SourceSID         = $sourceSID
            TargetSID         = $targetSID
            SourceProfilePath = $sourceProfilePath
            TargetProfilePath = $targetProfilePath
        }
    }
    catch {
        Log-Error "Domain to Local conversion failed: $_"
        
        # ROLLBACK LOGIC
        try {
            Log-Info "Attempting rollback..."
            if ($renameSuccess) {
                Log-Info "Rolling back RENAME operation..."
                Rename-Item -Path $targetProfilePath -NewName $sourceProfilePath -Force -ErrorAction Stop
                Log-Info "Rollback successful: Restored $sourceProfilePath"
            }
            elseif ($sourceProfilePath -ne $targetProfilePath -and (Test-Path $targetProfilePath)) {
                Log-Info "Rolling back COPY operation: Removing target $targetProfilePath"
                Remove-FolderRobust -Path $targetProfilePath -ErrorAction SilentlyContinue
            }
        }
        catch {
            Log-Warning "Rollback failed: $_"
        }

        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}

function Convert-AzureADToLocal {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AzureADUsername,
        
        [Parameter(Mandatory = $true)]
        [string]$LocalUsername,
        
        [Parameter(Mandatory = $true)]
        [SecureString]$LocalPassword,
        
        [Parameter(Mandatory = $false)]
        [bool]$MakeAdmin = $false,
        
        [Parameter(Mandatory = $false)]
        [bool]$UnjoinAzureAD = $false
    )
    
    try {
        Log-Info "Starting AzureAD to Local profile conversion..."
        Log-Info "AzureAD user: $AzureADUsername"
        Log-Info "Target local user: $LocalUsername"
        
        # Get source SID (AzureAD user)
        Log-Info "Resolving source AzureAD user SID..."
        $sourceSID = Get-LocalUserSID -UserName $AzureADUsername
        
        # Verify it's an AzureAD SID
        if (-not (Test-IsAzureADSID $sourceSID)) {
            throw "Source SID ($sourceSID) is not an AzureAD SID"
        }
        
        Log-Info "Source SID: $sourceSID (AzureAD)"
        
        # Create local user if needed
        try {
            $existingUser = Get-LocalUser -Name $LocalUsername -ErrorAction Stop
            Log-Info "Local user '$LocalUsername' already exists"
        }
        catch {
            Log-Info "Creating new local user '$LocalUsername'..."
            New-LocalUser -Name $LocalUsername -Password $LocalPassword -FullName $LocalUsername -Description "Converted from AzureAD profile" -ErrorAction Stop
            Log-Info "Local user created successfully"
            
            if ($MakeAdmin) {
                Add-LocalGroupMember -Group "Administrators" -Member $LocalUsername -ErrorAction Stop
                Log-Info "User added to Administrators group"
            }
            
            Add-LocalGroupMember -Group "Users" -Member $LocalUsername -ErrorAction SilentlyContinue
        }
        
        # Get target SID
        $targetSID = Get-LocalUserSID -UserName $LocalUsername
        Log-Info "Target SID: $targetSID"
        
        Update-ConversionProgress -PercentComplete 35
        
        # Get profile paths
        $profileListBase = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
        $sourceProfileKey = Get-ItemProperty -Path "$profileListBase\$sourceSID" -ErrorAction Stop
        $sourceProfilePath = $sourceProfileKey.ProfileImagePath
        
        $targetProfileKey = Get-ItemProperty -Path "$profileListBase\$targetSID" -ErrorAction SilentlyContinue
        $targetProfilePath = if ($targetProfileKey) { $targetProfileKey.ProfileImagePath } else { "C:\Users\$LocalUsername" }
        
        Log-Info "Source: $sourceProfilePath -> Target: $targetProfilePath"
        
        # Verify source profile integrity
        if (-not (Test-ValidProfilePath -Path $sourceProfilePath -RequireNTUSER)) {
            throw "Source profile corrupted or missing (NTUSER.DAT not found): $sourceProfilePath"
        }
        
        # Transfer profile
        if ($global:CancelRequested) { throw "Operation cancelled by user" }
        Log-Info "Transferring profile data..."
        if ($sourceProfilePath -ne $targetProfilePath) {
            Update-ConversionProgress -PercentComplete 50
            
            $renameSuccess = $false
            if ([System.IO.Path]::GetPathRoot($sourceProfilePath) -eq [System.IO.Path]::GetPathRoot($targetProfilePath)) {
                try {
                    if (Test-Path $targetProfilePath) { Remove-Item -Path $targetProfilePath -Recurse -Force }
                    Rename-Item -Path $sourceProfilePath -NewName (Split-Path $targetProfilePath -Leaf) -ErrorAction Stop
                    $renameSuccess = $true
                    Log-Info "Profile renamed"
                }
                catch { Log-Warning "Rename failed, using robocopy" }
            }
            
            if (-not $renameSuccess) {
                if (-not (Test-Path $targetProfilePath)) { New-Item -ItemType Directory -Path $targetProfilePath -Force | Out-Null }
                $robocopyArgs = @("`"$sourceProfilePath`"", "`"$targetProfilePath`"", "/E", "/COPYALL", "/R:1", "/W:1", "/MT:$($Config.RobocopyThreads)", "/NFL", "/NDL", "/NJH", "/NJS", "/DCOPY:DAT")
                $robocopyProcess = Start-Process -FilePath "robocopy.exe" -ArgumentList $robocopyArgs -NoNewWindow -Wait -PassThru
                if ($robocopyProcess.ExitCode -ge 8) { throw "Robocopy failed: $($robocopyProcess.ExitCode)" }
                Log-Info "Profile copied"
                Update-ConversionProgress -PercentComplete 70
            }
        }
        
        # Update registry
        if ($global:CancelRequested) { throw "Operation cancelled by user" }
        $regResult = Update-ProfileListRegistry -OldSID $sourceSID -NewSID $targetSID -NewProfilePath $targetProfilePath
        if (-not $regResult.Success) { throw "Registry update failed: $($regResult.Error)" }
        
        # Apply ACLs
        Update-ConversionProgress -PercentComplete 80
        # Redundant call removed: Set-ProfileFolderAcls is called internally by Set-ProfileAcls
        
        # Step 8 & 9: Apply hive ACLs AND Rewrite SIDs (Unified via Set-ProfileAcls)
        # This wrapper handles UsrClass.dat reset and other critical fixes
        Log-Message "Applying hive ACLs and rewriting SIDs..."
        Set-ProfileAcls -ProfileFolder $targetProfilePath -UserSID $targetSID -UserName $LocalUsername -SourceSID $sourceSID -OldProfilePath $sourceProfilePath -NewProfilePath $targetProfilePath
        
        Update-ConversionProgress -PercentComplete 100 -StatusMessage "AzureAD to Local conversion completed!"
        
        Log-Info "AzureAD to Local conversion completed!"
        
        # Handle optional AzureAD unjoin
        if ($UnjoinAzureAD) {
            Log-Info "User requested AzureAD unjoin..."
            
            # Show warning dialog
            $message = "You are about to unjoin this device from AzureAD.`r`n`r`n"
            $message += "This will:`r`n"
            $message += "- Remove device from your organization's management`r`n"
            $message += "- Disable conditional access policies`r`n"
            $message += "- Remove SSO capabilities`r`n`r`n"
            $message += "This action requires a reboot to complete.`r`n`r`n"
            $message += "Continue with unjoin?"
            
            $response = Show-ModernDialog -Message $message -Title "Confirm AzureAD Unjoin" -Type Warning -Buttons YesNo
            
            if ($response -eq 'Yes') {
                $unjoinResult = Invoke-AzureADUnjoin
                
                if ($unjoinResult.Success) {
                    Log-Info "AzureAD unjoin successful"
                }
                else {
                    Log-Warn "AzureAD unjoin failed: $($unjoinResult.Message)"
                    # Don't fail the entire conversion, just warn
                    Show-ModernDialog -Message "Profile conversion succeeded, but AzureAD unjoin failed:`r`n`r`n$($unjoinResult.Message)`r`n`r`nYou can manually unjoin using: dsregcmd /leave" -Title "Unjoin Warning" -Type Warning -Buttons OK
                }
            }
            else {
                Log-Info "User cancelled AzureAD unjoin"
            }
        }
        
        
        # AppX re-registration handled in main Convert-UserProfile function
        
        return @{
            Success           = $true
            SourceSID         = $sourceSID
            TargetSID         = $targetSID
            SourceProfilePath = $sourceProfilePath
            TargetProfilePath = $targetProfilePath
        }
    }
    catch {
        Log-Error "AzureAD to Local conversion failed: $_"
        
        # ROLLBACK LOGIC
        try {
            Log-Info "Attempting rollback..."
            if ($renameSuccess) {
                Log-Info "Rolling back RENAME operation..."
                Rename-Item -Path $targetProfilePath -NewName $sourceProfilePath -Force -ErrorAction Stop
                Log-Info "Rollback successful: Restored $sourceProfilePath"
            }
            elseif ($sourceProfilePath -ne $targetProfilePath -and (Test-Path $targetProfilePath)) {
                Log-Info "Rolling back COPY operation: Removing target $targetProfilePath"
                Remove-FolderRobust -Path $targetProfilePath -ErrorAction SilentlyContinue
            }
        }
        catch {
            Log-Warning "Rollback failed: $_"
        }
        
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Convert-LocalToAzureAD {
    param(
        [Parameter(Mandatory = $true)][string]$LocalUsername,
        [Parameter(Mandatory = $true)][string]$AzureADUsername,
        
        [Parameter(Mandatory = $false)]
        [bool]$UnjoinDomain = $false
    )
    
    try {
        Log-Info "=== Starting Local to AzureAD Profile Conversion ==="
        Log-Info "Source: $LocalUsername (Local)"
        Log-Info "Target: $AzureADUsername (AzureAD)"
        
        # CRITICAL: Check if device is domain joined before AzureAD join
        # Windows does not allow AzureAD join while domain joined
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        if ($computerSystem.PartOfDomain) {
            Log-Warning "Device is domain joined - must unjoin before AzureAD join"
            
            $message = "This device is currently joined to domain: $($computerSystem.Domain)`r`n`r`n"
            $message += "Windows REQUIRES unjoining from the domain before joining AzureAD.`r`n`r`n"
            $message += "This will:`r`n"
            $message += "- Remove device from domain management`r`n"
            $message += "- Disable domain group policies`r`n"
            $message += "- Remove domain authentication`r`n`r`n"
            $message += "After unjoin, you will be prompted to join AzureAD.`r`n`r`n"
            $message += "Continue with required domain unjoin?"
            
            $response = Show-ModernDialog -Message $message -Title "Domain Unjoin Required" -Type Warning -Buttons YesNo
            
            if ($response -eq 'Yes') {
                Log-Info "User confirmed - executing domain unjoin..."
                $unjoinResult = Invoke-DomainUnjoin
                
                if ($unjoinResult.Success) {
                    Log-Info "Domain unjoin successful - can now proceed with AzureAD join"
                }
                else {
                    $errorMsg = "Domain unjoin failed: $($unjoinResult.Message)`r`n`r`n"
                    $errorMsg += "Cannot proceed with AzureAD join while domain joined."
                    Show-ModernDialog -Message $errorMsg -Title "Unjoin Failed" -Type Error -Buttons OK
                    throw "Domain unjoin failed - cannot join AzureAD"
                }
            }
            else {
                Log-Info "User cancelled domain unjoin"
                throw "User cancelled domain unjoin - cannot join AzureAD while domain joined"
            }
        }
        
        # Verify computer is AzureAD-joined
        if (-not (Test-IsAzureADJoined)) {
            Log-Warning "Computer is not AzureAD-joined"
            
            # Show guidance dialog
            $joinMsg = "This computer is not joined to AzureAD/Entra ID.`r`n`r`n"
            $joinMsg += "To convert to an AzureAD profile, you must first join this computer to AzureAD.`r`n`r`n"
            $joinMsg += "Steps to join AzureAD:`r`n"
            $joinMsg += "1. Open Settings (Windows + I)`r`n"
            $joinMsg += "2. Go to Accounts > Access work or school`r`n"
            $joinMsg += "3. Click 'Connect'`r`n"
            $joinMsg += "4. Select 'Join this device to Azure Active Directory'`r`n"
            $joinMsg += "5. Sign in with your work or school account`r`n"
            $joinMsg += "6. Complete the join process`r`n`r`n"
            $joinMsg += "After joining, restart this tool and try the conversion again.`r`n`r`n"
            $joinMsg += "Would you like to open Settings now?"
            
            $response = Show-ModernDialog -Message $joinMsg -Title "AzureAD Join Required" -Type Warning -Buttons YesNo
            
            if ($response -eq "Yes") {
                try {
                    # Check connectivity before proceeding
                    do {
                        if (Test-InternetConnectivity) { break }
                        
                        Log-Warning "No internet connectivity detected"
                        $connResponse = Show-ModernDialog -Message "No internet connection detected.`r`n`r`nAzureAD join requires an active internet connection.`r`n`r`nPlease connect to the internet and click 'Yes' to check again, or 'No' to cancel." -Title "No Internet Connection" -Type Warning -Buttons YesNo
                        
                        if ($connResponse -eq "No") {
                            throw "AzureAD join requires internet connectivity. Conversion cancelled."
                        }
                        Log-Info "User retrying connectivity check..."
                    } until ($false)

                    # Open Settings to the Access work or school page
                    Log-Info "Opening Settings for AzureAD join..."
                    Start-Process "ms-settings:workplace"
                    
                    # Wait a moment for Settings to open, then show dialog on top
                    Start-Sleep -Seconds 2
                    
                    # Show instruction dialog
                    Log-Info "Showing instruction dialog..."
                    $instructionResult = Show-ModernDialog -Message "Settings has been opened to the 'Access work or school' page.`r`n`r`nPlease complete the AzureAD join process in Settings.`r`n`r`nWhen you're done, close Settings and click OK here to continue." -Title "Complete AzureAD Join" -Type Info -Buttons OK
                    Log-Info "User clicked OK on instruction dialog"
                    
                    # Now check if AzureAD join succeeded
                    Log-Info "Checking if AzureAD join succeeded..."
                    if (Test-IsAzureADJoined) {
                        Log-Info "AzureAD join detected successfully!"
                        Show-ModernDialog -Message "AzureAD join detected!`r`n`r`nThe device is now joined to AzureAD.`r`n`r`nConversion will now proceed." -Title "Join Successful" -Type Success -Buttons OK
                    }
                    else {
                        # Join not detected - give user option to retry
                        Log-Warning "AzureAD join not detected"
                        $retryResponse = Show-ModernDialog -Message "AzureAD join was not detected.`r`n`r`nThis could mean:`r`n- You haven't completed the join yet`r`n- The join is still processing`r`n- You closed Settings without joining`r`n`r`nWould you like to wait 5 seconds and check again?`r`n`r`r`nClick 'Yes' to retry, or 'No' to cancel conversion." -Title "Join Not Detected" -Type Warning -Buttons YesNo
                        
                        if ($retryResponse -eq "Yes") {
                            Log-Info "User requested retry - waiting 5 seconds..."
                            Start-Sleep -Seconds 5
                            
                            if (Test-IsAzureADJoined) {
                                Log-Info "AzureAD join detected on retry!"
                                Show-ModernDialog -Message "AzureAD join now detected!`r`n`r`nConversion will proceed." -Title "Join Successful" -Type Success -Buttons OK
                            }
                            else {
                                Log-Error "AzureAD join still not detected after retry"
                                throw "Computer is not AzureAD-joined. Please complete the join process and try again."
                            }
                        }
                        else {
                            Log-Info "User cancelled - AzureAD join not completed"
                            throw "AzureAD join was not completed. Conversion cancelled."
                        }
                    }
                }
                catch {
                    Log-Error "Error in AzureAD join process: $_"
                    throw
                }
            }
            else {
                # User declined to open Settings
                Log-Info "User declined to open Settings for AzureAD join"
                throw "Computer is not AzureAD-joined. Please join to AzureAD first."
            }
        }
        Log-Info "AzureAD join status: VERIFIED"
        
        # Get source SID
        $sourceSID = Get-LocalUserSID -UserName $LocalUsername
        Log-Info "Source SID: $sourceSID"
        
        # Verify AzureAD user exists (Get-LocalUserSID now handles Graph API fallback)
        Log-Info "Verifying AzureAD user..."
        try {
            $targetSID = Get-LocalUserSID -UserName $AzureADUsername
            
            if (-not (Test-IsAzureADSID $targetSID)) {
                throw "Retrieved SID is not an AzureAD SID format"
            }
            
            Log-Info "AzureAD user SID: $targetSID"
        }
        catch {
            $errorMsg = "Failed to verify AzureAD user '$AzureADUsername'.`r`n`r`n"
            $errorMsg += "Error: $_`r`n`r`n"
            $errorMsg += "Please ensure:`r`n"
            $errorMsg += "- The username is correct (email format)`r`n"
            $errorMsg += "- You have permissions to query Microsoft Graph`r`n"
            $errorMsg += "- The user exists in your AzureAD tenant"
            
            Show-ModernDialog -Message $errorMsg -Title "AzureAD User Verification Failed" -Type Error -Buttons OK
            throw "AzureAD user verification failed"
        }
        
        if ($global:ConversionProgressBar) { $global:ConversionProgressBar.Value = 35; [System.Windows.Forms.Application]::DoEvents() }
        
        # Get profile paths
        $profileListBase = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
        $sourceProfileKey = Get-ItemProperty -Path "$profileListBase\$sourceSID" -ErrorAction Stop
        $sourceProfilePath = $sourceProfileKey.ProfileImagePath
        
        $targetProfileKey = Get-ItemProperty -Path "$profileListBase\$targetSID" -ErrorAction SilentlyContinue
        # Extract just the username part (before @) for the folder name
        $azureUsername = $AzureADUsername.Split('\')[1]
        $folderName = if ($azureUsername -match '^([^@]+)@') { $matches[1] } else { $azureUsername }
        $targetProfilePath = if ($targetProfileKey) { $targetProfileKey.ProfileImagePath } else { "C:\Users\$folderName" }
        
        Log-Info "Source: $sourceProfilePath -> Target: $targetProfilePath"
        
        # Check if source and target paths are the same
        if ($sourceProfilePath -eq $targetProfilePath) {
            Log-Info "Source and target paths are identical - skipping file operations, will only update registry"
            # No file operations needed, just update registry later
        }
        else {
            if ($global:CancelRequested) { throw "Operation cancelled by user" }
            # Paths are different - need to transfer files
            # Check if target already has data
            if (Test-Path $targetProfilePath) {
                $targetHive = Join-Path $targetProfilePath "NTUSER.DAT"
                if (Test-Path $targetHive) {
                    $response = Show-ModernDialog -Message "The AzureAD user already has a profile.`r`n`r`nOverwrite it?`r`n`r`nWARNING: This will delete the existing AzureAD profile!" -Title "Overwrite Profile" -Type Warning -Buttons YesNo
                    if ($response -ne "Yes") { throw "User cancelled" }
                    Remove-Item -Path $targetProfilePath -Recurse -Force
                }
            }
            
            # Transfer profile
            if ($global:ConversionProgressBar) { $global:ConversionProgressBar.Value = 50; [System.Windows.Forms.Application]::DoEvents() }
            
            $renameSuccess = $false
            if ([System.IO.Path]::GetPathRoot($sourceProfilePath) -eq [System.IO.Path]::GetPathRoot($targetProfilePath)) {
                try {
                    Rename-Item -Path $sourceProfilePath -NewName (Split-Path $targetProfilePath -Leaf) -ErrorAction Stop
                    $renameSuccess = $true
                    Log-Info "Profile renamed"
                }
                catch { Log-Warning "Rename failed, using robocopy" }
            }
            
            if (-not $renameSuccess) {
                if (-not (Test-Path $targetProfilePath)) { New-Item -ItemType Directory -Path $targetProfilePath -Force | Out-Null }
                $robocopyArgs = @("`"$sourceProfilePath`"", "`"$targetProfilePath`"", "/E", "/COPYALL", "/R:1", "/W:1", "/MT:$($Config.RobocopyThreads)", "/NFL", "/NDL", "/NJH", "/NJS", "/DCOPY:DAT")
                $robocopyProcess = Start-Process -FilePath "robocopy.exe" -ArgumentList $robocopyArgs -NoNewWindow -Wait -PassThru
                if ($robocopyProcess.ExitCode -ge 8) { throw "Robocopy failed: $($robocopyProcess.ExitCode)" }
                Log-Info "Profile copied"
                if ($global:ConversionProgressBar) { $global:ConversionProgressBar.Value = 70; [System.Windows.Forms.Application]::DoEvents() }
            }
        }
        
        # Update registry
        if ($global:CancelRequested) { throw "Operation cancelled by user" }
        $regResult = Update-ProfileListRegistry -OldSID $sourceSID -NewSID $targetSID -NewProfilePath $targetProfilePath
        if (-not $regResult.Success) { throw "Registry update failed: $($regResult.Error)" }
        
        # Step 8 & 9: Apply hive ACLs AND Rewrite SIDs (Unified via Set-ProfileAcls)
        # This wrapper handles UsrClass.dat reset and other critical fixes
        Log-Message "Applying hive ACLs and rewriting SIDs..."
        Set-ProfileAcls -ProfileFolder $targetProfilePath -UserSID $targetSID -UserName $AzureADUsername -SourceSID $sourceSID -OldProfilePath $sourceProfilePath -NewProfilePath $targetProfilePath
        
        if ($global:ConversionProgressBar) { $global:ConversionProgressBar.Value = 100; [System.Windows.Forms.Application]::DoEvents() }
        
        Log-Info "Local to AzureAD conversion completed!"
        
        return @{
            Success           = $true
            SourceSID         = $sourceSID
            TargetSID         = $targetSID
            SourceProfilePath = $sourceProfilePath
            TargetProfilePath = $targetProfilePath
        }
    }
    catch {
        Log-Error "Local to AzureAD conversion failed: $_"
        
        # ROLLBACK LOGIC
        try {
            Log-Info "Attempting rollback..."
            if ($renameSuccess) {
                Log-Info "Rolling back RENAME operation..."
                Rename-Item -Path $targetProfilePath -NewName $sourceProfilePath -Force -ErrorAction Stop
                Log-Info "Rollback successful: Restored $sourceProfilePath"
            }
            elseif ($sourceProfilePath -ne $targetProfilePath -and (Test-Path $targetProfilePath)) {
                Log-Info "Rolling back COPY operation: Removing target $targetProfilePath"
                Remove-FolderRobust -Path $targetProfilePath -ErrorAction SilentlyContinue
            }
        }
        catch {
            Log-Warning "Rollback failed: $_"
        }
        
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

# Show Profile Conversion Dialog
function Show-ProfileConversionDialog {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    # Apply theme
    $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
    $theme = $Themes[$currentTheme]
    
    $convForm = New-Object System.Windows.Forms.Form
    $convForm.Text = "Convert Profile Type"
    $convForm.Size = New-Object System.Drawing.Size(650, 640)
    $convForm.StartPosition = "CenterScreen"
    $convForm.FormBorderStyle = "FixedDialog"
    $convForm.MaximizeBox = $false
    $convForm.MinimizeBox = $false
    $convForm.TopMost = $true
    $convForm.BackColor = $theme.FormBackColor
    $convForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Header panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(650, 70)
    $headerPanel.BackColor = $theme.HeaderBackColor
    $convForm.Controls.Add($headerPanel)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(610, 25)
    $lblTitle.Text = "Convert Profile Type"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $theme.HeaderTextColor
    $headerPanel.Controls.Add($lblTitle)
    
    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Location = New-Object System.Drawing.Point(22, 43)
    $lblSubtitle.Size = New-Object System.Drawing.Size(610, 20)
    $lblSubtitle.Text = "Convert an existing profile between Local and Domain types"
    $lblSubtitle.ForeColor = $theme.SubHeaderTextColor
    $headerPanel.Controls.Add($lblSubtitle)
    
    # Content panel
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Location = New-Object System.Drawing.Point(15, 85)
    $contentPanel.Size = New-Object System.Drawing.Size(610, 460)
    $contentPanel.BackColor = $theme.PanelBackColor
    $convForm.Controls.Add($contentPanel)
    
    # Source Profile Section
    $lblSource = New-Object System.Windows.Forms.Label
    $lblSource.Location = New-Object System.Drawing.Point(15, 15)
    $lblSource.Size = New-Object System.Drawing.Size(580, 20)
    $lblSource.Text = "Source Profile:"
    $lblSource.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblSource.ForeColor = $theme.LabelTextColor
    $contentPanel.Controls.Add($lblSource)
    
    $cmbSourceProfile = New-Object System.Windows.Forms.ComboBox
    $cmbSourceProfile.Location = New-Object System.Drawing.Point(15, 40)
    $cmbSourceProfile.Size = New-Object System.Drawing.Size(400, 25)
    $cmbSourceProfile.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbSourceProfile.BackColor = $theme.TextBoxBackColor
    $cmbSourceProfile.ForeColor = $theme.TextBoxForeColor
    $contentPanel.Controls.Add($cmbSourceProfile)
    
    # Populate source profiles
    try {
        $profiles = Get-ProfileDisplayEntries
        foreach ($p in $profiles) {
            $cmbSourceProfile.Items.Add($p.DisplayName) | Out-Null
        }
    }
    catch {
        Log-Error "Failed to load profiles: $_"
    }
    
    # Current Type Label
    $lblCurrentType = New-Object System.Windows.Forms.Label
    $lblCurrentType.Location = New-Object System.Drawing.Point(425, 43)
    $lblCurrentType.Size = New-Object System.Drawing.Size(170, 20)
    $lblCurrentType.Text = "Type: Unknown"
    $lblCurrentType.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblCurrentType.ForeColor = $theme.SubHeaderTextColor
    $contentPanel.Controls.Add($lblCurrentType)
    
    # Target Type Section
    $lblTarget = New-Object System.Windows.Forms.Label
    $lblTarget.Location = New-Object System.Drawing.Point(15, 80)
    $lblTarget.Size = New-Object System.Drawing.Size(580, 20)
    $lblTarget.Text = "Convert To:"
    $lblTarget.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblTarget.ForeColor = $theme.LabelTextColor
    $contentPanel.Controls.Add($lblTarget)
    
    # Radio buttons for target type
    $rbLocal = New-Object System.Windows.Forms.RadioButton
    $rbLocal.Location = New-Object System.Drawing.Point(30, 110)
    $rbLocal.Size = New-Object System.Drawing.Size(150, 25)
    $rbLocal.Text = "Local Profile"
    $rbLocal.ForeColor = $theme.LabelTextColor
    $rbLocal.Checked = $true
    $contentPanel.Controls.Add($rbLocal)
    
    $rbDomain = New-Object System.Windows.Forms.RadioButton
    $rbDomain.Location = New-Object System.Drawing.Point(200, 110)
    $rbDomain.Size = New-Object System.Drawing.Size(150, 25)
    $rbDomain.Text = "Domain Profile"
    $rbDomain.ForeColor = $theme.LabelTextColor
    $contentPanel.Controls.Add($rbDomain)
    
    $rbAzureAD = New-Object System.Windows.Forms.RadioButton
    $rbAzureAD.Location = New-Object System.Drawing.Point(370, 110)
    $rbAzureAD.Size = New-Object System.Drawing.Size(180, 25)
    $rbAzureAD.Text = "AzureAD Profile"
    $rbAzureAD.ForeColor = $theme.LabelTextColor
    $contentPanel.Controls.Add($rbAzureAD)
    
    # AzureAD username format hint (initially hidden)
    $lblAzureADHint = New-Object System.Windows.Forms.Label
    $lblAzureADHint.Location = New-Object System.Drawing.Point(15, 205)
    $lblAzureADHint.Size = New-Object System.Drawing.Size(550, 20)
    $lblAzureADHint.Text = "Note: For AzureAD, enter username in email format (e.g., user@domain.com)"
    $lblAzureADHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblAzureADHint.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)  # Blue info color
    $lblAzureADHint.Visible = $false
    $contentPanel.Controls.Add($lblAzureADHint)
    
    # Unjoin from AzureAD checkbox (only visible when source is AzureAD)
    $chkUnjoinAzureAD = New-Object System.Windows.Forms.CheckBox
    $chkUnjoinAzureAD.Location = New-Object System.Drawing.Point(430, 163)
    $chkUnjoinAzureAD.Size = New-Object System.Drawing.Size(180, 50)
    $chkUnjoinAzureAD.Text = "Unjoin from AzureAD after conversion"
    $chkUnjoinAzureAD.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $chkUnjoinAzureAD.ForeColor = $theme.LabelTextColor
    $chkUnjoinAzureAD.Visible = $false  # Only show when source is AzureAD
    $contentPanel.Controls.Add($chkUnjoinAzureAD)
    
    # Event handlers to show/hide password fields based on target type
    $rbLocal.Add_CheckedChanged({
            if ($rbLocal.Checked) {
                $lblPassword.Visible = $true
                $txtPassword.Visible = $true
                $txtPasswordConfirm.Visible = $true
                $chkShowPassword.Visible = $true
                $lblPasswordStrength.Visible = $true
                $chkMakeAdmin.Visible = $true
                $lblAzureADHint.Visible = $false
            }
        })
    
    $rbDomain.Add_CheckedChanged({
            if ($rbDomain.Checked) {
                $lblPassword.Visible = $false
                $txtPassword.Visible = $false
                $txtPasswordConfirm.Visible = $false
                $chkShowPassword.Visible = $false
                $lblPasswordStrength.Visible = $false
                $chkMakeAdmin.Visible = $false
                $lblAzureADHint.Visible = $false
            }
        })
    
    $rbAzureAD.Add_CheckedChanged({
            if ($rbAzureAD.Checked) {
                $lblPassword.Visible = $false
                $txtPassword.Visible = $false
                $txtPasswordConfirm.Visible = $false
                $chkShowPassword.Visible = $false
                $lblPasswordStrength.Visible = $false
                $chkMakeAdmin.Visible = $false
                $lblAzureADHint.Visible = $true
            }
        })
    
    # Target Username
    $lblTargetUser = New-Object System.Windows.Forms.Label
    $lblTargetUser.Location = New-Object System.Drawing.Point(15, 150)
    $lblTargetUser.Size = New-Object System.Drawing.Size(580, 20)
    $lblTargetUser.Text = "Target Username:"
    $lblTargetUser.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblTargetUser.ForeColor = $theme.LabelTextColor
    $contentPanel.Controls.Add($lblTargetUser)
    
    $txtTargetUser = New-Object System.Windows.Forms.TextBox
    $txtTargetUser.Location = New-Object System.Drawing.Point(15, 175)
    $txtTargetUser.Size = New-Object System.Drawing.Size(400, 25)
    $txtTargetUser.BackColor = $theme.TextBoxBackColor
    $txtTargetUser.ForeColor = $theme.TextBoxForeColor
    $contentPanel.Controls.Add($txtTargetUser)
    
    # Password fields (for local user creation)
    $lblPassword = New-Object System.Windows.Forms.Label
    $lblPassword.Location = New-Object System.Drawing.Point(15, 215)
    $lblPassword.Size = New-Object System.Drawing.Size(580, 20)
    $lblPassword.Text = "Password (for new local user):"
    $lblPassword.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblPassword.ForeColor = $theme.LabelTextColor
    $lblPassword.Visible = $true
    $contentPanel.Controls.Add($lblPassword)
    
    $txtPassword = New-Object System.Windows.Forms.TextBox
    $txtPassword.Location = New-Object System.Drawing.Point(15, 240)
    $txtPassword.Size = New-Object System.Drawing.Size(280, 25)
    $txtPassword.UseSystemPasswordChar = $true
    $txtPassword.BackColor = $theme.TextBoxBackColor
    $txtPassword.ForeColor = $theme.TextBoxForeColor
    $txtPassword.Visible = $true
    $contentPanel.Controls.Add($txtPassword)
    
    $txtPasswordConfirm = New-Object System.Windows.Forms.TextBox
    $txtPasswordConfirm.Location = New-Object System.Drawing.Point(315, 240)
    $txtPasswordConfirm.Size = New-Object System.Drawing.Size(280, 25)
    $txtPasswordConfirm.UseSystemPasswordChar = $true
    $txtPasswordConfirm.BackColor = $theme.TextBoxBackColor
    $txtPasswordConfirm.ForeColor = $theme.TextBoxForeColor
    $txtPasswordConfirm.Visible = $true
    $contentPanel.Controls.Add($txtPasswordConfirm)
    
    # Show Password checkbox
    $chkShowPassword = New-Object System.Windows.Forms.CheckBox
    $chkShowPassword.Location = New-Object System.Drawing.Point(15, 270)
    $chkShowPassword.Size = New-Object System.Drawing.Size(150, 25)
    $chkShowPassword.Text = "Show passwords"
    $chkShowPassword.ForeColor = $theme.LabelTextColor
    $chkShowPassword.Visible = $true
    $chkShowPassword.Add_CheckedChanged({
            $txtPassword.UseSystemPasswordChar = -not $chkShowPassword.Checked
            $txtPasswordConfirm.UseSystemPasswordChar = -not $chkShowPassword.Checked
        })
    $contentPanel.Controls.Add($chkShowPassword)
    
    # Password strength indicator
    $lblPasswordStrength = New-Object System.Windows.Forms.Label
    $lblPasswordStrength.Location = New-Object System.Drawing.Point(180, 273)
    $lblPasswordStrength.Size = New-Object System.Drawing.Size(415, 20)
    $lblPasswordStrength.Text = ""
    $lblPasswordStrength.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblPasswordStrength.Visible = $true
    $contentPanel.Controls.Add($lblPasswordStrength)
    
    # Password strength checker
    $txtPassword.Add_TextChanged({
            $len = $txtPassword.Text.Length
            if ($len -eq 0) {
                $lblPasswordStrength.Text = ""
                $lblPasswordStrength.ForeColor = $theme.LabelTextColor
            }
            elseif ($len -lt 8) {
                $lblPasswordStrength.Text = "Weak password (< 8 characters)"
                $lblPasswordStrength.ForeColor = [System.Drawing.Color]::OrangeRed
            }
            elseif ($len -lt 12) {
                $lblPasswordStrength.Text = "Moderate password"
                $lblPasswordStrength.ForeColor = [System.Drawing.Color]::Orange
            }
            else {
                $lblPasswordStrength.Text = "Strong password"
                $lblPasswordStrength.ForeColor = [System.Drawing.Color]::Green
            }
        })
    
    # Make administrator checkbox (for local user creation)
    $chkMakeAdmin = New-Object System.Windows.Forms.CheckBox
    $chkMakeAdmin.Location = New-Object System.Drawing.Point(315, 305)
    $chkMakeAdmin.Size = New-Object System.Drawing.Size(280, 25)
    $chkMakeAdmin.Text = "Make new local user an administrator"
    $chkMakeAdmin.Checked = $false
    $chkMakeAdmin.ForeColor = $theme.LabelTextColor
    $chkMakeAdmin.Visible = $true
    $contentPanel.Controls.Add($chkMakeAdmin)
    
    # Backup checkbox
    $chkBackup = New-Object System.Windows.Forms.CheckBox
    $chkBackup.Location = New-Object System.Drawing.Point(15, 305)
    $chkBackup.Size = New-Object System.Drawing.Size(580, 25)
    $chkBackup.Text = "Create backup before conversion (Recommended)"
    $chkBackup.Checked = $true
    $chkBackup.ForeColor = $theme.LabelTextColor
    $contentPanel.Controls.Add($chkBackup)
    
    # Status/Warning panel
    $statusPanel = New-Object System.Windows.Forms.Panel
    $statusPanel.Location = New-Object System.Drawing.Point(15, 340)
    $statusPanel.Size = New-Object System.Drawing.Size(580, 70)
    $statusPanel.BackColor = $theme.LogBoxBackColor
    $statusPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $contentPanel.Controls.Add($statusPanel)

    # Progress Bar
    $convProgressBar = New-Object System.Windows.Forms.ProgressBar
    $convProgressBar.Location = New-Object System.Drawing.Point(15, 425)
    $convProgressBar.Size = New-Object System.Drawing.Size(580, 20)
    $convProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $contentPanel.Controls.Add($convProgressBar)
    
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(10, 10)
    $lblStatus.Size = New-Object System.Drawing.Size(560, 50)
    $lblStatus.Text = "Select a source profile to begin"
    $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblStatus.ForeColor = $theme.SubHeaderTextColor
    $statusPanel.Controls.Add($lblStatus)
    
    # Event handlers
    $cmbSourceProfile.Add_SelectedIndexChanged({
            try {
                $selectedProfile = $cmbSourceProfile.SelectedItem
                if ($selectedProfile) {
                    # Extract username from display name
                    $username = $selectedProfile
                    if ($username -match '^(.+?)\s+-\s+\[.+\]$') {
                        $username = $matches[1]
                    }

                    # --- PROACTIVE CHECK: USER LOGGED IN? ---
                    if (-not (Invoke-ProactiveUserCheck -Username $username)) {
                        # Reset selection if user is logged in and declines logoff
                        $cmbSourceProfile.SelectedIndex = -1
                        $lblStatus.Text = "Selection cancelled - user is logged in"
                        return
                    }
                    # ----------------------------------------
                
                    # Get profile type
                    $profileType = Get-ProfileType -Username $username
                    $lblCurrentType.Text = "Type: $profileType"
                
                    # Show/hide unjoin checkbox based on source type
                    if ($profileType -eq "AzureAD") {
                        $chkUnjoinAzureAD.Visible = $true
                        $chkUnjoinAzureAD.Text = "Unjoin from AzureAD after conversion"
                    }
                    elseif ($profileType -eq "Domain") {
                        $chkUnjoinAzureAD.Visible = $true
                        $chkUnjoinAzureAD.Text = "Unjoin from Domain after conversion"
                    }
                    else {
                        $chkUnjoinAzureAD.Visible = $false
                        $chkUnjoinAzureAD.Checked = $false
                    }
                
                    # Update status
                    $lblStatus.Text = "Ready to convert $profileType profile: $username"
                    $lblStatus.ForeColor = $theme.LabelTextColor
                
                    # Auto-fill target username (strip domain if present)
                    if ($username -match '\\(.+)$') {
                        $txtTargetUser.Text = $matches[1]
                    }
                    else {
                        $txtTargetUser.Text = $username
                    }
                }
            }
            catch {
                $lblStatus.Text = "Error detecting profile type: $_"
                $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
            }
        })
    
    # Show/hide password fields based on target type
    $rbLocal.Add_CheckedChanged({
            if ($rbLocal.Checked) {
                $lblPassword.Visible = $true
                $txtPassword.Visible = $true
                $txtPasswordConfirm.Visible = $true
                $chkMakeAdmin.Visible = $true
            }
        })
    
    $rbDomain.Add_CheckedChanged({
            if ($rbDomain.Checked) {
                $lblPassword.Visible = $false
                $txtPassword.Visible = $false
                $txtPasswordConfirm.Visible = $false
                $chkMakeAdmin.Visible = $false
            }
        })
    
    $rbAzureAD.Add_CheckedChanged({
            if ($rbAzureAD.Checked) {
                $lblPassword.Visible = $false
                $txtPassword.Visible = $false
                $txtPasswordConfirm.Visible = $false
                $chkMakeAdmin.Visible = $false
            }
        })
    
    # Buttons
    $btnConvert = New-Object System.Windows.Forms.Button
    $btnConvert.Location = New-Object System.Drawing.Point(360, 560)
    $btnConvert.Size = New-Object System.Drawing.Size(120, 35)
    $btnConvert.Text = "Convert"
    $btnConvert.BackColor = $theme.ButtonPrimaryBackColor
    $btnConvert.ForeColor = $theme.ButtonPrimaryForeColor
    $btnConvert.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnConvert.FlatAppearance.BorderSize = 0
    $btnConvert.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnConvert.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnConvert.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
    $btnConvert.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
    $convForm.Controls.Add($btnConvert)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(490, 560)
    $btnCancel.Size = New-Object System.Drawing.Size(120, 35)
    $btnCancel.Text = "Cancel"
    $btnCancel.BackColor = $theme.ButtonSecondaryBackColor
    $btnCancel.ForeColor = $theme.ButtonSecondaryForeColor
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderSize = 0
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
    $btnCancel.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
    $convForm.Controls.Add($btnCancel)
    
    # Convert button click handler
    $btnConvert.Add_Click({
            try {
                # Validation
                if (-not $cmbSourceProfile.SelectedItem) {
                    Show-ModernDialog -Message "Please select a source profile" -Title "Validation Error" -Type Warning -Buttons OK
                    return
                }
            
                if (-not $txtTargetUser.Text) {
                    Show-ModernDialog -Message "Please enter a target username" -Title "Validation Error" -Type Warning -Buttons OK
                    return
                }
            
                # Extract source username
                $sourceUsername = $cmbSourceProfile.SelectedItem
                if ($sourceUsername -match '^(.+?)\s+-\s+\[.+\]$') {
                    $sourceUsername = $matches[1]
                }
                
                # --- PROACTIVE CHECK: USER LOGGED IN? ---
                if (-not (Invoke-ProactiveUserCheck -Username $sourceUsername)) {
                    return
                }
                # ----------------------------------------
            
                $targetUsername = $txtTargetUser.Text.Trim()
                $targetType = if ($rbLocal.Checked) { 
                    "Local" 
                }
                elseif ($rbDomain.Checked) { 
                    "Domain" 
                }
                else { 
                    "AzureAD" 
                }
            
                # Password validation for local conversion
                if ($rbLocal.Checked) {
                    if (-not $txtPassword.Text) {
                        Show-ModernDialog -Message "Please enter a password for the new local user" -Title "Validation Error" -Type Warning -Buttons OK
                        return
                    }
                
                    if ($txtPassword.Text -ne $txtPasswordConfirm.Text) {
                        Show-ModernDialog -Message "Passwords do not match" -Title "Validation Error" -Type Warning -Buttons OK
                        return
                    }
                }
            
                # Get source profile type
                $sourceType = Get-ProfileType -Username $sourceUsername
            
                # (Premature check removed) - Logic is handled in dispatch block


                # Create log file for this conversion (Early initialization to capture domain join)
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                # Sanitize usernames for filename (remove domain/computer prefix and invalid chars)
                $sourceUserSafe = ($sourceUsername -split '\\')[-1] -replace '[\\/:*?"<>|]', '_'
                $targetUserSafe = ($targetUsername -split '\\')[-1] -replace '[\\/:*?"<>|]', '_'
                $logFileName = "Conversion_${sourceUserSafe}_to_${targetType}_${timestamp}.log"
                $logsDir = Join-Path $PSScriptRoot "Logs"
                if (-not (Test-Path $logsDir)) {
                    New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
                }
                $conversionLogPath = Join-Path $logsDir $logFileName
                $global:ConversionLogPath = $conversionLogPath
                $conversionStartTime = Get-Date
            
                # Initialize log file
                $logHeader = @"
=============================================================================
PROFILE TYPE CONVERSION LOG
=============================================================================
Conversion Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Source User: $sourceUsername
Source Type: $sourceType
Target User: $targetUsername
Target Type: $targetType
Computer: $env:COMPUTERNAME
Script Version: $($Config.Version)
=============================================================================

"@
                $logHeader | Out-File -FilePath $conversionLogPath -Encoding UTF8
            
                # Run precondition checks
                $lblStatus.Text = "Checking preconditions..."
                $lblStatus.ForeColor = $theme.SubHeaderTextColor
                $convProgressBar.Value = 5
                [System.Windows.Forms.Application]::DoEvents()
            
                # Set global progress bar reference for called functions
                $global:ConversionProgressBar = $convProgressBar
                $global:ConversionStatusLabel = $lblStatus
            
                $preconditions = Test-ProfileConversionPreconditions -SourceUsername $sourceUsername -TargetType $targetType
            
                if (-not $preconditions.Success) {
                    # Check if the only error is domain join requirement
                    $domainJoinError = $preconditions.Errors | Where-Object { $_ -like "*not joined to a domain*" }
                
                    if ($domainJoinError -and $preconditions.Errors.Count -eq 1) {
                        # Try to extract NetBIOS domain name from target username (e.g., "CONTOSO\user")
                        $netbiosName = $null
                        $defaultFqdn = ""
                    
                        if ($targetUsername -match '^(.+?)\\') {
                            $netbiosName = $matches[1]
                            # Use Get-DomainFQDN to intelligently resolve FQDN
                            $defaultFqdn = Get-DomainFQDN -NetBIOSName $netbiosName
                        }
                    
                        # Prompt for FQDN (always needed for domain join)
                        $promptMessage = if ($netbiosName) {
                            "Enter the full domain name (FQDN) for domain: $netbiosName"
                        }
                        else {
                            "Enter the domain name to join this computer"
                        }
                    
                        $domainInput = Show-InputDialog -Title "Domain Name Required" -Message $promptMessage -DefaultValue $defaultFqdn -ExampleText "Example: contoso.local or contoso.com"
                    
                        if ($domainInput.Result -ne [System.Windows.Forms.DialogResult]::OK -or [string]::IsNullOrWhiteSpace($domainInput.Value)) {
                            Show-ModernDialog -Message "Domain name is required to join the domain." -Title "Cancelled" -Type Info -Buttons OK
                            return
                        }
                    
                        $domainNamePrompt = $domainInput.Value
                    
                        # Show custom themed credential dialog using shared helper
                        $domainCred = Get-DomainCredential -Domain $domainNamePrompt
                        
                        if (-not $domainCred) {
                            Show-ModernDialog -Message "Domain credentials are required to join the domain." -Title "Cancelled" -Type Info -Buttons OK
                            return
                        }
                        
                        $username = $domainCred.UserName
                        # Remove domain prefix for display if present
                        if ($username -match '\\') { $username = ($username -split '\\')[1] }
                    
                        # Validate target domain user exists in AD before domain join
                        Log-Info "Validating target domain user exists in Active Directory..."
                        try {
                            # Extract short username from DOMAIN\username format
                            $shortTargetName = if ($targetUsername -match '\\(.+)$') { $matches[1] } else { $targetUsername }
                            
                            # Use DirectoryServices with domain credentials to query AD (works before domain join)
                            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                            $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('Domain', $domainNamePrompt, $domainCred.UserName, $domainCred.GetNetworkCredential().Password)
                            $targetUser = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($ctx, $shortTargetName)
                            
                            if (-not $targetUser) {
                                throw "User '$shortTargetName' not found in domain '$domainNamePrompt'"
                            }
                            
                            $targetUserSID = $targetUser.Sid.Value
                            Log-Info "Target user validated successfully (SID: $targetUserSID)"
                            $ctx.Dispose()
                        }
                        catch {
                            Log-Warning "Could not validate target domain user: $_"
                            
                            # Show validation error with option to continue anyway
                            $validationMsg = "Warning: Could not verify that user '$targetUsername' exists in Active Directory.`r`n`r`n"
                            $validationMsg += "This could mean:`r`n"
                            $validationMsg += "1. The username is incorrect (typo)`r`n"
                            $validationMsg += "2. The user doesn't exist in Active Directory`r`n"
                            $validationMsg += "3. Temporary AD connectivity issue`r`n`r`n"
                            $validationMsg += "If you continue, the domain join will proceed, but the conversion will fail if the username is incorrect.`r`n`r`n"
                            $validationMsg += "Do you want to:`r`n"
                            $validationMsg += "- Click 'Yes' to continue anyway (if you're sure the username is correct)`r`n"
                            $validationMsg += "- Click 'No' to cancel and fix the username"
                            
                            $validationResponse = Show-ModernDialog -Message $validationMsg -Title "User Validation Warning" -Type Warning -Buttons YesNo
                            
                            if ($validationResponse -ne "Yes") {
                                Log-Info "User cancelled domain join due to validation warning"
                                Show-ModernDialog -Message "Domain join cancelled. Please verify the target username and try again." -Title "Cancelled" -Type Info -Buttons OK
                                return
                            }
                            
                            Log-Info "User chose to continue despite validation warning"
                        }
                    
                        # Confirm domain join

                        $joinMsg = "This computer is not currently joined to a domain.`r`n`r`n"
                        $joinMsg += "Domain Join Details:`r`n"
                        $joinMsg += "  - Domain: $domainNamePrompt`r`n"
                        $joinMsg += "  - Admin User: $username`r`n`r`n"
                        $joinMsg += "The computer will be joined to the domain and the`r`n"
                        $joinMsg += "conversion will continue automatically.`r`n`r`n"
                        $joinMsg += "Note: A restart will be required later for full`r`n"
                        $joinMsg += "domain integration.`r`n`r`n"
                        $joinMsg += "Do you want to proceed with the domain join?"
                    
                        $response = Show-ModernDialog -Message $joinMsg -Title "Confirm Domain Join" -Type Question -Buttons YesNo
                    
                        if ($response -eq "Yes") {
                            try {
                                # Store credentials globally for potential reuse
                                $global:DomainCredential = $domainCred
                            
                                Log-Info "Attempting to join domain: $domainNamePrompt"
                                $lblStatus.Text = "Joining domain..."
                                $lblStatus.ForeColor = $theme.SubHeaderTextColor
                                [System.Windows.Forms.Application]::DoEvents()
                            
                                # Call Join-Domain-Enhanced function with 'Never' restart behavior
                                Join-Domain-Enhanced -DomainName $domainNamePrompt -Credential $domainCred -RestartBehavior 'Never'
                            
                                # Verify domain join was successful
                                $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
                                if ($computerSystem.PartOfDomain) {
                                    Log-Info "Domain join successful! Continuing with conversion..."
                                    $lblStatus.Text = "Domain joined successfully - continuing conversion..."
                                    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
                                    [System.Windows.Forms.Application]::DoEvents()
                                
                                    # Don't return - let the conversion continue below
                                    # The precondition check will pass now that we're domain-joined
                                }
                                else {
                                    Show-ModernDialog -Message "Domain join process completed but computer is not yet domain-joined.`r`n`r`nPlease restart and try again." -Title "Restart Required" -Type Warning -Buttons OK
                                    return
                                }
                            }
                            catch {
                                Log-Error "Domain join failed: $_"
                                Show-ModernDialog -Message "Domain join failed:`r`n`r`n$_`r`n`r`nPlease check your domain name and credentials." -Title "Domain Join Failed" -Type Error -Buttons OK
                                return
                            }
                        }
                    }
                    # Check if the only error is AzureAD join requirement
                    elseif ($preconditions.Errors.Count -eq 1 -and $preconditions.Errors[0] -like "*not joined to AzureAD*") {
                        Log-Warning "AzureAD join required for conversion"
                    
                        # CRITICAL: Check if device is domain joined first
                        # Must unjoin from domain before joining AzureAD
                        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
                        if ($computerSystem.PartOfDomain) {
                            Log-Warning "Device is domain joined - must unjoin before AzureAD join"
                        
                            $message = "This device is currently joined to domain: $($computerSystem.Domain)`r`n`r`n"
                            $message += "Windows REQUIRES unjoining from the domain before joining AzureAD.`r`n`r`n"
                            $message += "This will:`r`n"
                            $message += "- Remove device from domain management`r`n"
                            $message += "- Disable domain group policies`r`n"
                            $message += "- Remove domain authentication`r`n`r`n"
                            $message += "After unjoin, you will be prompted to join AzureAD.`r`n`r`n"
                            $message += "Continue with required domain unjoin?"
                        
                            $response = Show-ModernDialog -Message $message -Title "Domain Unjoin Required" -Type Warning -Buttons YesNo
                        
                            if ($response -eq 'Yes') {
                                Log-Info "User confirmed - executing domain unjoin..."
                                $unjoinResult = Invoke-DomainUnjoin
                            
                                if (-not $unjoinResult.Success) {
                                    $errorMsg = "Domain unjoin failed: $($unjoinResult.Message)`r`n`r`n"
                                    $errorMsg += "Cannot proceed with AzureAD join while domain joined."
                                    Show-ModernDialog -Message $errorMsg -Title "Unjoin Failed" -Type Error -Buttons OK
                                    return
                                }
                            
                                Log-Info "Domain unjoin successful - can now proceed with AzureAD join"
                            }
                            else {
                                Log-Info "User cancelled domain unjoin"
                                return
                            }
                        }
                    
                        # Show guidance dialog for AzureAD join
                        $joinMsg = "This computer is not joined to AzureAD/Entra ID.`r`n`r`n"
                        $joinMsg += "To convert to an AzureAD profile, you must first join this computer to AzureAD.`r`n`r`n"
                        $joinMsg += "Steps to join AzureAD:`r`n"
                        $joinMsg += "1. Open Settings (Windows + I)`r`n"
                        $joinMsg += "2. Go to Accounts > Access work or school`r`n"
                        $joinMsg += "3. Click 'Connect'`r`n"
                        $joinMsg += "4. Select 'Join this device to Azure Active Directory'`r`n"
                        $joinMsg += "5. Sign in with your work or school account`r`n"
                        $joinMsg += "6. Complete the join process`r`n`r`n"
                        $joinMsg += "After joining, restart this tool and try the conversion again.`r`n`r`n"
                        $joinMsg += "Would you like to open Settings now?"
                    
                        $response = Show-ModernDialog -Message $joinMsg -Title "AzureAD Join Required" -Type Warning -Buttons YesNo
                    
                        if ($response -eq "Yes") {
                            try {
                                # Check connectivity before proceeding
                                do {
                                    if (Test-InternetConnectivity) { break }
                                    
                                    Log-Warning "No internet connectivity detected"
                                    $connResponse = Show-ModernDialog -Message "No internet connection detected.`r`n`r`nAzureAD join requires an active internet connection.`r`n`r`nPlease connect to the internet and click 'Yes' to check again, or 'No' to cancel." -Title "No Internet Connection" -Type Warning -Buttons YesNo
                                    
                                    if ($connResponse -eq "No") {
                                        Log-Info "User cancelled connectivity check"
                                        $lblStatus.Text = "AzureAD join cancelled (offline)"
                                        return
                                    }
                                    Log-Info "User retrying connectivity check..."
                                } until ($false)

                                # Open Settings to the Access work or school page
                                Log-Info "Opening Settings for AzureAD join..."
                                Start-Process "ms-settings:workplace"
                                
                                # Wait a moment for Settings to open, then show dialog on top
                                Start-Sleep -Seconds 2
                                
                                # Show instruction dialog
                                Log-Info "Showing instruction dialog..."
                                $instructionResult = Show-ModernDialog -Message "Settings has been opened to the 'Access work or school' page.`r`n`r`nPlease complete the AzureAD join process in Settings.`r`n`r`nWhen you're done, close Settings and click OK here to continue." -Title "Complete AzureAD Join" -Type Info -Buttons OK
                                Log-Info "User clicked OK on instruction dialog"
                                
                                # Now check if AzureAD join succeeded
                                Log-Info "Checking if AzureAD join succeeded..."
                                if (Test-IsAzureADJoined) {
                                    Log-Info "AzureAD join detected successfully!"
                                    $lblStatus.Text = "AzureAD join successful - click Convert to proceed"
                                    Show-ModernDialog -Message "AzureAD join detected!`r`n`r`nThe device is now joined to AzureAD.`r`n`r`nPlease click 'Convert' again to proceed with the conversion." -Title "Join Successful" -Type Success -Buttons OK
                                }
                                else {
                                    # Join not detected - give user option to retry
                                    Log-Warning "AzureAD join not detected"
                                    $retryResponse = Show-ModernDialog -Message "AzureAD join was not detected.`r`n`r`nThis could mean:`r`n- You haven't completed the join yet`r`n- The join is still processing`r`n- You closed Settings without joining`r`n`r`nWould you like to wait 5 seconds and check again?`r`n`r`r`nClick 'Yes' to retry, or 'No' to cancel." -Title "Join Not Detected" -Type Warning -Buttons YesNo
                                    
                                    if ($retryResponse -eq "Yes") {
                                        Log-Info "User requested retry - waiting 5 seconds..."
                                        Start-Sleep -Seconds 5
                                        
                                        if (Test-IsAzureADJoined) {
                                            Log-Info "AzureAD join detected on retry!"
                                            $lblStatus.Text = "AzureAD join successful - click Convert to proceed"
                                            Show-ModernDialog -Message "AzureAD join now detected!`r`n`r`nPlease click 'Convert' again to proceed with the conversion." -Title "Join Successful" -Type Success -Buttons OK
                                        }
                                        else {
                                            Log-Error "AzureAD join still not detected after retry"
                                            $lblStatus.Text = "AzureAD join failed - please join manually"
                                            Show-ModernDialog -Message "AzureAD join still not detected.`r`n`r`nPlease complete the join process in Settings and try again." -Title "Join Failed" -Type Error -Buttons OK
                                        }
                                    }
                                    else {
                                        # User clicked No - cancelled
                                        Log-Info "User cancelled AzureAD join verification"
                                        $lblStatus.Text = "AzureAD join cancelled"
                                    }
                                }
                            }
                            catch {
                                Log-Error "Error in AzureAD join process: $_"
                                $lblStatus.Text = "Error during AzureAD join process"
                            }
                        }
                    
                        return
                    }
                    else {
                        # Other errors - show standard error dialog
                        $errorMsg = "Conversion cannot proceed:`r`n`r`n" + ($preconditions.Errors -join "`r`n")
                        if ($preconditions.Warnings.Count -gt 0) {
                            $errorMsg += "`r`n`r`nWarnings:`r`n" + ($preconditions.Warnings -join "`r`n")
                        }
                        Show-ModernDialog -Message $errorMsg -Title "Precondition Check Failed" -Type Error -Buttons OK
                        $lblStatus.Text = "Precondition check failed"
                        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                        return
                    }
                }
            
                # Show warnings if any
                if ($preconditions.Warnings.Count -gt 0) {
                    $warnMsg = "Warnings detected:`n`n" + ($preconditions.Warnings -join "`n") + "`n`nDo you want to continue?"
                    $response = Show-ModernDialog -Message $warnMsg -Title "Warnings" -Type Warning -Buttons YesNo
                    if ($response -ne "Yes") {
                        return
                    }
                }
            

            
                # Note: Logging now handled by global Log-Message function via $global:ConversionLogPath
            
                Log-Info "=== PROFILE CONVERSION STARTED ==="
                Log-Info "Log file: $conversionLogPath"
                Log-Info "Source: $sourceUsername ($sourceType)"
                Log-Info "Target: $targetUsername ($targetType)"
                Log-Info ""
            
                # Log precondition check results
                Log-Info "=== PRECONDITION CHECKS ==="
                Log-Info "Checking if user is logged out..."
                Log-Info "User logged out: PASSED"
                Log-Info "Checking administrator privileges..."
                Log-Info "Administrator privileges: PASSED"
                Log-Info "Checking if profile exists..."
                Log-Info "Profile exists: PASSED"
            
                # Log domain connectivity
                Log-Info "Checking domain connectivity..."
                try {
                    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
                    if ($computerSystem.PartOfDomain) {
                        Log-Info "Computer is domain-joined: $($computerSystem.Domain)"
                    }
                    else {
                        Log-Info "Computer is not domain-joined (workgroup: $($computerSystem.Workgroup))"
                    }
                }
                catch {
                    Log-Info "Domain connectivity check: Unable to determine"
                }
            
                # Log disk space
                Log-Info "Checking disk space..."
                try {
                    $systemDrive = $env:SystemDrive
                    $drive = Get-PSDrive -Name $systemDrive.TrimEnd(':') -ErrorAction Stop
                    $freeSpaceGB = [Math]::Round($drive.Free / 1GB, 1)
                    Log-Info "Disk space check passed: $freeSpaceGB GB free on $systemDrive"
                }
                catch {
                    Log-Info "Disk space check: Unable to determine"
                }
                Log-Info ""
            
                $convProgressBar.Value = 10
                [System.Windows.Forms.Application]::DoEvents()
            
                # Create backup if requested
                $backupResult = $null
                if ($chkBackup.Checked) {
                    $lblStatus.Text = "Creating backup (this may take a while)..."
                    [System.Windows.Forms.Application]::DoEvents()
                
                    try {
                        $sourceSID = Get-LocalUserSID -UserName $sourceUsername
                        $profileListBase = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
                        $profileKey = Get-ItemProperty -Path "$profileListBase\$sourceSID" -ErrorAction Stop
                        $profilePath = $profileKey.ProfileImagePath
                    
                        # Backup is a long operation - use Export-UserProfile for comprehensive backup
                        $convProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
                        
                        # Determine save location for backup (Prompt User)
                        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                        $sanitizedUser = $sourceUsername -replace '[\\/:*?"<>| ]', '_'
                        $defaultFileName = "ProfileBackup_${sanitizedUser}_${timestamp}.zip"
                        $defaultDir = "C:\Users\ProfileBackups"
                        
                        # Ensure default backup directory exists
                        if (-not (Test-Path $defaultDir)) {
                            New-Item -ItemType Directory -Path $defaultDir -Force | Out-Null
                        }

                        # Prompt user for location
                        $sfd = New-Object System.Windows.Forms.SaveFileDialog
                        $sfd.Title = "Select Backup Location"
                        $sfd.Filter = "Zip Files (*.zip)|*.zip"
                        $sfd.FileName = $defaultFileName
                        $sfd.InitialDirectory = $defaultDir
                        
                        $backupZipPath = ""
                        $proceedWithBackup = $true

                        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                            $backupZipPath = $sfd.FileName
                        }
                        else {
                            # User cancelled - Ask if they want to skip or use default
                            $res = Show-ModernDialog -Message "Backup location not selected.`n`nDo you want to skip the backup entirely?" -Title "Backup Cancelled" -Type Question -Buttons YesNo
                            if ($res -eq "Yes") {
                                Log-Warning "User cancelled backup operation."
                                $proceedWithBackup = $false
                            }
                            else {
                                $backupZipPath = Join-Path $defaultDir $defaultFileName
                                Log-Info "User cancelled selection but chose to proceed. Using default: $backupZipPath"
                                Show-ModernDialog -Message "Using default backup location:`n$backupZipPath" -Title "Backup Default" -Type Information -Buttons OK
                            }
                        }

                        if ($proceedWithBackup) {
                            # Use Export-UserProfile for comprehensive backup (includes cleanup wizard)
                            try {
                                Export-UserProfile -Username $sourceUsername -ZipPath $backupZipPath
                            
                                # Export-UserProfile doesn't return a result object, check if file was created
                                if (Test-Path $backupZipPath) {
                                    $backupSize = (Get-Item $backupZipPath).Length
                                    $backupSizeMB = [Math]::Round($backupSize / 1MB, 2)
                                
                                    # Also backup registry key (Export-UserProfile doesn't do this)
                                    $regBackupPath = "$backupZipPath.reg"
                                    try {
                                        $profileListKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sourceSID"
                                        reg export $profileListKey $regBackupPath /y | Out-Null
                                        Log-Info "Registry backup created: $regBackupPath"
                                    }
                                    catch {
                                        Log-Warning "Could not backup registry key: $_"
                                    }
                                
                                    $backupResult = @{
                                        Success      = $true
                                        BackupPath   = $backupZipPath
                                        BackupSizeMB = $backupSizeMB
                                        RegBackup    = $regBackupPath
                                    }
                                }
                                else {
                                    $backupResult = @{
                                        Success = $false
                                        Error   = "Export completed but backup file not found"
                                    }
                                }
                            }
                            catch {
                                $backupResult = @{
                                    Success = $false
                                    Error   = $_.Exception.Message
                                }
                            }
                        
                            $convProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                            $convProgressBar.Value = 30
                    
                            if (-not $backupResult.Success) {
                                $response = Show-ModernDialog -Message "Backup failed: $($backupResult.Error)`r`n`r`nDo you want to continue without backup?" -Title "Backup Failed" -Type Warning -Buttons YesNo
                                if ($response -ne "Yes") {
                                    return
                                }
                            }
                        }
                    }
                    catch {
                        $response = Show-ModernDialog -Message "Backup error: $_`r`n`r`nDo you want to continue without backup?" -Title "Backup Error" -Type Warning -Buttons YesNo
                        if ($response -ne "Yes") {
                            return
                        }
                    }
                }
            
                # CRITICAL: For AzureAD→Domain conversions, find source AzureAD profile SID
                # We can't use Get-LocalUserSID because after domain join it resolves to domain SID
                # Instead, search the registry ProfileList for the AzureAD user's profile
                $sourceAzureADSID = $null
                if ($sourceType -eq "AzureAD" -and $targetType -eq "Domain") {
                    Log-Info "=== RESOLVING SOURCE AZUREAD SID ==="
                    Log-Info ">>> Searching for AzureAD user profile in registry..."
                
                    try {
                        # Extract username from AzureAD\username or TENANT\username format
                        $azureUsername = if ($sourceUsername -match '\\(.+)$') { $matches[1] } else { $sourceUsername }
                        Log-Info ">>> Looking for AzureAD user: $azureUsername"
                    
                        # Search ProfileList for AzureAD SIDs (S-1-12-1-...)
                        $profileListBase = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
                        $profiles = Get-ChildItem $profileListBase -ErrorAction Stop
                    
                        foreach ($profile in $profiles) {
                            $sid = $profile.PSChildName
                        
                            # Check if this is an AzureAD SID (starts with S-1-12-1)
                            if ($sid -match '^S-1-12-1-') {
                                $profilePath = (Get-ItemProperty -Path $profile.PSPath -ErrorAction SilentlyContinue).ProfileImagePath
                            
                                # Check if the profile path matches the username
                                if ($profilePath -and $profilePath -match "\\Users\\(.+)$") {
                                    $profileUsername = $matches[1]
                                
                                    # Match by username (case-insensitive)
                                    if ($profileUsername -eq $azureUsername) {
                                        $sourceAzureADSID = $sid
                                        Log-Info ">>> Found AzureAD profile: $profilePath"
                                        Log-Info ">>> Source AzureAD SID: $sourceAzureADSID"
                                        break
                                    }
                                }
                            }
                        }
                    
                        if (-not $sourceAzureADSID) {
                            throw "Could not find AzureAD profile for user '$azureUsername' in registry. The user may not have logged in to this device yet."
                        }
                    }
                    catch {
                        Log-Error ">>> Failed to find source AzureAD profile: $_"
                        Show-ModernDialog -Message "Failed to find source AzureAD user profile:`r`n`r`n$_`r`n`r`nCannot proceed with conversion." -Title "Profile Not Found" -Type Error -Buttons OK
                        $convForm.Close()
                        return
                    }
                }
            
                # CRITICAL: Handle AzureAD unjoin BEFORE conversion if needed
                # Windows does not allow domain join while AzureAD joined
                Log-Info "=== UNJOIN CHECK START ==="
                Log-Info "Source Type: $sourceType"
                Log-Info "Target Type: $targetType"
            
                if ($sourceType -eq "AzureAD" -and $targetType -eq "Domain") {
                    Log-Info ">>> AzureAD to Domain conversion detected - checking AzureAD join status..."
                
                    $isAzureADJoined = Test-IsAzureADJoined
                    Log-Info ">>> Test-IsAzureADJoined result: $isAzureADJoined"
                
                    if ($isAzureADJoined) {
                        Log-Info ">>> Device IS AzureAD joined - unjoin is REQUIRED before domain join"
                    
                        $lblStatus.Text = "AzureAD unjoin required..."
                        [System.Windows.Forms.Application]::DoEvents()
                    
                        $message = "This device is currently joined to AzureAD.`r`n`r`n"
                        $message += "Windows REQUIRES unjoining from AzureAD before joining a domain.`r`n`r`n"
                        $message += "This will:`r`n"
                        $message += "- Remove device from your organization's management`r`n"
                        $message += "- Disable conditional access policies`r`n"
                        $message += "- Remove SSO capabilities`r`n`r`n"
                        $message += "After unjoin, the conversion will proceed with domain join.`r`n`r`n"
                        $message += "Continue with required AzureAD unjoin?"
                    
                        Log-Info ">>> Showing unjoin confirmation dialog..."
                        $response = Show-ModernDialog -Message $message -Title "AzureAD Unjoin Required" -Type Warning -Buttons YesNo
                        Log-Info ">>> User response: $response"
                    
                        if ($response -eq 'Yes') {
                            Log-Info ">>> User confirmed - executing unjoin..."
                            $lblStatus.Text = "Unjoining from AzureAD..."
                            $convProgressBar.Value = 40
                            [System.Windows.Forms.Application]::DoEvents()
                        
                            $unjoinResult = Invoke-AzureADUnjoin
                            Log-Info ">>> Unjoin result - Success: $($unjoinResult.Success), Message: $($unjoinResult.Message)"
                        
                            if ($unjoinResult.Success) {
                                Log-Info ">>> AzureAD unjoin SUCCESSFUL - can now proceed with domain join"
                                $lblStatus.Text = "AzureAD unjoin successful"
                                [System.Windows.Forms.Application]::DoEvents()
                                Start-Sleep -Milliseconds 500
                            
                                # Now join the domain before converting profile
                                Log-Info ">>> Proceeding with domain join..."
                                $lblStatus.Text = "Joining domain..."
                                $convProgressBar.Value = 45
                                [System.Windows.Forms.Application]::DoEvents()
                            
                                try {
                                    # Extract domain name from target username (DOMAIN\username)
                                    if ($targetUsername -match '^([^\\]+)\\') {
                                        $domainName = $matches[1]
                                        Log-Info ">>> Domain name extracted: $domainName"
                                    
                                        # Prompt for domain credentials using Get-Credential
                                        Log-Info ">>> Prompting for domain administrator credentials..."
                                        $domainCred = Get-Credential -Message "Enter domain administrator credentials to join $domainName" -UserName "$domainName\"
                                    
                                        if ($domainCred) {
                                            Log-Info ">>> Joining domain $domainName with provided credentials..."
                                            Add-Computer -DomainName $domainName -Credential $domainCred -Force -ErrorAction Stop
                                            Log-Info ">>> Domain join SUCCESSFUL"
                                            $lblStatus.Text = "Domain join successful"
                                            [System.Windows.Forms.Application]::DoEvents()
                                            Start-Sleep -Milliseconds 500
                                        }
                                        else {
                                            throw "User cancelled domain join"
                                        }
                                    }
                                    else {
                                        throw "Could not extract domain name from username: $targetUsername"
                                    }
                                }
                                catch {
                                    Log-Error ">>> Domain join FAILED: $_"
                                    $errorMsg = "Domain join failed: $_`r`n`r`n"
                                    $errorMsg += "Cannot convert to domain profile without joining domain.`r`n`r`n"
                                    $errorMsg += "Please join the domain manually and retry."
                                    Show-ModernDialog -Message $errorMsg -Title "Domain Join Failed" -Type Error -Buttons OK
                                    $convForm.Close()
                                    return
                                }
                            }
                            else {
                                Log-Error ">>> AzureAD unjoin FAILED - cannot proceed"
                                $errorMsg = "AzureAD unjoin failed: $($unjoinResult.Message)`r`n`r`n"
                                $errorMsg += "Cannot proceed with domain join while AzureAD joined.`r`n`r`n"
                                $errorMsg += "You can manually unjoin using: dsregcmd /leave"
                                Show-ModernDialog -Message $errorMsg -Title "Unjoin Failed" -Type Error -Buttons OK
                                $convForm.Close()
                                return
                            }
                        }
                        else {
                            Log-Info ">>> User CANCELLED required AzureAD unjoin - aborting conversion"
                            Show-ModernDialog -Message "Cannot join domain while AzureAD joined.`r`n`r`nConversion cancelled." -Title "Conversion Cancelled" -Type Warning -Buttons OK
                            $convForm.Close()
                            return
                        }
                    }
                    else {
                        Log-Info ">>> Device is NOT AzureAD joined - proceeding with conversion"
                    }
                }
                elseif ($sourceType -eq "Domain" -and $targetType -eq "AzureAD") {
                    Log-Info ">>> Domain to AzureAD conversion detected - checking domain join status..."
                
                    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
                    $isDomainJoined = $computerSystem.PartOfDomain
                    Log-Info ">>> Is domain joined: $isDomainJoined"
                
                    if ($isDomainJoined) {
                        Log-Info ">>> Device IS domain joined - unjoin is REQUIRED before AzureAD join"
                    
                        $lblStatus.Text = "Domain unjoin required..."
                        [System.Windows.Forms.Application]::DoEvents()
                    
                        $message = "This device is currently joined to domain: $($computerSystem.Domain)`r`n`r`n"
                        $message += "Windows REQUIRES unjoining from the domain before joining AzureAD.`r`n`r`n"
                        $message += "This will:`r`n"
                        $message += "- Remove device from domain management`r`n"
                        $message += "- Disable domain group policies`r`n"
                        $message += "- Remove domain authentication`r`n`r`n"
                        $message += "After unjoin, the conversion will proceed with AzureAD join.`r`n`r`n"
                        $message += "Continue with required domain unjoin?"
                    
                        Log-Info ">>> Showing domain unjoin confirmation dialog..."
                        $response = Show-ModernDialog -Message $message -Title "Domain Unjoin Required" -Type Warning -Buttons YesNo
                        Log-Info ">>> User response: $response"
                    
                        if ($response -eq 'Yes') {
                            Log-Info ">>> User confirmed - executing domain unjoin..."
                            $lblStatus.Text = "Unjoining from domain..."
                            $convProgressBar.Value = 40
                            [System.Windows.Forms.Application]::DoEvents()
                        
                            $unjoinResult = Invoke-DomainUnjoin
                            Log-Info ">>> Unjoin result - Success: $($unjoinResult.Success), Message: $($unjoinResult.Message)"
                        
                            if ($unjoinResult.Success) {
                                Log-Info ">>> Domain unjoin SUCCESSFUL - can now proceed with AzureAD join"
                                $lblStatus.Text = "Domain unjoin successful"
                                [System.Windows.Forms.Application]::DoEvents()
                                Start-Sleep -Milliseconds 500
                            
                                # Now join AzureAD before converting profile
                                Log-Info ">>> Proceeding with AzureAD join..."
                                $lblStatus.Text = "Joining AzureAD..."
                                $convProgressBar.Value = 45
                                [System.Windows.Forms.Application]::DoEvents()
                            
                                # Show AzureAD join dialog
                                Show-ModernDialog -Message "After clicking OK, you will be prompted to join AzureAD.`r`n`r`nPlease sign in with your AzureAD credentials when prompted.`r`n`r`nAfter joining AzureAD, the profile conversion will continue." -Title "AzureAD Join Required" -Type Info -Buttons OK
                            
                                # Launch AzureAD join with robust verification
                                try {
                                    # Check connectivity before proceeding
                                    do {
                                        if (Test-InternetConnectivity) { break }
                                            
                                        Log-Warning "No internet connectivity detected"
                                        $connResponse = Show-ModernDialog -Message "No internet connection detected.`r`n`r`nAzureAD join requires an active internet connection.`r`n`r`nPlease connect to the internet and click 'Yes' to check again, or 'No' to cancel." -Title "No Internet Connection" -Type Warning -Buttons YesNo
                                            
                                        if ($connResponse -eq "No") {
                                            Log-Info "User cancelled connectivity check"
                                            $lblStatus.Text = "AzureAD join cancelled (offline)"
                                            return
                                        }
                                        Log-Info "User retrying connectivity check..."
                                    } until ($false)
    
                                    # Open Settings to the Access work or school page
                                    Log-Info "Opening Settings for AzureAD join..."
                                    Start-Process "ms-settings:workplace"
                                        
                                    # Wait a moment for Settings to open, then show dialog on top
                                    Start-Sleep -Seconds 2
                                        
                                    # Show instruction dialog
                                    Log-Info "Showing instruction dialog..."
                                    $instructionResult = Show-ModernDialog -Message "Settings has been opened to the 'Access work or school' page.`r`n`r`nPlease complete the AzureAD join process in Settings.`r`n`r`nWhen you're done, close Settings and click OK here to continue." -Title "Complete AzureAD Join" -Type Info -Buttons OK
                                    Log-Info "User clicked OK on instruction dialog"
                                        
                                    # Verify AzureAD join
                                    Start-Sleep -Seconds 2
                                    if (Test-IsAzureADJoined) {
                                        Log-Info ">>> AzureAD join SUCCESSFUL"
                                        $lblStatus.Text = "AzureAD join successful"
                                        [System.Windows.Forms.Application]::DoEvents()
                                        Start-Sleep -Milliseconds 500
                                    }
                                    else {
                                        # Join not detected - give user option to retry
                                        Log-Warning "AzureAD join not detected"
                                        $retryResponse = Show-ModernDialog -Message "AzureAD join was not detected.`r`n`r`nThis could mean:`r`n- You haven't completed the join yet`r`n- The join is still processing`r`n- You closed Settings without joining`r`n`r`nWould you like to wait 5 seconds and check again?`r`n`r`r`nClick 'Yes' to retry, or 'No' to cancel." -Title "Join Not Detected" -Type Warning -Buttons YesNo
                                            
                                        if ($retryResponse -eq "Yes") {
                                            Log-Info "User requested retry - waiting 5 seconds..."
                                            Start-Sleep -Seconds 5
                                                
                                            if (Test-IsAzureADJoined) {
                                                Log-Info "AzureAD join detected on retry!"
                                                $lblStatus.Text = "AzureAD join successful"
                                                [System.Windows.Forms.Application]::DoEvents()
                                            }
                                            else {
                                                Log-Error "AzureAD join still not detected after retry"
                                                throw "AzureAD join was not completed"
                                            }
                                        }
                                        else {
                                            Log-Info "User cancelled AzureAD join verification"
                                            $lblStatus.Text = "AzureAD join cancelled"
                                            return
                                        }
                                    }
                                }
                                catch {
                                    Log-Error ">>> AzureAD join FAILED: $_"
                                    $errorMsg = "AzureAD join failed: $_`r`n`r`n"
                                    $errorMsg += "Cannot convert to AzureAD profile without joining AzureAD.`r`n`r`n"
                                    $errorMsg += "Please join AzureAD manually and retry."
                                    Show-ModernDialog -Message $errorMsg -Title "AzureAD Join Failed" -Type Error -Buttons OK
                                    $convForm.Close()
                                    return
                                }
                            }
                            else {
                                Log-Error ">>> Domain unjoin FAILED - cannot proceed"
                                $errorMsg = "Domain unjoin failed: $($unjoinResult.Message)`r`n`r`n"
                                $errorMsg += "Cannot proceed with AzureAD join while domain joined.`r`n`r`n"
                                $errorMsg += "You can manually unjoin using System Settings."
                                Show-ModernDialog -Message $errorMsg -Title "Unjoin Failed" -Type Error -Buttons OK
                                $convForm.Close()
                                return
                            }
                        }
                        else {
                            Log-Info ">>> User CANCELLED required domain unjoin - aborting conversion"
                            Show-ModernDialog -Message "Cannot join AzureAD while domain joined.`r`n`r`nConversion cancelled." -Title "Conversion Cancelled" -Type Warning -Buttons OK
                            $convForm.Close()
                            return
                        }
                    }
                    else {
                        Log-Info ">>> Device is NOT domain joined - proceeding with conversion"
                    }
                }
                else {
                    Log-Info ">>> Not an AzureAD->Domain or Domain->AzureAD conversion - skipping unjoin check"
                }
                Log-Info "=== UNJOIN CHECK END ==="
            
                # Perform conversion
                # Perform conversion
                $lblStatus.Text = "Converting profile..."
                $convProgressBar.Value = 35
                [System.Windows.Forms.Application]::DoEvents()
            
                $conversionResult = $null
            
                # Universal Repair Check (Source == Target)
                # Handles Domain->Domain and AzureAD->AzureAD repair scenarios
                # Local->Local is handled separately below to preserve existing password/admin features
                if ($sourceUsername -eq $targetUsername -and $sourceType -eq $targetType -and $sourceType -ne "Local") {
                    Log-Info "Universal Repair triggered for $sourceType user: $sourceUsername"
                    $conversionResult = Repair-UserProfile -Username $sourceUsername -UserType $sourceType
                    $lblStatus.Text = "Finalizing..."
                    [System.Windows.Forms.Application]::DoEvents()
                }
                elseif ($sourceType -eq "Local" -and $targetType -eq "Domain") {
                    # Add domain prefix if not present
                    if ($targetUsername -notmatch '\\') {
                        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
                        $targetUsername = "$($computerSystem.Domain)\$targetUsername"
                    }
                    $conversionResult = Convert-LocalToDomain -LocalUsername $sourceUsername -DomainUsername $targetUsername
                    $lblStatus.Text = "Finalizing..."
                    [System.Windows.Forms.Application]::DoEvents()
                }
                elseif ($sourceType -eq "Domain" -and $targetType -eq "Local") {
                    $password = ConvertTo-SecureString $txtPassword.Text -AsPlainText -Force
                    $makeAdmin = $chkMakeAdmin.Checked
                    $unjoinDomain = $chkUnjoinAzureAD.Checked  # Reuse same checkbox
                    $conversionResult = Convert-DomainToLocal -DomainUsername $sourceUsername -LocalUsername $targetUsername -LocalPassword $password -MakeAdmin $makeAdmin -UnjoinDomain $unjoinDomain
                    $lblStatus.Text = "Finalizing..."
                    [System.Windows.Forms.Application]::DoEvents()
                }
                elseif ($sourceType -eq "Local" -and $targetType -eq "Local") {
                    # Local to Local conversion
                    if ($sourceUsername -eq $targetUsername) {
                        # REPAIR MODE: Source == Target
                        
                        $repairMsg = "Source and Target types are the same (Local).`r`n`r`n"
                        $repairMsg += "Do you want to run an In-Place Repair?`r`n`r`n"
                        $repairMsg += "This will re-process the profile permissions and SIDs without changing the account type."
                        
                        $repairResponse = Show-ModernDialog -Message $repairMsg -Title "Run In-Place Repair?" -Type Question -Buttons YesNo
                        
                        if ($repairResponse -eq "Yes") {
                            Log-Info "Local Repair triggered for user: $sourceUsername"
                            $conversionResult = Repair-UserProfile -Username $sourceUsername -UserType "Local"
                            $lblStatus.Text = "Finalizing..."
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                        else {
                            Log-Info "User cancelled In-Place Repair."
                            return
                        }
                    }
                    else {
                        # MIGRATION MODE: Source != Target
                        $password = ConvertTo-SecureString $txtPassword.Text -AsPlainText -Force
                        $makeAdmin = $chkMakeAdmin.Checked
                        
                        # Reuse DomainToLocal logic (it handles local SIDs correctly via Get-LocalUserSID)
                        $conversionResult = Convert-DomainToLocal -DomainUsername $sourceUsername -LocalUsername $targetUsername -LocalPassword $password -MakeAdmin $makeAdmin
                    }
                }
                elseif ($sourceType -eq "AzureAD" -and $targetType -eq "Local") {
                    $password = ConvertTo-SecureString $txtPassword.Text -AsPlainText -Force
                    $makeAdmin = $chkMakeAdmin.Checked
                    $unjoinAzureAD = $chkUnjoinAzureAD.Checked
                    $conversionResult = Convert-AzureADToLocal -AzureADUsername $sourceUsername -LocalUsername $targetUsername -LocalPassword $password -MakeAdmin $makeAdmin -UnjoinAzureAD $unjoinAzureAD
                    $lblStatus.Text = "Finalizing..."
                    [System.Windows.Forms.Application]::DoEvents()
                }
                elseif ($sourceType -eq "Local" -and $targetType -eq "AzureAD") {
                    # Ensure AzureAD\ prefix
                    if ($targetUsername -notmatch '^AzureAD\\') {
                        $targetUsername = "AzureAD\$targetUsername"
                    }
                    $conversionResult = Convert-LocalToAzureAD -LocalUsername $sourceUsername -AzureADUsername $targetUsername
                }
                elseif ($sourceType -eq "Domain" -and $targetType -eq "AzureAD") {
                    # Ensure AzureAD\ prefix
                    if ($targetUsername -notmatch '^AzureAD\\') {
                        $targetUsername = "AzureAD\$targetUsername"
                    }
                    # Domain to AzureAD: Similar to Local to AzureAD
                    $conversionResult = Convert-LocalToAzureAD -LocalUsername $sourceUsername -AzureADUsername $targetUsername
                }
                elseif ($sourceType -eq "AzureAD" -and $targetType -eq "Domain") {
                    # AzureAD to Domain: Must unjoin from AzureAD first
                    # Add domain prefix if not present
                    if ($targetUsername -notmatch '\\') {
                        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
                        $targetUsername = "$($computerSystem.Domain)\$targetUsername"
                    }
                    # CRITICAL: Pass pre-resolved source SID (captured before unjoin)
                    # After unjoin, we can't resolve AzureAD SIDs anymore
                    $conversionResult = Convert-LocalToDomain -LocalUsername $sourceUsername -DomainUsername $targetUsername -UnjoinAzureAD $true -SourceSID $sourceAzureADSID
                }
                elseif ($sourceType -eq "AzureAD" -and $targetType -eq "AzureAD") {
                    # AzureAD to AzureAD (Repair Only)
                    
                    # Ensure AzureAD\ prefix for target
                    if ($targetUsername -notmatch '^AzureAD\\') {
                        $targetUsername = "AzureAD\$targetUsername"
                    }

                    if ($sourceUsername -eq $targetUsername) {
                        Log-Info "AzureAD Repair triggered (Explicit Match) for: $sourceUsername"
                        $conversionResult = Repair-UserProfile -Username $sourceUsername -UserType "AzureAD"
                        $lblStatus.Text = "Finalizing..."
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                    else {
                        Show-ModernDialog -Message "AzureAD to AzureAD *migration* (different users) is not supported.`r`n`r`nOnly In-Place Repair (Same User) is supported." -Title "Not Supported" -Type Info -Buttons OK
                        return
                    }
                }
                else {
                    $supportedMsg = "Unsupported conversion:`r`n`r`n"
                    $supportedMsg += "Attempted: $sourceType to $targetType`r`n`r`n"
                    $supportedMsg += "Supported conversions:`r`n"
                    $supportedMsg += "- Local <-> Domain`r`n"
                    $supportedMsg += "- Local <-> Local (Repair/Migration)`r`n"
                    $supportedMsg += "- Local <-> AzureAD`r`n"
                    $supportedMsg += "- Domain <-> AzureAD`r`n"
                    $supportedMsg += "- AzureAD <-> Local`r`n"
                
                    Show-ModernDialog -Message $supportedMsg -Title "Conversion Not Supported" -Type Info -Buttons OK
                    return
                }
            
                # Show results
                if ($conversionResult.Success) {
                    # Write completion to log
                    $logFooter = @"

=============================================================================
CONVERSION COMPLETED SUCCESSFULLY
=============================================================================
Completed: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Source SID: $($conversionResult.SourceSID)
Target SID: $($conversionResult.TargetSID)
=============================================================================
"@
                    $logFooter | Out-File -FilePath $conversionLogPath -Append -Encoding UTF8
                
                    Log-Info "=== PROFILE CONVERSION COMPLETED SUCCESSFULLY ==="
                    Log-Info "Source SID: $($conversionResult.SourceSID)"

                    Log-Info "Target SID: $($conversionResult.TargetSID)"
                    
                    $lblStatus.Text = "Conversion complete!"
                    if ($global:StatusText) { $global:StatusText.Text = "Conversion complete!" }
                    if ($global:ConversionProgressBar) { $global:ConversionProgressBar.Value = 100 }
                    [System.Windows.Forms.Application]::DoEvents()
                
                    $successMsg = "Profile conversion completed successfully!`r`n`r`n"
                    $successMsg += "Source: $sourceUsername ($sourceType)`r`n"
                    $successMsg += "Target: $targetUsername ($targetType)`r`n`r`n"
                    $successMsg += "Source SID: $($conversionResult.SourceSID)`r`n"
                    $successMsg += "Target SID: $($conversionResult.TargetSID)`r`n`r`n"
                
                    if ($backupResult -and $backupResult.Success) {
                        $successMsg += "Backup: $($backupResult.BackupPath) ($($backupResult.BackupSizeMB) MB)`r`n`r`n"
                    }
                
                    $successMsg += "Log file: $conversionLogPath`r`n`r`n"
                    $successMsg += "IMPORTANT: Restart the computer before logging in with the converted profile."
                
                    
                    # --- NEW FEATURE: Option to delete source user ---
                    $userDeleted = $false
                    if ($sourceType -eq "Local" -and $sourceUsername -ne $env:USERNAME) {
                        try {
                            $delResponse = Show-ModernDialog -Message "Conversion was successful.`r`n`r`nDo you want to DELETE the source user '$sourceUsername'?`r`n`r`nWARNING: This will permanently delete the source account and its data (if not fully migrated).`r`n`r`nWe recommend keeping it as a backup for now." -Title "Delete Source User?" -Type Warning -Buttons YesNo
                            
                            if ($delResponse -eq "Yes") {
                                Log-Info "User requested deletion of source user: $sourceUsername"
                                
                                # Sanitize username for Remove-LocalUser (strip computer/domain prefix)
                                $userToDelete = $sourceUsername
                                if ($userToDelete -match '\\') {
                                    $userToDelete = ($userToDelete -split '\\')[-1]
                                }
                                
                                Remove-LocalUser -Name $userToDelete -ErrorAction Stop
                                Log-Info "Source user '$userToDelete' deleted successfully."
                                Show-ModernDialog -Message "User '$userToDelete' has been deleted." -Title "User Deleted" -Type Success -Buttons OK
                                $userDeleted = $true
                            }
                        }
                        catch {
                            Log-Error "Failed to delete source user: $_"
                            Show-ModernDialog -Message "Failed to delete user '$sourceUsername':`r`n$_" -Title "Deletion Failed" -Type Error -Buttons OK
                        }
                    }

                    # Offer domain unjoin option for domain-to-local conversions (only if user wasn't deleted)
                    $unjoinStatus = "Not Attempted"
                    if ($sourceType -eq "Domain" -and $targetType -eq "Local") {
                        try {
                            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
                            if ($computerSystem.PartOfDomain) {
                                $unjoinPrompt = "The profile has been successfully converted to a local user.`r`n`r`n"
                                $unjoinPrompt += "This computer is currently joined to domain: $($computerSystem.Domain)`r`n`r`n"
                                $unjoinPrompt += "Would you like to unjoin this computer from the domain and move it to a workgroup?`r`n`r`n"
                                $unjoinPrompt += "Note: This will require domain admin credentials and a restart."
                            
                                $unjoinResponse = Show-ModernDialog -Message $unjoinPrompt -Title "Unjoin from Domain?" -Type Question -Buttons YesNo
                            
                                if ($unjoinResponse -eq "Yes") {
                                    # Collect domain admin credentials
                                    $credForm = New-Object System.Windows.Forms.Form
                                    $credForm.Text = "Domain Admin Credentials Required"
                                    $credForm.Size = New-Object System.Drawing.Size(500, 340)
                                    $credForm.StartPosition = "CenterScreen"
                                    $credForm.FormBorderStyle = "FixedDialog"
                                    $credForm.MaximizeBox = $false
                                    $credForm.MinimizeBox = $false
                                    $credForm.TopMost = $true
                                    $credForm.BackColor = $theme.FormBackColor
                                    $credForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

                                    # Header panel
                                    $headerPanel = New-Object System.Windows.Forms.Panel
                                    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
                                    $headerPanel.Size = New-Object System.Drawing.Size(500, 70)
                                    $headerPanel.BackColor = $theme.HeaderBackColor
                                    $credForm.Controls.Add($headerPanel)

                                    $lblTitle = New-Object System.Windows.Forms.Label
                                    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
                                    $lblTitle.Size = New-Object System.Drawing.Size(460, 25)
                                    $lblTitle.Text = "Domain Admin Credentials Required"
                                    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
                                    $lblTitle.ForeColor = $theme.HeaderTextColor
                                    $headerPanel.Controls.Add($lblTitle)

                                    $lblInfo = New-Object System.Windows.Forms.Label
                                    $lblInfo.Location = New-Object System.Drawing.Point(22, 43)
                                    $lblInfo.Size = New-Object System.Drawing.Size(460, 20)
                                    $lblInfo.Text = "Enter credentials to unjoin from domain: $($computerSystem.Domain)"
                                    $lblInfo.ForeColor = $theme.SubHeaderTextColor
                                    $headerPanel.Controls.Add($lblInfo)

                                    # Main content card
                                    $contentCard = New-Object System.Windows.Forms.Panel
                                    $contentCard.Location = New-Object System.Drawing.Point(15, 85)
                                    $contentCard.Size = New-Object System.Drawing.Size(460, 160)
                                    $contentCard.BackColor = $theme.PanelBackColor
                                    $credForm.Controls.Add($contentCard)

                                    $lblUser = New-Object System.Windows.Forms.Label
                                    $lblUser.Location = New-Object System.Drawing.Point(20, 25)
                                    $lblUser.Size = New-Object System.Drawing.Size(100, 20)
                                    $lblUser.Text = "Username:"
                                    $lblUser.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                                    $lblUser.ForeColor = $theme.LabelTextColor
                                    $contentCard.Controls.Add($lblUser)

                                    $txtUser = New-Object System.Windows.Forms.TextBox
                                    $txtUser.Location = New-Object System.Drawing.Point(130, 23)
                                    $txtUser.Size = New-Object System.Drawing.Size(300, 25)
                                    $txtUser.Text = "$($computerSystem.Domain)\"
                                    $txtUser.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                                    $txtUser.BackColor = $theme.TextBoxBackColor
                                    $txtUser.ForeColor = $theme.TextBoxForeColor
                                    $contentCard.Controls.Add($txtUser)

                                    $lblPass = New-Object System.Windows.Forms.Label
                                    $lblPass.Location = New-Object System.Drawing.Point(20, 65)
                                    $lblPass.Size = New-Object System.Drawing.Size(100, 20)
                                    $lblPass.Text = "Password:"
                                    $lblPass.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                                    $lblPass.ForeColor = $theme.LabelTextColor
                                    $contentCard.Controls.Add($lblPass)

                                    $txtPass = New-Object System.Windows.Forms.TextBox
                                    $txtPass.Location = New-Object System.Drawing.Point(130, 63)
                                    $txtPass.Size = New-Object System.Drawing.Size(300, 25)
                                    $txtPass.UseSystemPasswordChar = $true
                                    $txtPass.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                                    $txtPass.BackColor = $theme.TextBoxBackColor
                                    $txtPass.ForeColor = $theme.TextBoxForeColor
                                    $contentCard.Controls.Add($txtPass)

                                    $chkShowPass = New-Object System.Windows.Forms.CheckBox
                                    $chkShowPass.Location = New-Object System.Drawing.Point(130, 100)
                                    $chkShowPass.Size = New-Object System.Drawing.Size(150, 23)
                                    $chkShowPass.Text = "Show password"
                                    $chkShowPass.ForeColor = $theme.LabelTextColor
                                    $chkShowPass.Add_CheckedChanged({ $txtPass.UseSystemPasswordChar = -not $chkShowPass.Checked })
                                    $contentCard.Controls.Add($chkShowPass)

                                    $btnOK = New-Object System.Windows.Forms.Button
                                    $btnOK.Location = New-Object System.Drawing.Point(250, 260)
                                    $btnOK.Size = New-Object System.Drawing.Size(110, 35)
                                    $btnOK.Text = "Unjoin"
                                    $btnOK.BackColor = $theme.ButtonPrimaryBackColor
                                    $btnOK.ForeColor = $theme.ButtonPrimaryForeColor
                                    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                                    $btnOK.FlatAppearance.BorderSize = 0
                                    $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                                    $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
                                    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
                                    $btnOK.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
                                    $btnOK.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
                                    $credForm.Controls.Add($btnOK)

                                    $btnCancel = New-Object System.Windows.Forms.Button
                                    $btnCancel.Location = New-Object System.Drawing.Point(370, 260)
                                    $btnCancel.Size = New-Object System.Drawing.Size(105, 35)
                                    $btnCancel.Text = "Skip"
                                    $btnCancel.BackColor = $theme.ButtonSecondaryBackColor
                                    $btnCancel.ForeColor = $theme.ButtonSecondaryForeColor
                                    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                                    $btnCancel.FlatAppearance.BorderSize = 0
                                    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                                    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
                                    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                                    $btnCancel.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
                                    $btnCancel.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
                                    $credForm.Controls.Add($btnCancel)

                                    $credForm.AcceptButton = $btnOK
                                    $credForm.CancelButton = $btnCancel

                                    $credResult = $credForm.ShowDialog()
                                    if ($credResult -eq [System.Windows.Forms.DialogResult]::OK) {
                                        $username = $txtUser.Text.Trim()
                                        $password = $txtPass.Text
                                        $credForm.Dispose()

                                        if ([string]::IsNullOrWhiteSpace($username) -or [string]::IsNullOrWhiteSpace($password)) {
                                            Log-Warning "Domain unjoin skipped - credentials not provided"
                                            $unjoinStatus = "Skipped"
                                        }
                                        else {
                                            try {
                                                Log-Info "Attempting to unjoin from domain: $($computerSystem.Domain)"
                                                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                                                $domainCred = New-Object System.Management.Automation.PSCredential($username, $securePassword)
                                            
                                                # Unjoin from domain and move to workgroup
                                                $workgroupName = "WORKGROUP"
                                                Remove-Computer -UnjoinDomainCredential $domainCred -WorkgroupName $workgroupName -Force -ErrorAction Stop
                                            
                                                Log-Info "Successfully unjoined from domain. Computer will be in workgroup: $workgroupName"
                                                $unjoinStatus = "Success"
                                                $successMsg += "`r`n`r`nDomain Unjoin: SUCCESS - Computer moved to workgroup '$workgroupName'"
                                                $successMsg += "`r`n`r`nRESTART REQUIRED to complete domain unjoin."
                                            }
                                            catch {
                                                Log-Error "Failed to unjoin from domain: $_"
                                                $unjoinStatus = "Failed"
                                                $successMsg += "`r`n`r`nDomain Unjoin: FAILED - $($_.Exception.Message)"
                                            }
                                        }
                                    }
                                    else {
                                        $credForm.Dispose()
                                        Log-Info "Domain unjoin skipped by user"
                                        $unjoinStatus = "Skipped"
                                    }
                                }
                                else {
                                    Log-Info "Domain unjoin declined by user"
                                    $unjoinStatus = "Declined"
                                }
                            }
                        }
                        catch {
                            Log-Warning "Could not check domain membership: $_"
                        }
                    }
                
                    # Finalize Status Strings for Report
                    $domainStatus = "Not Changed"
                    $azureStatusVal = "Not Changed"
                    
                    # Determine domain status
                    if ($unjoinStatus -eq "Success") {
                        $domainStatus = "Unjoined (Workgroup)"
                    }
                    elseif ($unjoinStatus -eq "Failed") {
                        $domainStatus = "Unjoin Failed"
                    }
                    elseif ($sourceType -eq "Local" -and $targetType -eq "Domain") {
                        # We effectively joined a domain profile context, though machine join happens separately
                        $domainStatus = "Target is Domain Profile" 
                    }
                    elseif ($sourceType -eq "Domain" -and $targetType -eq "Local") {
                        # Domain to Local conversion - domain was unjoined
                        $domainStatus = "Unjoined Domain"
                    }
                    
                    # Determine AzureAD status
                    if ($chkUnjoinAzureAD.Checked) {
                        $azureStatusVal = "Unjoined AzureAD"
                    }
                    elseif ($targetType -eq "AzureAD") {
                        $azureStatusVal = "Target is AzureAD"
                    }
                    elseif ($sourceType -eq "AzureAD" -and $targetType -eq "Local") {
                        # AzureAD to Local conversion - AzureAD was unjoined
                        $azureStatusVal = "Unjoined AzureAD"
                    }
                    elseif ($conversionResult.AzureADStatus -eq "Unjoined") {
                        # Dynamic unjoin handled by conversion function (e.g. Local->Domain)
                        $azureStatusVal = "Unjoined AzureAD"
                    }


                    # Register AppX fix for next login (Active Setup)
                    # This will run before the desktop loads to prevent file lock errors
                    Add-AppxReregistrationActiveSetup -Username $targetUsername -OperationType "Conversion"

                    # Generate Final HTML Report
                    $conversionEndTime = Get-Date
                    $reportPath = New-ConversionReport -SourceUser $sourceUsername -TargetUser $targetUsername -SourceType $sourceType -TargetType $targetType -Status "Success" -StartTime $conversionStartTime -EndTime $conversionEndTime -BackupPath $backupResult.BackupPath -BackupSizeMB $backupResult.BackupSizeMB -LogPath $conversionLogPath -JoinedDomain $domainStatus -AzureStatus $azureStatusVal
                    
                    if ($reportPath) {
                        $successMsg += "`r`n`r`nHTML Report: $reportPath"
                        Log-Info "HTML Report generated: $reportPath"
                        # Auto-launch report if configured
                        if ($global:Config.AutoOpenReports) {
                            try {
                                Start-Process $reportPath
                                Log-Info "Opened HTML report in browser"
                            }
                            catch {
                                Log-Warning "Could not auto-open report: $_"
                            }
                        }
                    }
                
                    Show-ModernDialog -Message $successMsg -Title "Conversion Successful" -Type Success -Buttons OK
                    $convForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $convForm.Close()
                }
                else {
                    # Write failure to log
                    $logFooter = @"

=============================================================================
CONVERSION FAILED
=============================================================================
Failed: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Error: $($conversionResult.Error)
=============================================================================
"@
                    $logFooter | Out-File -FilePath $conversionLogPath -Append -Encoding UTF8
                
                    Log-Info "=== PROFILE CONVERSION FAILED ===" "ERROR"
                    Log-Info "Error: $($conversionResult.Error)" "ERROR"
                
                    # Generate HTML report for failure
                    $conversionEndTime = Get-Date
                    $reportPath = New-ConversionReport -SourceUser $sourceUsername -TargetUser $targetUsername -SourceType $sourceType -TargetType $targetType -Status "Failed" -StartTime $conversionStartTime -EndTime $conversionEndTime -BackupPath "" -BackupSizeMB "" -LogPath $conversionLogPath
                    
                    # Auto-launch failure report if configured
                    if ($reportPath -and $global:Config.AutoOpenReports) {
                        try {
                            Start-Process $reportPath
                            Log-Info "Opened failure report in browser"
                        }
                        catch {
                            Log-Warning "Could not auto-open report: $_"
                        }
                    }
                
                    Show-ModernDialog -Message "Conversion failed:`n`n$($conversionResult.Error)`r`n`r`nLog file: $conversionLogPath" -Title "Conversion Failed" -Type Error -Buttons OK
                    $lblStatus.Text = "Conversion failed"
                    if ($global:StatusText) { $global:StatusText.Text = "Conversion failed" }
                    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                }
            }
            catch {
                Log-Error "Conversion dialog error: $_"
                Show-ModernDialog -Message "An error occurred during conversion:`r`n`r`n$_" -Title "Error" -Type Error -Buttons OK
                $lblStatus.Text = "Error: $_"
                $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
            }
        })
    
    $convForm.CancelButton = $btnCancel
    $convForm.CancelButton = $btnCancel
    
    # Cleanup global reference on close
    $convForm.Add_FormClosed({
            $global:ConversionProgressBar = $null
            $global:ConversionLogPath = $null
        })
    
    $convForm.ShowDialog()
}

# =============================================================================
# HANDLE RESTART - GUI AWARE WITH MULTIPLE BEHAVIORS
# =============================================================================
function Handle-Restart {
    param(
        [ValidateSet('Prompt', 'Immediate', 'Never', 'Delayed')]
        [string]$Behavior = 'Prompt',
        [int]$DelaySeconds = 30,
        [string]$Reason = "System configuration changed"
    )
    switch ($Behavior) {
        'Immediate' {
            Log-Message "Restarting immediately (Behavior: Immediate)"
            Show-ModernDialog -Message "$Reason`r`n`r`nRestarting NOW!" -Title "Restarting" -Type Warning -Buttons OK
            Log-Message "Executing forced restart via shutdown.exe"
            shutdown /r /f /t 0 /c "Profile Migration Tool - Completing operation" /d p:4:1
        }
        'Delayed' {
            Log-Message "Restarting in $DelaySeconds seconds (Behavior: Delayed)"
            $response = Show-ModernDialog -Message "$Reason`r`n`r`nComputer will restart in $DelaySeconds seconds.`r`n`r`nClick OK to restart now, Cancel to wait the full countdown." -Title "Restart Scheduled" -Type Warning -Buttons OKCancel

            if ($response -eq "OK") {
                Log-Message "User clicked OK - restarting immediately"
                shutdown /r /f /t 0 /c "Profile Migration Tool - User triggered restart" /d p:4:1
            }
            else {
                $global:StatusText.Text = "Restart scheduled in $DelaySeconds seconds... (cannot cancel)"
                Log-Message "Starting $DelaySeconds-second countdown"
                for ($i = $DelaySeconds; $i -gt 0; $i--) {
                    $global:StatusText.Text = "Restarting in $i seconds... (close window to attempt cancel)"
                    $global:ProgressBar.Value = [int](($DelaySeconds - $i) / $DelaySeconds * 100)
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Seconds 1
                }
                Log-Message "Countdown finished - forcing restart"
                shutdown /r /f /t 0 /c "Profile Migration Tool - Countdown complete" /d p:4:1
            }
        }
        'Never' {
            Log-Message "Restart required but automatic restart disabled (Behavior: Never)"
            $global:StatusText.Text = "Manual restart required"
            Show-ModernDialog -Message "$Reason`r`n`r`nRESTART REQUIRED but auto-restart is disabled.`r`n`r`nPlease restart manually when ready." -Title "Manual Restart Required" -Type Info -Buttons OK
        }
        'Prompt' {
            Log-Message "Prompting user for restart (Behavior: Prompt)"
            $response = Show-ModernDialog -Message "$Reason`r`n`r`nA restart is REQUIRED to complete the operation.`r`n`r`nRestart now?" -Title "Restart Required" -Type Question -Buttons YesNo

            if ($response -eq "Yes") {
                Log-Message "User approved restart - starting 10-second forced countdown"

                # This message box is purely informational - we ignore its result
                Show-ModernDialog -Message "Computer will restart in 10 seconds.`r`n`r`nSave all work NOW!`r`n`r`nThe restart cannot be cancelled." -Title "Restarting in 10 seconds" -Type Warning -Buttons OK | Out-Null

                # 10-second visible countdown + forced restart
                for ($i = 10; $i -gt 0; $i--) {
                    $global:StatusText.Text = "Restarting in $i seconds..."
                    $global:ProgressBar.Value = [int](100 - ($i * 10))
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Seconds 1
                }

                Log-Message "10-second countdown finished - forcing restart"
                shutdown /r /f /t 0 /c "Profile Migration Tool - User approved restart" /d p:4:1
            }
            else {
                Log-Message "User clicked No - forcing restart anyway (domain join requires it)"
                Show-ModernDialog -Message "You clicked No, but a restart is REQUIRED after domain join.`r`n`r`nRestarting in 15 seconds anyway..." -Title "Restart Mandatory" -Type Warning -Buttons OK | Out-Null

                for ($i = 15; $i -gt 0; $i--) {
                    $global:StatusText.Text = "Restart MANDATORY - restarting in $i seconds..."
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Seconds 1
                }

                shutdown /r /f /t 0 /c "Profile Migration Tool - Forced restart (domain join)" /d p:4:1
            }
        }
    }
}

# =============================================================================
# PRE-FLIGHT CHECKS
# =============================================================================
function Join-Domain-Enhanced {
    param(
        [Parameter(Mandatory = $false)][ValidateScript({ if ([string]::IsNullOrWhiteSpace($_)) { $true } else { if ($_ -notmatch '^[a-zA-Z0-9-]{1,15}$') { throw "Computer name must be 1-15 characters (alphanumeric and hyphens only)" } else { $true } } })][string]$ComputerName,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DomainName,
        [Parameter(Mandatory = $false)][ValidateSet('Prompt', 'Immediate', 'Never', 'Delayed')][string]$RestartBehavior = 'Prompt',
        [Parameter(Mandatory = $false)][ValidateRange(5, 300)][int]$DelaySeconds = 30,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential = $null
    )
    try {
        if ([string]::IsNullOrWhiteSpace($DomainName)) { throw "Domain name is required." }
        $cs = Get-WmiObject Win32_ComputerSystem
        $currentComputerName = $env:COMPUTERNAME
        Log-Message "Current State: Computer='$currentComputerName', Domain='$($cs.Domain)', PartOfDomain=$($cs.PartOfDomain)"
        $targetName = $null
        if (![string]::IsNullOrWhiteSpace($ComputerName)) {
            $targetName = $ComputerName.Trim()
            if ($targetName.Length -gt 15) { throw "Computer name cannot exceed 15 characters. Provided: $($targetName.Length) chars" }
            if ($targetName -match '[^a-zA-Z0-9-]') { throw "Computer name contains invalid characters. Only letters, numbers, and hyphens are allowed." }
            if ($targetName.StartsWith('-') -or $targetName.EndsWith('-')) { throw "Computer name cannot start or end with a hyphen." }
        }
        if ($cs.PartOfDomain -and $cs.Domain -ieq $DomainName) {
            Log-Message "Already member of domain '$DomainName'"
            if ($targetName -and $targetName.ToUpper() -ne $currentComputerName.ToUpper()) {
                $response = Show-ModernDialog -Message "Already in domain '$DomainName'.`r`n`r`nRename computer?`nFrom: $currentComputerName`nTo: $targetName" -Title "Rename Computer" -Type Question -Buttons YesNo
                if ($response -eq "Yes") {
                    $cred = Get-Credential -Message "Domain credentials to rename computer" -UserName "$DomainName\"
                    if (-not $cred) { throw "Credentials cancelled." }
                    Log-Message "Renaming: '$currentComputerName' to '$targetName'"
                    try {
                        Rename-Computer -NewName $targetName -DomainCredential $cred -Force -ErrorAction Stop
                        Handle-Restart -Behavior $RestartBehavior -DelaySeconds $DelaySeconds -Reason "Computer renamed to '$targetName'"
                    }
                    catch {
                        $errorDetails = Get-DomainJoinErrorDetails -ErrorRecord $_
                        Log-Message "Rename failed: $($errorDetails.UserFriendlyMessage)"
                        throw "Computer rename failed: $($errorDetails.UserFriendlyMessage)`n`nSuggestion: $($errorDetails.Suggestion)"
                    }
                }
            }
            else {
                Log-Message "No changes needed. Computer name: $currentComputerName"
            }
            return
        }
        if ($cs.PartOfDomain) {
            $currentDomain = $cs.Domain
            $response = Show-ModernDialog -Message "Currently in domain: $currentDomain`r`n`r`nLeave and join: $DomainName`r`n`r`nThis requires TWO RESTARTS.`r`n`r`r`nContinue?" -Title "Leave Domain" -Type Warning -Buttons YesNo
            if ($response -ne "Yes") {
                Log-Message "Operation cancelled."
                return
            }
            $disjoinCred = Get-Credential -Message "Credentials to leave '$currentDomain'" -UserName "$currentDomain\"
            if (-not $disjoinCred) { throw "Disjoin credentials cancelled." }
            Log-Message "Leaving domain '$currentDomain'"
            try {
                Remove-Computer -UnjoinDomainCredential $disjoinCred -Force -ErrorAction Stop
                Show-ModernDialog -Message "Left domain successfully.`r`n`r`nAfter restart, run this tool again to join '$DomainName'." -Title "Step 1 Complete" -Type Info -Buttons OK
                Handle-Restart -Behavior $RestartBehavior -DelaySeconds $DelaySeconds -Reason "Left domain '$currentDomain'"
            }
            catch {
                $errorDetails = Get-DomainJoinErrorDetails -ErrorRecord $_
                Log-Message "Failed to leave domain: $($errorDetails.UserFriendlyMessage)"
                throw "Failed to leave domain: $($errorDetails.UserFriendlyMessage)`r`n`r`nSuggestion: $($errorDetails.Suggestion)"
            }
            return
        }
        
        # CRITICAL: Check if device is AzureAD joined before attempting domain join
        # Windows does not allow domain join while AzureAD joined
        Log-Message "=== CHECKING AZUREAD JOIN STATUS ==="
        if (Test-IsAzureADJoined) {
            Log-Message ">>> Device IS AzureAD joined - unjoin is REQUIRED before domain join"
            
            $global:StatusText.Text = "AzureAD unjoin required..."
            [System.Windows.Forms.Application]::DoEvents()
            
            $message = "This device is currently joined to AzureAD.`r`n`r`n"
            $message += "Windows REQUIRES unjoining from AzureAD before joining a domain.`r`n`r`n"
            $message += "This will:`r`n"
            $message += "- Remove device from your organization's management`r`n"
            $message += "- Disable conditional access policies`r`n"
            $message += "- Remove SSO capabilities`r`n`r`n"
            $message += "After unjoin, domain join will proceed.`r`n`r`n"
            $message += "Continue with required AzureAD unjoin?"
            
            $response = Show-ModernDialog -Message $message -Title "AzureAD Unjoin Required" -Type Warning -Buttons YesNo
            
            if ($response -eq 'Yes') {
                Log-Message ">>> User confirmed - executing AzureAD unjoin..."
                $global:StatusText.Text = "Unjoining from AzureAD..."
                [System.Windows.Forms.Application]::DoEvents()
                
                $unjoinResult = Invoke-AzureADUnjoin
                Log-Message ">>> Unjoin result - Success: $($unjoinResult.Success), Message: $($unjoinResult.Message)"
                
                if ($unjoinResult.Success) {
                    Log-Message ">>> AzureAD unjoin SUCCESSFUL - proceeding with domain join"
                    $global:StatusText.Text = "AzureAD unjoin successful"
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Milliseconds 1000
                }
                else {
                    Log-Message ">>> AzureAD unjoin FAILED - cannot proceed with domain join"
                    $errorMsg = "AzureAD unjoin failed: $($unjoinResult.Message)`r`n`r`n"
                    $errorMsg += "Cannot proceed with domain join while AzureAD joined.`r`n`r`n"
                    $errorMsg += "You can manually unjoin using: dsregcmd /leave"
                    throw $errorMsg
                }
            }
            else {
                Log-Message ">>> User CANCELLED required AzureAD unjoin - aborting domain join"
                throw "Cannot join domain while AzureAD joined. Domain join cancelled."
            }
        }
        else {
            Log-Message ">>> Device is NOT AzureAD joined - proceeding with domain join"
        }
        
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "=== STARTING PRE-FLIGHT CHECKS ==="
        $global:StatusText.Text = "Checking domain reachability..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # --- REACHABILITY RETRY LOOP ---
        while ($true) {
            $reachabilityTest = Test-DomainReachability -DomainName $DomainName
            if ($reachabilityTest.Success) {
                break # Success, proceed
            }
            
            Log-Message "Pre-flight check failed: $($reachabilityTest.Error)"
            $res = Show-ModernDialog -Message "Domain Reachability Check Failed:`r`n`r`n$($reachabilityTest.Error)`r`n`r`nPlease verify:`r`n- Domain name is correct`r`n- Network connection is working`r`n- DNS is properly configured`r`n- Firewalls allow domain traffic`r`n`r`nRetry connection?" -Title "Domain Join Error" -Type Error -Buttons YesNo
            
            if ($res -eq "No") {
                Log-Message "User cancelled during reachability check."
                return # Exit cleanly, don't throw exception to avoid double error dialog
            }
            # Loop continues on Retry
            $global:StatusText.Text = "Retrying domain check..."
            [System.Windows.Forms.Application]::DoEvents()
        }
        Log-Message "Joining domain: $DomainName"
        # --- NEW CENTRALIZED CREDENTIAL LOGIC ---
        $joinCred = Get-DomainAdminCredential -DomainName $DomainName -InitialCredential $Credential
        
        if (-not $joinCred) {
            throw "Join credentials cancelled."
        }
        Log-Message "=== PRE-FLIGHT CHECKS COMPLETE ==="
        $params = @{ DomainName = $DomainName; Credential = $joinCred; Force = $true }
        $actionDescription = "Join domain '$DomainName'"
        if ($targetName -and $targetName.ToUpper() -ne $currentComputerName.ToUpper()) {
            $params['NewName'] = $targetName
            $actionDescription += "`r`nRename: $currentComputerName to $targetName"
            Log-Message "Will rename computer during domain join."
        }
        else {
            $actionDescription += "`r`nKeep name: $currentComputerName"
            Log-Message "Computer name will remain: $currentComputerName"
        }
        $confirmation = Show-ModernDialog -Message "$actionDescription`r`n`r`nA restart will be required.`r`n`r`nContinue?" -Title "Confirm Domain Join" -Type Question -Buttons YesNo
        if ($confirmation -ne "Yes") {
            Log-Message "Domain join cancelled."
            return
        }
        $global:StatusText.Text = "Joining domain..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Executing Add-Computer with parameters: Domain=$DomainName, NewName=$($params.NewName)"
        try {
            Add-Computer @params -ErrorAction Stop
            Log-Message "Successfully joined domain '$DomainName'!"
            Handle-Restart -Behavior $RestartBehavior -DelaySeconds $DelaySeconds -Reason "Joined domain '$DomainName'"
        }
        catch {
            $errorDetails = Get-DomainJoinErrorDetails -ErrorRecord $_
            Log-Message "Domain join failed: [$($errorDetails.ErrorCode)] $($errorDetails.UserFriendlyMessage)"
            Log-Message "Original error: $($errorDetails.OriginalMessage)"
            $errorMsg = "Domain Join Failed`n`nError: $($errorDetails.UserFriendlyMessage)`n`n"
            if ($errorDetails.Suggestion) { $errorMsg += "Suggestion: $($errorDetails.Suggestion)`n`n" }
            $errorMsg += "Error Code: $($errorDetails.ErrorCode)"
            throw $errorMsg
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Log-Message "ERROR: $errorMsg"
        $global:StatusText.Text = "Domain join failed"
        Show-ModernDialog -Message $errorMsg -Title "Domain Join Error" -Type Error -Buttons OK
    }
}

# === WINGET IMPORT USING NATIVE WINGET IMPORT COMMAND ===

function Show-ProfileCleanupWizard {
    param([string]$ProfilePath)
    
    Log-Message "=== PROFILE CLEANUP WIZARD ==="
    Log-Message "Analyzing profile: $ProfilePath"
    
    # Initialize theme
    $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
    $theme = $Themes[$currentTheme]
    
    # Create wizard form
    $wizForm = New-Object System.Windows.Forms.Form
    $wizForm.Text = "Profile Cleanup Wizard - Optimize Export Size"
    $wizForm.Size = New-Object System.Drawing.Size(950, 750)
    $wizForm.StartPosition = "CenterScreen"
    $wizForm.FormBorderStyle = "Sizable"
    $wizForm.MaximizeBox = $false
    $wizForm.MinimizeBox = $false
    $wizForm.TopMost = $true
    $wizForm.BackColor = $theme.FormBackColor
    $wizForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Header
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(950, 80)
    $headerPanel.BackColor = $theme.HeaderBackColor
    $wizForm.Controls.Add($headerPanel)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(900, 30)
    $lblTitle.Text = "Reduce Export Size - Cleanup Recommendations"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $theme.HeaderTextColor
    $headerPanel.Controls.Add($lblTitle)
    
    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Location = New-Object System.Drawing.Point(22, 48)
    $lblSubtitle.Size = New-Object System.Drawing.Size(900, 25)
    $lblSubtitle.Text = "Analyzing profile for temporary files, caches, and large files that can be safely excluded..."
    $lblSubtitle.ForeColor = $theme.SubHeaderTextColor
    $headerPanel.Controls.Add($lblSubtitle)
    
    # Progress bar for analysis
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 95)
    $progressBar.Size = New-Object System.Drawing.Size(900, 20)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.ForeColor = $theme.ProgressBarForeColor
    $wizForm.Controls.Add($progressBar)
    
    # Results panel
    $resultsPanel = New-Object System.Windows.Forms.Panel
    $resultsPanel.Location = New-Object System.Drawing.Point(15, 125)
    $resultsPanel.Size = New-Object System.Drawing.Size(910, 520)
    $resultsPanel.BackColor = $theme.PanelBackColor
    $resultsPanel.AutoScroll = $true
    $wizForm.Controls.Add($resultsPanel)
    
    # Show the form early to display analysis progress
    $wizForm.Show()
    [System.Windows.Forms.Application]::DoEvents()
    
    # Analyze profile (populate cleanup categories)
    $cleanupItems = @()
    $totalSavings = 0
    $yPos = 20
    
    # Category 1: Browser Caches
    $progressBar.Value = 10
    $lblSubtitle.Text = "Scanning browser caches..."
    [System.Windows.Forms.Application]::DoEvents()
    
    $browserPaths = @(
        @{Name = "Chrome Cache"; Path = "AppData\Local\Google\Chrome\User Data\Default\Cache" },
        @{Name = "Chrome Code Cache"; Path = "AppData\Local\Google\Chrome\User Data\Default\Code Cache" },
        @{Name = "Edge Cache"; Path = "AppData\Local\Microsoft\Edge\User Data\Default\Cache" },
        @{Name = "Edge Code Cache"; Path = "AppData\Local\Microsoft\Edge\User Data\Default\Code Cache" },
        @{Name = "Firefox Cache"; Path = "AppData\Local\Mozilla\Firefox\Profiles\*\cache2" },
        @{Name = "IE Cache"; Path = "AppData\Local\Microsoft\Windows\INetCache" }
    )
    
    $browserCacheSize = 0
    $browserCachePaths = @()
    foreach ($bp in $browserPaths) {
        $fullPath = Join-Path $ProfilePath $bp.Path
        if ($bp.Path -like "*`**") {
            # Wildcard path - expand it
            $parentPath = Split-Path (Join-Path $ProfilePath $bp.Path) -Parent
            if (Test-Path $parentPath) {
                $matchingDirs = Get-ChildItem $parentPath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.FullName -like $fullPath }
                foreach ($dir in $matchingDirs) {
                    try {
                        $size = (Get-ChildItem $dir.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                        if ($size -gt 0) {
                            $browserCacheSize += $size
                            $browserCachePaths += $dir.FullName
                        }
                    }
                    catch {}
                }
            }
        }
        else {
            if (Test-Path $fullPath) {
                try {
                    $size = (Get-ChildItem $fullPath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                    if ($size -gt 0) {
                        $browserCacheSize += $size
                        $browserCachePaths += $fullPath
                    }
                }
                catch {}
            }
        }
    }
    
    if ($browserCacheSize -gt 0) {
        $cleanupItems += New-CleanupItem `
            -Category "Browser Caches" `
            -Description "Temporary browser data (Chrome, Edge, Firefox, IE) - Safe to delete" `
            -Size $browserCacheSize `
            -Paths $browserCachePaths `
            -DefaultChecked $true
    }
    
    # Category 2: Windows Temp Files
    $progressBar.Value = 25
    $lblSubtitle.Text = "Scanning temporary files..."
    [System.Windows.Forms.Application]::DoEvents()
    
    $tempSize = 0
    $tempPaths = @()
    $tempLocations = @(
        "AppData\Local\Temp",
        "AppData\Local\Microsoft\Windows\Temporary Internet Files",
        "AppData\Local\Microsoft\Windows\INetCache\Content.IE5",
        "AppData\Local\Microsoft\Windows\WebCache"
    )
    
    foreach ($tl in $tempLocations) {
        $fullPath = Join-Path $ProfilePath $tl
        if (Test-Path $fullPath) {
            try {
                $size = (Get-ChildItem $fullPath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                if ($size -gt 0) {
                    $tempSize += $size
                    $tempPaths += $fullPath
                }
            }
            catch {}
        }
    }
    
    if ($tempSize -gt 0) {
        $cleanupItems += New-CleanupItem `
            -Category "Temporary Files" `
            -Description "Windows temporary files and cache - Safe to delete" `
            -Size $tempSize `
            -Paths $tempPaths `
            -DefaultChecked $true
    }
    
    # Category 3: Large Files (>100MB)
    $progressBar.Value = 50
    $lblSubtitle.Text = "Finding large files (>100MB)..."
    [System.Windows.Forms.Application]::DoEvents()
    
    $largeFiles = @()
    $largeFilesSize = 0
    # Define Downloads path early for exclusion
    $downloadsPath = Join-Path $ProfilePath 'Downloads'
    
    try {
        # Force array @() to handle scalar return when only 1 file is found
        $largeFiles = @(Get-ChildItem $ProfilePath -Recurse -File -ErrorAction SilentlyContinue | 
            Where-Object { 
                $_.Length -gt 100MB -and 
                -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -and
                $_.FullName -notmatch 'OneDrive|SharePoint|Cloud|Storage' -and
                $_.FullName -notlike "$downloadsPath\*"
            } | 
            Sort-Object Length -Descending |
            Select-Object -First 50 FullName, @{N = 'SizeMB'; E = { [math]::Round($_.Length / 1MB, 2) } })
        
        Log-Debug "Cleanup Wizard: Found $($largeFiles.Count) large files > 100MB"
        
        foreach ($lf in $largeFiles) {
            $largeFilesSize += ($lf.SizeMB * 1MB)
        }
    }
    catch {
        Log-Error "Cleanup Wizard: Error finding large files: $_"
    }
    
    if ($largeFiles.Count -gt 0) {
        $largeFileObjects = @()
        foreach ($lf in $largeFiles) {
            $largeFileObjects += [pscustomobject]@{
                Path        = $lf.FullName
                SizeMB      = $lf.SizeMB
                Size        = $lf.Length  # Added raw size for accurate calculation
                Selected    = $false
                DisplayName = "$(Split-Path $lf.FullName -Leaf) - $($lf.SizeMB) MB"
            }
        }
        
        $cleanupItems += New-CleanupItem `
            -Category "Large Files" `
            -Description "$($largeFiles.Count) files over 100MB - Review and exclude if not needed" `
            -Size $largeFilesSize `
            -Paths ($largeFiles | ForEach-Object { $_.FullName }) `
            -DefaultChecked $false `
            -Details ($largeFileObjects | ForEach-Object { $_.DisplayName }) `
            -HasIndividualSelection $true `
            -IndividualItems $largeFileObjects
    }
    
    # Category 4: Duplicate Files
    $progressBar.Value = 70
    $lblSubtitle.Text = "Detecting duplicate files..."
    [System.Windows.Forms.Application]::DoEvents()
    
    $duplicates = @{}
    $duplicateSize = 0
    try {
        $allFiles = Get-ChildItem $ProfilePath -Recurse -File -ErrorAction SilentlyContinue | 
        Where-Object { 
            $_.Length -gt 1MB -and $_.Length -lt 1GB -and
            -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -and
            $_.FullName -notmatch 'OneDrive|SharePoint|Cloud|Storage'
        } |
        Select-Object -First 1000 FullName, Length, @{N = 'Hash'; E = { $null } }
        
        # Group by size first (faster than hashing everything)
        $sizeGroups = $allFiles | Group-Object Length | Where-Object { $_.Count -gt 1 }
        
        $hashCount = 0
        foreach ($group in $sizeGroups) {
            if ($hashCount -gt 100) { break } # Limit hashing for performance
            [System.Windows.Forms.Application]::DoEvents()
            foreach ($file in $group.Group) {
                try {
                    $hash = (Get-FileHash $file.FullName -Algorithm MD5 -ErrorAction SilentlyContinue).Hash
                    if ($hash) {
                        if ($duplicates.ContainsKey($hash)) {
                            $duplicates[$hash] += @{Path = $file.FullName; Size = $file.Length }
                            $duplicateSize += $file.Length
                        }
                        else {
                            $duplicates[$hash] = @(@{Path = $file.FullName; Size = $file.Length })
                        }
                    }
                    $hashCount++
                }
                catch {}
            }
        }
        
        # Filter to only groups with duplicates
        $duplicates = $duplicates.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
    }
    catch {}
    
    if ($duplicates.Count -gt 0) {
        # Build duplicate groups with ALL files (including the one to keep)
        $duplicateGroups = @()
        $totalDuplicateFiles = 0
        $totalSavings = 0
        
        foreach ($dup in $duplicates) {
            $groupFiles = @()
            $fileName = Split-Path $dup.Value[0].Path -Leaf
            
            # Add ALL files in this duplicate group
            for ($i = 0; $i -lt $dup.Value.Count; $i++) {
                $file = $dup.Value[$i]
                $sizeMB = [math]::Round($file.Size / 1MB, 2)
                
                $groupFiles += [pscustomobject]@{
                    Path        = $file.Path
                    DisplayName = "$($file.Path) - $sizeMB MB"
                    Size        = $file.Size
                    IsKept      = ($i -eq 0)  # First one is default "keep"
                    Selected    = $false       # User can change selection
                }
                
                $totalDuplicateFiles++
                if ($i -gt 0) {
                    $totalSavings += $file.Size  # Potential savings if we exclude duplicates
                }
            }
            
            $duplicateGroups += @{
                FileName  = $fileName
                Hash      = $dup.Key
                Files     = $groupFiles
                KeptIndex = 0  # Index of file to keep (user can change)
                GroupSize = $groupFiles[0].Size * $groupFiles.Count
            }
        }
        
        $cleanupItems += New-CleanupItem `
            -Category "Duplicate Files" `
            -Description "$totalDuplicateFiles files in $($duplicateGroups.Count) groups - Select files to EXCLUDE" `
            -Size $totalSavings `
            -Paths @() `
            -DefaultChecked $false `
            -HasIndividualSelection $true `
            -DuplicateGroups $duplicateGroups

    }
    
    # Category 5: Recycle Bin
    $progressBar.Value = 85
    $lblSubtitle.Text = "Checking Recycle Bin..."
    
    # Enhanced Recycle Bin detection using SID
    $recyclePath = $null
    try {
        $username = Split-Path $ProfilePath -Leaf
        $userSid = Get-LocalUserSID -Username $username -ErrorAction SilentlyContinue
        if ($userSid) {
            $driveRoot = Split-Path $ProfilePath -Qualifier
            if (-not $driveRoot) { $driveRoot = "C:" } # Default if unc/unknown
            $sysRecyclePath = Join-Path "$driveRoot\" "`$RECYCLE.BIN\$userSid"
            if (Test-Path $sysRecyclePath) {
                $recyclePath = $sysRecyclePath
                Log-Debug "Found system Recycle Bin for user $username ($userSid) at $recyclePath"
            }
        }
    }
    catch {}

    if (-not $recyclePath) {
        $recyclePath = Join-Path $ProfilePath '$RECYCLE.BIN'
    }
    $recycleSize = 0
    if (Test-Path $recyclePath) {
        try {
            $recycleSize = (Get-ChildItem $recyclePath -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            if ($recycleSize -gt 0) {
                $cleanupItems += New-CleanupItem `
                    -Category "Recycle Bin" `
                    -Description "Deleted files in Recycle Bin - Safe to delete" `
                    -Size $recycleSize `
                    -Paths @($recyclePath) `
                    -DefaultChecked $true
            }
        }
        catch {}
    }
    
    # Category 6: Downloads folder (large files only)
    $progressBar.Value = 95
    $lblSubtitle.Text = "Analyzing Downloads folder..."
    [System.Windows.Forms.Application]::DoEvents()
    
    # $downloadsPath is already defined above
    $downloadSize = 0
    $downloadLargeFiles = @()
    if (Test-Path $downloadsPath) {
        try {
            # Force array @()
            $downloadLargeFiles = @(Get-ChildItem $downloadsPath -File -ErrorAction SilentlyContinue | 
                Where-Object { 
                    $_.Length -gt 50MB -and
                    -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -and
                    $_.FullName -notmatch 'OneDrive|SharePoint|Cloud|Storage'
                } |
                Sort-Object Length -Descending |
                Select-Object FullName, @{N = 'SizeMB'; E = { [math]::Round($_.Length / 1MB, 2) } })
            
            foreach ($dlf in $downloadLargeFiles) {
                $downloadSize += ($dlf.SizeMB * 1MB)
            }
            
            if ($downloadLargeFiles.Count -gt 0) {
                $cleanupItems += New-CleanupItem `
                    -Category "Large Downloads" `
                    -Description "$($downloadLargeFiles.Count) large files in Downloads (>50MB) - Review for exclusion" `
                    -Size $downloadSize `
                    -Paths ($downloadLargeFiles | ForEach-Object { $_.FullName }) `
                    -DefaultChecked $false `
                    -Details ($downloadLargeFiles | ForEach-Object { "$(Split-Path $_.FullName -Leaf) - $($_.SizeMB) MB" }) `
                    -HasIndividualSelection $true `
                    -IndividualItems ($downloadLargeFiles | ForEach-Object { 
                        [pscustomobject]@{
                            Path        = $_.FullName
                            DisplayName = "$(Split-Path $_.FullName -Leaf) - $($_.SizeMB) MB"
                            Size        = $_.Length  # Added raw size for accurate calculation
                            Selected    = $false
                        }
                    })
            }
        }

        catch {}
    }
    
    $progressBar.Value = 100
    $lblSubtitle.Text = "Analysis complete!"
    
    # Display results
    if ($cleanupItems.Count -eq 0) {
        $lblNoItems = New-Object System.Windows.Forms.Label
        $lblNoItems.Location = New-Object System.Drawing.Point(20, 20)
        $lblNoItems.Size = New-Object System.Drawing.Size(870, 60)
        $lblNoItems.Text = "No significant cleanup opportunities found.`n`nProfile appears to be already optimized."
        $lblNoItems.Font = New-Object System.Drawing.Font("Segoe UI", 11)
        $lblNoItems.ForeColor = $theme.ButtonSuccessBackColor
        $resultsPanel.Controls.Add($lblNoItems)
    }
    else {
        $checkboxes = @()
        
        foreach ($item in $cleanupItems) {
            # Category checkbox
            $chk = New-Object System.Windows.Forms.CheckBox
            $chk.Location = New-Object System.Drawing.Point(20, $yPos)
            $chk.Size = New-Object System.Drawing.Size(850, 25)
            $chk.Checked = $item.Checked
            $chk.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $chk.ForeColor = $theme.LabelTextColor
            
            $sizeMB = [math]::Round($item.Size / 1MB, 2)
            $sizeGB = [math]::Round($item.Size / 1GB, 2)
            $sizeStr = if ($sizeGB -gt 1) { "$sizeGB GB" } else { "$sizeMB MB" }
            
            $chk.Text = "$($item.Category) - $sizeStr - $($item.Description)"
            # Store base text for later label updates
            $item.BaseText = $chk.Text
            $chk.Tag = $item
            
            # Add Click handler for Select All / Deselect All behavior on main checkbox
            $chk.Add_Click({
                    $thisItem = $this.Tag
                    if ($thisItem.HasIndividualSelection) {
                        if ($this.Checked) {
                            # Select All
                            foreach ($subItem in $thisItem.IndividualItems) { $subItem.Selected = $true }
                            $paths = @($thisItem.IndividualItems | ForEach-Object { $_.Path })
                            $thisItem.SelectedIndividualPaths = $paths
                            $count = $paths.Count
                            Log-Debug "Select All clicked. Found $count paths. First: $($paths[0])"
                            $this.Text = "$($thisItem.BaseText) ($count selected)"
                        }
                        else {
                            # Deselect All
                            foreach ($subItem in $thisItem.IndividualItems) { $subItem.Selected = $false }
                            $thisItem.SelectedIndividualPaths = @()
                            $this.Text = $thisItem.BaseText
                        }
                    }
                })
            
            $resultsPanel.Controls.Add($chk)
            $checkboxes += $chk
            $yPos += 30
            
            # Individual file selection if available (including duplicate groups)
            if ($item.HasIndividualSelection -and ($item.IndividualItems -or $item.HasDuplicateGroups)) {
                # Individual checkboxes removed - selection handled via View Details dialog

                
                # Add View Details button for individual selection categories
                $btnDetails = New-Object System.Windows.Forms.Button
                $btnDetails.Location = New-Object System.Drawing.Point(40, $yPos)
                $btnDetails.Size = New-Object System.Drawing.Size(120, 25)
                $btnDetails.Text = "View Details"
                $btnDetails.BackColor = $theme.ButtonPrimaryBackColor
                $btnDetails.ForeColor = $theme.ButtonPrimaryForeColor
                $btnDetails.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $btnDetails.FlatAppearance.BorderSize = 0
                $btnDetails.Font = New-Object System.Drawing.Font("Segoe UI", 8)
                $btnDetails.Cursor = [System.Windows.Forms.Cursors]::Hand
                $btnDetails.Tag = @{ Item = $item; Checkbox = $chk; BaseText = $chk.Text }
                $btnDetails.Add_Click({
                        $ctx = $this.Tag
                        $currentItem = $ctx.Item
                        $currentChk = $ctx.Checkbox
                        
                        # Create a dialog with checkboxes for individual file selection
                        $detailsForm = New-Object System.Windows.Forms.Form
                        $detailsForm.Text = "$($currentItem.Category) - Select Files"
                        $detailsForm.Size = New-Object System.Drawing.Size(600, 520)
                        $detailsForm.StartPosition = "CenterScreen"
                        $detailsForm.FormBorderStyle = "FixedDialog"
                        $detailsForm.MaximizeBox = $false
                        $detailsForm.MinimizeBox = $false
                        $detailsForm.TopMost = $true
                        $detailsForm.BackColor = $theme.FormBackColor
                        $detailsForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                    
                        # Header
                        $headerPanel = New-Object System.Windows.Forms.Panel
                        $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
                        $headerPanel.Size = New-Object System.Drawing.Size(600, 60)
                        $headerPanel.BackColor = $theme.HeaderBackColor
                        $detailsForm.Controls.Add($headerPanel)
                    
                        $lblTitle = New-Object System.Windows.Forms.Label
                        $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
                        $lblTitle.Size = New-Object System.Drawing.Size(560, 30)
                        $lblTitle.Text = "$($currentItem.Category) - Select Files to Clean"
                        $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
                        $lblTitle.ForeColor = $theme.HeaderTextColor
                        $headerPanel.Controls.Add($lblTitle)
                    
                        # Content panel
                        $contentPanel = New-Object System.Windows.Forms.Panel
                        $contentPanel.Location = New-Object System.Drawing.Point(15, 75)
                        $contentPanel.Size = New-Object System.Drawing.Size(560, 350)
                        $contentPanel.BackColor = $theme.PanelBackColor
                        $contentPanel.AutoScroll = $true
                        $detailsForm.Controls.Add($contentPanel)
                    
                        # Check if this is a duplicate files category (special rendering)
                        if ($currentItem.HasDuplicateGroups -and $currentItem.DuplicateGroups) {
                            # SPECIAL RENDERING FOR DUPLICATE GROUPS
                            $yPosDetails = 15
                            
                            # Add helper text
                            $lblHelper = New-Object System.Windows.Forms.Label
                            $lblHelper.Location = New-Object System.Drawing.Point(15, $yPosDetails)
                            $lblHelper.Size = New-Object System.Drawing.Size(530, 40)
                            $lblHelper.Text = "For each group, check the boxes for files you want to EXCLUDE from export. Unchecked files will be kept."
                            $lblHelper.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
                            $lblHelper.ForeColor = $theme.ButtonPrimaryBackColor
                            $contentPanel.Controls.Add($lblHelper)
                            $yPosDetails += 50
                            
                            $fileCheckboxes = @()
                            $groupIndex = 0
                            
                            foreach ($group in $currentItem.DuplicateGroups) {
                                # Group header
                                $lblGroup = New-Object System.Windows.Forms.Label
                                $lblGroup.Location = New-Object System.Drawing.Point(15, $yPosDetails)
                                $lblGroup.Size = New-Object System.Drawing.Size(530, 25)
                                $lblGroup.Text = "Group $($groupIndex + 1): $($group.FileName) ($($group.Files.Count) copies)"
                                $lblGroup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                                $lblGroup.ForeColor = $theme.LabelTextColor
                                $contentPanel.Controls.Add($lblGroup)
                                $yPosDetails += 30
                                
                                # Create checkboxes for each file
                                for ($i = 0; $i -lt $group.Files.Count; $i++) {
                                    $file = $group.Files[$i]
                                    
                                    # Checkbox for "exclude" selection
                                    $chkFile = New-Object System.Windows.Forms.CheckBox
                                    $chkFile.Location = New-Object System.Drawing.Point(30, $yPosDetails)
                                    $chkFile.Size = New-Object System.Drawing.Size(500, 20)
                                    $chkFile.Checked = $file.Selected
                                    $chkFile.Text = $file.DisplayName
                                    $chkFile.Font = New-Object System.Drawing.Font("Segoe UI", 8)
                                    $chkFile.ForeColor = $theme.LabelTextColor
                                    $chkFile.Tag = @{ GroupIndex = $groupIndex; FileIndex = $i; File = $file }
                                    $contentPanel.Controls.Add($chkFile)
                                    $fileCheckboxes += $chkFile
                                    
                                    $yPosDetails += 25
                                }
                                
                                $yPosDetails += 10  # Space between groups
                                $groupIndex++
                            }
                        }
                        else {
                            # REGULAR RENDERING FOR NON-DUPLICATE CATEGORIES
                            # Checkboxes for individual files
                            $fileCheckboxes = @()
                            $yPosDetails = 15
                            foreach ($fileItem in $currentItem.IndividualItems) {
                                $chkFile = New-Object System.Windows.Forms.CheckBox
                                $chkFile.Location = New-Object System.Drawing.Point(15, $yPosDetails)
                                $chkFile.Size = New-Object System.Drawing.Size(530, 20)
                                $chkFile.Checked = $fileItem.Selected
                                $chkFile.Text = $fileItem.DisplayName
                                $chkFile.Font = New-Object System.Drawing.Font("Segoe UI", 8)
                                $chkFile.ForeColor = $theme.LabelTextColor
                                $chkFile.Tag = $fileItem
                                $contentPanel.Controls.Add($chkFile)
                                $fileCheckboxes += $chkFile
                                $yPosDetails += 25
                            }
                        }
                    
                        # Select All / Select None buttons
                        $btnSelectAll = New-Object System.Windows.Forms.Button
                        $btnSelectAll.Location = New-Object System.Drawing.Point(15, 430)
                        $btnSelectAll.Size = New-Object System.Drawing.Size(100, 35)
                        $btnSelectAll.Text = "Select All"
                        $btnSelectAll.BackColor = $theme.ButtonPrimaryBackColor
                        $btnSelectAll.ForeColor = $theme.ButtonPrimaryForeColor
                        $btnSelectAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                        $btnSelectAll.FlatAppearance.BorderSize = 0
                        $btnSelectAll.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                        $btnSelectAll.Cursor = [System.Windows.Forms.Cursors]::Hand
                        $btnSelectAll.Add_Click({ foreach ($chk in $fileCheckboxes) { $chk.Checked = $true } })
                        $detailsForm.Controls.Add($btnSelectAll)
                    
                        $btnSelectNone = New-Object System.Windows.Forms.Button
                        $btnSelectNone.Location = New-Object System.Drawing.Point(125, 430)
                        $btnSelectNone.Size = New-Object System.Drawing.Size(100, 35)
                        $btnSelectNone.Text = "Select None"
                        $btnSelectNone.BackColor = $theme.ButtonSecondaryBackColor
                        $btnSelectNone.ForeColor = $theme.ButtonSecondaryForeColor
                        $btnSelectNone.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                        $btnSelectNone.FlatAppearance.BorderSize = 0
                        $btnSelectNone.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                        $btnSelectNone.Cursor = [System.Windows.Forms.Cursors]::Hand
                        $btnSelectNone.Add_Click({ foreach ($chk in $fileCheckboxes) { $chk.Checked = $false } })
                        $detailsForm.Controls.Add($btnSelectNone)
                    
                        # OK and Cancel buttons (aligned to right with 15px margin to match Select All)
                        $btnOK = New-Object System.Windows.Forms.Button
                        $btnOK.Location = New-Object System.Drawing.Point(359, 430)
                        $btnOK.Size = New-Object System.Drawing.Size(100, 35)
                        $btnOK.Text = "OK"
                        $btnOK.BackColor = $theme.ButtonSuccessBackColor
                        $btnOK.ForeColor = $theme.ButtonSuccessForeColor
                        $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                        $btnOK.FlatAppearance.BorderSize = 0
                        $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                        $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
                        $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
                        $detailsForm.Controls.Add($btnOK)
                    
                        $btnCancel = New-Object System.Windows.Forms.Button
                        $btnCancel.Location = New-Object System.Drawing.Point(469, 430)
                        $btnCancel.Size = New-Object System.Drawing.Size(100, 35)
                        $btnCancel.Text = "Cancel"
                        $btnCancel.BackColor = $theme.ButtonSecondaryBackColor
                        $btnCancel.ForeColor = $theme.ButtonSecondaryForeColor
                        $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                        $btnCancel.FlatAppearance.BorderSize = 0
                        $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                        $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
                        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                        $detailsForm.Controls.Add($btnCancel)
                    
                        $detailsForm.AcceptButton = $btnOK
                        $detailsForm.CancelButton = $btnCancel
                    
                        $result = $detailsForm.ShowDialog()
                        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                            # Check if this is a duplicate files category
                            if ($currentItem.HasDuplicateGroups -and $currentItem.DuplicateGroups) {
                                # Process duplicate groups
                                $currentItem.SelectedIndividualPaths = @()
                                $selectedPaths = @()
                                
                                # First, update the selection state based on checkboxes
                                foreach ($chk in $fileCheckboxes) {
                                    $tag = $chk.Tag
                                    if ($tag -and $tag.GroupIndex -ge 0 -and $tag.FileIndex -ge 0) {
                                        # Update the specific file in the duplicate groups structure
                                        $currentItem.DuplicateGroups[$tag.GroupIndex].Files[$tag.FileIndex].Selected = $chk.Checked
                                    }
                                }
                                
                                # Now rebuild the list of selected paths (exclusions)
                                foreach ($group in $currentItem.DuplicateGroups) {
                                    for ($i = 0; $i -lt $group.Files.Count; $i++) {
                                        $file = $group.Files[$i]
                                        # If selected, add to exclusion list
                                        if ($file.Selected) {
                                            $selectedPaths += $file.Path
                                        }
                                    }
                                }
                                
                                $currentItem.SelectedIndividualPaths = @($selectedPaths)
                                $currentItem.Paths = @($selectedPaths)  # Update Paths for exclusion
                                
                                Log-Debug "Duplicate groups processed. Excluding $($currentItem.SelectedIndividualPaths.Count) files"
                            }
                            else {
                                # Regular processing for non-duplicate categories
                                $currentItem.SelectedIndividualPaths = @()
                                $selectedPaths = @()
                                for ($i = 0; $i -lt $fileCheckboxes.Count; $i++) {
                                    $currentItem.IndividualItems[$i].Selected = $fileCheckboxes[$i].Checked
                                    if ($fileCheckboxes[$i].Checked) {
                                        $selectedPaths += $currentItem.IndividualItems[$i].Path
                                    }
                                }
                                # Force array type to prevent string decay
                                $currentItem.SelectedIndividualPaths = @($selectedPaths)
                                
                                Log-Debug "Details dialog OK. Selected $($currentItem.SelectedIndividualPaths.Count) paths for $($currentItem.Category)"
                            }
                            
                            # Update the main checkbox state based on individual selections
                            $hasSelection = ($currentItem.SelectedIndividualPaths.Count -gt 0)
                            $currentItem.Checked = $hasSelection
                            $currentChk.Checked = $hasSelection
                            
                            # Update label to show count
                            if ($currentItem.SelectedIndividualPaths.Count -gt 0) {
                                $currentChk.Text = "$($ctx.BaseText) ($($currentItem.SelectedIndividualPaths.Count) selected)"
                            }
                            else {
                                $currentChk.Text = $ctx.BaseText
                            }
                        }
                        $detailsForm.Dispose()
                    })
                $resultsPanel.Controls.Add($btnDetails)
                $yPos += 35
                # For categories without individual selection, show View Details button
                if (!$item.HasIndividualSelection) {
                    $btnDetails = New-Object System.Windows.Forms.Button
                    $btnDetails.Location = New-Object System.Drawing.Point(40, $yPos)
                    $btnDetails.Size = New-Object System.Drawing.Size(120, 25)
                    $btnDetails.Text = "View Details"
                    $btnDetails.BackColor = $theme.ButtonPrimaryBackColor
                    $btnDetails.ForeColor = $theme.ButtonPrimaryForeColor
                    $btnDetails.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $btnDetails.FlatAppearance.BorderSize = 0
                    $btnDetails.Font = New-Object System.Drawing.Font("Segoe UI", 8)
                    $btnDetails.Cursor = [System.Windows.Forms.Cursors]::Hand
                    $btnDetails.Tag = $item
                    $btnDetails.Add_Click({
                            $currentItem = $this.Tag
                            # Create an advanced dialog with scrollable details
                            $detailsForm = New-Object System.Windows.Forms.Form
                            $detailsForm.Text = "$($currentItem.Category) - Details"
                            $detailsForm.Size = New-Object System.Drawing.Size(600, 500)
                            $detailsForm.StartPosition = "CenterScreen"
                            $detailsForm.FormBorderStyle = "FixedDialog"
                            $detailsForm.MaximizeBox = $false
                            $detailsForm.MinimizeBox = $false
                            $detailsForm.TopMost = $true
                            $detailsForm.BackColor = $theme.FormBackColor
                            $detailsForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                        
                            # Header panel
                            $headerPanel = New-Object System.Windows.Forms.Panel
                            $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
                            $headerPanel.Size = New-Object System.Drawing.Size(600, 60)
                            $headerPanel.BackColor = $theme.HeaderBackColor
                            $detailsForm.Controls.Add($headerPanel)
                        
                            $lblTitle = New-Object System.Windows.Forms.Label
                            $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
                            $lblTitle.Size = New-Object System.Drawing.Size(560, 30)
                            $lblTitle.Text = "$($currentItem.Category) - File Details"
                            $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
                            $lblTitle.ForeColor = $theme.HeaderTextColor
                            $headerPanel.Controls.Add($lblTitle)
                        
                            # Content panel
                            $contentPanel = New-Object System.Windows.Forms.Panel
                            $contentPanel.Location = New-Object System.Drawing.Point(15, 75)
                            $contentPanel.Size = New-Object System.Drawing.Size(560, 350)
                            $contentPanel.BackColor = $theme.PanelBackColor
                            $detailsForm.Controls.Add($contentPanel)
                        
                            # Scrollable text box for details
                            $txtDetails = New-Object System.Windows.Forms.TextBox
                            $txtDetails.Multiline = $true
                            $txtDetails.ReadOnly = $true
                            $txtDetails.ScrollBars = "Vertical"
                            $txtDetails.Location = New-Object System.Drawing.Point(15, 15)
                            $txtDetails.Size = New-Object System.Drawing.Size(530, 320)
                            $txtDetails.Font = New-Object System.Drawing.Font("Consolas", 9)
                            $txtDetails.BackColor = $theme.LogBoxBackColor
                            $txtDetails.ForeColor = $theme.LogBoxForeColor
                            $txtDetails.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                            $txtDetails.Text = $this.Tag.Details -join "`r`n"
                            $contentPanel.Controls.Add($txtDetails)
                        
                            # Close button
                            $btnClose = New-Object System.Windows.Forms.Button
                            $btnClose.Location = New-Object System.Drawing.Point(475, 430)
                            $btnClose.Size = New-Object System.Drawing.Size(100, 35)
                            $btnClose.Text = "Close"
                            $btnClose.BackColor = $theme.ButtonSecondaryBackColor
                            $btnClose.ForeColor = $theme.ButtonSecondaryForeColor
                            $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                            $btnClose.FlatAppearance.BorderSize = 0
                            $btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                            $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
                            $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
                            $detailsForm.Controls.Add($btnClose)
                        
                            $detailsForm.AcceptButton = $btnClose
                        
                            $detailsForm.ShowDialog() | Out-Null
                            $detailsForm.Dispose()
                        })
                    $resultsPanel.Controls.Add($btnDetails)
                    $yPos += 35
                }
            }
            else {
                $yPos += 10
            }
            
            $totalSavings += $item.Size
        }
        
        # Summary at top
        $savingsGB = [math]::Round($totalSavings / 1GB, 2)
        $savingsMB = [math]::Round($totalSavings / 1MB, 2)
        $savingsStr = if ($savingsGB -gt 1) { "$savingsGB GB" } else { "$savingsMB MB" }
        
        $lblSummary = New-Object System.Windows.Forms.Label
        $lblSummary.Location = New-Object System.Drawing.Point(20, 655)
        $lblSummary.Size = New-Object System.Drawing.Size(500, 30)
        $lblSummary.Text = "Potential savings: $savingsStr ($($cleanupItems.Count) categories)"
        $lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
        $lblSummary.ForeColor = $theme.ButtonSuccessBackColor
        $wizForm.Controls.Add($lblSummary)
    }
    
    # Action buttons
    $btnClean = New-Object System.Windows.Forms.Button
    $btnClean.Location = New-Object System.Drawing.Point(550, 655)
    $btnClean.Size = New-Object System.Drawing.Size(180, 40)
    $btnClean.Text = "Clean & Export"
    $btnClean.BackColor = $theme.ButtonSuccessBackColor
    $btnClean.ForeColor = $theme.ButtonSuccessForeColor
    $btnClean.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClean.FlatAppearance.BorderSize = 0
    $btnClean.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnClean.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClean.DialogResult = "OK"
    $btnClean.Add_MouseEnter({ $this.BackColor = $theme.ButtonSuccessHoverBackColor })
    $btnClean.Add_MouseLeave({ $this.BackColor = $theme.ButtonSuccessBackColor })
    $wizForm.Controls.Add($btnClean)
    
    $btnSkip = New-Object System.Windows.Forms.Button
    $btnSkip.Location = New-Object System.Drawing.Point(740, 655)
    $btnSkip.Size = New-Object System.Drawing.Size(180, 40)
    $btnSkip.Text = "Skip Cleanup"
    $btnSkip.BackColor = $theme.ButtonSecondaryBackColor
    $btnSkip.ForeColor = $theme.ButtonSecondaryForeColor
    $btnSkip.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSkip.FlatAppearance.BorderSize = 0
    $btnSkip.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnSkip.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnSkip.DialogResult = "Ignore"
    $btnSkip.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
    $btnSkip.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
    $wizForm.Controls.Add($btnSkip)
    
    $wizForm.AcceptButton = $btnClean
    $wizForm.CancelButton = $btnSkip
    
    # After analysis is complete and UI is ready, hide and show as modal
    $wizForm.Hide()
    $result = $wizForm.ShowDialog()
    
    # Process cleanup if user clicked Clean & Export
    $cleanupPaths = @()
    $selectedCategories = @()
    $totalSelectedSavings = 0
    $categoryExclusions = @{}
    
    if ($result -eq "OK") {
        foreach ($chk in $checkboxes) {
            if ($chk.Checked -and $chk.Tag) {
                $item = $chk.Tag
                $selectedCategories += $item.Category
                
                # Calculate True Selected Size
                if ($item.HasIndividualSelection) {
                    $selectedSize = 0
                    if ($item.HasDuplicateGroups) {
                        # Iterate duplicate groups
                        foreach ($group in $item.DuplicateGroups) {
                            foreach ($file in $group.Files) {
                                if ($file.Selected) { $selectedSize += $file.Size }
                            }
                        }
                    }
                    elseif ($item.IndividualItems) {
                        # Iterate individual items (Large Files, Downloads)
                        foreach ($subItem in $item.IndividualItems) {
                            if ($subItem.Selected) { $selectedSize += $subItem.Size }
                        }
                    }
                    $totalSelectedSavings += $selectedSize
                }
                else {
                    # No individual selection - take whole category
                    $totalSelectedSavings += $item.Size
                }

                Log-Message "Selected for cleanup: $($item.Category) - $([math]::Round($item.Size/1MB,2)) MB (Actual selected: $([math]::Round($selectedSize/1MB,2)) MB)"
                
                # For categories with individual selection, collect only selected files
                if ($item.HasIndividualSelection -and $item.SelectedIndividualPaths) {
                    $cleanupPaths += $item.SelectedIndividualPaths
                    $categoryExclusions[$item.Category] = $item.SelectedIndividualPaths
                    Log-Message "  - Total individual files selected: $($item.SelectedIndividualPaths.Count)"
                    foreach ($path in $item.SelectedIndividualPaths) {
                        Log-Message "  - Selected individual file: $(Split-Path $path -Leaf)"
                    }
                }
                else {
                    # For categories without individual selection, include all paths
                    $cleanupPaths += $item.Paths
                    $categoryExclusions[$item.Category] = $item.Paths
                }
            }
        }
        
        if ($cleanupPaths.Count -gt 0) {
            Log-Message "Total paths to exclude from export: $($cleanupPaths.Count)"
            Log-Message "Total space to be saved: $([math]::Round($totalSelectedSavings/1MB,2)) MB"
        }
    }
    
    $wizForm.Dispose()
    
    return @{
        Proceed            = ($result -eq "OK" -or $result -eq "Ignore")
        CleanupPaths       = $cleanupPaths
        CleanupCategories  = $selectedCategories
        CategoryExclusions = $categoryExclusions
        TotalSavingsMB     = [math]::Round($totalSelectedSavings / 1MB, 2)
    }
}

# =============================================================================
# EXPORT 
# =============================================================================
function Export-UserProfile {
    param([string]$Username, [string]$ZipPath)

    $global:ProgressBar.Value = 0
    $global:StatusText.Text = "Preparing export..."
    $global:CancelRequested = $false
    $global:CurrentOperation = "Export"
    if ($global:CancelButton) { $global:CancelButton.Enabled = $true }
    [System.Windows.Forms.Application]::DoEvents()

    # Start persistent logging for export operation
    Start-OperationLog -Operation "Export" -Username $Username

    try {
        $global:ExportStartTime = [DateTime]::Now
        Log-Message "=== EXPORT START: $Username -> $ZipPath"
        Log-Message "Operation started at: $($global:ExportStartTime.ToString('yyyy-MM-dd HH:mm:ss'))"

        # Extract short name from DOMAIN\username or COMPUTERNAME\username for folder lookup
        $shortName = if ($Username -match '\\') { ($Username -split '\\', 2)[1] } else { $Username }
        
        $userProfile = Get-LocalProfiles | Where-Object Username -eq $shortName | Select-Object -First 1
        if (-not $userProfile) { throw "Profile not found: $shortName" }
        
        # Store source path for reporting
        $source = $userProfile.Path

        if ($global:CancelRequested) { throw "Operation cancelled by user" }

        # === PROFILE SIZE OPTIMIZATION - CLEANUP WIZARD ===
        # Show cleanup wizard to let user reduce export size
        $cleanupResult = Show-ProfileCleanupWizard -ProfilePath $userProfile.Path
        
        if (-not $cleanupResult.Proceed) {
            Log-Message "Export cancelled - user declined cleanup wizard"
            throw "Export cancelled by user"
        }
        
        $userCleanupPaths = $cleanupResult.CleanupPaths
        $cleanupCategories = $cleanupResult.CleanupCategories  # Store for report
        $categoryExclusions = $cleanupResult.CategoryExclusions # Store for detailed reporting
        $cleanupSavingsMB = $cleanupResult.TotalSavingsMB      # Store for report
        
        if ($userCleanupPaths.Count -gt 0) {
            Log-Message "User selected $($userCleanupPaths.Count) paths for exclusion from export"
            Log-Message "Estimated space savings: $cleanupSavingsMB MB"
        }
        else {
            Log-Message "No cleanup paths selected - proceeding with standard export"
        }

        # Validate Profile Size and Disk Space
        $estSize = 5GB # Fallback
        try {
            $estSize = Get-ChildItem $userProfile.Path -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum | Select-Object -ExpandProperty Sum
        }
        catch { }

        # Check destination space
        try {
            $destRoot = [IO.Path]::GetPathRoot($ZipPath)
            $destDrive = Get-PSDrive -Name $destRoot.TrimEnd('\', ':') -ErrorAction Stop
             
            # Risk factor: 1.5x profile size (compression uses temp space, but specific drive depends on temp config)
            # We check TEMP drive and Destination drive.
             
            $destFree = $destDrive.Free
            Log-Info "Disk space check: $([Math]::Round($destFree/1GB,2)) GB free on $destRoot"

            if ($destFree -lt ($estSize * 0.5)) {
                # conservative check for final ZIP
                throw "Insufficient disk space on destination '$destRoot'. Available: $([Math]::Round($destFree/1GB,2)) GB. Required (approx): $([Math]::Round(($estSize*0.5)/1GB,2)) GB."
            }
        }
        catch {
            if ($_.Exception.Message -like "*Insufficient*") { throw $_ }
            Log-Warning "Could not verify destination disk space: $_"
        }

        $ts = Get-Date -f yyyyMMdd_HHmmss
        $tmp = Join-Path ([IO.Path]::GetDirectoryName($ZipPath)) "$shortName-Export-$ts"
        New-Item -ItemType Directory $tmp -Force | Out-Null
        Log-Message "Temporary export directory: $tmp"

        # Use 7-Zip to compress profile folder directly (handles locked files better than robocopy)
        $global:StatusText.Text = "Compressing profile with 7-Zip..."
        Log-Message "Compressing profile folder directly with 7-Zip (no robocopy)"
        $global:ProgressBar.Value = 10
        [System.Windows.Forms.Application]::DoEvents()

        # Build exclusion filters for 7-Zip
        $exclusions = Get-RobocopyExclusions -Mode 'Export'
        $7zExclusions = @()      # Non-recursive patterns (-x)
        $7zExclusionsRec = @()   # Recursive patterns (-xr)
        
        # Add user-selected cleanup paths to exclusions
        foreach ($cleanupPath in $userCleanupPaths) {
            if (-not [string]::IsNullOrWhiteSpace($cleanupPath)) {
                $relPath = $cleanupPath
                if ($cleanupPath -like "$($userProfile.Path)\*") {
                    $relPath = $cleanupPath.Substring($userProfile.Path.Length + 1)
                }
                
                # Handle both files and directories
                if (Test-Path $cleanupPath -PathType Container) {
                    # Directory - exclude it and all contents recursively
                    $7zExclusionsRec += "$relPath\*"
                    $7zExclusions += $relPath
                    Log-Message "Excluding cleanup directory: $relPath"
                }
                else {
                    # File - exclude just the file
                    $7zExclusions += $relPath
                    Log-Message "Excluding cleanup file: $relPath"
                }
            }
        }
        
        # Add standard exclusions
        foreach ($file in $exclusions.Files) {
            if (-not [string]::IsNullOrWhiteSpace($file)) {
                $relFile = $file
                if ($userProfile -and $userProfile.Path -and $file -like ("$($userProfile.Path)\*")) {
                    $relFile = $file.Substring($userProfile.Path.Length + 1)
                }
                $7zExclusions += $relFile
            }
        }
        foreach ($dir in $exclusions.Dirs) {
            if (-not [string]::IsNullOrWhiteSpace($dir)) {
                $relDir = $dir
                if ($userProfile -and $userProfile.Path -and $dir -like ("$($userProfile.Path)\*")) {
                    $relDir = $dir.Substring($userProfile.Path.Length + 1)
                }
                $patternContents = if ($relDir.EndsWith('*')) { $relDir } else { "$relDir\*" }
                # Exclude the directory entry and all its contents recursively
                $7zExclusionsRec += $patternContents
                $7zExclusions += $relDir
            }
        }

        # Explicitly exclude common legacy junctions under Documents
        $junctions = @(
            'Documents\My Music',
            'Documents\My Pictures',
            'Documents\My Videos'
        )
        foreach ($j in $junctions) {
            # Exclude junction directory entry and its contents recursively
            $7zExclusionsRec += "$j\*"
            $7zExclusions += $j
        }

        # Discover and exclude all reparse-point directories (junctions) under the profile root
        try {
            $rpItems = Get-ChildItem -Path $userProfile.Path -Directory -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }
            foreach ($item in $rpItems) {
                # Ensure we get path relative to the profile root
                $rel = $item.FullName.Substring($userProfile.Path.Length).TrimStart('\')
                if ([string]::IsNullOrWhiteSpace($rel)) { continue }
                if ($7zExclusionsRec -notcontains "$rel\*") { $7zExclusionsRec += "$rel\*" }
                if ($7zExclusions -notcontains $rel) { $7zExclusions += $rel }
            }
            if ($rpItems.Count -gt 0) {
                Log-Message "7-Zip export: excluding $($rpItems.Count) reparse points (junctions)"
            }
        }
        catch {
            Log-Message "WARNING: Junction discovery failed: $_"
        }
        
        # Exclude OneDrive/SharePoint/ReparsePoint items (prevents "cloud file provider is not running" errors)
        try {
            Log-Message "Scanning for cloud provider folders and reparse points..."
            $foundCloudExcludes = @()
            $topLevelCloudFolders = @()
            $scanQueue = New-Object System.Collections.Generic.Queue[string]
            $scanQueue.Enqueue($userProfile.Path)
            
            $cloudItemsCount = 0
            # Manual recursion to ensure we catch reparse-point directories (which Get-ChildItem -Recurse skips)
            while ($scanQueue.Count -gt 0) {
                $currPath = $scanQueue.Dequeue()
                $items = Get-ChildItem -Path $currPath -Force -ErrorAction SilentlyContinue
                foreach ($item in $items) {
                    $isReparse = ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)
                    $isCloudName = ($item.FullName -match 'OneDrive|SharePoint|Cloud|Storage')
                    
                    if ($isReparse -or $isCloudName) {
                        $foundCloudExcludes += $item
                        # If it's a directory, do NOT recurse into it
                        
                        # Use relative path for reporting (only for user-facing cloud folders)
                        # Skip internal AppData, hidden technical folders, and standard system junctions
                        if ($item.PSIsContainer -and ($isCloudName -or $isReparse)) {
                            $rel = $item.FullName.Substring($userProfile.Path.Length).TrimStart('\')
                            $systemJunctions = '^(AppData|Local Settings|Application Data|Cookies|Templates|NetHood|PrintHood|SendTo|Start Menu|Recent|Links|Searches|My Documents|LocalState|Documents\\My)'
                            if (-not [string]::IsNullOrWhiteSpace($rel) -and $rel -notmatch "^(\.|$systemJunctions)") {
                                $topLevelCloudFolders += $rel
                            }
                        }
                    }
                    elseif ($item.PSIsContainer) {
                        $scanQueue.Enqueue($item.FullName)
                    }
                    if ($cloudItemsCount++ % 100 -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
                }
            }
            
            $finalExclusionCount = 0
            foreach ($item in $foundCloudExcludes) {
                # Ensure we get path relative to the profile root
                $rel = $item.FullName.Substring($userProfile.Path.Length).TrimStart('\\')
                if ([string]::IsNullOrWhiteSpace($rel)) { continue }
                
                # We add multiple pattern variations to ensure 7-Zip matches regardless of how it sees the path
                if ($item.PSIsContainer) {
                    # Recursive patterns (folders) - relative paths are best for 7z
                    foreach ($p in @("$rel\*", "$rel", "*\$rel\*", "*\$rel")) {
                        if ($7zExclusionsRec -notcontains $p) { $7zExclusionsRec += $p }
                    }
                    $finalExclusionCount++
                }
                else {
                    # File patterns
                    foreach ($p in @("$rel", "*\$rel")) {
                        if ($7zExclusions -notcontains $p) { $7zExclusions += $p }
                    }
                    $finalExclusionCount++
                }
            }
            
            if ($finalExclusionCount -gt 0) {
                Log-Message "Excluded $finalExclusionCount cloud-linked root items from 7-Zip"
            }
        }
        catch {
            Log-Message "WARNING: Cloud file detection failed: $_"
        }
        
        # Detect and log PST files (Outlook personal folders/archives)
        try {
            $pstFiles = Get-ChildItem -Path $userProfile.Path -Filter "*.pst" -Recurse -Force -ErrorAction SilentlyContinue
            if ($pstFiles.Count -gt 0) {
                Log-Message "========================================="
                Log-Message "OUTLOOK PST FILES DETECTED: $($pstFiles.Count) file(s)"
                Log-Message "========================================="
                foreach ($pst in $pstFiles) {
                    $pstSize = [math]::Round($pst.Length / 1MB, 2)
                    $relPath = $pst.FullName.Substring($userProfile.Path.Length).TrimStart('\\')
                    Log-Message "PST: $relPath ($pstSize MB)"
                }
                Log-Message "NOTE: PST files in profile will be migrated."
                Log-Message "PST files on network drives or custom locations must be copied manually."
                Log-Message "========================================="
            }
        }
        catch {
            Log-Message "WARNING: PST file detection failed: $_"
        }

        # === GENERATE MANIFEST AND WINGET FIRST (before compression) ===
        $global:StatusText.Text = "Generating manifest..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Get SID first for profile mounted check
        $sid = (Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' |
            Where-Object { (Get-ItemProperty $_.PSPath -Name ProfileImagePath -EA SilentlyContinue).ProfileImagePath -like "*\$shortName" }).PSChildName

        # PRE-FLIGHT CHECK: Detect if profile is currently mounted/logged in
        if ($sid -and (Test-ProfileMounted $sid)) {
            Log-Message "WARNING: User profile is currently mounted (user may be logged in via HKU)"
            $res = Show-ModernDialog -Message "User profile '$shortName' is currently mounted in registry (HKU\$sid).`n`nThis indicates the user is likely logged in or services are using the profile.`r`nExporting while logged in may result in locked files being skipped.`r`n`r`nContinue anyway?" -Title "Profile Mounted" -Type Warning -Buttons YesNo
            if ($res -ne "Yes") {
                throw "Cancelled - profile is mounted (user logged in)"
            }
        }
        
        # Derive domain vs local from SID -> NTAccount

        # Derive domain vs local from SID -> NTAccount
        $derivedDomain = $null
        $derivedUsername = $Username
        $isAzureAD = $false
        try {
            if ($sid) {
                # Check if this is an AzureAD/Entra ID account
                $isAzureAD = Test-IsAzureADSID -SID $sid
                
                $sidObj = [System.Security.Principal.SecurityIdentifier]::new($sid)
                $nt = $sidObj.Translate([System.Security.Principal.NTAccount])
                $parts = $nt.Value -split '\\', 2
                if ($parts.Count -ge 2) { 
                    $derivedDomain = $parts[0]
                    $derivedUsername = $parts[1]
                    
                    # For AzureAD accounts, normalize domain to "AzureAD"
                    if ($isAzureAD) {
                        $derivedDomain = "AzureAD"
                        Log-Message "Detected AzureAD/Entra ID account: $derivedUsername"
                    }
                }
            }
        }
        catch {
            Log-Message "WARNING: Could not translate SID to NTAccount: $_"
        }
        
        
        # Profile timestamp for diagnostics
        $profileTimestamp = (Get-Item $userProfile.Path -Force -ErrorAction SilentlyContinue).LastWriteTimeUtc.ToString('o')
        
        # Create manifest in temp folder (will be included in compression)
        $manifest = [pscustomobject]@{
            ExportedAt       = (Get-Date).ToString('o')
            Username         = $derivedUsername
            ProfilePath      = $userProfile.Path
            SourceSID        = $sid
            IsAzureADUser    = $isAzureAD
            IsDomainUser     = if ($isAzureAD) { $false } elseif ($derivedDomain) { ($derivedDomain -ine $env:COMPUTERNAME) } else { $Username -match '\\' }
            Domain           = if ($isAzureAD) { "AzureAD" } elseif ($derivedDomain -and ($derivedDomain -ine $env:COMPUTERNAME)) { $derivedDomain } elseif ($Username -match '\\') { ($Username -split '\\')[0] } else { $null }
            ProfileTimestamp = $profileTimestamp
        }
        $manifest | ConvertTo-Json -Depth 5 | Out-File "$tmp\manifest.json" -Encoding UTF8
        Log-Message "Manifest created in temp folder"

        # Winget export in temp folder (will be included in compression)
        $wingetFile = "$tmp\Winget-Packages.json"
        try {
            Log-Message "Exporting Winget apps..."
            $wingetProc = Start-Process "winget.exe" -ArgumentList "export", "-o", "`"$wingetFile`"", "--accept-source-agreements" -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop

            if ($wingetProc.ExitCode -eq 0) {
                Log-Message "Winget export succeeded"
            }
            else {
                Log-Message "Winget completed with exit code $($wingetProc.ExitCode)"
                "{}" | Out-File $wingetFile -Encoding UTF8
            }
        }
        catch {
            Log-Message "Winget FAILED: $_ - creating empty placeholder"
            "{}" | Out-File $wingetFile -Encoding UTF8
        }

        # Compress profile folder contents + manifest + winget in ONE operation
        # Multi-threading: Use all CPU cores for maximum compression speed (2-3x faster on multi-core systems)
        $threadCount = $Config.SevenZipThreads
        Log-Message "7-Zip compression using $threadCount threads (detected $cpuCores CPU cores)"
        
        # Check if debug mode is enabled
        $debugMode = $false
        if ($global:DebugCheckBox -and $global:DebugCheckBox.Checked) {
            $debugMode = $true
            Log-Message "DEBUG MODE ENABLED: Verbose 7-Zip logging and artifact preservation active"
        }
        
        # If exclusion lists are not empty, write to temp files and use -x@ and -xr@
        $exclusionFile = $null
        $exclusionFileRec = $null
        $7zArgs = @('a', '-tzip', "`"$ZipPath`"", "`"$($userProfile.Path)\*`"", "`"$tmp\manifest.json`"", "`"$tmp\Winget-Packages.json`"", '-mx=5', "-mmt=$threadCount", '-bsp1')
        
        # Add verbose logging if debug mode is enabled
        if ($debugMode) {
            $7zArgs += '-bb3'  # Maximum verbosity - shows every file processed
            Log-Message "7-Zip verbose logging enabled (-bb3)"
        }
        
        if ($7zExclusions.Count -gt 0) {
            $exclusionFile = [System.IO.Path]::GetTempFileName()
            Set-Content -Path $exclusionFile -Value ($7zExclusions -join "`r`n") -Encoding UTF8
            $7zArgs += ("-x@" + $exclusionFile)
            Log-Message "Added $($7zExclusions.Count) non-recursive exclusions via listfile"
            
            if ($debugMode) {
                $copyPath = Join-Path $PSScriptRoot "Logs\Exclusions-NonRecursive-$($shortName).txt"
                if (-not (Test-Path "$PSScriptRoot\Logs")) { New-Item -ItemType Directory -Path "$PSScriptRoot\Logs" -Force | Out-Null }
                Copy-Item -Path $exclusionFile -Destination $copyPath -Force
                Log-Message "Exclusion list (non-recursive) copied to $copyPath" -Level DEBUG
            }
        }
        if ($7zExclusionsRec.Count -gt 0) {
            $exclusionFileRec = [System.IO.Path]::GetTempFileName()
            Set-Content -Path $exclusionFileRec -Value ($7zExclusionsRec -join "`r`n") -Encoding UTF8
            $7zArgs += ("-xr@" + $exclusionFileRec)
            Log-Message "Added $($7zExclusionsRec.Count) recursive exclusions via listfile"
            
            if ($debugMode) {
                $copyPath = Join-Path $PSScriptRoot "Logs\Exclusions-Recursive-$($shortName).txt"
                if (-not (Test-Path "$PSScriptRoot\Logs")) { New-Item -ItemType Directory -Path "$PSScriptRoot\Logs" -Force | Out-Null }
                Copy-Item -Path $exclusionFileRec -Destination $copyPath -Force
                Log-Message "Exclusion list (recursive) copied to $copyPath" -Level DEBUG
            }
        }
        
        # Setup 7-Zip logging implementation
        # Setup 7-Zip log file
        $7zLogFile = $null
        $isTemporaryLog = $false
        
        if ($debugMode) {
            $logsDir = Join-Path $PSScriptRoot "Logs"
            if (-not (Test-Path $logsDir)) {
                New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
            }
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $7zLogFile = Join-Path $logsDir "Export-7Zip-Log-$timestamp.txt"
            Log-Message "7-Zip detailed log will be saved to: $7zLogFile"
        }
        else {
            # Normal mode: use a temporary file for progress tracking to avoid stream deadlocks
            $7zLogFile = [System.IO.Path]::GetTempFileName()
            $isTemporaryLog = $true
        }
        
        $stderrFile = "$7zLogFile.err"
        
        # Start the process with output redirection
        Log-Message "Starting 7-Zip compression..."
        try {
            $zip = Start-Process -FilePath $global:SevenZipPath -ArgumentList $7zArgs -NoNewWindow -PassThru -Wait:$false -RedirectStandardOutput $7zLogFile -RedirectStandardError $stderrFile
        }
        catch {
            throw "Failed to start 7-Zip: $($_.Exception.Message)"
        }
        
        # Unified wait loop for both modes
        $lastUpdate = [DateTime]::Now
        while (-not $zip.HasExited) {
            if ($global:CancelRequested) {
                try { $zip.Kill() } catch { }
                throw "Export cancelled by user"
            }
            
            # Update progress from log file
            try {
                if (Test-Path $7zLogFile) {
                    $tailContent = Get-Content $7zLogFile -Tail 5 -ErrorAction SilentlyContinue
                    if ($tailContent -match '(\d+)%') {
                        $pct = 0
                        $progressMatches = [regex]::Matches($tailContent, '(\d+)%')
                        foreach ($m in $progressMatches) {
                            $val = [int]$m.Groups[1].Value
                            if ($val -gt $pct) { $pct = $val }
                        }
                        
                        if ($pct -gt $maxPct) { $maxPct = $pct }
                        $global:ProgressBar.Value = [Math]::Min(90, 10 + [int]($maxPct * 0.8))
                        $global:StatusText.Text = "Compressing - $maxPct%"
                    }
                }
            }
            catch { }
            
            # Periodically update generic status
            if (([DateTime]::Now - $lastUpdate).TotalSeconds -ge 2) {
                $lastUpdate = [DateTime]::Now
                if ($debugMode) {
                    $global:StatusText.Text = "Compressing (debug mode - check log for details)..."
                }
                elseif ($maxPct -eq 0) {
                    $global:StatusText.Text = "Compressing..."
                }
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 500
        }
        
        # End of wait loop logic - proceeding to cleanup and finalization
        if ($maxPct -gt 0) { $global:StatusText.Text = "Compressing - 100%" }
        $global:ProgressBar.Value = 90
        
        # Capture exit code immediately and ensure it's valid
        $exitCode = 0
        try {
            $exitCode = $zip.ExitCode
        }
        catch { 
            Log-Message "WARNING: Could not capture 7-Zip exit code via direct access: $_"
            # Fallback check if ZIP actually exists and has size
            if (Test-Path $ZipPath) { $exitCode = 0 } else { $exitCode = 2 }
        }
        if ($null -eq $exitCode) { $exitCode = 0 } # Assume success if no code but log says Ok
        
        # Merge stderr processing
        Start-Sleep -Milliseconds 500
        if (Test-Path $stderrFile) {
            $stderrContent = Get-Content $stderrFile -Raw -ErrorAction SilentlyContinue
            if ($stderrContent) {
                if ($debugMode) {
                    Add-Content -Path $7zLogFile -Value "`r`n=== STDERR OUTPUT ===`r`n$stderrContent"
                }
                else {
                    Log-Message "7-Zip encountered issues (non-fatal):`r`n$stderrContent"
                }
            }
            Remove-Item $stderrFile -Force -ErrorAction SilentlyContinue
        }
        
        if ($debugMode) {
            Add-Content -Path $7zLogFile -Value "`r`n=== EXIT CODE: $exitCode ==="
            Log-Message "7-Zip log file saved: $7zLogFile"
        }
        elseif ($isTemporaryLog) {
            Remove-Item $7zLogFile -Force -ErrorAction SilentlyContinue
        }

        # Check for incomplete/failed export
        if ($exitCode -notin 0, 1) {
            if ($exclusionFile -and (Test-Path $exclusionFile)) { Remove-Item $exclusionFile -Force -ErrorAction SilentlyContinue }
            
            # Build detailed error message - different for debug vs normal mode
            if ($debugMode) {
                # Debug mode: no progress tracking, focus on log file
                $errorMsg = "7-Zip compression failed with exit code $exitCode.`r`n`r`n"
                $errorMsg += "DEBUG MODE ACTIVE - Troubleshooting information:`r`n`r`n"
                $errorMsg += "Detailed 7-Zip log saved to:`r`n$7zLogFile`r`n`r`n"
                $errorMsg += "Temporary staging folder preserved at:`r`n$tmp`r`n`r`n"
                
                # Check if partial ZIP was created
                if (Test-Path $ZipPath) {
                    $partialSize = [math]::Round((Get-Item $ZipPath).Length / 1MB, 1)
                    $errorMsg += "Partial ZIP file preserved at:`r`n$ZipPath ($partialSize MB)`r`n`r`n"
                    Log-Message "DEBUG MODE: Partial ZIP file preserved for analysis: $ZipPath ($partialSize MB)"
                }
                
                $errorMsg += "Check the log file for details about which file caused the failure."
                Log-Message "DEBUG MODE: Preserving temp folder for troubleshooting: $tmp"
            }
            else {
                # Normal mode: include progress information
                $errorMsg = "7-Zip compression failed with exit code $exitCode"
                if ($maxPct -lt 80) {
                    $errorMsg += " after only reaching $maxPct% progress.`r`n`r`nThe export was incomplete."
                }
                else {
                    $errorMsg += " after reaching $maxPct% progress.`r`n`r`nThe export may be incomplete."
                }
            }
            
            # Check if USB or network drive
            $usbOrNetwork = $false
            try {
                $drive = ([System.IO.FileInfo]$ZipPath).Directory.Root.FullName
                $driveType = (Get-PSDrive | Where-Object { $_.Root -eq $drive }).Provider.Name
                $isRemovable = $false
                $isNetwork = $false
                $driveLetter = $drive.Substring(0, 1)
                $wmi = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='${driveLetter}:'"
                if ($wmi.DriveType -eq 2) { $isRemovable = $true }
                if ($wmi.DriveType -eq 4) { $isNetwork = $true }
                if ($isRemovable -or $isNetwork) { $usbOrNetwork = $true }
            }
            catch {}
            
            if ($usbOrNetwork) {
                $errorMsg += "`n`nExporting to a removable or network drive can fail if the drive is slow, disconnected, or full. Try exporting to a local disk (like C:) if problems persist."
            }
            
            if ($stderrOutput) {
                $errorMsg += "`n`n7-Zip Error Output:`n$stderrOutput"
            }
            
            throw $errorMsg
        }

        # Cleanup and finalize
        $zip.Dispose()
        if ($exclusionFile -and (Test-Path $exclusionFile)) { Remove-Item $exclusionFile -Force -ErrorAction SilentlyContinue }
        if ($exclusionFileRec -and (Test-Path $exclusionFileRec)) { Remove-Item $exclusionFileRec -Force -ErrorAction SilentlyContinue }

        Log-Message "Profile compressed successfully"
        $global:ProgressBar.Value = 90
        
        if ($global:CancelRequested) { throw "Operation cancelled by user" }
        
        # === HASH VERIFICATION (Export) ===
        if ($Config.HashVerificationEnabled) {
            $global:StatusText.Text = "Verifying ZIP integrity..."
            [System.Windows.Forms.Application]::DoEvents()
            Log-Message "Calculating ZIP hash for verification..."
            try {
                $zipHashResult = Get-FileHash -Path $ZipPath -Algorithm SHA256 -ErrorAction Stop
                if ($zipHashResult -and $zipHashResult.Hash) {
                    Log-Message "ZIP Hash (SHA256): $($zipHashResult.Hash)"
                    # Store hash in a sidecar file for later verification
                    "$($zipHashResult.Hash)  $(Split-Path $ZipPath -Leaf)" | Out-File "$ZipPath.sha256" -Encoding ASCII
                    Log-Message "Hash saved to: $ZipPath.sha256"
                }
                else {
                    Log-Message "WARNING: Hash calculation returned null"
                }
            }
            catch {
                Log-Message "WARNING: Hash verification failed: $_"
            }
        }

        # === COLLECT DATA FOR REPORT (BEFORE CLEANUP) ===
        # Parse installed programs from Winget export for report
        $installedPrograms = @()
        $wingetExportPath = "$tmp\Winget-Packages.json"
        if (Test-Path $wingetExportPath) {
            try {
                Log-Message "Parsing Winget export for installed programs list..."
                $wingetData = Get-Content $wingetExportPath -Raw | ConvertFrom-Json
                if ($wingetData.Sources) {
                    foreach ($wingetSource in $wingetData.Sources) {
                        if ($wingetSource.Packages) {
                            foreach ($pkg in $wingetSource.Packages) {
                                $installedPrograms += @{
                                    Name    = $pkg.PackageIdentifier
                                    Version = if ($pkg.Version) { $pkg.Version } else { "Latest" }
                                    Source  = if ($wingetSource.SourceDetails.Name) { $wingetSource.SourceDetails.Name } else { "Unknown" }
                                }
                            }
                        }
                    }
                }
                Log-Message "Collected $($installedPrograms.Count) installed programs for report"
            }
            catch {
                Log-Message "Could not parse Winget export for report: $_"
            }
        }
        else {
            Log-Message "WARNING: Winget-Packages.json not found at $wingetExportPath"
        }

        # Cleanup temp folder (skip if debug mode to preserve artifacts)
        if ($debugMode) {
            Log-Message "DEBUG MODE: Preserving temp folder for analysis: $tmp"
            $global:StatusText.Text = "Debug mode: Temp files preserved"
        }
        else {
            $global:StatusText.Text = "Cleaning up temp files"
            Log-Message "Cleaning up temp files"
            Remove-FolderRobust $tmp
        }

        $sizeMB = [math]::Round((Get-Item $ZipPath).Length / 1MB, 1)
        $global:ProgressBar.Value = 100
        $global:StatusText.Text = "[OK] Export complete!"
        $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)

        # Log elapsed time for diagnostics
        if ($global:ExportStartTime) {
            $elapsed = [DateTime]::Now - $global:ExportStartTime
            Log-Message "Operation elapsed time: $($elapsed.TotalMinutes.ToString('F2')) minutes ($($elapsed.TotalSeconds.ToString('F0')) seconds)"
        }

        Log-Message "EXPORT SUCCESS: $ZipPath ($sizeMB MB)"
        
        # Generate migration report
        try {
            # Collect ZIP file statistics for report
            $zipFileCount = 'N/A'
            $zipFolderCount = 'N/A'
            $uncompressedSizeMB = 'N/A'
            $compressionRatio = 'N/A'
            
            try {
                if (Test-Path $ZipPath) {
                    Log-Message "Collecting ZIP archive statistics for report..."
                    
                    # Use 7-Zip with -slt (technical listing) for reliable parsing
                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName = $global:SevenZipPath
                    $psi.Arguments = "l -slt `"$ZipPath`""
                    $psi.CreateNoWindow = $true
                    $psi.UseShellExecute = $false
                    $psi.RedirectStandardOutput = $true
                    
                    $listProc = [System.Diagnostics.Process]::Start($psi)
                    $output = $listProc.StandardOutput.ReadToEnd()
                    $listProc.WaitForExit()
                    
                    if ($listProc.ExitCode -eq 0) {
                        Log-Message "7-Zip technical listing received, parsing..."
                        
                        # Count files and folders, sum sizes from technical output
                        $files = 0
                        $folders = 0
                        $totalSize = 0
                        
                        $lines = $output -split "`n"
                        $isFolder = $false
                        $currentSize = 0
                        
                        foreach ($line in $lines) {
                            $line = $line.Trim()
                            
                            # Check if this entry is a folder
                            if ($line -match '^Attributes = .+D') {
                                $isFolder = $true
                            }
                            
                            # Get the size
                            if ($line -match '^Size = (\d+)$') {
                                $currentSize = [int64]$matches[1]
                            }
                            
                            # When we hit a Path entry, we've finished reading an item
                            if ($line -match '^Path = (.+)$' -and $line -notmatch '^Path\s*=\s*$') {
                                if ($isFolder) {
                                    $folders++
                                }
                                else {
                                    $files++
                                    $totalSize += $currentSize
                                }
                                # Reset for next item
                                $isFolder = $false
                                $currentSize = 0
                            }
                        }
                        
                        if ($files -gt 0) {
                            $zipFileCount = $files
                            $zipFolderCount = $folders
                            $uncompressedSizeMB = [math]::Round($totalSize / 1MB, 2)
                            
                            # Calculate compression ratio
                            if ($totalSize -gt 0) {
                                $zipFileObj = Get-Item $ZipPath
                                $compressionRatio = "$([math]::Round(($zipFileObj.Length / $totalSize) * 100, 1))%"
                            }
                            
                            Log-Message "ZIP statistics: $zipFileCount files, $zipFolderCount folders, $uncompressedSizeMB MB uncompressed, compression: $compressionRatio"
                        }
                        else {
                            Log-Message "WARNING: No files found in 7-Zip output"
                        }
                    }
                    else {
                        Log-Message "7-Zip list command failed with exit code: $($listProc.ExitCode)"
                    }
                }
            }
            catch {
                Log-Message "Could not collect ZIP statistics: $_"
            }
            
            $reportData = @{
                Username           = $importTargetUser
                SourceSID          = $sid
                SourcePath         = $source
                ZipPath            = $ZipPath
                ZipSizeMB          = $sizeMB
                ElapsedMinutes     = if ($global:ExportStartTime) { ([DateTime]::Now - $global:ExportStartTime).TotalMinutes.ToString('F2') } else { 'N/A' }
                ElapsedSeconds     = if ($global:ExportStartTime) { ([DateTime]::Now - $global:ExportStartTime).TotalSeconds.ToString('F0') } else { 'N/A' }
                Timestamp          = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                FileCount          = $zipFileCount
                FolderCount        = $zipFolderCount
                UncompressedSizeMB = $uncompressedSizeMB
                CompressionRatio   = $compressionRatio
                Exclusions         = @('Temp folders', 'Cache files', 'Log files', 'OST files')
                HashEnabled        = $Config.HashVerificationEnabled
                InstalledPrograms  = $installedPrograms
                CleanupCategories  = $cleanupCategories
                CleanupSavingsMB   = $cleanupSavingsMB
                CategoryExclusions = $categoryExclusions
                CloudExclusions    = $topLevelCloudFolders
            }
            
            $reportPath = Generate-MigrationReport -OperationType 'Export' -ReportData $reportData
            if ($reportPath) {
                Log-Info "Migration report available: $reportPath"
            }
        }
        catch {
            Log-Warning "Could not generate migration report: $_"
        }
        
        Show-ModernDialog -Message "Export completed!`r`n`r`n$ZipPath`r`nSize: $sizeMB MB" -Title "Success" -Type Success -Buttons OK

    }
    catch {
        # Log elapsed time for diagnostics
        if ($global:ExportStartTime) {
            $elapsed = [DateTime]::Now - $global:ExportStartTime
            Log-Message "Operation elapsed time: $($elapsed.TotalMinutes.ToString('F2')) minutes ($($elapsed.TotalSeconds.ToString('F0')) seconds)"
        }
        
        Log-Message "EXPORT FAILED: $_"
        
        # Cleanup partial ZIP if cancelled
        if ($global:CancelRequested -or $_.Exception.Message -match "cancelled") {
            Log-Message "Export cancelled - cleaning up partial files..."
            if (Test-Path $ZipPath) {
                Remove-Item $ZipPath -Force -ErrorAction SilentlyContinue
                Log-Message "Deleted partial ZIP: $ZipPath"
            }
        }
        
        # Generate failure report
        $errorMessage = $_.Exception.Message
        
        try {
            $elapsedMins = if ($global:ExportStartTime) { ([DateTime]::Now - $global:ExportStartTime).TotalMinutes.ToString('F2') } else { 'N/A' }
            $elapsedSecs = if ($global:ExportStartTime) { ([DateTime]::Now - $global:ExportStartTime).TotalSeconds.ToString('F0') } else { 'N/A' }
            
            $reportData = @{
                Username       = if ($importTargetUser) { $importTargetUser } else { 'N/A' }
                SourceSID      = if ($sid) { $sid } else { 'N/A' }
                SourcePath     = if ($source) { $source } else { 'N/A' }
                ZipPath        = $ZipPath
                ZipSizeMB      = 'N/A'
                ElapsedMinutes = $elapsedMins
                ElapsedSeconds = $elapsedSecs
                Timestamp      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
            
            $reportPath = Generate-MigrationReport -OperationType 'Export' -ReportData $reportData -Success $false -ErrorMessage $errorMessage
            if ($reportPath) {
                Log-Message "Failure report generated: $reportPath"
                # Report auto-opens via AutoOpenReports config in Generate-MigrationReport
            }
        }
        catch {
            Log-Message "Could not generate failure report: $_"
        }
        
        Show-ModernDialog -Message "Export failed:`r`n$_" -Title "Error" -Type Error -Buttons OK
    }
    finally {
        Stop-OperationLog
        $global:StatusText.Text = "Idle"
        if ($global:ProgressBar.Value -ne 100) { $global:ProgressBar.Value = 0 }
    }
}

# =============================================================================
# IMPORT-USERPROFILE FUNCTION
# =============================================================================
function Import-UserProfile {
    param([string]$ZipPath, [string]$Username)
    $global:ProgressBar.Value = 0
    $global:StatusText.Text = "Starting import..."
    if ($global:CancelButton) { $global:CancelButton.Enabled = $true }
    [System.Windows.Forms.Application]::DoEvents()
    
    # Start persistent logging for import operation
    Start-OperationLog -Operation "Import" -Username $Username
    
    try {
        $global:CancelRequested = $false
        $global:CurrentOperation = "Import"
        
        Log-Message "=== IMPORT to $Username ==="
        if (-not (Test-Path $ZipPath)) { throw "ZIP not found: $ZipPath" }
        Log-Message "ZIP file verified: $ZipPath"
        
        # === HASH VERIFICATION (Import) ===
        # Track hash verification status for reporting
        $hashVerificationStatus = "Not Performed"
        $hashVerificationDetails = "Hash verification disabled in configuration"
        
        if ($Config.HashVerificationEnabled) {
            $hashFile = "$ZipPath.sha256"
            if (Test-Path $hashFile) {
                $global:StatusText.Text = "Verifying ZIP integrity..."
                [System.Windows.Forms.Application]::DoEvents()
                Log-Message "Verifying ZIP hash..."
                try {
                    $expectedHash = (Get-Content $hashFile -First 1).Split()[0]
                    $actualHashResult = Get-FileHash -Path $ZipPath -Algorithm SHA256 -ErrorAction Stop
                    if ($actualHashResult -and $actualHashResult.Hash) {
                        if ($actualHashResult.Hash -eq $expectedHash) {
                            Log-Message "ZIP hash verification: PASSED"
                            $hashVerificationStatus = "Passed"
                            $hashVerificationDetails = "SHA256 hash matched expected value"
                        }
                        else {
                            Log-Message "WARNING: ZIP hash mismatch!"
                            Log-Message "Expected: $expectedHash"
                            Log-Message "Actual:   $($actualHashResult.Hash)"
                            $hashResult = Show-ModernDialog -Message "ZIP file hash verification FAILED!`r`n`r`nThe file may be corrupted or tampered with.`r`n`r`nContinue anyway?" -Title "Hash Verification Failed" -Type Warning -Buttons YesNo
                            if ($hashResult -ne "Yes") {
                                throw "Import cancelled - hash verification failed"
                            }
                            $hashVerificationStatus = "Failed (User Bypassed)"
                            $hashVerificationDetails = "Hash mismatch detected - user chose to continue"
                        }
                    }
                }
                catch {
                    # If it was our cancellation throw, re-throw it
                    if ($_.Exception.Message -like "*Import cancelled*") {
                        throw $_
                    }
                    Log-Message "WARNING: Hash verification error: $_"
                    $hashVerificationStatus = "Error"
                    $hashVerificationDetails = "Hash verification error: $_"
                }
            }
            else {
                Log-Message "No hash file found - skipping verification"
                # Warn user that hash verification cannot be performed
                $noHashResult = Show-ModernDialog -Message "SHA256 hash file not found!`r`n`r`nThe ZIP file cannot be verified for integrity or tampering.`r`n`r`nContinue import without verification?" -Title "Hash File Missing" -Type Warning -Buttons YesNo
                if ($noHashResult -ne "Yes") {
                    throw "Import cancelled - hash file missing"
                }
                Log-Message "User chose to continue without hash verification"
                $hashVerificationStatus = "Skipped (No Hash File)"
                $hashVerificationDetails = "Hash file not found - user chose to continue"
            }
        }
        
        $global:StatusText.Text = "Validating ZIP file..."
        [System.Windows.Forms.Application]::DoEvents()
        
        if ($global:CancelRequested) { throw "Operation cancelled by user" }
        
        # Use account type and username from Set Target User if available
        # Always use the exact Set Target User value
        $importTargetUser = $global:SelectedTargetUser
        if (-not $importTargetUser) { $importTargetUser = $Username }
        $domain = $null
        $shortName = $importTargetUser
        $isDomain = $false
        $isAzureAD = $false
        # Only parse account type for AzureAD check
        if ($importTargetUser -match '^AzureAD\\(.+)$') {
            $isAzureAD = $true
            $shortName = $matches[1].Trim()
        }
        elseif ($importTargetUser -match '^(?<domain>[^\\]+)\\(?<user>.+)$') {
            $parsedDomain = $matches['domain']
            $shortName = $matches['user']
            if ($parsedDomain -ieq $env:COMPUTERNAME) {
                $isDomain = $false
            }
            else {
                $isDomain = $true
                $domain = $parsedDomain  # Use exactly as entered
            }
        }
        else {
            $isDomain = $false
        }
        $isAzureAD = $false
        # CRITICAL: Check for AzureAD format FIRST before general domain parsing
        if ($Username -match '^AzureAD\\(.+)$') {
            $shortName = $matches[1].Trim()
            $isAzureAD = $true
            $isDomain = $false
            Log-Message "Parsed as AZUREAD user (explicit format): User='$shortName'"
        }
        # Check for UPN format (user@domain.com)
        elseif ($Username -match '@') {
            # Extract username part (before @) for shortName
            $shortName = ($Username -split '@')[0].Trim()
            $isAzureAD = $true
            $isDomain = $false
            Log-Message "Parsed as AZUREAD user (UPN format): User='$shortName', Full UPN='$Username'"
        }
        elseif ($Username -match '\\') {
            $parts = $Username -split '\\', 2
            $parsedDomain = $parts[0].Trim().ToUpper()
            $shortName = $parts[1].Trim()
            
            # Check if domain is actually the computer name (local user)
            if ($parsedDomain -ieq $env:COMPUTERNAME) {
                $isDomain = $false
                Log-Message "Parsed as LOCAL user (computer name match): User='$shortName'"
            }
            else {
                # Before treating as domain user, check if this is actually an AzureAD user
                # AzureAD users show up as TENANTNAME\username in the profile list
                # We can detect them by checking their SID pattern (S-1-12-1-...)
                try {
                    $testProfile = Get-WmiObject Win32_UserProfile | Where-Object {
                        $_.LocalPath -like "*\$shortName"
                    } | Select-Object -First 1
                    
                    if ($testProfile -and (Test-IsAzureADSID $testProfile.SID)) {
                        # This is an AzureAD user, not a domain user
                        $isAzureAD = $true
                        $isDomain = $false
                        Log-Message "Detected as AZUREAD user (by SID pattern): User='$shortName', TenantName='$parsedDomain'"
                        Log-Message "Note: AzureAD users appear as TENANTNAME\username in Windows"
                    }
                    else {
                        # This is a real domain user
                        $isDomain = $true
                        $domain = Get-DomainFQDN -NetBIOSName $parsedDomain
                        Log-Message "Parsed as DOMAIN user: Domain='$domain' (FQDN), User='$shortName'"
                    }
                }
                catch {
                    # If we can't check the profile, assume it's a domain user
                    $isDomain = $true
                    $domain = Get-DomainFQDN -NetBIOSName $parsedDomain
                    Log-Message "Parsed as DOMAIN user (profile check failed): Domain='$domain' (FQDN), User='$shortName'"
                }
            }
        }
        else {
            $isDomain = $false
            Log-Message "Parsed as LOCAL user (no domain): User='$shortName'"
        }
        
        # === AZUREAD/ENTRA ID VALIDATION (if TARGET user is AzureAD) ===
        # Check for both AzureAD\username and UPN (user@domain.com) formats
        $targetIsAzureAD = ($importTargetUser -match '^AzureAD\\(.+)$') -or ($importTargetUser -match '@')
        if ($targetIsAzureAD) {
            # --- HYBRID JOIN DETECTION ---
            # If the device is Hybrid Joined (Domain + AzureAD), we MUST use the AD SID (S-1-5-...)
            # The Cloud SID (S-1-12-...) is not used for profile mapping in Hybrid scenarios.
            $isHybrid = $false
            try {
                $compSys = Get-CimInstance Win32_ComputerSystem
                if ($compSys.PartOfDomain -and (Test-IsAzureADJoined)) {
                    $isHybrid = $true
                    Log-Message "Hybrid Join detected (Domain: $($compSys.Domain) + AzureAD)."
                    Log-Message "Using Active Directory SID logic (S-1-5-...) instead of Cloud SID logic (S-1-12-...)."
                    
                    # Resolve using Get-LocalUserSID with the shortname (will use NTAccount translation)
                    $targetSID = Get-LocalUserSID -UserName $shortName
                    Log-Message "Resolved Hybrid SID: $targetSID"
                }
            }
            catch {
                Log-Warning "Hybrid detection failed: $_"
            }

            if (-not $isHybrid) {
                Log-Message "Target user is AzureAD (Cloud Only) - validating device join status..."
                # Check if system is AzureAD joined
                if (-not (Test-IsAzureADJoined)) {
                    Log-Message "ERROR: System is not AzureAD joined but trying to import to AzureAD profile"
                    # Show guidance dialog
                    $dlgResult = Show-AzureADJoinDialog -Username $shortName
                    if ($dlgResult -eq "Cancel") {
                        throw "Import cancelled - AzureAD join required for this profile"
                    }
                    # User clicked Continue Import - re-check join status
                    if (-not (Test-IsAzureADJoined)) {
                        throw "System is still not AzureAD joined. Please complete the join process in Settings > Accounts > Access work or school and try again."
                    }
                    Log-Message "AzureAD join status verified after user completed join process"
                }
                Log-Message "AzureAD join status: VERIFIED"
            
                # Verify/resolve AzureAD user SID using Microsoft Graph
                # No longer requires user to have signed in first!
                try {
                    Log-Message "Resolving AzureAD user SID for '$shortName'..."
                
                    # Check if user has already signed in (profile exists)
                    $azureProfile = Get-WmiObject Win32_UserProfile | Where-Object {
                        $_.LocalPath -like "*\$shortName" -and (Test-IsAzureADSID $_.SID)
                    }
                
                    if ($azureProfile) {
                        Log-Message "AzureAD user '$shortName' has already signed in - profile exists"
                        $targetSID = $azureProfile.SID
                        Log-Message "Using existing SID: $targetSID"
                    }
                    else {
                        Log-Message "AzureAD user '$shortName' has not signed in yet - using Microsoft Graph to resolve SID"
                    
                        # Determine UPN format
                        $upn = if ($shortName -match '@') {
                            # Username is already in UPN format
                            Log-Message "Username is already in UPN format: $shortName"
                            $shortName
                        }
                        elseif ($global:TargetUserUPN) {
                            # UPN was stored from Set Target User
                            Log-Message "Using UPN stored from Set Target User: $global:TargetUserUPN"
                            $global:TargetUserUPN
                        }
                        else {
                            # Need to prompt for UPN
                            Log-Message "Prompting user for UPN..."
                            $inputResult = Show-InputDialog -Title "AzureAD User Principal Name Required" `
                                -Message "Enter the full email address (UPN) for AzureAD user '$shortName':" `
                                -DefaultValue "$shortName@" `
                                -ExampleText "Example: $shortName@company.com"
                        
                            # Check if user cancelled
                            if ($inputResult.Result -ne [System.Windows.Forms.DialogResult]::OK) {
                                throw "User cancelled UPN input"
                            }
                        
                            # Return the value (don't assign to $upn here - it's already being assigned by the outer if)
                            $inputResult.Value
                        }
                    
                        # Validate the UPN
                        if ([string]::IsNullOrWhiteSpace($upn)) {
                            throw "UPN cannot be empty"
                        }
                    
                        Log-Message "Using UPN: $upn"
                    
                        # Use Microsoft Graph to get SID
                        $targetSID = Get-AzureADUserSID -UserPrincipalName $upn
                    
                        if (-not $targetSID) {
                            throw "Failed to resolve AzureAD SID for user '$upn' via Microsoft Graph"
                        }
                    
                        Log-Message "Successfully resolved AzureAD SID via Microsoft Graph: $targetSID"
                    }
                }
                catch {
                    Log-Message "ERROR: AzureAD user validation failed: $_"
                    throw $_
                }
            }
        }
        else {
            # If manifest indicates AzureAD but target is not AzureAD, just warn
            if ($manifest.IsAzureADUser -and -not $isAzureAD) {
                Log-Message "WARNING: Manifest indicates AzureAD profile but username was not detected as AzureAD user"
                Log-Message "Current username format: $importTargetUser (parsed as $([string]::Join('/',@('DOMAIN') * $isDomain + @('LOCAL') * (-not $isDomain) )))"
                # Optionally, show a warning dialog (not blocking)
                Show-ModernDialog -Message "You are importing an AzureAD/Entra ID profile into a non-AzureAD account. This is supported, but some settings may not transfer perfectly. Continue?" -Title "AzureAD Profile Mismatch" -Type Warning -Buttons OK
            }
        }
        
        $global:StatusText.Text = "Checking if user is logged on..."
        [System.Windows.Forms.Application]::DoEvents()
        $loggedOn = qwinsta | Select-String "\b$shortName\b" -Quiet
        if ($loggedOn) {
            Log-Message "WARNING: User '$shortName' appears to be logged on!"
            $res = Show-ModernDialog -Message "User '$shortName' appears to be logged on.`r`n`r`nRisk: Corrupted profile! Continue anyway?" -Title "User Logged On" -Type Warning -Buttons YesNo
            if ($res -ne "Yes") { throw "Cancelled by user." }
        }
        else {
            Log-Message "User is not logged on - safe to proceed"
        }
        
        # Create local user if needed (skip for domain and AzureAD users)
        if (-not $isDomain -and -not $isAzureAD -and -not (Get-LocalUser -Name $shortName -ErrorAction SilentlyContinue)) {
            Log-Message "Creating local user '$shortName'..."
            $global:StatusText.Text = "Creating local user..."
            [System.Windows.Forms.Application]::DoEvents()
            # Apply theme
            $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
            $theme = $Themes[$currentTheme]

            $passForm = New-Object System.Windows.Forms.Form
            $passForm.Text = "Create Local User"
            $passForm.Size = New-Object System.Drawing.Size(480, 420)
            $passForm.StartPosition = "CenterScreen"
            $passForm.FormBorderStyle = "FixedDialog"
            $passForm.MaximizeBox = $false
            $passForm.MinimizeBox = $false
            $passForm.TopMost = $true
            $passForm.BackColor = $theme.FormBackColor
            $passForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

            # Header panel
            $headerPanel = New-Object System.Windows.Forms.Panel
            $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
            $headerPanel.Size = New-Object System.Drawing.Size(480, 70)
            $headerPanel.BackColor = $theme.HeaderBackColor
            $passForm.Controls.Add($headerPanel)

            $lblTitle = New-Object System.Windows.Forms.Label
            $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
            $lblTitle.Size = New-Object System.Drawing.Size(440, 25)
            $lblTitle.Text = "Create User: $shortName"
            $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
            $lblTitle.ForeColor = $theme.HeaderTextColor
            $headerPanel.Controls.Add($lblTitle)

            $lblInfo = New-Object System.Windows.Forms.Label
            $lblInfo.Location = New-Object System.Drawing.Point(22, 43)
            $lblInfo.Size = New-Object System.Drawing.Size(440, 20)
            $lblInfo.Text = "User does not exist - Set a password to create"
            $lblInfo.ForeColor = $theme.SubHeaderTextColor
            $headerPanel.Controls.Add($lblInfo)

            # Main content card
            $contentCard = New-Object System.Windows.Forms.Panel
            $contentCard.Location = New-Object System.Drawing.Point(15, 85)
            $contentCard.Size = New-Object System.Drawing.Size(440, 240)
            $contentCard.BackColor = $theme.PanelBackColor
            $passForm.Controls.Add($contentCard)
            
            $lblPass1 = New-Object System.Windows.Forms.Label
            $lblPass1.Location = New-Object System.Drawing.Point(20, 20)
            $lblPass1.Size = New-Object System.Drawing.Size(100, 20)
            $lblPass1.Text = "Password:"
            $lblPass1.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $lblPass1.ForeColor = $theme.LabelTextColor
            $contentCard.Controls.Add($lblPass1)
            
            $txtPass1 = New-Object System.Windows.Forms.TextBox
            $txtPass1.Location = New-Object System.Drawing.Point(130, 18)
            $txtPass1.Size = New-Object System.Drawing.Size(220, 25)
            $txtPass1.UseSystemPasswordChar = $true
            $txtPass1.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $txtPass1.BackColor = $theme.TextBoxBackColor
            $txtPass1.ForeColor = $theme.TextBoxForeColor
            $contentCard.Controls.Add($txtPass1)
            
            $lblPass2 = New-Object System.Windows.Forms.Label
            $lblPass2.Location = New-Object System.Drawing.Point(20, 55)
            $lblPass2.Size = New-Object System.Drawing.Size(100, 20)
            $lblPass2.Text = "Confirm:"
            $lblPass2.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $lblPass2.ForeColor = $theme.LabelTextColor
            $contentCard.Controls.Add($lblPass2)
            
            $txtPass2 = New-Object System.Windows.Forms.TextBox
            $txtPass2.Location = New-Object System.Drawing.Point(130, 53)
            $txtPass2.Size = New-Object System.Drawing.Size(220, 25)
            $txtPass2.UseSystemPasswordChar = $true
            $txtPass2.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $txtPass2.BackColor = $theme.TextBoxBackColor
            $txtPass2.ForeColor = $theme.TextBoxForeColor
            $contentCard.Controls.Add($txtPass2)
            
            $chkShow1 = New-Object System.Windows.Forms.CheckBox
            $chkShow1.Location = New-Object System.Drawing.Point(360, 18)
            $chkShow1.Size = New-Object System.Drawing.Size(70, 23)
            $chkShow1.Text = "Show"
            $chkShow1.ForeColor = $theme.LabelTextColor
            $chkShow1.Add_CheckedChanged({ $txtPass1.UseSystemPasswordChar = -not $chkShow1.Checked; $txtPass2.UseSystemPasswordChar = -not $chkShow1.Checked })
            $contentCard.Controls.Add($chkShow1)

            $lblStrength = New-Object System.Windows.Forms.Label
            $lblStrength.Location = New-Object System.Drawing.Point(130, 85)
            $lblStrength.Size = New-Object System.Drawing.Size(290, 20)
            $lblStrength.Text = ""
            $lblStrength.Font = New-Object System.Drawing.Font("Segoe UI", 8)
            $contentCard.Controls.Add($lblStrength)

            $chkAdmin = New-Object System.Windows.Forms.CheckBox
            $chkAdmin.Location = New-Object System.Drawing.Point(20, 115)
            $chkAdmin.Size = New-Object System.Drawing.Size(400, 25)
            $chkAdmin.Text = "Add this user to Administrators group"
            $chkAdmin.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $chkAdmin.ForeColor = $theme.LabelTextColor
            $chkAdmin.Checked = $true
            $contentCard.Controls.Add($chkAdmin)

            $lblError = New-Object System.Windows.Forms.Label
            $lblError.Location = New-Object System.Drawing.Point(20, 150)
            $lblError.Size = New-Object System.Drawing.Size(400, 40)
            $lblError.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
            $lblError.Text = ""
            $lblError.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $contentCard.Controls.Add($lblError)

            $txtPass1.Add_TextChanged({
                    $len = $txtPass1.Text.Length
                    if ($len -eq 0) {
                        $lblStrength.Text = ""
                        $lblStrength.ForeColor = $theme.LabelTextColor
                    }
                    elseif ($len -lt 8) {
                        $lblStrength.Text = "Weak password (< 8 characters)"
                        $lblStrength.ForeColor = [System.Drawing.Color]::OrangeRed
                    }
                    elseif ($len -lt 12) {
                        $lblStrength.Text = "Moderate password"
                        $lblStrength.ForeColor = [System.Drawing.Color]::Orange
                    }
                    else {
                        $lblStrength.Text = "Strong password"
                        $lblStrength.ForeColor = [System.Drawing.Color]::Green
                    }
                })

            # Warning area
            $lblWarning = New-Object System.Windows.Forms.Label
            $lblWarning.Location = New-Object System.Drawing.Point(20, 195) # Adjusted Y position
            $lblWarning.Size = New-Object System.Drawing.Size(400, 40) # Adjusted height
            $lblWarning.Text = "Note: The user should change their password after first logon." # Simplified text
            $lblWarning.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
            $lblWarning.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
            $contentCard.Controls.Add($lblWarning)

            # Action buttons
            $btnOK = New-Object System.Windows.Forms.Button
            $btnOK.Location = New-Object System.Drawing.Point(240, 340)
            $btnOK.Size = New-Object System.Drawing.Size(100, 32)
            $btnOK.Text = "Create"
            $btnOK.BackColor = $theme.ButtonSuccessBackColor
            $btnOK.ForeColor = $theme.ButtonSuccessForeColor
            $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnOK.FlatAppearance.BorderSize = 0
            $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnOK.Add_MouseEnter({ $this.BackColor = $theme.ButtonSuccessHoverBackColor })
            $btnOK.Add_MouseLeave({ $this.BackColor = $theme.ButtonSuccessBackColor })
            $btnOK.Add_Click({
                    if ($txtPass1.Text -ne $txtPass2.Text) {
                        $lblError.Text = "Passwords do not match!"
                    }
                    elseif ($txtPass1.Text.Length -lt 1) {
                        $lblError.Text = "Password cannot be empty!"
                    }
                    else {
                        $passForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                        $passForm.Close()
                    }
                })
            $passForm.Controls.Add($btnOK)

            $btnCancel = New-Object System.Windows.Forms.Button
            $btnCancel.Location = New-Object System.Drawing.Point(350, 340)
            $btnCancel.Size = New-Object System.Drawing.Size(100, 32)
            $btnCancel.Text = "Cancel"
            $btnCancel.BackColor = $theme.ButtonSecondaryBackColor
            $btnCancel.ForeColor = $theme.ButtonSecondaryForeColor
            $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnCancel.FlatAppearance.BorderSize = 0
            $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnCancel.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
            $btnCancel.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
            $btnCancel.Add_Click({ $passForm.Close() })
            $passForm.Controls.Add($btnCancel)
            $passForm.AcceptButton = $btnOK
            $passForm.CancelButton = $btnCancel
            $txtPass1.Focus()
            
            $result = $passForm.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $passwordText = $txtPass1.Text
                $pass = ConvertTo-SecureString $passwordText -AsPlainText -Force
                try {
                    New-LocalUser -Name $shortName -Password $pass -FullName $shortName -Description "Profile Import" -ErrorAction Stop
                    Log-Message "Local user '$shortName' created successfully"
                    if ($chkAdmin.Checked) {
                        Add-LocalGroupMember -Group "Administrators" -Member $shortName -ErrorAction Stop
                        Log-Message "Added '$shortName' to Administrators group"
                    }
                }
                catch {
                    throw "Failed to create local user or add to Administrators: $_"
                }
            }
            else {
                throw "User creation cancelled - cannot proceed with import"
            }
            $passForm.Dispose()
        }
        # Target profile directory
        $target = "C:\Users\$shortName"
        
        # Check if user profile already exists and prompt for merge/replace option
        $targetExistsBeforeResolveSID = Test-Path $target
        $mergeMode = $false
        # Show merge/replace dialog for local and AzureAD users (not domain users, as they need domain join first)
        if ($targetExistsBeforeResolveSID -and -not $isDomain) {
            Log-Message "User profile already exists: $target"
            
            # Apply theme
            $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
            $theme = $Themes[$currentTheme]

            # Create custom dialog with Merge/Replace buttons
            $choiceForm = New-Object System.Windows.Forms.Form
            $choiceForm.Text = "User Profile Exists"
            $choiceForm.Size = New-Object System.Drawing.Size(550, 360)
            $choiceForm.StartPosition = "CenterScreen"
            $choiceForm.FormBorderStyle = "FixedDialog"
            $choiceForm.MaximizeBox = $false
            $choiceForm.MinimizeBox = $false
            $choiceForm.TopMost = $true
            $choiceForm.BackColor = $theme.FormBackColor
            $choiceForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

            # Header panel
            $headerPanel = New-Object System.Windows.Forms.Panel
            $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
            $headerPanel.Size = New-Object System.Drawing.Size(550, 60)
            $headerPanel.BackColor = $theme.HeaderBackColor
            $choiceForm.Controls.Add($headerPanel)

            $lblTitle = New-Object System.Windows.Forms.Label
            $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
            $lblTitle.Size = New-Object System.Drawing.Size(510, 30)
            $lblTitle.Text = "User '$shortName' already exists"
            $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
            $lblTitle.ForeColor = $theme.HeaderTextColor
            $headerPanel.Controls.Add($lblTitle)

            # Content panel
            $contentPanel = New-Object System.Windows.Forms.Panel
            $contentPanel.Location = New-Object System.Drawing.Point(15, 75)
            $contentPanel.Size = New-Object System.Drawing.Size(510, 180)
            $contentPanel.BackColor = $theme.PanelBackColor
            $choiceForm.Controls.Add($contentPanel)

            $lblInfo = New-Object System.Windows.Forms.Label
            $lblInfo.Location = New-Object System.Drawing.Point(15, 15)
            $lblInfo.Size = New-Object System.Drawing.Size(480, 150)
            $lblInfo.Text = "Choose how to handle the existing profile:`n`nMERGE: Keep existing profile and merge imported files`n               (preserves user settings)`n`nREPLACE: Backup existing profile and replace with`n                  imported profile"
            $lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $lblInfo.ForeColor = $theme.LabelTextColor
            $contentPanel.Controls.Add($lblInfo)

            # Merge button
            $btnMerge = New-Object System.Windows.Forms.Button
            $btnMerge.Location = New-Object System.Drawing.Point(180, 265)
            $btnMerge.Size = New-Object System.Drawing.Size(110, 35)
            $btnMerge.Text = "Merge"
            $btnMerge.BackColor = $theme.ButtonPrimaryBackColor
            $btnMerge.ForeColor = $theme.ButtonPrimaryForeColor
            $btnMerge.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnMerge.FlatAppearance.BorderSize = 0
            $btnMerge.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $btnMerge.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnMerge.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
            $btnMerge.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
            $btnMerge.Add_Click({ $choiceForm.Tag = "Merge"; $choiceForm.Close() })
            $choiceForm.Controls.Add($btnMerge)

            # Replace button
            $btnReplace = New-Object System.Windows.Forms.Button
            $btnReplace.Location = New-Object System.Drawing.Point(300, 265)
            $btnReplace.Size = New-Object System.Drawing.Size(110, 35)
            $btnReplace.Text = "Replace"
            $btnReplace.BackColor = $theme.ButtonDangerBackColor
            $btnReplace.ForeColor = $theme.ButtonDangerForeColor
            $btnReplace.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnReplace.FlatAppearance.BorderSize = 0
            $btnReplace.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $btnReplace.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnReplace.Add_MouseEnter({ $this.BackColor = $theme.ButtonDangerHoverBackColor })
            $btnReplace.Add_MouseLeave({ $this.BackColor = $theme.ButtonDangerBackColor })
            $btnReplace.Add_Click({ $choiceForm.Tag = "Replace"; $choiceForm.Close() })
            $choiceForm.Controls.Add($btnReplace)

            $choiceForm.ShowDialog() | Out-Null
            $choice = $choiceForm.Tag
            $choiceForm.Dispose()

            if ($choice -eq "Merge") {
                $mergeMode = $true
                Log-Message "MERGE MODE: Will merge imported files into existing profile"
            }
            elseif ($choice -eq "Replace") {
                $mergeMode = $false
                Log-Message "REPLACE MODE: Will backup and replace existing profile"
            }
            else {
                throw "Operation cancelled - no option selected"
            }
        }
        
        $global:StatusText.Text = "Resolving target SID..."
        [System.Windows.Forms.Application]::DoEvents()

        # RESOLVE SID FIRST - needed for profile mounted checks and registry operations
        if ($isDomain) {
            # --- DOMAIN REACHABILITY CHECK ---
            # --- DOMAIN REACHABILITY RETRY LOOP ---
            while ($true) {
                $reach = Test-DomainReachability -DomainName $domain
                if ($reach.Success) {
                    break # Success, proceed
                }
                
                Log-Message "Domain reachability check failed: $($reach.Error)"
                $res = Show-ModernDialog -Message "Domain Reachability Check Failed:`r`n`r`n$($reach.Error)`r`n`r`nImport cannot continue until the domain is reachable.`r`n`r`nRetry connection?" -Title "Domain Connection Error" -Type Error -Buttons YesNo
                
                if ($res -eq "No") {
                    Log-Message "User cancelled during reachability check."
                    throw "Domain controller unreachable: $($reach.Error)"
                }
                # Loop continues on "Yes" (Retry)
                $global:StatusText.Text = "Retrying domain check..."
                [System.Windows.Forms.Application]::DoEvents()
            }

            # Apply theme
            $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
            $theme = $Themes[$currentTheme]

            # Modern domain credential prompt
            # Modern domain credential prompt using shared helper
            # Modern domain credential prompt using shared helper with retry logic
            $cred = Get-DomainAdminCredential -DomainName $domain -InitialCredential $null
            
            if (-not $cred) {
                # Clean up and exit if cancelled
                throw "Credentials cancelled"
            }
            
            $username = $cred.UserName
            # Remove domain prefix for display if needed (though we use full cred object below)
            if ($username -match '\\') { $username = ($username -split '\\')[1] }
            
            # Store credentials globally for optional reuse in domain join
            $global:DomainCredential = $cred
            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
            try {
                $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('Domain', $domain, $cred.UserName, $cred.GetNetworkCredential().Password)
                $user = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($ctx, $shortName)
                if (-not $user) { throw "User '$shortName' not found in domain '$domain'" }
                $sid = $user.Sid.Value
            }
            catch {
                $exMsg = $_.Exception.Message
                # Clean up "Exception calling..." wrapper
                if ($exMsg -match 'Exception calling "FindByIdentity" with "2" argument\(s\): "(?<inner>.*)"') {
                    $exMsg = $matches['inner']
                }
                
                throw "Failed to verify user '$shortName' in domain: $exMsg"
            }
        }
        elseif ($isAzureAD) {
            # AzureAD user - use SID already resolved via Microsoft Graph
            if ($targetSID) {
                # We already resolved the SID via Microsoft Graph earlier
                $sid = $targetSID
                Log-Message "Using AzureAD SID from Microsoft Graph: $sid"
            }
            else {
                # Fallback: try to get SID from existing profile (if user has signed in)
                Log-Message "Resolving AzureAD user SID from Win32_UserProfile..."
                try {
                    $azureProfile = Get-WmiObject Win32_UserProfile | Where-Object {
                        $_.LocalPath -like "*\$shortName" -and (Test-IsAzureADSID $_.SID)
                    }
                    
                    if (-not $azureProfile) {
                        throw "AzureAD user '$shortName' profile not found. User must sign in at least once to create profile."
                    }
                    
                    $sid = $azureProfile.SID
                    Log-Message "AzureAD user SID resolved from profile: $sid"
                }
                catch {
                    throw "Failed to resolve AzureAD SID for user '$shortName': $_"
                }
            }
        }
        else {
            # Local user - use NTAccount translation
            try {
                $nt = New-Object System.Security.Principal.NTAccount($shortName)
                $sid = $nt.Translate([System.Security.Principal.SecurityIdentifier]).Value
            }
            catch {
                throw "Failed to resolve SID for user '$shortName'. Ensure the user exists: $_"
            }
        }
        Log-Message "Target SID resolved: $sid"

        # === DOMAIN JOIN ===
        if ($isDomain -and $global:DomainCheckBox -and $global:DomainCheckBox.Checked) {
            try {
                Log-Message "Domain join requested. Initiating join..."
                $domainName = $global:DomainNameTextBox.Text.Trim()
                $computerName = $global:ComputerNameTextBox.Text.Trim()
                $restartMode = if ($global:RestartComboBox.SelectedItem) { $global:RestartComboBox.SelectedItem } else { 'Prompt' }
                $delaySeconds = 30
                if ($global:DelayTextBox.Enabled -and $global:DelayTextBox.Text -match '^[0-9]+$') {
                    $delaySeconds = [int]$global:DelayTextBox.Text
                }
                $cred = $null
                if ($global:DomainCredential) { $cred = $global:DomainCredential }
                Join-Domain-Enhanced -ComputerName $computerName -DomainName $domainName -RestartBehavior $restartMode -DelaySeconds $delaySeconds -Credential $cred
            }
            catch {
                Log-Error "Domain join failed: $_"
                Show-ModernDialog -Message "Domain join failed:`n$_" -Title "Domain Join Error" -Type Error -Buttons OK
            }
        }
        
        $global:StatusText.Text = "Preparing target directory..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # PRE-FLIGHT VALIDATION: Record whether target existed before validation and check writeability
        $targetExistedBefore = Test-Path $target
        Log-Message "Running pre-flight validation... (target existed before: $targetExistedBefore)"
        if ($targetExistedBefore -and -not (Test-ProfilePathWriteable $target)) {
            throw "Target path '$target' is not writable. Check permissions."
        }
        if ($targetExistedBefore) {
            Log-Message "Target path writeability verified"
        }

        # === PRE-FLIGHT DISK CHECK ===
        # Perform this BEFORE backup to avoid time/IO waste if space is insufficient
        $global:StatusText.Text = "Checking disk space..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            # 1. Get Uncompressed Size
            $checkZipSize = Get-ZipUncompressedSize -SevenZipPath $global:SevenZipPath -ZipPath $ZipPath
            if ($checkZipSize -gt 0) {
                # Determine where we are putting the files
                # Loop variable re-use issue - renamed to checkZipSize
                $targetDriveRoot = [IO.Path]::GetPathRoot($target) # Import uses $target
                
                # Add 10% safety buffer
                $requiredSpace = $checkZipSize * 1.1 
                
                $destDrive = Get-PSDrive -Name $targetDriveRoot.TrimEnd('\', ':') -ErrorAction Stop
                $destFree = $destDrive.Free
                
                # MOCK FOR TESTING: Force low space
                # $destFree = 1MB
                
                Log-Info "Disk check: Need $([Math]::Round($requiredSpace/1GB,2)) GB, Have $([Math]::Round($destFree/1GB,2)) GB on $targetDriveRoot"
                
                if ($destFree -lt $requiredSpace) {
                    throw "Insufficient disk space on '$targetDriveRoot'. Available: $([Math]::Round($destFree/1GB,2)) GB. Required: $([Math]::Round($requiredSpace/1GB,2)) GB."
                }
            }
            else {
                Log-Warning "Could not determine ZIP uncompressed size. Skipping disk check."
            }
        }
        catch {
            if ($_.Exception.Message -like "*Insufficient*") { throw $_ }
            Log-Warning "Disk space check failed to execute: $_"
        }
        
        # PRE-FLIGHT CHECK: Detect if profile is currently mounted/logged in
        if (Test-ProfileMounted $sid) {
            Log-Message "WARNING: User profile is currently mounted (user may be logged in via HKU)"
            $res = Show-ModernDialog -Message "User profile is currently mounted in registry (HKU\$sid).`n`nThis may indicate the user is logged in. Importing now may cause data loss or corruption.`r`n`r`nContinue anyway?" -Title "Profile Mounted" -Type Warning -Buttons YesNo
            if ($res -ne "Yes") {
                throw "Cancelled - profile is mounted"
            }
        }
        else {
            Log-Message "Profile not mounted - safe to proceed"
        }
        
        # BACKUP EXISTING PROFILE: Create timestamped backup before importing
        if ($targetExistedBefore) {
            Log-Message "Creating backup for existing profile: $target"
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $backup = "$target.backup_$timestamp"
            $global:ImportBackupPath = $backup
            
            if ($mergeMode) {
                Log-Message "MERGE MODE: Backup is for safety only (will not replace profile)"
            }
            else {
                Log-Message "REPLACE MODE: Backup before replacement"
            }
            
            $global:StatusText.Text = "Backing up existing profile..."
            $global:ProgressBar.Value = 1
            [System.Windows.Forms.Application]::DoEvents()
            $backupSuccessful = $false
            
            # OPTIMIZATION: Try fast rename first (Replace Mode Only)
            # This is O(1) instead of O(N) and saves disk space
            if (-not $mergeMode) {
                try {
                    Log-Message "Attempting fast backup (rename/move)..."
                    Rename-Item -Path $target -NewName (Split-Path $backup -Leaf) -ErrorAction Stop
                    
                    # Recreate empty target folder for the import
                    New-Item -Path $target -ItemType Directory -Force | Out-Null
                    
                    # Inherit permissions from parent temporarily (Set-ProfileAcls fixes this later)
                    
                    Log-Message "Fast backup successful"
                    $backupSuccessful = $true
                }
                catch {
                    Log-Warning "Fast backup failed (folder locked?), falling back to robust copy. Error: $_"
                }
            }

            if (-not $backupSuccessful) {
                try {
                    Log-Message "Backup path: $backup"
                    # Use robocopy for faster backup (multi-threaded, handles locked files)
                    # Dynamic thread count: 8-32 threads based on CPU cores for optimal performance
                    $robocopyThreads = $Config.RobocopyThreads
                    Log-Message "Robocopy backup using $robocopyThreads threads (detected $cpuCores CPU cores)"
                    $robocopyArgs = @(
                        "`"$target`"",
                        "`"$backup`"",
                        '/E',           # Copy subdirectories including empty
                        '/COPYALL',     # Copy all file info (timestamps, ACLs, owner, auditing)
                        '/XJ',          # Exclude junction points (prevent infinite loops)
                        '/XD', 'WindowsApps', 'SFAP', 'Packages',  # Exclude system-managed directories
                        '/R:2',         # Retry 2 times on failed copies
                        '/W:1',         # Wait 1 second between retries
                        "/MT:$robocopyThreads",  # Multi-threaded (dynamic based on CPU cores)
                        '/256',         # Enable long path support (>260 chars)
                        '/NP',          # No progress per file (reduces output)
                        '/NFL',         # No file list
                        '/NDL'          # No directory list
                    )
                    
                    Log-Message "Starting robocopy backup..."
                    $robocopyProc = Start-Process 'robocopy.exe' -ArgumentList $robocopyArgs -PassThru -NoNewWindow
                    
                    # Monitor backup progress by checking destination folder growth
                    $startTime = [DateTime]::Now
                    while (-not $robocopyProc.HasExited) {
                        Start-Sleep -Milliseconds 500
                        
                        # Update UI to show we're still working
                        $elapsed = ([DateTime]::Now - $startTime).TotalSeconds
                        $global:StatusText.Text = "Backing up existing profile... $([int]$elapsed)s"
                        
                        # Pulse progress bar (1-4 range)
                        $pulseValue = 1 + (($elapsed % 3) / 3 * 3)
                        $global:ProgressBar.Value = [Math]::Min(4, $pulseValue)
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                    
                    $robocopyProc.WaitForExit()
                    
                    # Robocopy exit codes: 0-7 are success (0=no changes, 1=files copied, 2=extra files, etc.)
                    if ($robocopyProc.ExitCode -gt 7) {
                        throw "Robocopy failed with exit code $($robocopyProc.ExitCode)"
                    }
                    Log-Message "Backup completed (robocopy exit code: $($robocopyProc.ExitCode))"
                    $backupSuccessful = $true
                }
                catch {
                    Log-Message "WARNING: Backup failed: $_"
                }
            }
            
            # Only remove profile directory if not in merge mode
            if (-not $mergeMode) {
                Log-Message "Removing existing profile directory: $target"
                $global:StatusText.Text = "Removing existing profile directory..."
                [System.Windows.Forms.Application]::DoEvents()
                Remove-FolderRobust -Path $target
                Log-Message "Existing directory removed"
            }
            else {
                Log-Message "MERGE MODE: Keeping existing profile directory for merge"
            }
        }
        else {
            Log-Message "No existing profile found at $target"
            $global:ImportBackupPath = $null
        }
        
        # Start import timer for elapsed time logging
        $global:ImportStartTime = [DateTime]::Now
        Log-Message "Import operation started at: $($global:ImportStartTime.ToString('yyyy-MM-dd HH:mm:ss'))"

        # Determine extraction location
        # In merge mode: extract to temp location first, then merge
        # In replace mode: extract directly to target
        $extractionTarget = $target
        $tempMergeLocation = $null
        if ($mergeMode) {
            $tempMergeLocation = Join-Path $env:TEMP ("ProfileMerge_$shortName_$(Get-Random)")
            $extractionTarget = $tempMergeLocation
            Log-Message "MERGE MODE: Will extract to temp location first: $tempMergeLocation"
        }

        # Create target directory for extraction (clean directory after backup/removal)
        if (-not (Test-Path $extractionTarget)) { 
            New-Item -ItemType Directory -Path $extractionTarget -Force | Out-Null 
        }
        
        if ($mergeMode) {
            Log-Message "MERGE MODE: Extracting to temp location for merging: $extractionTarget"
        }
        else {
            Log-Message "REPLACE MODE: Extracting directly to final profile location: $extractionTarget"
        }
        
        $global:ProgressBar.Value = 5
        $global:StatusText.Text = "Extracting ZIP archive..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Starting 7-Zip extraction..."

        # === PRE-FLIGHT ZIP INTEGRITY CHECK ===
        $global:StatusText.Text = "Verifying archive integrity..."
        [System.Windows.Forms.Application]::DoEvents()
        if (-not (Test-ZipIntegrity -SevenZipPath $global:SevenZipPath -ZipPath $ZipPath)) {
            throw "The ZIP archive is corrupted or invalid. Import cannot proceed."
        }
        
        # === PRE-FLIGHT DISK CHECK ===
        $global:StatusText.Text = "Checking disk space..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            # 1. Get Uncompressed Size
            $uncompressedSize = Get-ZipUncompressedSize -SevenZipPath $global:SevenZipPath -ZipPath $ZipPath
            if ($uncompressedSize -gt 0) {
                $requiredSpace = $uncompressedSize * 1.1 # 10% safety buffer
                
                # 2. Get Free Space
                $destRoot = [IO.Path]::GetPathRoot($extractionTarget)
                $destDrive = Get-PSDrive -Name $destRoot.TrimEnd('\', ':') -ErrorAction Stop
                
                $destFree = $destDrive.Free
                

                
                Log-Info "Disk check: Need $([Math]::Round($requiredSpace/1GB,2)) GB, Have $([Math]::Round($destFree/1GB,2)) GB on $destRoot"
                
                if ($destFree -lt $requiredSpace) {
                    # If we are in REPLACE mode, we might be freeing up space by deleting the old profile first?
                    # But we extract first, THEN swap (verify this logic). 
                    # Current logic: Extract to target (Replace) or Temp (Merge). 
                    # For Replace: we extract safely to target? 
                    # Wait, earlier logic says: "Extract directly to final profile location" if Replace.
                    # If target exists and we are replacing, lines 9792 removed it?
                    # Yes: "Only remove profile directory if not in merge mode" -> Line 9792.
                    # So space IS available.
                     
                    # But we must ensure specific logic matches.
                    throw "Insufficient disk space on '$destRoot'. Available: $([Math]::Round($destFree/1GB,2)) GB. Required: $([Math]::Round($requiredSpace/1GB,2)) GB."
                }
            }
            else {
                Log-Warning "Could not determine ZIP uncompressed size. Skipping disk check."
            }
        }
        catch {
            if ($_.Exception.Message -like "*Insufficient*") { throw $_ }
            Log-Warning "Disk space check failed to execute: $_"
        }

        # Build 7-Zip exclusion filters for Import mode
        $importExclusions = Get-RobocopyExclusions -Mode 'Import'
        $7zExclusions = @()
        
        # Exclude LOG files explicitly (not wildcards that might match NTUSER.DAT)
        $7zExclusions += "-x!ntuser.dat.LOG1"
        $7zExclusions += "-x!ntuser.dat.LOG2"
        $7zExclusions += "-x!NTUSER.DAT.LOG1"
        $7zExclusions += "-x!NTUSER.DAT.LOG2"
        $7zExclusions += "-x!UsrClass.dat.LOG1"
        $7zExclusions += "-x!UsrClass.dat.LOG2"
        
        # Exclude temp/cache files
        $7zExclusions += "-x!Thumbs.db"
        $7zExclusions += "-x!desktop.ini"
        
        # Exclude temp directories
        $7zExclusions += "-xr!AppData/Local/Temp"
        $7zExclusions += "-xr!AppData/LocalLow"
        
        # Exclude Outlook OST files (cached Exchange data - rebuilds automatically)
        $7zExclusions += "-x!*.ost"
        Log-Message "Excluding Outlook OST files (will rebuild automatically from Exchange)"
        
        Log-Message "Import exclusions: $($7zExclusions.Count) filters applied"
        
        # Extract with exclusion filters
        # Multi-threading: Use all CPU cores for maximum extraction speed (2-3x faster on multi-core systems)
        $threadCount = $Config.SevenZipThreads
        Log-Message "7-Zip extraction using $threadCount threads (detected $cpuCores CPU cores)"
        $args = @('x', "`"$ZipPath`"", "-o`"$extractionTarget`"", '-y', "-mmt=$threadCount", '-bsp1') + $7zExclusions
        
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $global:SevenZipPath
        $psi.Arguments = $args -join ' '
        $psi.CreateNoWindow = $true
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        
        Log-Message "7-Zip command: $global:SevenZipPath $($args -join ' ')"
        
        Log-Message "Starting 7-Zip extraction with progress monitoring..."
        $extractProcess = [System.Diagnostics.Process]::Start($psi)
        
        # Monitor extraction progress
        $lastUpdate = [DateTime]::Now
        $buffer = New-Object System.Text.StringBuilder
        while (-not $extractProcess.HasExited) {
            if ($global:CancelRequested) {
                try { $extractProcess.Kill() } catch { }
                throw "Import cancelled by user during extraction"
            }
            $line = $extractProcess.StandardOutput.ReadLine()
            if ($line) {
                $buffer.AppendLine($line) | Out-Null
                # Parse progress from 7-Zip output (format: "  5%")
                if ($line -match '^\s*(\d+)%') {
                    $percent = [int]$matches[1]
                    # Map 7-Zip progress (0-100%) to our progress range (5-20%)
                    $mappedProgress = 5 + ($percent * 0.15)
                    $global:ProgressBar.Value = [Math]::Min(20, $mappedProgress)
                    $global:StatusText.Text = "Extracting files... $percent%"
                    [System.Windows.Forms.Application]::DoEvents()
                }
            }
            elseif (([DateTime]::Now - $lastUpdate).TotalSeconds -ge 1) {
                $lastUpdate = [DateTime]::Now
                [System.Windows.Forms.Application]::DoEvents()
            }
            Start-Sleep -Milliseconds 100
        }
        
        # Consume any remaining output
        $buffer.Append($extractProcess.StandardOutput.ReadToEnd()) | Out-Null
        
        $extractProcess.WaitForExit()
        Log-Message "7-Zip exit code: $($extractProcess.ExitCode)"
        if ($extractProcess.ExitCode -ne 0) { 
            throw "Extract failed (ExitCode: $($extractProcess.ExitCode))"
        }
        
        if ($mergeMode) {
            Log-Message "ZIP extracted successfully to temp location for merging"
        }
        else {
            Log-Message "ZIP extracted successfully to final profile location"
        }
        
        # === CRITICAL: Verify extraction integrity ===
        $global:ProgressBar.Value = 20
        $global:StatusText.Text = "Verifying extracted files..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Post-extraction verification..."
        
        # List what was actually extracted
        $extractedItems = @()
        try {
            $extractedItems = Get-ChildItem -Path $extractionTarget -Force -ErrorAction SilentlyContinue | Select-Object -First 20
            Log-Message "Items in extraction directory:"
            foreach ($item in $extractedItems) {
                Log-Message "  - $($item.Name) $(if ($item.PSIsContainer) { '[DIR]' } else { "($($item.Length) bytes)" })"
            }
            if ((Get-ChildItem -Path $extractionTarget -Force -ErrorAction SilentlyContinue | Measure-Object).Count -gt 20) {
                Log-Message "  ... and more"
            }
        }
        catch {
            Log-Message "WARNING: Could not list extraction directory: $_"
        }
        
        # Check NTUSER.DAT exists and has reasonable size
        # SKIP THIS CHECK IN MERGE MODE - we delete NTUSER.DAT anyway to preserve target user's registry
        if (-not $mergeMode) {
            # Retry a few times in case of timing issues
            $hiveSrc = "$extractionTarget\NTUSER.DAT"
            $hiveFound = $false
            $hiveSize = 0
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    if (Test-Path $hiveSrc -ErrorAction Stop) {
                        $fileObj = Get-Item $hiveSrc -Force -ErrorAction Stop
                        $hiveSize = $fileObj.Length
                        $hiveFound = $true
                        Log-Message "NTUSER.DAT found on attempt $($attempt): $hiveSize bytes"
                        break
                    }
                }
                catch {
                    Log-Message "Attempt $($attempt) to access NTUSER.DAT failed: $_"
                    if ($attempt -lt 3) {
                        Start-Sleep -Milliseconds 500
                    }
                }
            }
            
            if (-not $hiveFound) {
                throw "CRITICAL: NTUSER.DAT missing or inaccessible after extraction! ZIP may be corrupted or extraction failed."
            }
            
            if ($hiveSize -lt 100000) {
                throw "CRITICAL: NTUSER.DAT is suspiciously small ($hiveSize bytes) - extraction corrupted?"
            }
        }
        else {
            Log-Message "MERGE MODE: Skipping NTUSER.DAT validation (will be deleted to preserve target user's registry)"
        }
        
        # Check manifest
        $manifestPath = "$extractionTarget\manifest.json"
        if (-not (Test-Path $manifestPath)) {
            throw "manifest.json missing after extraction"
        }
        Log-Message "Manifest found"
        
        # Parse manifest for source SID
        $manifest = $null
        try {
            $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
            Log-Message "Manifest parsed successfully"
        }
        catch {
            Log-Message "WARNING: Could not parse manifest: $_"
        }
        
        # === AZUREAD/ENTRA ID VALIDATION ===
        if ($manifest -and $manifest.IsAzureADUser) {
            Log-Message "Detected AzureAD/Entra ID profile in manifest"
            
            # Verify that the username was correctly identified as AzureAD
            # (Either explicit "AzureAD\username" format OR auto-detected by SID pattern)
            if (-not $isAzureAD) {
                Log-Message "WARNING: Manifest indicates AzureAD profile but username was not detected as AzureAD user"
                Log-Message "Current username format: $importTargetUser (parsed as $(if ($isDomain) { 'DOMAIN' } else { 'LOCAL' }) user)"
                
                # Suggest correction
                $suggestFormat = "AzureAD\$($manifest.Username)"
                $formatResult = Show-ModernDialog -Message "Import mismatch detected!`n`nThe profile was exported from an AzureAD/Entra ID user, but the target username '$importTargetUser' was not recognized as an AzureAD account.`n`nDid you mean: $suggestFormat ?`r`n`r`nClick Yes to correct the username format, or No to continue (may fail)." -Title "AzureAD Profile Mismatch" -Type Warning -Buttons YesNo
                if ($formatResult -eq "Yes") {
                    # Force correct AzureAD format
                    $Username = $suggestFormat
                    $shortName = $manifest.Username
                    $isAzureAD = $true
                    $isDomain = $false
                    Log-Message "Username corrected to: $Username (AzureAD format)"
                }
                else {
                    Log-Message "User chose to continue with current username format - import may fail"
                }
            }
            else {
                Log-Message "AzureAD user correctly identified - proceeding with import"
            }
            
            # Only require AzureAD join if the target user is AzureAD (after possible correction)
            if ($isAzureAD) {
                if (-not (Test-IsAzureADJoined)) {
                    Log-Message "WARNING: Importing AzureAD profile but system is not AzureAD joined"
                    # Show guidance dialog
                    $dlgResult = Show-AzureADJoinDialog -Username $manifest.Username
                    if ($dlgResult -eq "Cancel") {
                        throw "Import cancelled - AzureAD join required for this profile"
                    }
                    # User clicked Continue Import - re-check join status
                    if (-not (Test-IsAzureADJoined)) {
                        throw "System is still not AzureAD joined. Please complete the join process in Settings > Accounts > Access work or school and try again."
                    }
                    Log-Message "AzureAD join status verified after user completed join process"
                }
                Log-Message "System is AzureAD joined - proceeding with import"
            }
            else {
                Log-Message "Target user is not AzureAD after mismatch dialog. Skipping AzureAD join check."
            }
        }
        
        $global:ProgressBar.Value = 25
        # Install Winget apps from the correct location based on mode
        if ($mergeMode) {
            # In merge mode, Winget JSON is in temp extraction location
            Install-WingetAppsFromExport -TargetProfilePath $extractionTarget
        }
        else {
            # In replace mode, Winget JSON is in final profile location
            Install-WingetAppsFromExport -TargetProfilePath $target
        }
        $global:StatusText.Text = "Resolving user SID..."
        [System.Windows.Forms.Application]::DoEvents()
        Log-Message "Resolving SID for user..."
        $sourceSID = $null
        if ($manifest) {
            $sourceSID = $manifest.SourceSID
            $oldProfilePath = $manifest.ProfilePath
            Log-Message "Source SID from manifest: $sourceSID"
            Log-Message "Source Profile Path from manifest: $oldProfilePath"
            
            
            # Profile timestamp available for diagnostics
            if ($manifest.ProfileTimestamp) {
                Log-Message "Profile timestamp from export: $($manifest.ProfileTimestamp)"
            }
        }
        else {
            Log-Message "WARNING: manifest.json not available - cannot verify source SID"
        }

        # === MERGE MODE: Merge extracted files into existing profile ===
        if ($mergeMode) {
            $global:ProgressBar.Value = 28
            $global:StatusText.Text = "Merging profile files..."
            [System.Windows.Forms.Application]::DoEvents()
            Log-Message "MERGE MODE: Merging files from extracted profile into existing profile..."
            Log-Message "MERGE MODE: Preserving existing user's NTUSER.DAT (keeping user settings/preferences)"
            
            try {
                # In merge mode, we want to preserve the existing user's NTUSER.DAT hive
                # This keeps their current registry settings, preferences, and avoids SID rewriting
                # We'll delete NTUSER.DAT from temp extraction before merging
                $hiveFileToSkip = Join-Path $tempMergeLocation "NTUSER.DAT"
                if (Test-Path $hiveFileToSkip) {
                    Log-Message "Removing NTUSER.DAT from extracted profile (will use existing user's hive)"
                    Remove-Item $hiveFileToSkip -Force -ErrorAction SilentlyContinue
                }
                
                # Also skip UsrClass.dat (user class registry hive)
                $userClassToSkip = Join-Path $tempMergeLocation "AppData\Local\Microsoft\Windows\UsrClass.dat"
                if (Test-Path $userClassToSkip) {
                    Log-Message "Removing UsrClass.dat from extracted profile (will use existing user's registry)"
                    Remove-Item $userClassToSkip -Force -ErrorAction SilentlyContinue
                }
                
                # Use robocopy to merge remaining files, skipping identical files and not overwriting newer files
                # in existing profile (to preserve any recent customizations)
                # Dynamic thread count: 8-32 threads based on CPU cores for optimal performance
                $robocopyThreads = $Config.RobocopyThreads
                Log-Message "Robocopy merge using $robocopyThreads threads (detected $cpuCores CPU cores)"
                $mergeArgs = @(
                    "`"$tempMergeLocation`"",      # Source profile root
                    "`"$target`"",                  # Destination profile root

                    '/E',                           # Include subdirectories (even empty)

                    # Copy ONLY files that do not already exist in destination
                    '/XO',                          # Exclude older files
                    '/XN',                          # Exclude newer files
                    '/XC',                          # Exclude changed files

                    # Safe metadata handling (do NOT touch ownership or ACLs)
                    '/COPY:DAT',                    # Data, Attributes, Timestamps only
                    '/DCOPY:DAT',                   # Preserve directory timestamps

                    # Exclude system- and app-managed folders
                    '/XD',
                    'WindowsApps',
                    'SFAP',
                    'Packages',
                    'AppData\Local\Microsoft\Windows',
                    'AppData\Local\Temp',
                    'AppData\Local\Microsoft\Edge',
                    'AppData\Local\Microsoft\OneDrive',

                    # Reliability & performance
                    '/R:2',                         # Retry twice on failure
                    '/W:1',                         # Wait 1 second between retries
                    "/MT:$robocopyThreads",         # Multithreaded copy
                    '/256',                         # Long path support

                    # Quiet output
                    '/NP',                          # No progress percentage
                    '/NFL',                         # No file list
                    '/NDL'                          # No directory list
                )
                
                Log-Message "Robocopy merge command: robocopy $($mergeArgs -join ' ')"
                $mergeProcess = Start-Process -FilePath 'robocopy.exe' -ArgumentList $mergeArgs -Wait -NoNewWindow -PassThru
                
                # Robocopy exit codes: 0-7 are success, 8+ are errors
                if ($mergeProcess.ExitCode -gt 7) {
                    throw "Robocopy merge failed with exit code $($mergeProcess.ExitCode)"
                }
                Log-Message "Robocopy merge completed with exit code $($mergeProcess.ExitCode)"
                
                # Cleanup temp merge location
                Log-Message "Cleaning up temporary merge location: $tempMergeLocation"
                Remove-FolderRobust -Path $tempMergeLocation
                Log-Message "Temporary merge location cleaned up successfully"
                
                # Verify key files exist in target after merge
                # Note: We skip NTUSER.DAT verification in merge mode since we're keeping the existing one
                $keyFiles = @('AppData', 'Desktop', 'Documents')
                $missingFiles = @()
                foreach ($keyFile in $keyFiles) {
                    if (-not (Test-Path (Join-Path $target $keyFile) -ErrorAction SilentlyContinue)) {
                        $missingFiles += $keyFile
                    }
                }
                if ($missingFiles.Count -gt 0) {
                    Log-Message "WARNING: Some folders missing after merge (may be OK): $($missingFiles -join ', ')"
                }
                else {
                    Log-Message "Merge verification: All key profile folders present in target"
                }
                
            }
            catch {
                Log-Message "ERROR during merge: $_"
                throw "Profile merge failed: $_"
            }
        }

        $global:ProgressBar.Value = 30
        $global:StatusText.Text = "Cleaning stale registry entries..."
        [System.Windows.Forms.Application]::DoEvents()
        # Only clean stale registry entries in REPLACE mode
        # In MERGE mode, we're merging into an existing profile, so we must preserve the target user's registry entry
        if (-not $mergeMode) {
            Log-Message "Checking for stale profile registry entries..."
            $base = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
            $staleKeys = Get-ChildItem $base -ErrorAction SilentlyContinue | Where-Object {
                $_.PSChildName -like "$sid*" -or ((Get-ItemProperty $_.PSPath -Name ProfileImagePath -EA SilentlyContinue).ProfileImagePath -like "*\$shortName*")
            }
            if ($staleKeys.Count -gt 0) {
                Log-Message "Found $($staleKeys.Count) stale registry entries - removing..."
                foreach ($k in $staleKeys) {
                    Remove-Item $k.PSPath -Force -Recurse -ErrorAction SilentlyContinue
                    Log-Message "Removed: $($k.PSChildName)"
                }
            }
            else {
                Log-Message "No stale registry entries found"
            }
        }
        else {
            Log-Message "MERGE MODE: Skipping stale registry cleanup (preserving existing user's ProfileList entry)"
        }

        # Also remove any abandoned Temp* profile folders (excluding current target user)
        try {
            # Match any profile folder starting with Temp (e.g. Temp, Temp., Temp-XYZ, Temp123, temp.anything)
            $tempFolders = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^Temp.*' }
            foreach ($tf in $tempFolders) {
                if ($tf.Name -ieq $shortName) { continue } # Never remove the target user folder if it matches
                Log-Message "Removing temporary profile folder: $($tf.FullName)"
                Remove-FolderRobust -Path $tf.FullName
            }
        }
        catch {
            Log-Message "WARNING: Failed Temp* profile cleanup: $_"
        }
        $global:ProgressBar.Value = 35
        
        # Log appropriate completion message based on mode
        if ($mergeMode) {
            Log-Message "MERGE MODE: Profile files successfully merged into existing profile at $target"
        }
        else {
            Log-Message "REPLACE MODE: Files extracted directly to target: $target"
        }
        $global:ProgressBar.Value = 95
        $global:StatusText.Text = "Applying permissions..."
        [System.Windows.Forms.Application]::DoEvents()
        
        if (-not $isDomain) {
            if ($mergeMode) {
                Log-Message "MERGE MODE: Skipping SID rewriting (using existing user's NTUSER.DAT)"
                Log-Message "MERGE MODE: Only applying folder ACLs (not touching existing hive)"
                
                # In merge mode, only set folder permissions without SID rewriting
                $global:StatusText.Text = "Applying folder permissions (merge mode)..."
                [System.Windows.Forms.Application]::DoEvents()
                
                # Apply basic folder ACLs without hive manipulation
                try {
                    $acl = Get-Acl $target
                    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($shortName, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
                    $acl.SetAccessRule($accessRule)
                    Set-Acl -Path $target -AclObject $acl
                    Log-Message "Basic folder permissions applied for merge mode"
                }
                catch {
                    Log-Message "WARNING: Could not set basic folder ACLs: $_"
                }
            }
            else {
                Log-Message "Setting ACLs for $(if ($isAzureAD) { 'AzureAD' } else { 'local' }) user $shortName (SID: $sid)"
                Set-ProfileAcls -ProfileFolder $target -UserName $shortName -SourceSID $sourceSID -UserSID $sid -OldProfilePath $oldProfilePath -NewProfilePath $target
            }
        }
        else {
            # DOMAIN USER - Set-ProfileAcls handles SID translation + ACLs
            $global:StatusText.Text = "Applying DOMAIN user permissions + SID translation..."
            [System.Windows.Forms.Application]::DoEvents()
            Log-Message "DOMAIN USER: Applying profile ACLs for $importTargetUser..."
            Set-ProfileAcls -ProfileFolder $target -UserName $importTargetUser -SourceSID $sourceSID -UserSID $sid -OldProfilePath $oldProfilePath -NewProfilePath $target
            Log-Message "DOMAIN user profile fully prepared - first logon will succeed"
        }
        if (-not $isDomain) {
            $global:StatusText.Text = "Configuring user for login screen..."
            [System.Windows.Forms.Application]::DoEvents()
            Log-Message "Configuring local user for interactive login..."
            try {
                $localUser = Get-LocalUser -Name $shortName -ErrorAction Stop
                if ($localUser.PasswordExpires) {
                    Set-LocalUser -Name $shortName -PasswordNeverExpires $true
                    Log-Message "Set password to never expire"
                }
                if (-not $localUser.Enabled) {
                    Enable-LocalUser -Name $shortName
                    Log-Message "Enabled user account"
                }
                try {
                    Add-LocalGroupMember -Group "Users" -Member $shortName -ErrorAction SilentlyContinue
                    Log-Message "Confirmed user is in Users group"
                }
                catch {
                    Log-Message "User already in Users group"
                }
                Log-Message "User account configured for login"
            }
            catch {
                Log-Message "WARNING: Could not fully configure user account: $_"
            }
            Log-Message "Configuring registry settings for login screen visibility..."
            try {
                $winlogonPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                if (Test-Path $winlogonPath) {
                    Set-ItemProperty -Path $winlogonPath -Name "DontDisplayLastUserName" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                    Log-Message "Configured Winlogon to show user list"
                }
                $specialAccountsPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                if (-not (Test-Path $specialAccountsPath)) { New-Item -Path $specialAccountsPath -Force | Out-Null }
                Set-ItemProperty -Path $specialAccountsPath -Name "DontDisplayLastUserName" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                $hiddenAccountsPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList"
                if (Test-Path $hiddenAccountsPath) {
                    Remove-ItemProperty -Path $hiddenAccountsPath -Name $shortName -ErrorAction SilentlyContinue
                    Log-Message "Ensured user is not in hidden accounts list"
                }
                Log-Message "Registry configured for login screen visibility"
            }
            catch {
                Log-Message "WARNING: Could not configure registry for login screen: $_"
            }
        }
        
        # === PROFILE REGISTRATION (REPLACE MODE ONLY) ===
        # In merge mode, the target user already exists with a valid registry entry
        # We should NOT touch the registry at all - just copy files
        if (-not $mergeMode) {
            $global:ProgressBar.Value = 93
            $global:StatusText.Text = "Registering profile in Windows..."
            [System.Windows.Forms.Application]::DoEvents()
            Log-Message "Registering profile in ProfileList..."
            $profileKey = "$base\$sid"
            New-Item -Path $profileKey -Force | Out-Null
            Set-ItemProperty -Path $profileKey -Name ProfileImagePath -Value $target -Type ExpandString -Force
            $dwordProps = "Flags", "State", "RefCount", "ProfileLoadTimeLow", "ProfileLoadTimeHigh"
            foreach ($p in $dwordProps) { Set-ItemProperty -Path $profileKey -Name $p -Value 0 -Type DWord -Force }
            Set-ItemProperty -Path $profileKey -Name CentralProfile -Value "" -Type String -Force
            Log-Message "Creating ProfileList SID binary value..."
            $userSidObj = [System.Security.Principal.SecurityIdentifier]::new($sid)
            $sidBytes = [byte[]]::new($userSidObj.BinaryLength)
            $userSidObj.GetBinaryForm($sidBytes, 0)
            Set-ItemProperty -Path $profileKey -Name Sid -Value $sidBytes -Type Binary -Force
            Log-Message "SID binary value created successfully ($($sidBytes.Length) bytes)"
            Log-Message "Profile registered: $profileKey"
            Log-Message "ProfileImagePath: $target"
        }
        else {
            Log-Message "MERGE MODE: Skipping registry operations - target user already exists"
            $global:ProgressBar.Value = 93
            $global:StatusText.Text = "Merge complete - skipping registry..."
            [System.Windows.Forms.Application]::DoEvents()
        }

        # === VERIFY: Ensure NTUSER.DAT can be loaded (prevent temporary profile issues) ===
        function Test-AndFix-ProfileHive {
            param(
                [string]$HivePath,
                [string]$UserId,
                [string]$UserName
            )

            $testMount = "HKU\ProfileCheck_$(Get-Random)"
            try {
                Log-Message "Attempting to load hive for verification: $HivePath"
                $out = reg load "$testMount" "$HivePath" 2>&1
                if ($LASTEXITCODE -ne 0) {
                    Log-Message "Initial hive load failed: $out"
                    Log-Message "Attempting to repair NTUSER.DAT ownership and ACLs..."

                    # Take ownership as Administrators and grant full control, then set to user
                    try {
                        takeown /F "$HivePath" /A /R /D Y >$null 2>&1
                        icacls "$HivePath" /grant "Administrators:(F)" /C /Q >$null 2>&1
                        icacls "$HivePath" /setowner *"$UserName" /Q >$null 2>&1
                        icacls "$HivePath" /grant:r *"${UserName}:(F)" /Q >$null 2>&1
                        Start-Sleep -Milliseconds 500
                    }
                    catch {
                        Log-Message "Failed to adjust NTUSER.DAT ACLs: $_"
                    }

                    # Retry load
                    $out2 = reg load "$testMount" "$HivePath" 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Log-Message "Retry hive load failed: $out2"
                        return $false
                    }
                }

                # Successfully loaded - unload and return success
                reg unload "$testMount" | Out-Null
                Log-Message "Hive verification succeeded"
                return $true
            }
            catch {
                Log-Message "Exception during hive verification: $_"
                try { reg unload "$testMount" | Out-Null } catch {}
                return $false
            }
        }

        # Run verification and attempt fixes if needed
        # In merge mode, we skip hive verification since we kept the existing user's NTUSER.DAT
        if (-not $mergeMode) {
            $hivePathToTest = "$target\NTUSER.DAT"
            $verifyOk = Test-AndFix-ProfileHive -HivePath $hivePathToTest -UserId $sid -UserName $shortName
            if (-not $verifyOk) {
                Log-Message "CRITICAL: NTUSER.DAT verification failed - profile cannot be used!"
                Log-Message "The hive file cannot be loaded into registry. This will result in a temporary profile on login."
                Log-Message "Possible causes:"
                Log-Message "  1. Hive file corrupted during extraction"
                Log-Message "  2. Owner/permissions incorrect (cannot be fixed with icacls)"
                Log-Message "  3. Hive file has incompatible format"
                Log-Message "Attempting additional diagnostics..."
                
                # Additional diagnostics
                $fileInfo = Get-Item $hivePathToTest -Force -ErrorAction SilentlyContinue
                if ($fileInfo) {
                    Log-Message "File exists: Yes"
                    Log-Message "File size: $($fileInfo.Length) bytes"
                    Log-Message "File ACL: $(icacls $hivePathToTest 2>&1 | Select-Object -First 3)"
                }
                else {
                    Log-Message "File exists: NO - extraction failed!"
                }
                
                throw "CRITICAL: NTUSER.DAT cannot be loaded. Import failed. Check the log for diagnostics."
            }
        }
        else {
            Log-Message "MERGE MODE: Skipping hive verification (using existing user's NTUSER.DAT)"
        }

        # IMPORTANT: Do NOT modify any registry settings here!
        # The extracted hive already contains all settings exactly as they were on source.
        # Any modifications to shell folders, navigation pane, search, or theme
        # will break UX fidelity. The only transformations applied were:
        # 1. Binary SID replacement (handled in Rewrite-HiveSID)
        # 2. Path string replacement for embedded paths (handled in Rewrite-HiveSID)
        Log-Message "Hive extraction complete - all settings preserved from source"

        # --------------------------------------------------------------
        # FINAL SUCCESS
        # --------------------------------------------------------------
        $global:ProgressBar.Value = 100
        $global:StatusText.Text = "[OK] Migration completed successfully!"
        $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)

        # AppX re-registration handled later in this function via Active Setup

        Log-Message "=================================================="
        Log-Message "PROFILE MIGRATION COMPLETED SUCCESSFULLY"
        Log-Message "=================================================="

        # Your original cleanup (if you still want it)
        $global:StatusText.Text = "Cleaning up temporary files..."
			
        $global:StatusText.Text = "Cleaning up temporary files..."
        [System.Windows.Forms.Application]::DoEvents()

        # === RECREATE STANDARD JUNCTIONS ===
        # These were excluded during export and need to be recreated pointing to new paths
        Log-Message "Recreating standard profile junctions..."
        $junctionsToCreate = @(
            @{ Source = "$target\Documents\My Music"; Target = "$target\Music"; Name = "My Music" }
            @{ Source = "$target\Documents\My Pictures"; Target = "$target\Pictures"; Name = "My Pictures" }
            @{ Source = "$target\Documents\My Videos"; Target = "$target\Videos"; Name = "My Videos" }
        )
        
        foreach ($junc in $junctionsToCreate) {
            try {
                # Check if target directory exists
                if (-not (Test-Path $junc.Target)) {
                    Log-Message "Junction target does not exist (OK if user customized profile): $($junc.Target)"
                    continue
                }
                
                # Remove if junction already exists (shouldn't, but be safe)
                if (Test-Path $junc.Source) {
                    $item = Get-Item $junc.Source -Force
                    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                        # For junctions, use rmdir /s /q to avoid prompts - removes only the link, not target
                        cmd /c rmdir /s /q "$($junc.Source)" >$null 2>&1
                        if ($LASTEXITCODE -eq 0) {
                            Log-Message "Removed stale junction: $($junc.Source)"
                        }
                        else {
                            Log-Message "WARNING: Failed to remove junction at $($junc.Source) (will attempt to recreate)"
                        }
                    }
                }
                
                # Create junction
                cmd /c mklink /j "$($junc.Source)" "$($junc.Target)" | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Log-Message "Created junction: $($junc.Name) -> $($junc.Target)"
                }
                else {
                    Log-Message "WARNING: Failed to create junction $($junc.Name)"
                }
            }
            catch {
                Log-Message "WARNING: Junction creation error for $($junc.Name): $_"
            }
        }

        # Remove Winget JSON artifacts from target profile
        try {
            $wingetJsons = @()
            $wingetJsons += (Join-Path $target 'Winget-Packages.json')
            $wingetJsons += (Join-Path $target 'Winget-Packages-Selected.json')
            foreach ($wj in $wingetJsons) {
                if (Test-Path $wj) {
                    Remove-Item $wj -Force -ErrorAction SilentlyContinue
                    Log-Message "Removed Winget file: $wj"
                }
            }
        }
        catch {
            Log-Message "WARNING: Failed to remove Winget JSON files: $_"
        }

        # Remove manifest.json from profile folder
        try {
            $manifestPath = Join-Path $target "manifest.json"
            if (Test-Path $manifestPath) {
                Remove-Item $manifestPath -Force -ErrorAction Stop
                Log-Message "Removed manifest.json from profile folder"
            }
        }
        catch {
            Log-Message "WARNING: Failed to remove manifest.json: $_"
        }

        # Clean up old backup profile folders
        try {
            Log-Message "Checking for old profile backup folders to clean up..."
            $backupPattern = "$target.backup_*"
            $backupFolders = Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -match "^$([regex]::Escape($shortName))\.backup_\d{8}_\d{6}$" } |
            Sort-Object Name -Descending
            
            if ($backupFolders.Count -gt 0) {
                Log-Message "Found $($backupFolders.Count) backup folder(s) for user '$shortName'"
                
                # Keep the 2 most recent backups, delete older ones
                $keepCount = 2
                $foldersToDelete = $backupFolders | Select-Object -Skip $keepCount
                
                if ($foldersToDelete.Count -gt 0) {
                    Log-Message "Keeping $keepCount most recent backup(s), removing $($foldersToDelete.Count) older backup(s)..."
                    foreach ($oldBackup in $foldersToDelete) {
                        try {
                            $backupAge = [DateTime]::Now - $oldBackup.CreationTime
                            Log-Message "Removing old backup: $($oldBackup.Name) (created $([int]$backupAge.TotalDays) days ago)"
                            Remove-FolderRobust -Path $oldBackup.FullName
                            Log-Message "Successfully removed: $($oldBackup.Name)"
                        }
                        catch {
                            Log-Message "WARNING: Could not remove backup folder $($oldBackup.Name): $_"
                        }
                    }
                    Log-Message "Old backup cleanup complete"
                }
                else {
                    Log-Message "All backups are recent - nothing to clean up"
                }
            }
            else {
                Log-Message "No backup folders found for cleanup"
            }
        }
        catch {
            Log-Message "WARNING: Error during backup folder cleanup: $_"
        }

        Log-Message "Cleanup complete"
        
        $global:ProgressBar.Value = 100
        Log-Message "=========================================="
        Log-Message "IMPORT SUCCESSFUL for user '$Username'"
        Log-Message "Profile SID: $sid"
        Log-Message "Profile Path: $target"
        Log-Message "User Type: $(if ($isDomain) { 'Domain' } else { 'Local' })"
        if (-not $isDomain) { Log-Message "GPSVC Support: ENABLED (Users group permissions applied)" }
        
        # (Removed) ACL diagnostics output
        
        # Log elapsed time for diagnostics
        if ($global:ImportStartTime) {
            $elapsed = [DateTime]::Now - $global:ImportStartTime
            Log-Message "Operation elapsed time: $($elapsed.TotalMinutes.ToString('F2')) minutes ($($elapsed.TotalSeconds.ToString('F0')) seconds)"
        }
        
        Log-Message "=========================================="
        
        # Register AppX fix for next login (Active Setup)
        # This will run before the desktop loads to prevent file lock errors
        Add-AppxReregistrationActiveSetup -Username $importTargetUser -OperationType "Import"

        # Generate HTML migration report
        try {
            # Collect ZIP file statistics for report
            $zipSizeMB = 'N/A'
            $zipFileCount = 'N/A'
            $zipFolderCount = 'N/A'
            $uncompressedSizeMB = 'N/A'
            $compressionRatio = 'N/A'
            
            try {
                if (Test-Path $ZipPath) {
                    $zipFile = Get-Item $ZipPath -ErrorAction Stop
                    $zipSizeMB = [math]::Round($zipFile.Length / 1MB, 2)
                    
                    # Get basic ZIP statistics using 7-Zip
                    Log-Message "Collecting ZIP archive statistics for report..."
                    
                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName = $global:SevenZipPath
                    $psi.Arguments = "l -slt `"$ZipPath`""
                    $psi.CreateNoWindow = $true
                    $psi.UseShellExecute = $false
                    $psi.RedirectStandardOutput = $true
                    
                    $listProc = [System.Diagnostics.Process]::Start($psi)
                    $output = $listProc.StandardOutput.ReadToEnd()
                    $listProc.WaitForExit()
                    
                    if ($listProc.ExitCode -eq 0) {
                        Log-Message "7-Zip technical listing received, parsing..."
                        
                        # Count files and folders, sum sizes from technical output
                        $files = 0
                        $folders = 0
                        $totalSize = 0
                        
                        $lines = $output -split "`n"
                        $isFolder = $false
                        $currentSize = 0
                        
                        foreach ($line in $lines) {
                            $line = $line.Trim()
                            
                            # Check if this entry is a folder
                            if ($line -match '^Attributes = .+D') {
                                $isFolder = $true
                            }
                            
                            # Get the size
                            if ($line -match '^Size = (\d+)$') {
                                $currentSize = [int64]$matches[1]
                            }
                            
                            # When we hit a Path entry, we've finished reading an item
                            if ($line -match '^Path = (.+)$' -and $line -notmatch '^Path\s*=\s*$') {
                                if ($isFolder) {
                                    $folders++
                                }
                                else {
                                    $files++
                                    $totalSize += $currentSize
                                }
                                # Reset for next item
                                $isFolder = $false
                                $currentSize = 0
                            }
                        }
                        
                        if ($files -gt 0) {
                            $zipFileCount = $files
                            $zipFolderCount = $folders
                            $uncompressedSizeMB = [math]::Round($totalSize / 1MB, 2)
                            
                            # Calculate compression ratio
                            if ($totalSize -gt 0) {
                                $compressionRatio = "$([math]::Round(($zipFile.Length / $totalSize) * 100, 1))%"
                            }
                            
                            Log-Message "ZIP statistics: $zipFileCount files, $zipFolderCount folders, $uncompressedSizeMB MB uncompressed, compression: $compressionRatio"
                        }
                        else {
                            Log-Message "WARNING: No files found in 7-Zip output"
                        }
                    }
                    else {
                        Log-Message "7-Zip list command failed with exit code: $($listProc.ExitCode)"
                    }
                }
            }
            catch {
                Log-Message "Could not collect ZIP statistics: $_"
            }
            
            # Get installed apps from global variable if available
            $installedAppsList = @()
            if ($global:InstalledAppsList) {
                $installedAppsList = $global:InstalledAppsList
                Log-Message "Including $($installedAppsList.Count) installed apps in report"
            }
            
            $reportData = @{
                Username                = $importTargetUser
                UserType                = if ($isDomain) { 'Domain' } else { 'Local' }
                TargetSID               = $sid
                SourceSID               = if ($sourceSID) { $sourceSID } else { 'N/A' }
                ProfilePath             = $target
                ZipPath                 = $ZipPath
                ZipSizeMB               = $zipSizeMB
                ImportMode              = if ($mergeMode) { 'Merge' } else { 'Replace' }
                ElapsedMinutes          = if ($global:ImportStartTime) { ([DateTime]::Now - $global:ImportStartTime).TotalMinutes.ToString('F2') } else { 'N/A' }
                ElapsedSeconds          = if ($global:ImportStartTime) { ([DateTime]::Now - $global:ImportStartTime).TotalSeconds.ToString('F0') } else { 'N/A' }
                Timestamp               = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                BackupPath              = $global:ImportBackupPath
                HashVerified            = $Config.HashVerificationEnabled
                HashVerificationStatus  = $hashVerificationStatus
                HashVerificationDetails = $hashVerificationDetails
                FileCount               = $zipFileCount
                FolderCount             = $zipFolderCount
                UncompressedSizeMB      = $uncompressedSizeMB
                CompressionRatio        = $compressionRatio
                Warnings                = @()
                InstalledApps           = $installedAppsList
            }
            
            $reportPath = Generate-MigrationReport -OperationType 'Import' -ReportData $reportData
            if ($reportPath) {
                Log-Info "Migration report available: $reportPath"
            }
        }
        catch {
            Log-Warning "Could not generate migration report: $_"
        }
        
        $userType = if ($isDomain) { 'Domain' } else { 'Local' }
        $successMsg = @"
IMPORT SUCCESSFUL

User: $importTargetUser
Type: $userType
SID: $sid
Path: $target

IMPORTANT: REBOOT NOW, THEN LOG IN

The profile has been successfully imported.
You must restart the computer before logging in with this user.
"@
        Show-ModernDialog -Message $successMsg -Title "SUCCESS" -Type Success -Buttons OK

    }
    catch {
        Log-Message "=========================================="
        Log-Message "IMPORT FAILED: $_"
        Log-Message "=========================================="
        
        # Cleanup temp merge location if it exists (Merge Mode)
        if ($tempMergeLocation -and (Test-Path $tempMergeLocation)) {
            Log-Message "Cleaning up temp merge location after failure: $tempMergeLocation"
            Remove-FolderRobust -Path $tempMergeLocation -ErrorAction SilentlyContinue
        }
        
        
        # ROLLBACK: Restore from backup if it exists and operation had started
        $rollbackSucceeded = $false
        if ($global:ImportBackupPath -and (Test-Path $global:ImportBackupPath)) {
            Log-Message "Rollback: restoring from backup"
            try {
                if (Test-Path $target) {
                    Log-Message "ROLLBACK: Removing failed import at $target"
                    Remove-FolderRobust -Path $target -ErrorAction SilentlyContinue
                }
                Log-Message "Rollback: source $($global:ImportBackupPath)"
                Log-Message "Rollback: source $($global:ImportBackupPath)"
                
                # Hybrid Rollback: Try Rename first (Fast), then Robocopy (Robust)
                $rollbackDone = $false
                try {
                    Log-Message "Attempting fast rollback (Rename)..."
                    Rename-Item -Path $global:ImportBackupPath -NewName (Split-Path $target -Leaf) -ErrorAction Stop
                    $rollbackDone = $true
                    Log-Message "Fast rollback successful"
                }
                catch {
                    Log-Warning "Fast rollback failed, falling back to robust copy: $_"
                }
                
                if (-not $rollbackDone) {
                    # Robocopy Fallback - handles junctions/permissions better than Copy-Item
                    Log-Message "Attempting robust rollback (Robocopy)..."
                    $roboArgs = @(
                        "`"$global:ImportBackupPath`"", "`"$target`"", "/E", "/COPY:DAT", "/DCOPY:DAT", 
                        "/XJ", # Exclude junctions to prevent recursion loops
                        "/R:2", "/W:1", "/MT:8", "/NP", "/NFL", "/NDL"
                    )
                    $p = Start-Process -FilePath "robocopy.exe" -ArgumentList $roboArgs -Wait -NoNewWindow -PassThru
                    if ($p.ExitCode -gt 7) {
                        throw "Robocopy rollback failed with exit code $($p.ExitCode)"
                    }
                    Log-Message "Robust rollback completed"
                }
                Log-Message "ROLLBACK: Restoration completed successfully"
                
                # CRITICAL: Ensure registry entry exists and points to correct path
                # BUT ONLY IN REPLACE MODE - in merge mode, registry should be untouched
                if (-not $mergeMode) {
                    if ($sid) {
                        $base = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
                        $profileKey = "$base\$sid"
                        if (Test-Path $profileKey) {
                            # Verify ProfileImagePath is correct
                            $currentPath = (Get-ItemProperty -Path $profileKey -Name ProfileImagePath -ErrorAction SilentlyContinue).ProfileImagePath
                            if ($currentPath -ne $target) {
                                Log-Message "ROLLBACK: Correcting ProfileImagePath from '$currentPath' to '$target'"
                                Set-ItemProperty -Path $profileKey -Name ProfileImagePath -Value $target -Type ExpandString -Force
                            }
                            # Clear any error state flags
                            Set-ItemProperty -Path $profileKey -Name State -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                            Log-Message "ROLLBACK: Registry entry verified and corrected"
                        }
                        else {
                            Log-Message "WARNING: Registry entry missing after rollback - user may experience temporary profile"
                        }
                    }
                }
                else {
                    Log-Message "ROLLBACK (MERGE MODE): Skipping registry operations - target user registry unchanged"
                }
                
                $rollbackSucceeded = $true
                # Remove the backup folder after successful restoration to avoid leaving stale backups
                try {
                    if (Test-Path $global:ImportBackupPath) {
                        Log-Message "Rollback: removing $($global:ImportBackupPath)"
                        Remove-FolderRobust -Path $global:ImportBackupPath -ErrorAction SilentlyContinue
                        Log-Message "Rollback: backup folder removed"
                        $global:ImportBackupPath = $null
                    }
                }
                catch {
                    Log-Message "WARNING: Could not remove backup folder: $_"
                }
            }
            catch {
                Log-Message "WARNING: Rollback failed: $_. Manual recovery may be needed."
                Log-Message "Backup available at: $($global:ImportBackupPath)"
            }
        }
        else {
            Log-Message "No backup available for rollback"
        }
        
        # Log elapsed time for diagnostics
        if ($global:ImportStartTime) {
            $elapsed = [DateTime]::Now - $global:ImportStartTime
            Log-Message "Operation elapsed time: $($elapsed.TotalMinutes.ToString('F2')) minutes ($($elapsed.TotalSeconds.ToString('F0')) seconds)"
        }
        
        # Generate failure report
        $errorMessage = $_.Exception.Message
        $rollbackStatus = if ($rollbackSucceeded) {
            "Rollback completed successfully - profile restored from backup"
        }
        elseif ($global:ImportBackupPath) {
            "Rollback attempted but may have failed - backup at: $($global:ImportBackupPath)"
        }
        else {
            "No rollback performed - no backup available"
        }
        
        try {
            $elapsedMins = if ($global:ImportStartTime) { ([DateTime]::Now - $global:ImportStartTime).TotalMinutes.ToString('F2') } else { 'N/A' }
            $elapsedSecs = if ($global:ImportStartTime) { ([DateTime]::Now - $global:ImportStartTime).TotalSeconds.ToString('F0') } else { 'N/A' }
            
            $reportData = @{
                Username       = $Username
                ZipPath        = $ZipPath
                TargetPath     = if ($target) { $target } else { 'N/A' }
                ElapsedMinutes = $elapsedMins
                ElapsedSeconds = $elapsedSecs
                Timestamp      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                MergeMode      = if ($mergeMode) { 'Yes' } else { 'No' }
            }
            
            $reportPath = Generate-MigrationReport -OperationType 'Import' -ReportData $reportData -Success $false -ErrorMessage $errorMessage -RollbackStatus $rollbackStatus
            if ($reportPath) {
                Log-Message "Failure report generated: $reportPath"
                # Report auto-opens via AutoOpenReports config in Generate-MigrationReport
            }
        }
        catch {
            Log-Message "Could not generate failure report: $_"
        }
        
        Show-ModernDialog -Message "Import failed: $_" -Title "Error" -Type Error -Buttons OK
    }
    finally {
        Stop-OperationLog
        # Note: Temp folder cleanup is handled in try/catch blocks where needed
        $global:StatusText.Text = "Idle"
        $global:ProgressBar.Value = 0
        $global:CurrentOperation = $null
        if ($global:CancelButton) { $global:CancelButton.Enabled = $false }
    }
}

# =============================================================================
# GUI - MODERN FLAT DESIGN
# =============================================================================
$global:Form = New-Object System.Windows.Forms.Form
$global:Form.Text = "Profile Migration Tool $($Config.Version)"
$global:Form.Size = New-Object System.Drawing.Size(920, 720)
$global:Form.StartPosition = "CenterScreen"
$global:Form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$global:Form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$global:Form.MinimumSize = New-Object System.Drawing.Size(920, 720)

# Header section with modern styling
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(920, 80)
$headerPanel.BackColor = [System.Drawing.Color]::White
$global:Form.Controls.Add($headerPanel)

$lblHeader = New-Object System.Windows.Forms.Label
$lblHeader.Location = New-Object System.Drawing.Point(20, 15)
$lblHeader.Size = New-Object System.Drawing.Size(500, 35)
$lblHeader.Text = "Windows Profile Migration"
$lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblHeader.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$headerPanel.Controls.Add($lblHeader)

$lblSubheader = New-Object System.Windows.Forms.Label
$lblSubheader.Location = New-Object System.Drawing.Point(22, 48)
$lblSubheader.Size = New-Object System.Drawing.Size(500, 20)
$lblSubheader.Text = "Export, import, and merge user profiles with ease"
$lblSubheader.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$headerPanel.Controls.Add($lblSubheader)

# Settings button in header
$btnSettings = New-Object System.Windows.Forms.Button
$btnSettings.Location = New-Object System.Drawing.Point(850, 25)
$btnSettings.Size = New-Object System.Drawing.Size(40, 40)
$btnSettings.Text = "..."
$btnSettings.BackColor = [System.Drawing.Color]::White
$btnSettings.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnSettings.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSettings.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$btnSettings.FlatAppearance.BorderSize = 1
$btnSettings.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$btnSettings.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnSettings.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    })
$btnSettings.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::White
    })
$btnSettings.Add_Click({
        # Settings dialog
        $settingsForm = New-Object System.Windows.Forms.Form
        $settingsForm.Text = "Settings"
        $settingsForm.Size = New-Object System.Drawing.Size(550, 200)
        $settingsForm.StartPosition = "CenterParent"
        $settingsForm.FormBorderStyle = "FixedDialog"
        $settingsForm.MaximizeBox = $false
        $settingsForm.MinimizeBox = $false
        
        # Apply theme
        $currentTheme = if ($global:CurrentTheme) { $global:CurrentTheme } else { "Light" }
        $theme = $Themes[$currentTheme]
        $settingsForm.BackColor = $theme.FormBackColor
        $settingsForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
        # Content panel
        $panel = New-Object System.Windows.Forms.Panel
        $panel.Location = New-Object System.Drawing.Point(15, 15)
        $panel.Size = New-Object System.Drawing.Size(505, 90)
        $panel.BackColor = $theme.PanelBackColor
        $settingsForm.Controls.Add($panel)
    
        # 7-Zip path label
        $lbl7z = New-Object System.Windows.Forms.Label
        $lbl7z.Location = New-Object System.Drawing.Point(15, 15)
        $lbl7z.Size = New-Object System.Drawing.Size(100, 23)
        $lbl7z.Text = "7-Zip Path:"
        $lbl7z.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $lbl7z.ForeColor = $theme.LabelTextColor
        $panel.Controls.Add($lbl7z)
    
        # 7-Zip path textbox
        $txt7zPath = New-Object System.Windows.Forms.TextBox
        $txt7zPath.Location = New-Object System.Drawing.Point(15, 40)
        $txt7zPath.Size = New-Object System.Drawing.Size(390, 25)
        $txt7zPath.Text = $global:SevenZipPath
        $txt7zPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $txt7zPath.ReadOnly = $true
        $txt7zPath.BackColor = $theme.TextBoxBackColor
        $txt7zPath.ForeColor = $theme.TextBoxForeColor
        $panel.Controls.Add($txt7zPath)
    
        # Browse button for 7-Zip
        $btnBrowse7z = New-Object System.Windows.Forms.Button
        $btnBrowse7z.Location = New-Object System.Drawing.Point(415, 38)
        $btnBrowse7z.Size = New-Object System.Drawing.Size(75, 28)
        $btnBrowse7z.Text = "Browse..."
        $btnBrowse7z.BackColor = $theme.ButtonPrimaryBackColor
        $btnBrowse7z.ForeColor = $theme.ButtonPrimaryForeColor
        $btnBrowse7z.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnBrowse7z.FlatAppearance.BorderSize = 0
        $btnBrowse7z.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnBrowse7z.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
        $btnBrowse7z.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
        $btnBrowse7z.Add_Click({
                $openFile = New-Object System.Windows.Forms.OpenFileDialog
                $openFile.Filter = "7-Zip Executable (7z.exe)|7z.exe|All Files (*.*)|*.*"
                $openFile.Title = "Locate 7z.exe"
                if ($openFile.ShowDialog() -eq "OK") {
                    if (Test-Path $openFile.FileName) {
                        $txt7zPath.Text = $openFile.FileName
                    }
                }
            })
        $panel.Controls.Add($btnBrowse7z)
    
        # OK button
        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Location = New-Object System.Drawing.Point(310, 120)
        $btnOK.Size = New-Object System.Drawing.Size(100, 32)
        $btnOK.Text = "Save"
        $btnOK.BackColor = $theme.ButtonPrimaryBackColor
        $btnOK.ForeColor = $theme.ButtonPrimaryForeColor
        $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnOK.FlatAppearance.BorderSize = 0
        $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnOK.Add_MouseEnter({ $this.BackColor = $theme.ButtonPrimaryHoverBackColor })
        $btnOK.Add_MouseLeave({ $this.BackColor = $theme.ButtonPrimaryBackColor })
        $btnOK.Add_Click({
                $newPath = $txt7zPath.Text
                if (Test-Path $newPath) {
                    $global:SevenZipPath = $newPath
                    Show-ModernDialog -Message "7-Zip path updated successfully!`r`n`r`nNew path: $newPath" -Title "Settings Saved" -Type Success -Buttons OK
                    $settingsForm.Close()
                }
                else {
                    Show-ModernDialog -Message "Invalid path! Please select a valid 7z.exe file." -Title "Invalid Path" -Type Error -Buttons OK
                }
            })
        $settingsForm.Controls.Add($btnOK)
    
        # Cancel button
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(420, 120)
        $btnCancel.Size = New-Object System.Drawing.Size(100, 32)
        $btnCancel.Text = "Cancel"
        $btnCancel.BackColor = $theme.ButtonSecondaryBackColor
        $btnCancel.ForeColor = $theme.ButtonSecondaryForeColor
        $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnCancel.FlatAppearance.BorderSize = 0
        $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnCancel.Add_MouseEnter({ $this.BackColor = $theme.ButtonSecondaryHoverBackColor })
        $btnCancel.Add_MouseLeave({ $this.BackColor = $theme.ButtonSecondaryBackColor })
        $btnCancel.Add_Click({ $settingsForm.Close() })
        $settingsForm.Controls.Add($btnCancel)
    
        $settingsForm.ShowDialog() | Out-Null
    })
$headerPanel.Controls.Add($btnSettings)

# Theme toggle button in header
$global:ThemeToggleButton = New-Object System.Windows.Forms.Button
$global:ThemeToggleButton.Location = New-Object System.Drawing.Point(800, 25)
$global:ThemeToggleButton.Size = New-Object System.Drawing.Size(40, 40)
$global:ThemeToggleButton.Text = "L"  # Toggle theme (L=light, D=dark)
$global:ThemeToggleButton.BackColor = [System.Drawing.Color]::White
$global:ThemeToggleButton.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$global:ThemeToggleButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:ThemeToggleButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$global:ThemeToggleButton.FlatAppearance.BorderSize = 1
$global:ThemeToggleButton.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$global:ThemeToggleButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:ThemeToggleButton.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    })
$global:ThemeToggleButton.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::White
    })
$global:ThemeToggleButton.Add_Click({
        Toggle-Theme
    })
$headerPanel.Controls.Add($global:ThemeToggleButton)

# Convert Profile Type Button (in header, left of theme button)
$global:ConvertButton = New-Object System.Windows.Forms.Button
$global:ConvertButton.Location = New-Object System.Drawing.Point(650, 25)
$global:ConvertButton.Size = New-Object System.Drawing.Size(140, 40)
$global:ConvertButton.Text = "Convert"
$global:ConvertButton.BackColor = [System.Drawing.Color]::FromArgb(142, 68, 173)
$global:ConvertButton.ForeColor = [System.Drawing.Color]::White
$global:ConvertButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:ConvertButton.FlatAppearance.BorderSize = 0
$global:ConvertButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:ConvertButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:ConvertButton.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(125, 60, 152)
    })
$global:ConvertButton.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(142, 68, 173)
    })
$global:ConvertButton.Add_Click({
        try {
            Show-ProfileConversionDialog
        }
        catch {
            Log-Error "Error opening conversion dialog: $_"
            Show-ModernDialog -Message "Error opening conversion dialog:`r`n`r`n$_" -Title "Error" -Type Error -Buttons OK
        }
    })
$headerPanel.Controls.Add($global:ConvertButton)

# Tooltip for Convert button
$convertTooltip = New-Object System.Windows.Forms.ToolTip
$convertTooltip.SetToolTip($global:ConvertButton, "Convert profile between Local and Domain types")

# Tooltip for settings button
$settingsTooltip = New-Object System.Windows.Forms.ToolTip
$settingsTooltip.SetToolTip($btnSettings, "Settings - Configure 7-Zip path")

# Tooltip for theme toggle button
$themeTooltip = New-Object System.Windows.Forms.ToolTip
$themeTooltip.SetToolTip($global:ThemeToggleButton, "Toggle theme: L=Light mode, D=Dark mode")

# Main content card
$cardPanel = New-Object System.Windows.Forms.Panel
$cardPanel.Location = New-Object System.Drawing.Point(15, 95)
$cardPanel.Size = New-Object System.Drawing.Size(880, 100)
$cardPanel.BackColor = [System.Drawing.Color]::White
$global:Form.Controls.Add($cardPanel)

$lblU = New-Object System.Windows.Forms.Label
$lblU.Location = New-Object System.Drawing.Point(15, 15)
$lblU.Size = New-Object System.Drawing.Size(100, 23)
$lblU.Text = "Username:"
$lblU.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$cardPanel.Controls.Add($lblU)
$global:UserComboBox = New-Object System.Windows.Forms.ComboBox
$global:UserComboBox.Location = New-Object System.Drawing.Point(120, 12)
$global:UserComboBox.Size = New-Object System.Drawing.Size(280, 25)
$global:UserComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$global:UserComboBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
Get-ProfileDisplayEntries | ForEach-Object { $global:UserComboBox.Items.Add($_.DisplayName) | Out-Null }
$cardPanel.Controls.Add($global:UserComboBox)

# Proactive Check: Warn if selected user is logged in

            


# Set Target User button
$btnSetTargetUser = New-Object System.Windows.Forms.Button
$btnSetTargetUser.Location = New-Object System.Drawing.Point(15, 45)
$btnSetTargetUser.Size = New-Object System.Drawing.Size(120, 28)
$btnSetTargetUser.Text = "Set Target User"
$btnSetTargetUser.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnSetTargetUser.ForeColor = [System.Drawing.Color]::White
$btnSetTargetUser.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSetTargetUser.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSetTargetUser.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnSetTargetUser.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190) })
$btnSetTargetUser.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212) })
$cardPanel.Controls.Add($btnSetTargetUser)

# Status label for user type
$lblUserTypeStatus = New-Object System.Windows.Forms.Label
$lblUserTypeStatus.Location = New-Object System.Drawing.Point(145, 45)
$lblUserTypeStatus.Size = New-Object System.Drawing.Size(260, 28)
$lblUserTypeStatus.Text = ""
$lblUserTypeStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cardPanel.Controls.Add($lblUserTypeStatus)

$btnSetTargetUser.Add_Click({
        $selectedUser = $global:UserComboBox.Text.Trim()
        # Strip size info if present (e.g. "User - [50 GB]")
        if ($selectedUser -match '^(.+?)\s+-\s+\[.+\]$') { $selectedUser = $matches[1] }
        
        # --- PROACTIVE CHECK: USER LOGGED IN? ---
        if ($selectedUser) {
            # Normalize username for check
            $checkName = $selectedUser
            
            # Resolve SID for check
            $checkSid = $null
            try {
                $obj = New-Object System.Security.Principal.NTAccount($checkName)
                $checkSid = $obj.Translate([System.Security.Principal.SecurityIdentifier]).Value
            }
            catch {
                # Fallback
                try {
                    $sName = if ($checkName -match '\\') { ($checkName -split '\\', 2)[1] } else { $checkName }
                    $obj = New-Object System.Security.Principal.NTAccount($sName)
                    $checkSid = $obj.Translate([System.Security.Principal.SecurityIdentifier]).Value
                }
                catch {}
            }
            
            if ($checkSid) {
                if (Test-ProfileMounted -UserSID $checkSid) {
                    $res = Show-ModernDialog -Message "The user '$checkName' appears to be logged in (profile registry hive is mounted).`r`n`r`nImporting or Exporting with a logged-in user profile WILL result in locked files and data corruption.`r`n`r`nWould you like to force log off this user now and continue?" -Title "User Logged In Warning" -Type Warning -Buttons YesNo
                    
                    if ($res -eq "Yes") {
                        $logoffSuccess = Invoke-ForceUserLogoff -Username $checkName
                        if ($logoffSuccess) {
                            if (Test-ProfileMounted -UserSID $checkSid) {
                                Show-ModernDialog -Message "User was logged off but profile hive is still mounted. Please wait a moment and try again." -Title "Profile Still Mounted" -Type Warning -Buttons OK
                                return 
                            }
                            # Success - proceed
                        }
                        else {
                            Show-ModernDialog -Message "Failed to force log off the user. Please log them off manually." -Title "Logoff Failed" -Type Error -Buttons OK
                            return 
                        }
                    }
                    else {
                        # User said No
                        return
                    }
                }
            }
        }
        # ----------------------------------------
        
        $isDomain = $false
        $isAzureAD = $false
        $domain = ""
        
        # Check if username is in UPN format (user@domain.com)
        if ($selectedUser -match '@') {
            $isAzureAD = $true
            $lblUserTypeStatus.Text = "Target user is AzureAD/Entra ID (UPN format)"
            $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $global:DomainCheckBox.Checked = $false
            $global:DomainNameTextBox.Text = ""
            
            # Store the UPN globally for later use during import
            $global:TargetUserUPN = $selectedUser
            Log-Info "Set Target User: Detected UPN format '$selectedUser' - stored for import"
            
            # --- AzureAD/Entra ID checks ---
            if (-not (Test-IsAzureADJoined)) {
                $dlgResult = Show-AzureADJoinDialog -Username $selectedUser
                
                # Check if user cancelled
                if ($dlgResult -eq "Cancel") {
                    $lblUserTypeStatus.Text = "Device is NOT AzureAD joined!"
                    $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                    $global:SelectedTargetUser = $null
                    $global:TargetUserUPN = $null
                    Log-Info "Set Target User: AzureAD join cancelled by user. Aborted."
                    [System.Windows.Forms.Application]::DoEvents()
                    return
                }
                
                # User clicked Continue - re-check join status
                if (-not (Test-IsAzureADJoined)) {
                    $lblUserTypeStatus.Text = "Device is NOT AzureAD joined!"
                    $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                    $global:SelectedTargetUser = $null
                    $global:TargetUserUPN = $null
                    Log-Info "Set Target User: AzureAD join not completed. Aborted."
                    [System.Windows.Forms.Application]::DoEvents()
                    return
                }
                
                # Join verified - update status to show success
                Log-Info "AzureAD join verified after user completed join process"
                $lblUserTypeStatus.Text = "Target user is AzureAD/Entra ID (joined successfully)"
                $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            }
            
            # AzureAD user detected in UPN format - UPN already stored
            Log-Info "AzureAD user selected (UPN format) - UPN stored, will use Microsoft Graph for SID resolution during import"
        }
        # AzureAD\user
        elseif ($selectedUser -match '^AzureAD\\(.+)$') {
            $isAzureAD = $true
            $lblUserTypeStatus.Text = "Target user is AzureAD/Entra ID"
            $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $global:DomainCheckBox.Checked = $false
            $global:DomainNameTextBox.Text = ""
            # --- AzureAD/Entra ID checks ---
            if (-not (Test-IsAzureADJoined)) {
                $dlgResult = Show-AzureADJoinDialog -Username $selectedUser
                
                # Check if user cancelled
                if ($dlgResult -eq "Cancel") {
                    $lblUserTypeStatus.Text = "Device is NOT AzureAD joined!"
                    $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                    $global:SelectedTargetUser = $null
                    Log-Info "Set Target User: AzureAD join cancelled by user. Aborted."
                    [System.Windows.Forms.Application]::DoEvents()
                    return
                }
                
                # User clicked Continue - re-check join status
                if (-not (Test-IsAzureADJoined)) {
                    $lblUserTypeStatus.Text = "Device is NOT AzureAD joined!"
                    $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                    $global:SelectedTargetUser = $null
                    Log-Info "Set Target User: AzureAD join not completed. Aborted."
                    [System.Windows.Forms.Application]::DoEvents()
                    return
                }
                
                # Join verified - update status to show success
                Log-Info "AzureAD join verified after user completed join process"
                $lblUserTypeStatus.Text = "Target user is AzureAD/Entra ID (joined successfully)"
                $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            }
            # AzureAD user detected - no need to check if signed in
            # Microsoft Graph will resolve SID during import
            Log-Info "AzureAD user selected - will use Microsoft Graph for SID resolution during import"
        }
        # DOMAIN\user
        elseif ($selectedUser -match '^(?<domain>[^\\]+)\\(?<user>.+)$') {
            $parsedDomain = $matches['domain'].Trim().ToUpper()
            # Check if domain is actually the computer name (local user)
            if ($parsedDomain -ieq $env:COMPUTERNAME) {
                $lblUserTypeStatus.Text = "Target user is Local account"
                $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
                $global:DomainCheckBox.Checked = $false
                $global:DomainNameTextBox.Text = ""
            }
            else {
                # Check all profiles for AzureAD SID pattern
                try {
                    $allProfiles = Get-WmiObject Win32_UserProfile
                    $profiles = $allProfiles | Where-Object {
                        $_.LocalPath -like "*\$($matches['user'])"
                    }
                    $isAzureAD = $false
                    $allSids = @()
                    $azureSids = @()
                    foreach ($p in $profiles) {
                        $allSids += $p.SID
                        if (Test-IsAzureADSID $p.SID) {
                            $azureSids += $p.SID
                            $isAzureAD = $true
                        }
                    }
                    Log-Info ("Set Target User: Username={0}, Domain={1}, AllSIDs={2}, AzureADSIDs={3}, IsAzureAD={4}" -f $matches['user'], $parsedDomain, ($allSids -join ', '), ($azureSids -join ', '), $isAzureAD)
                    if ($isAzureAD) {
                        $lblUserTypeStatus.Text = "Target user is AzureAD/Entra ID (tenant: $parsedDomain)"
                        $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
                        $global:DomainCheckBox.Checked = $false
                        $global:DomainNameTextBox.Text = ""
                        # --- AzureAD/Entra ID checks ---
                        if (-not (Test-IsAzureADJoined)) {
                            $dlgResult = Show-AzureADJoinDialog -Username $selectedUser
                            
                            # Check if user cancelled
                            if ($dlgResult -eq "Cancel") {
                                $lblUserTypeStatus.Text = "Device is NOT AzureAD joined!"
                                $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                                $global:SelectedTargetUser = $null
                                Log-Info "Set Target User: AzureAD join cancelled by user. Aborted."
                                [System.Windows.Forms.Application]::DoEvents()
                                return
                            }
                            
                            # User clicked Continue - re-check join status
                            if (-not (Test-IsAzureADJoined)) {
                                $lblUserTypeStatus.Text = "Device is NOT AzureAD joined!"
                                $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                                $global:SelectedTargetUser = $null
                                Log-Info "Set Target User: AzureAD join not completed. Aborted."
                                [System.Windows.Forms.Application]::DoEvents()
                                return
                            }
                            
                            # Join verified - update status to show success
                            Log-Info "AzureAD join verified after user completed join process"
                            $lblUserTypeStatus.Text = "Target user is AzureAD/Entra ID (joined successfully)"
                            $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
                        }
                        # AzureAD user detected (DOMAIN\user format) - no need to check if signed in
                        # Microsoft Graph will resolve SID during import
                        Log-Info "AzureAD user selected (DOMAIN\user format) - will use Microsoft Graph for SID resolution during import"
                    }
                    else {
                        $isDomain = $true
                        $domain = $parsedDomain
                        $lblUserTypeStatus.Text = "Target user is Domain account ($domain)"
                        $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                        # Only check the Join Domain box if not already domain-joined
                        $partOfDomain = $false
                        try {
                            $cs = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
                            $partOfDomain = $cs.PartOfDomain
                        }
                        catch {}
                        if (-not $partOfDomain) {
                            $global:DomainCheckBox.Checked = $true
                        }
                        else {
                            $global:DomainCheckBox.Checked = $false
                        }
                        $global:RestartComboBox.SelectedItem = "Never"
                        # Use FQDN resolution for domain name (.local/.com/.net/etc)
                        $fqdn = $null
                        try {
                            $fqdn = Get-DomainFQDN -NetBIOSName $domain
                        }
                        catch {}
                        if ($fqdn) {
                            $global:DomainNameTextBox.Text = $fqdn
                        }
                        else {
                            $global:DomainNameTextBox.Text = $domain
                        }
                    }
                }
                catch {
                    $isDomain = $true
                    $domain = $parsedDomain
                    $lblUserTypeStatus.Text = "Target user is Domain account ($domain)"
                    $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                    # Only check the Join Domain box if not already domain-joined
                    $partOfDomain = $false
                    try {
                        $cs = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
                        $partOfDomain = $cs.PartOfDomain
                    }
                    catch {}
                    if (-not $partOfDomain) {
                        $global:DomainCheckBox.Checked = $true
                    }
                    else {
                        $global:DomainCheckBox.Checked = $false
                    }
                    $global:RestartComboBox.SelectedItem = "Never"
                    # Use FQDN resolution for domain name (.local/.com/.net/etc)
                    $fqdn = $null
                    try {
                        $fqdn = Get-DomainFQDN -NetBIOSName $domain
                    }
                    catch {}
                    if ($fqdn) {
                        $global:DomainNameTextBox.Text = $fqdn
                    }
                    else {
                        $global:DomainNameTextBox.Text = $domain
                    }
                }
            }
        }
        else {
            $lblUserTypeStatus.Text = "Target user is Local account"
            $lblUserTypeStatus.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
            $global:DomainCheckBox.Checked = $false
            $global:DomainNameTextBox.Text = ""
        }
        $global:SelectedTargetUser = $selectedUser
        Log-Info "Set Target User: $selectedUser (stored in global:SelectedTargetUser)"
        [System.Windows.Forms.Application]::DoEvents()
    })

# Refresh Profiles button
$btnRefreshProfiles = New-Object System.Windows.Forms.Button
$btnRefreshProfiles.Location = New-Object System.Drawing.Point(405, 10)
$btnRefreshProfiles.Size = New-Object System.Drawing.Size(30, 28)
$btnRefreshProfiles.Text = "Refresh"
$btnRefreshProfiles.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnRefreshProfiles.ForeColor = [System.Drawing.Color]::White
$btnRefreshProfiles.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRefreshProfiles.FlatAppearance.BorderSize = 0
$btnRefreshProfiles.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$btnRefreshProfiles.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnRefreshProfiles.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190)
    })
$btnRefreshProfiles.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    })
$btnRefreshProfiles.Add_Click({
        # Refresh profile list with progress indication
        $btnRefreshProfiles.Enabled = $false
        $originalText = $btnRefreshProfiles.Text
        $btnRefreshProfiles.Text = "..."
    
        # Show scanning status
        $originalStatusText = $global:StatusText.Text
        $originalStatusColor = $global:StatusText.ForeColor
        $global:StatusText.Text = "Scanning profiles and calculating sizes..."
        $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        [System.Windows.Forms.Application]::DoEvents()
    
        try {
            $currentSelection = $global:UserComboBox.Text
            $global:UserComboBox.Items.Clear()
        
            # Calculate sizes during refresh (this is when user explicitly requests it)
            $allProfiles = Get-LocalProfiles
            $profileCount = $allProfiles.Count
            $currentIndex = 0
        
            $profiles = Get-ProfileDisplayEntries -CalculateSizes
        
            foreach ($p in $profiles) {
                $global:UserComboBox.Items.Add($p.DisplayName) | Out-Null
            }
        
            # Restore selection if it still exists
            if ($currentSelection) {
                $matchingItem = $global:UserComboBox.Items | Where-Object { $_ -like "$currentSelection*" } | Select-Object -First 1
                if ($matchingItem) {
                    $global:UserComboBox.Text = $matchingItem
                }
            }
        
            # Success message with size info
            $global:StatusText.Text = "Profile scan complete - $($profiles.Count) profiles with size estimates"
            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
            [System.Windows.Forms.Application]::DoEvents()
        
            # Brief pause to show success message
            Start-Sleep -Milliseconds 800
        
        }
        catch {
            $global:StatusText.Text = "Error scanning profiles: $_"
            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
            Show-ModernDialog -Message "Error refreshing profiles: $_" -Title "Error" -Type Error -Buttons OK
        }
        finally {
            # Restore original status
            $global:StatusText.Text = $originalStatusText
            $global:StatusText.ForeColor = $originalStatusColor
            $btnRefreshProfiles.Text = $originalText
            $btnRefreshProfiles.Enabled = $true
        }
    })
$cardPanel.Controls.Add($btnRefreshProfiles)

# Tooltip for refresh button
$refreshTooltip = New-Object System.Windows.Forms.ToolTip
$refreshTooltip.SetToolTip($btnRefreshProfiles, "Refresh profile list and scan sizes")

# Modern flat buttons with icons and hover effects
$global:BrowseButton = New-Object System.Windows.Forms.Button
$global:BrowseButton.Location = New-Object System.Drawing.Point(465, 8)
$global:BrowseButton.Size = New-Object System.Drawing.Size(120, 32)
$global:BrowseButton.Text = "[...] Browse"
$global:BrowseButton.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$global:BrowseButton.ForeColor = [System.Drawing.Color]::White
$global:BrowseButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:BrowseButton.FlatAppearance.BorderSize = 0
$global:BrowseButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$global:BrowseButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:BrowseButton.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    })
$global:BrowseButton.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    })
$cardPanel.Controls.Add($global:BrowseButton)

# View Log Button (below Browse button)
$global:ViewLogButton = New-Object System.Windows.Forms.Button
$global:ViewLogButton.Location = New-Object System.Drawing.Point(465, 48)
$global:ViewLogButton.Size = New-Object System.Drawing.Size(35, 32)
$global:ViewLogButton.Text = "..."
$global:ViewLogButton.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$global:ViewLogButton.ForeColor = [System.Drawing.Color]::White
$global:ViewLogButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:ViewLogButton.FlatAppearance.BorderSize = 0
$global:ViewLogButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:ViewLogButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:ViewLogButton.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    })
$global:ViewLogButton.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    })
$global:ViewLogButton.Add_Click({
        if ($global:CurrentLogFile -and (Test-Path $global:CurrentLogFile)) {
            Show-LogViewer -LogPath $global:CurrentLogFile -Title "Current Operation Log"
        }
        else {
            # Show most recent log file
            $logsDir = Join-Path $PSScriptRoot "Logs"
            if (Test-Path $logsDir) {
                $latestLog = Get-ChildItem -Path $logsDir -Filter "*.log" -File -ErrorAction SilentlyContinue | 
                Sort-Object LastWriteTime -Descending | 
                Select-Object -First 1
                if ($latestLog) {
                    Show-LogViewer -LogPath $latestLog.FullName -Title "Latest Log: $($latestLog.Name)"
                }
                else {
                    Show-ModernDialog -Message "No log files found.`r`n`r`nLog files will be created in the Logs folder when you run Export or Import operations." -Title "No Logs" -Type Info -Buttons OK
                }
            }
            else {
                Show-ModernDialog -Message "No log files found.`r`n`r`nLog files will be created in the Logs folder when you run Export or Import operations." -Title "No Logs" -Type Info -Buttons OK
            }
        }
    })
$cardPanel.Controls.Add($global:ViewLogButton)

# Tooltip for View Log button
$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.SetToolTip($global:ViewLogButton, "View Full Log")

# Log Level Selector (next to View Log button)
$lblLogLevel = New-Object System.Windows.Forms.Label
$lblLogLevel.Location = New-Object System.Drawing.Point(510, 51)
$lblLogLevel.Size = New-Object System.Drawing.Size(70, 23)
$lblLogLevel.Text = "Log Level:"
$lblLogLevel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblLogLevel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$cardPanel.Controls.Add($lblLogLevel)

$global:LogLevelComboBox = New-Object System.Windows.Forms.ComboBox
$global:LogLevelComboBox.Location = New-Object System.Drawing.Point(510, 70)
$global:LogLevelComboBox.Size = New-Object System.Drawing.Size(75, 25)
$global:LogLevelComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$global:LogLevelComboBox.Font = New-Object System.Drawing.Font("Segoe UI", 8)
@('DEBUG', 'INFO', 'WARN', 'ERROR') | ForEach-Object { $global:LogLevelComboBox.Items.Add($_) | Out-Null }
$global:LogLevelComboBox.SelectedItem = $Config.LogLevel
$global:LogLevelComboBox.Add_SelectedIndexChanged({
        $oldLevel = $Config.LogLevel
        $Config.LogLevel = $global:LogLevelComboBox.SelectedItem
        # Refresh display to show filtered logs
        Refresh-LogDisplay
        # Log the change (will appear if new level allows INFO)
        Log-Info "Log level changed: $oldLevel -> $($Config.LogLevel)"
    })
$cardPanel.Controls.Add($global:LogLevelComboBox)

# Debug Mode Checkbox (for verbose 7-Zip logging)
$global:DebugCheckBox = New-Object System.Windows.Forms.CheckBox
$global:DebugCheckBox.Location = New-Object System.Drawing.Point(595, 48)
$global:DebugCheckBox.Size = New-Object System.Drawing.Size(120, 32)
$global:DebugCheckBox.Text = "Debug Mode"
$global:DebugCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$global:DebugCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$global:DebugCheckBox.Checked = $false
$cardPanel.Controls.Add($global:DebugCheckBox)

# Tooltip for Debug Mode checkbox
$debugTooltip = New-Object System.Windows.Forms.ToolTip
$debugTooltip.SetToolTip($global:DebugCheckBox, "Enables detailed 7-Zip logging and preserves temporary files on failure")


$global:ExportButton = New-Object System.Windows.Forms.Button
$global:ExportButton.Location = New-Object System.Drawing.Point(595, 8)
$global:ExportButton.Size = New-Object System.Drawing.Size(120, 32)
$global:ExportButton.Text = "[>>] Export"
$global:ExportButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$global:ExportButton.ForeColor = [System.Drawing.Color]::White
$global:ExportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:ExportButton.FlatAppearance.BorderSize = 0
$global:ExportButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:ExportButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:ExportButton.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190)
    })
$global:ExportButton.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    })
$cardPanel.Controls.Add($global:ExportButton)
$global:ExportButton.Add_Click({
        try {
            # Always uncheck Join Domain After Import when exporting
            if ($global:DomainCheckBox) { $global:DomainCheckBox.Checked = $false }
            # Validate user selection
            $username = $global:SelectedTargetUser
            if (-not $username -or $username -eq "") {
                # Show modern dialog for No User Selected (Export)
                Show-ModernDialog -Message "Please use the 'Set Target User' button to select and confirm the user profile to export." -Title "No User Selected" -Type Warning -Buttons OK
                return
            }
            # Validate ZIP path
            $zipPath = $global:SelectedZipPath
            # Restore previous export filename format: username-Export-yyyyMMdd_HHmmss.zip
            $defaultFileName = ""
            if ($zipPath -and $zipPath -ne "") {
                $defaultFileName = [System.IO.Path]::GetFileName($zipPath)
            }
            else {
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                # Remove domain prefix if present
                $shortName = $username
                if ($username -match "^.+\\(.+)$") { $shortName = $matches[1] }
                # Sanitize filename (replace invalid chars and dots with underscores)
                $shortName = $shortName -replace '[^a-zA-Z0-9_\-]', '_'
                $defaultFileName = "$shortName-Export-$timestamp.zip"
            }
            if (-not $zipPath -or $zipPath -eq "") {
                # Prompt user for save location
                $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
                $saveDialog.Title = "Select Export Destination"
                $saveDialog.Filter = "Profile Archive (*.zip)|*.zip|All Files (*.*)|*.*"
                $saveDialog.FileName = $defaultFileName
                $saveDialog.InitialDirectory = $null
                if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $zipPath = $saveDialog.FileName
                    $global:SelectedZipPath = $zipPath
                }
                else {
                    return
                }
            }
            # Confirm export (ensure full path is visible)
            $msg = "You are about to export the profile for user:" + [Environment]::NewLine +
            "    $username" + [Environment]::NewLine + [Environment]::NewLine +
            "To archive (full path):" + [Environment]::NewLine +
            "    " + $zipPath + [Environment]::NewLine + [Environment]::NewLine +
            "Continue with export?"
            # Use Show-ModernDialog, and if possible, set a wider minimum width for the dialog
            $result = Show-ModernDialog -Message $msg -Title "Confirm Export" -Type Question -Buttons YesNo
            if ($result -ne "Yes") { return }
            # Start export
            Export-UserProfile -Username $username -ZipPath $zipPath
        }
        catch {
            Log-Error "Export failed: $_"
            Show-ModernDialog -Message "Export failed: $_" -Title "Export Error" -Type Error -Buttons OK
        }
    })

$global:ImportButton = New-Object System.Windows.Forms.Button
$global:ImportButton.Location = New-Object System.Drawing.Point(725, 8)
$global:ImportButton.Size = New-Object System.Drawing.Size(120, 32)
$global:ImportButton.Text = "[<<] Import"
$global:ImportButton.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$global:ImportButton.ForeColor = [System.Drawing.Color]::White
$global:ImportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:ImportButton.FlatAppearance.BorderSize = 0
$global:ImportButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:ImportButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:ImportButton.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(12, 100, 12)
    })
$global:ImportButton.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
    })
$cardPanel.Controls.Add($global:ImportButton)

# Add Click handler for ImportButton
$global:ImportButton.Add_Click({
        try {
            # Validate ZIP selection
            $zipPath = $global:SelectedZipPath
            if (-not $zipPath -or -not (Test-Path $zipPath)) {
                Show-ModernDialog -Message "Please select a valid profile archive (ZIP) before importing." -Title "No ZIP Selected" -Type Warning -Buttons OK
                return
            }
            # Use the target user set by the Set Target User button
            $username = $global:SelectedTargetUser
            if (-not $username -or $username -eq "") {
                Show-ModernDialog -Message "Please use the 'Set Target User' button to select and confirm the target user before importing." -Title "No User Selected" -Type Warning -Buttons OK
                return
            }
            # Confirm import
            $msg = "You are about to import the profile for user:" + [Environment]::NewLine +
            "    $username" + [Environment]::NewLine + [Environment]::NewLine +
            "From archive:" + [Environment]::NewLine +
            "    $zipPath" + [Environment]::NewLine + [Environment]::NewLine +
            "Continue with import?"
            $result = Show-ModernDialog -Message $msg -Title "Confirm Import" -Type Question -Buttons YesNo
            if ($result -ne "Yes") { return }
            # Start import
            Import-UserProfile -ZipPath $zipPath -Username $username
        }
        catch {
            Log-Error "Import failed: $_"
            Show-ModernDialog -Message "Import failed: $_" -Title "Import Error" -Type Error -Buttons OK
        }
    })

# Cancel Button (in second row, rightmost position)
$global:CancelButton = New-Object System.Windows.Forms.Button
$global:CancelButton.Location = New-Object System.Drawing.Point(725, 48)
$global:CancelButton.Size = New-Object System.Drawing.Size(120, 32)
$global:CancelButton.Text = "[X] Cancel"
$global:CancelButton.BackColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
$global:CancelButton.ForeColor = [System.Drawing.Color]::White
$global:CancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:CancelButton.FlatAppearance.BorderSize = 0
$global:CancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:CancelButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:CancelButton.Enabled = $false
$global:CancelButton.Add_MouseEnter({
        if ($this.Enabled) { $this.BackColor = [System.Drawing.Color]::FromArgb(200, 15, 30) }
    })
$global:CancelButton.Add_MouseLeave({
        if ($this.Enabled) { $this.BackColor = [System.Drawing.Color]::FromArgb(232, 17, 35) }
    })
$global:CancelButton.Add_Click({
        if ($global:CurrentOperation) {
            $result = Show-ModernDialog -Message "Cancel current $($global:CurrentOperation) operation?`r`n`r`nThis may leave the profile in an inconsistent state." -Title "Cancel Operation" -Type Warning -Buttons YesNo
            if ($result -eq "Yes") {
                $global:CancelRequested = $true
                Log-Message "User requested cancellation of $($global:CurrentOperation) operation"
                $global:StatusText.Text = "Cancelling operation..."
                $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
                $global:CancelButton.Enabled = $false
            }
        }
    })
$cardPanel.Controls.Add($global:CancelButton)

# File selection indicator
$global:FileLabel = New-Object System.Windows.Forms.Label
$global:FileLabel.Location = New-Object System.Drawing.Point(120, 85)
$global:FileLabel.Size = New-Object System.Drawing.Size(725, 35)
$global:FileLabel.Text = "No file selected - Use Browse to select a profile archive"
$global:FileLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$global:FileLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$cardPanel.Controls.Add($global:FileLabel)

# Domain Join Card
$domainCard = New-Object System.Windows.Forms.Panel
$domainCard.Location = New-Object System.Drawing.Point(15, 210)
$domainCard.Size = New-Object System.Drawing.Size(875, 180)
$domainCard.BackColor = [System.Drawing.Color]::White
$global:Form.Controls.Add($domainCard)

$lblDomainHeader = New-Object System.Windows.Forms.Label
$lblDomainHeader.Location = New-Object System.Drawing.Point(15, 10)
$lblDomainHeader.Size = New-Object System.Drawing.Size(250, 25)
$lblDomainHeader.Text = "Domain Join Settings"
$lblDomainHeader.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblDomainHeader.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$domainCard.Controls.Add($lblDomainHeader)

$global:DomainCheckBox = New-Object System.Windows.Forms.CheckBox
$global:DomainCheckBox.Location = New-Object System.Drawing.Point(25, 45)
$global:DomainCheckBox.Size = New-Object System.Drawing.Size(200, 23)
$global:DomainCheckBox.Text = "Join Domain"
$global:DomainCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$domainCard.Controls.Add($global:DomainCheckBox)
$lblC = New-Object System.Windows.Forms.Label
$lblC.Location = New-Object System.Drawing.Point(25, 75)
$lblC.Size = New-Object System.Drawing.Size(120, 23)
$lblC.Text = "Computer Name:"
$domainCard.Controls.Add($lblC)
$global:ComputerNameTextBox = New-Object System.Windows.Forms.TextBox
$global:ComputerNameTextBox.Location = New-Object System.Drawing.Point(150, 73)
$global:ComputerNameTextBox.Size = New-Object System.Drawing.Size(200, 25)
$global:ComputerNameTextBox.Text = $env:COMPUTERNAME
$global:ComputerNameTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$domainCard.Controls.Add($global:ComputerNameTextBox)

$lblD = New-Object System.Windows.Forms.Label
$lblD.Location = New-Object System.Drawing.Point(380, 75)
$lblD.Size = New-Object System.Drawing.Size(100, 23)
$lblD.Text = "Domain Name:"
$domainCard.Controls.Add($lblD)
$global:DomainNameTextBox = New-Object System.Windows.Forms.TextBox
$global:DomainNameTextBox.Location = New-Object System.Drawing.Point(485, 73)
$global:DomainNameTextBox.Size = New-Object System.Drawing.Size(250, 25)
$global:DomainNameTextBox.Text = "corp.example.com"
$global:DomainNameTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$domainCard.Controls.Add($global:DomainNameTextBox)

$lblRestart = New-Object System.Windows.Forms.Label
$lblRestart.Location = New-Object System.Drawing.Point(25, 110)
$lblRestart.Size = New-Object System.Drawing.Size(120, 23)
$lblRestart.Text = "Restart Mode:"
$domainCard.Controls.Add($lblRestart)
$global:RestartComboBox = New-Object System.Windows.Forms.ComboBox
$global:RestartComboBox.Location = New-Object System.Drawing.Point(150, 108)
$global:RestartComboBox.Size = New-Object System.Drawing.Size(200, 25)
$global:RestartComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$global:RestartComboBox.Items.AddRange(@('Prompt', 'Delayed', 'Never', 'Immediate'))
$global:RestartComboBox.SelectedIndex = 0
$global:RestartComboBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$domainCard.Controls.Add($global:RestartComboBox)

$lblDelay = New-Object System.Windows.Forms.Label
$lblDelay.Location = New-Object System.Drawing.Point(380, 110)
$lblDelay.Size = New-Object System.Drawing.Size(100, 23)
$lblDelay.Text = "Delay (sec):"
$domainCard.Controls.Add($lblDelay)
$global:DelayTextBox = New-Object System.Windows.Forms.TextBox
$global:DelayTextBox.Location = New-Object System.Drawing.Point(485, 108)
$global:DelayTextBox.Size = New-Object System.Drawing.Size(80, 25)
$global:DelayTextBox.Text = "30"
$global:DelayTextBox.Enabled = $false
$global:DelayTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$domainCard.Controls.Add($global:DelayTextBox)

$global:DomainJoinButton = New-Object System.Windows.Forms.Button
$global:DomainJoinButton.Location = New-Object System.Drawing.Point(25, 143)
$global:DomainJoinButton.Size = New-Object System.Drawing.Size(150, 30)
$global:DomainJoinButton.Text = "Join Now"
$global:DomainJoinButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$global:DomainJoinButton.ForeColor = [System.Drawing.Color]::White
$global:DomainJoinButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:DomainJoinButton.FlatAppearance.BorderSize = 0
$global:DomainJoinButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:DomainJoinButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:DomainJoinButton.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(16, 110, 190)
    })
$global:DomainJoinButton.Add_Click({
        $domain = $global:DomainNameTextBox.Text.Trim()
        if (-not $domain) {
            Show-ModernDialog -Message "Please enter a valid domain name." -Title "Input Required" -Type Warning -Buttons OK
            return
        }

        $computerName = $global:ComputerNameTextBox.Text.Trim()
        
        # Confirm details
        $msg = "Ready to join domain: $domain"
        if ($computerName -ne $env:COMPUTERNAME) {
            $msg += "`r`nNew Computer Name: $computerName"
            $msg += "`r`n`r`nNOTE: Changing computer name requires a restart."
        }
        $msg += "`r`n`r`nContinue?"
        
        if ((Show-ModernDialog -Message $msg -Title "Confirm Domain Join" -Type Question -Buttons YesNo) -ne "Yes") { return }

        # Get Credentials
        $cred = Get-DomainCredential -Domain $domain
        if (-not $cred) { return }

        # Join
        try {
            $global:StatusText.Text = "Joining domain '$domain'..."
            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            [System.Windows.Forms.Application]::DoEvents()

            $restartMode = $global:RestartComboBox.SelectedItem
            $delay = 30
            if ($global:DelayTextBox.Enabled -and $global:DelayTextBox.Text -match '^\d+$') {
                $delay = [int]$global:DelayTextBox.Text
            }

            Join-Domain-Enhanced -DomainName $domain -ComputerName $computerName -Credential $cred -RestartBehavior $restartMode -DelaySeconds $delay
            
            # Status will be updated by function or restart triggered
            $global:StatusText.Text = "Domain join initiated successfully."
            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
        }
        catch {
            Log-Error "Domain join failed: $_"
            $global:StatusText.Text = "Domain join failed"
            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
            Show-ModernDialog -Message "Failed to join domain:`r`n`r`n$_" -Title "Join Error" -Type Error -Buttons OK
        }
    })
$global:DomainJoinButton.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    })
$domainCard.Controls.Add($global:DomainJoinButton)

# Retry domain reachability button

$global:DomainRetryButton = New-Object System.Windows.Forms.Button
$global:DomainRetryButton.Location = New-Object System.Drawing.Point(185, 143)
$global:DomainRetryButton.Size = New-Object System.Drawing.Size(165, 30)
$global:DomainRetryButton.Text = "Retry Domain Check"
$global:DomainRetryButton.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$global:DomainRetryButton.ForeColor = [System.Drawing.Color]::White
$global:DomainRetryButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$global:DomainRetryButton.FlatAppearance.BorderSize = 0
$global:DomainRetryButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$global:DomainRetryButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$global:DomainRetryButton.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    })
$global:DomainRetryButton.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    })
$global:DomainRetryButton.Add_Click({
        $domainName = $global:DomainNameTextBox.Text.Trim()
        if (-not $domainName) {
            Show-ModernDialog -Message "Please enter a domain name to check." -Title "Domain Name Required" -Type Warning -Buttons OK
            return
        }
        $global:StatusText.Text = "Checking domain reachability..."
        $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        [System.Windows.Forms.Application]::DoEvents()
        $result = Test-DomainReachability -DomainName $domainName
        if ($result.Success) {
            $global:StatusText.Text = "Domain '$domainName' is reachable."
            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
            Show-ModernDialog -Message "Domain '$domainName' is reachable!" -Title "Domain Check" -Type Success -Buttons OK
        }
        else {
            $global:StatusText.Text = "Domain '$domainName' unreachable: $($result.Error)"
            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
            Show-ModernDialog -Message "Domain '$domainName' is unreachable.`r`n`r`n$result.Error" -Title "Domain Check Failed" -Type Error -Buttons OK
        }
        [System.Windows.Forms.Application]::DoEvents()
    })
$domainCard.Controls.Add($global:DomainRetryButton)

$global:RestartComboBox.Add_SelectedIndexChanged({ $global:DelayTextBox.Enabled = ($global:RestartComboBox.SelectedItem -eq 'Delayed') })
$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.SetToolTip($global:RestartComboBox, "Prompt: Ask before restarting`nDelayed: Countdown before restart`nNever: Manual restart required`nImmediate: Restart without asking")

# Progress and Status Card
$progressCard = New-Object System.Windows.Forms.Panel
$progressCard.Location = New-Object System.Drawing.Point(15, 405)
$progressCard.Size = New-Object System.Drawing.Size(875, 75)
$progressCard.BackColor = [System.Drawing.Color]::White
$global:Form.Controls.Add($progressCard)

$global:ProgressBar = New-Object System.Windows.Forms.ProgressBar
$global:ProgressBar.Location = New-Object System.Drawing.Point(15, 15)
$global:ProgressBar.Size = New-Object System.Drawing.Size(845, 25)
$global:ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$global:ProgressBar.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$progressCard.Controls.Add($global:ProgressBar)

$global:StatusText = New-Object System.Windows.Forms.Label
$global:StatusText.Location = New-Object System.Drawing.Point(15, 45)
$global:StatusText.Size = New-Object System.Drawing.Size(845, 23)
$global:StatusText.Text = "[Ready]"
$global:StatusText.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$progressCard.Controls.Add($global:StatusText)

# Activity Log Section
$lblLogHeader = New-Object System.Windows.Forms.Label
$lblLogHeader.Location = New-Object System.Drawing.Point(20, 490)
$lblLogHeader.Size = New-Object System.Drawing.Size(200, 25)
$lblLogHeader.Text = "Activity Log"
$lblLogHeader.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblLogHeader.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$global:Form.Controls.Add($lblLogHeader)

$global:LogBox = New-Object System.Windows.Forms.TextBox
$global:LogBox.Location = New-Object System.Drawing.Point(15, 520)
$global:LogBox.Size = New-Object System.Drawing.Size(875, 140)
$global:LogBox.Multiline = $true
$global:LogBox.ScrollBars = "Vertical"
$global:LogBox.ReadOnly = $true
$global:LogBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$global:LogBox.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$global:LogBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$global:Form.Controls.Add($global:LogBox)

# =============================================================================
# BUTTON HANDLERS
# =============================================================================
$global:BrowseButton.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "ZIP Files (*.zip)|*.zip"
        if ($dlg.ShowDialog() -eq "OK") {

            $global:SelectedZipPath = $dlg.FileName
        
            # Update file selection indicator
            $fileInfo = Get-Item $dlg.FileName
            $sizeKB = [math]::Round($fileInfo.Length / 1KB, 2)
            $sizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
            $sizeDisplay = if ($sizeMB -gt 1) { "$sizeMB MB" } else { "$sizeKB KB" }

            $global:FileLabel.Text = "Selected: $($fileInfo.Name) ($sizeDisplay)"
            $global:FileLabel.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
        
            $global:ProgressBar.Value = 5
            $global:StatusText.Text = "Analyzing ZIP file..."
            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            [System.Windows.Forms.Application]::DoEvents()
            Log-Message "Selected: $global:SelectedZipPath"

            try {
                $global:ProgressBar.Value = 10
                $global:StatusText.Text = "Reading manifest from ZIP..."
                [System.Windows.Forms.Application]::DoEvents()

                Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
                $zip = [IO.Compression.ZipFile]::OpenRead($global:SelectedZipPath)
                $entry = $zip.Entries | Where-Object { $_.FullName -ieq 'manifest.json' }

                if ($entry) {

                    $sr = New-Object IO.StreamReader($entry.Open())
                    $jsonContent = $sr.ReadToEnd()
                    $sr.Dispose()
                    $zip.Dispose()

                    try {
                        $manifestObj = $jsonContent | ConvertFrom-Json -ErrorAction Stop

                        # AzureAD?
                        if ($manifestObj.IsAzureADUser) {
                            Log-Message "Manifest found for AzureAD profile: $($manifestObj.Username)"
                            $global:ProgressBar.Value = 100
                            $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf) - AzureAD profile"
                            $global:StatusText.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
                        }
                        elseif ($manifestObj.IsDomainUser -and $manifestObj.Domain) {
                            $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf) - Domain profile"
                        }
                        else {
                            Log-Message "Manifest found but not a domain profile (IsDomainUser=$($manifestObj.IsDomainUser))"
                            $global:ProgressBar.Value = 0
                            $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf)"
                        }
                    }
                    catch {
                        Log-Message "Domain auto-detect parse failed: $($_.Exception.Message)"
                        $global:ProgressBar.Value = 0
                        $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf)"
                    }

                }
                else {
                    $zip.Dispose()
                    Log-Message "No manifest.json in ZIP - skipping auto-detect"
                    $global:ProgressBar.Value = 0
                    $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf)"
                }
            }
            catch {
                Log-Message "Auto-detect skipped: $($_.Exception.Message)"
                $global:ProgressBar.Value = 0
                $global:StatusText.Text = "ZIP: $(Split-Path $global:SelectedZipPath -Leaf)"
            }
        }
    })




# =============================================================================
# LAUNCH
# =============================================================================


Log-Message "Profile Transfer Tool $($Config.Version)"
Log-Message "IMPORTANT: After import, RESTART computer before logging in!"
Log-Message "Restart Mode Options: Prompt (ask), Delayed (countdown), Never (manual), Immediate (auto)"

# Detect Windows theme and apply initial theme
$initialTheme = Get-WindowsTheme
Apply-Theme -ThemeName $initialTheme

$global:Form.ShowDialog() | Out-Null
