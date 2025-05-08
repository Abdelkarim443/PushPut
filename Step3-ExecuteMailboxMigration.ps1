#==========================================================================
# Script: Step3-ExecuteMailboxMigration.ps1
# Author: Manus
# Date: 04/11/2025 (Modified: 2025-05-08)
# Description: Script to execute mailbox migration via Mailbox Restore Request
#              - Finds mailboxes with CustomAttribute6 containing "STEP2;OK"
#              - Disables on-premises mailboxes (and attempts rollback on failure)
#              - Executes New-MailboxRestoreRequest using ExchangeGuid and ArchiveGuid (and attempts rollback)
#              - Updates CustomAttribute6 to "INITIATED" on success, or "KO" on failure for this step.
#==========================================================================

#==========================================================================
# Configuration - Predefined Paths
#==========================================================================
$BasePath = "C:\ExchangeMigration"
$LogPath = "$BasePath\Logs"
$CSVPath = "$BasePath\CSV"
$ReportsPath = "$BasePath\Reports"

$LogFile = Join-Path -Path $LogPath -ChildPath "Step3-ExecuteMailboxMigration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$StatusReportFile = Join-Path -Path $ReportsPath -ChildPath "MigrationInitiated_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

$DefaultBadItemLimit = 10

#==========================================================================
# Logging Function
#==========================================================================
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    if (-not (Test-Path -Path $LogPath)) {
        try { New-Item -Path $LogPath -ItemType Directory -Force | Out-Null; Write-Output "Log folder created: $LogPath" }
        catch { Write-Error "Unable to create log folder: $($_.Exception.Message)"; return }
    }
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$TimeStamp] [$Level] $Message"
    try { Add-Content -Path $LogFile -Value $LogMessage -ErrorAction Stop }
    catch { Write-Error "Unable to write to log file: $($_.Exception.Message)" }
    switch ($Level) {
        "INFO"    { Write-Host $LogMessage -ForegroundColor Cyan }
        "WARNING" { Write-Host $LogMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $LogMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $LogMessage -ForegroundColor Green }
        default   { Write-Host $LogMessage }
    }
}

#==========================================================================
# Function to establish Exchange sessions
#==========================================================================
function Connect-ExchangeSessions {
    Write-Log "Setting up Exchange sessions" -Level "INFO"
    $SessionsCreated = $true
    if (-not (Get-Command Get-OnpremMailbox -ErrorAction SilentlyContinue)) {
        Write-Log "On-premises Exchange commands not found. Prefixing..." -Level "INFO"
        if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
            Write-Log "Exchange module not loaded. Attempting to load..." -Level "WARNING"
            try { Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop; Write-Log "Exchange module loaded." -Level "SUCCESS" }
            catch { Write-Log "Unable to load Exchange module." -Level "ERROR"; $SessionsCreated = $false }
        }
        if ($SessionsCreated) {
            try {
                $OnpremSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
                if ($null -eq $OnpremSession) { Write-Log "No active Exchange session. Run in EMS." -Level "ERROR"; $SessionsCreated = $false }
                else { Import-PSSession $OnpremSession -Prefix "Onprem" -DisableNameChecking -AllowClobber | Out-Null; Write-Log "On-prem Exchange cmdlets prefixed." -Level "SUCCESS" }
            } catch { Write-Log "Error setting up On-prem session: $($_.Exception.Message)" -Level "ERROR"; $SessionsCreated = $false }
        }
    }
    if (-not (Get-Command Get-CloudMailbox -ErrorAction SilentlyContinue)) {
        Write-Log "Exchange Online commands not found. Prefixing..." -Level "INFO"
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) { Write-Log "ExchangeOnlineManagement module not installed." -Level "ERROR"; $SessionsCreated = $false }
        else {
            try {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop; Write-Log "ExchangeOnlineManagement module imported." -Level "SUCCESS"
                try {
                    $CredentialPath = "$BasePath\ExchangeOnlineCredential.xml"
                    if (Test-Path $CredentialPath) {
                        $CloudCredential = Import-Clixml -Path $CredentialPath
                        Connect-ExchangeOnline -Credential $CloudCredential -ShowBanner:$false -ErrorAction Stop
                    } else { Write-Log "Cloud credentials not found. Default connection (may prompt)." -Level "WARNING"; Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop }
                    $CloudSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" -and $_.ComputerName -like "*.outlook.com" }
                    if ($null -eq $CloudSession) { Write-Log "No active EXO session." -Level "ERROR"; $SessionsCreated = $false }
                    else { Import-PSSession $CloudSession -Prefix "Cloud" -DisableNameChecking -AllowClobber | Out-Null; Write-Log "EXO cmdlets prefixed." -Level "SUCCESS" }
                } catch { Write-Log "Error connecting to EXO: $($_.Exception.Message)" -Level "ERROR"; $SessionsCreated = $false }
            } catch { Write-Log "Error importing EXO module: $($_.Exception.Message)" -Level "ERROR"; $SessionsCreated = $false }
        }
    }
    return $SessionsCreated
}

#==========================================================================
# Function to find eligible mailboxes
#==========================================================================
function Find-EligibleMailboxes {
    Write-Log "Finding eligible mailboxes with CustomAttribute6 containing 'STEP2;OK'" -Level "INFO"
    # Adjust filter if using MIGWAVE format: Where-Object { $_.CustomAttribute6 -eq "MIGWAVE:$TargetWaveDate;STEP:2;STATUS:OK" }
    try {
        $OnPremMailboxes = Get-OnpremMailbox -ResultSize Unlimited | Where-Object { $_.CustomAttribute6 -like "*DEL_MIG;STEP2;OK;*" }
        if ($OnPremMailboxes.Count -eq 0) { Write-Log "No on-prem mailboxes found for Step 3." -Level "WARNING"; return $null }
        Write-Log "Found $($OnPremMailboxes.Count) eligible on-prem mailboxes." -Level "SUCCESS"
        $MailboxPairs = @()
        foreach ($OnPremMailbox in $OnPremMailboxes) {
            Write-Log "Processing pair for on-prem: $($OnPremMailbox.Identity)" -Level "INFO"
            try {
                $CloudMailboxes = Get-CloudMailbox -ResultSize Unlimited | Where-Object { $_.DisplayName -eq $OnPremMailbox.DisplayName }
                if ($CloudMailboxes.Count -eq 0) { Write-Log "No cloud match for $($OnPremMailbox.Identity) (DisplayName: $($OnPremMailbox.DisplayName))" -Level "WARNING"; continue }
                if ($CloudMailboxes.Count -gt 1) { Write-Log "Multiple cloud matches for $($OnPremMailbox.DisplayName). Using first. VERIFY." -Level "WARNING" }
                $CloudMailbox = $CloudMailboxes[0]
                $MailboxPairs += [PSCustomObject]@{
                    OnPremIdentityObject    = $OnPremMailbox # Pass the whole object
                    CloudIdentity           = $CloudMailbox.Identity
                    CloudPrimarySmtpAddress = $CloudMailbox.PrimarySmtpAddress
                }
                Write-Log "Added pair: $($OnPremMailbox.Identity) -> $($CloudMailbox.Identity)" -Level "SUCCESS"
            } catch { Write-Log "Error finding cloud match for $($OnPremMailbox.Identity): $($_.Exception.Message)" -Level "ERROR" }
        }
        Write-Log "Found $($MailboxPairs.Count) mailbox pairs for migration." -Level "INFO"
        return $MailboxPairs
    } catch { Write-Log "Error in Find-EligibleMailboxes: $($_.Exception.Message)" -Level "ERROR"; return $null }
}

#==========================================================================
# Function to disable on-premises mailbox (Renamed)
#==========================================================================
function Invoke-MyDisableOnPremMailbox {
    param ([Parameter(Mandatory = $true)][string]$Identity)
    Write-Log "Disabling on-prem mailbox: $Identity" -Level "INFO"
    try {
        Disable-OnpremMailbox -Identity $Identity -Confirm:$false -ErrorAction Stop
        Write-Log "Successfully disabled on-prem mailbox: $Identity" -Level "SUCCESS"
        return $true
    } catch {
        Write-Log "Error disabling on-prem mailbox $Identity: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

#==========================================================================
# Function to execute mailbox restore request with rollback
#==========================================================================
function Start-MailboxRestoreRequest {
    param (
        [Parameter(Mandatory = $true)] [object]$MailboxPair,
        [Parameter(Mandatory = $false)] [int]$BadItemLimit = $DefaultBadItemLimit
    )

    $OriginalOnPremMailbox = $MailboxPair.OnPremIdentityObject # Full object from Find-EligibleMailboxes
    $OnPremIdentity = $OriginalOnPremMailbox.Identity
    $CloudIdentity = $MailboxPair.CloudIdentity

    $OnPremMailboxDisabled = $false
    $RestoreRequestSubmitted = $false
    $RestoreRequestIdentity = $null
    $CurrentDateForStatus = Get-Date -Format "yyyy-MM-dd" # Or your fixed wave date
    
    # Define CA6 status values (adjust if using MIGWAVE format, e.g., "MIGWAVE:$WaveDate;STEP:3;STATUS:KO")
    $StatusKO = "DEL_MIG;STEP3;KO;$CurrentDateForStatus"
    $StatusInitiated = "DEL_MIG;STEP3;INITIATED;$CurrentDateForStatus"

    Write-Log "Processing Step 3 for $OnPremIdentity -> $CloudIdentity" -Level "INFO"

    try {
        # Details already in $OriginalOnPremMailbox from Find-EligibleMailboxes
        Write-Log "Using On-prem details: Name '$($OriginalOnPremMailbox.Name)', ExchangeGuid '$($OriginalOnPremMailbox.ExchangeGuid)', ArchiveGuid '$($OriginalOnPremMailbox.ArchiveGuid)'" -Level "INFO"

        # 1. Disable On-Prem Mailbox
        Write-Log "Attempting to disable on-prem mailbox $OnPremIdentity..." -Level "INFO"
        if (-not (Invoke-MyDisableOnPremMailbox -Identity $OnPremIdentity)) {
            throw "Invoke-MyDisableOnPremMailbox returned false for $OnPremIdentity." # This will be caught by the main catch
        }
        $OnPremMailboxDisabled = $true
        Write-Log "Successfully disabled on-prem mailbox $OnPremIdentity." -Level "SUCCESS"

        # 2. Prepare and Start Mailbox Restore Request
        $RequestName = "MIG_$($OriginalOnPremMailbox.Alias)_$($OriginalOnPremMailbox.ExchangeGuid)_$(Get-Date -Format 'yyyyMMddHHmmss')"
        $RestoreParams = @{
            SourceStoreMailbox    = $OriginalOnPremMailbox.ExchangeGuid
            TargetMailbox         = $CloudIdentity
            Name                  = $RequestName
            AllowLegacyDNMismatch = $true
            BadItemLimit          = $BadItemLimit
        }

        if ($OriginalOnPremMailbox.ArchiveGuid -and ($OriginalOnPremMailbox.ArchiveStatus -eq "Active" -or $OriginalOnPremMailbox.ArchiveState -eq "Local")) {
            Write-Log "On-prem mailbox $OnPremIdentity has an active archive (GUID: $($OriginalOnPremMailbox.ArchiveGuid)). Including in restore." -Level "INFO"
            $RestoreParams.Add("SourceArchiveGUID", $OriginalOnPremMailbox.ArchiveGuid)
            # Note: Ensure target cloud shared mailbox has an archive enabled if you intend to restore archive data.
        }

        Write-Log "Attempting to create restore request for $OnPremIdentity with parameters: $($RestoreParams | Out-String)" -Level "INFO"
        $RestoreRequest = New-CloudMailboxRestoreRequest @RestoreParams -ErrorAction Stop
        if (-not $RestoreRequest) { # Should not happen with ErrorAction Stop, but as a safeguard.
            throw "New-CloudMailboxRestoreRequest returned null or failed silently for $OnPremIdentity."
        }
        $RestoreRequestSubmitted = $true
        $RestoreRequestIdentity = $RestoreRequest.Identity
        Write-Log "Mailbox restore request submitted successfully. Request ID: $($RestoreRequestIdentity)" -Level "SUCCESS"

        # 3. Update CustomAttribute6 on Cloud Mailbox to INITIATED
        Write-Log "Attempting to update CustomAttribute6 for cloud mailbox $CloudIdentity to $StatusInitiated..." -Level "INFO"
        Set-CloudMailbox -Identity $CloudIdentity -CustomAttribute6 $StatusInitiated -ErrorAction Stop
        Write-Log "Successfully updated cloud mailbox $CloudIdentity CustomAttribute6 to $StatusInitiated." -Level "SUCCESS"

        return [PSCustomObject]@{
            RequestId    = $RestoreRequestIdentity; RequestName = $RequestName
            OnPremIdentity = $OnPremIdentity; CloudIdentity = $CloudIdentity
            Status       = $RestoreRequest.Status; StartTime = $RestoreRequest.RequestQueue
        }

    } catch {
        Write-Log "An error occurred during Step 3 for $OnPremIdentity: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Performing rollback actions for $OnPremIdentity..." -Level "WARNING"

        try {
            Write-Log "Setting CustomAttribute6 on cloud mailbox $CloudIdentity to $StatusKO." -Level "INFO"
            Set-CloudMailbox -Identity $CloudIdentity -CustomAttribute6 $StatusKO -ErrorAction SilentlyContinue # Allow other rollbacks even if this fails
            Write-Log "Attempted to set CustomAttribute6 on cloud mailbox $CloudIdentity to $StatusKO." -Level "INFO" # Changed to INFO as it's an attempt
        } catch {
            Write-Log "Failed to set CustomAttribute6 to $StatusKO for cloud mailbox $CloudIdentity. Error: $($_.Exception.Message)" -Level "ERROR"
        }

        if ($RestoreRequestSubmitted) {
            Write-Log "Rollback: Restore request was submitted ($RestoreRequestIdentity). Attempting to remove it." -Level "WARNING"
            try {
                Get-CloudMailboxRestoreRequest -Identity $RestoreRequestIdentity -ErrorAction SilentlyContinue | Remove-CloudMailboxRestoreRequest -Confirm:$false -ErrorAction Stop
                Write-Log "Successfully removed restore request $RestoreRequestIdentity." -Level "SUCCESS"
            } catch {
                Write-Log "Failed to remove restore request $RestoreRequestIdentity. Manual cleanup may be required. Error: $($_.Exception.Message)" -Level "ERROR"
            }
        }

        if ($OnPremMailboxDisabled) {
            Write-Log "Rollback: On-prem mailbox $OnPremIdentity was disabled. Attempting to re-enable it." -Level "WARNING"
            try {
                if ($OriginalOnPremMailbox -and $OriginalOnPremMailbox.ExchangeGuid) {
                    Enable-OnpremMailbox -Identity $OriginalOnPremMailbox.UserPrincipalName -DisconnectedMailbox $OriginalOnPremMailbox.ExchangeGuid -ErrorAction Stop
                    Write-Log "Successfully re-enabled (reconnected) on-prem mailbox $OnPremIdentity using its original ExchangeGuid." -Level "SUCCESS"
                } else {
                     Write-Log "Original ExchangeGuid for $OnPremIdentity not available. Attempting Enable-Mailbox by identity only (may create new empty mailbox)." -Level "WARNING"
                    Enable-OnpremMailbox -Identity $OnPremIdentity -ErrorAction Stop # Fallback
                    Write-Log "Attempted to re-enable on-prem mailbox $OnPremIdentity by identity." -Level "SUCCESS"
                }
            } catch {
                Write-Log "Rollback FAILED: Unable to re-enable on-prem mailbox $OnPremIdentity. Manual intervention required. Error: $($_.Exception.Message)" -Level "ERROR"
            }
        }
        return $null # Indicate failure
    }
}

#==========================================================================
# Function to verify and create required folders
#==========================================================================
function Initialize-Environment {
    Write-Log "Initializing environment and verifying paths" -Level "INFO"
    $Folders = @($BasePath, $LogPath, $CSVPath, $ReportsPath)
    foreach ($Folder in $Folders) {
        if (-not (Test-Path -Path $Folder)) {
            try { New-Item -Path $Folder -ItemType Directory -Force | Out-Null; Write-Log "Folder created: $Folder" -Level "SUCCESS" }
            catch { Write-Log "Error creating folder $Folder: $($_.Exception.Message)" -Level "ERROR"; return $false }
        }
    }
    return $true
}

#==========================================================================
# Main function
#==========================================================================
function Start-MailboxMigration {
    Write-Log "Starting mailbox migration process (Step 3)" -Level "INFO"
    Write-Log "Using BadItemLimit: $DefaultBadItemLimit. Restore includes archive if present." -Level "INFO"

    $MailboxPairs = Find-EligibleMailboxes
    if ($null -eq $MailboxPairs -or $MailboxPairs.Count -eq 0) { Write-Log "No eligible mailbox pairs found. Script ending." -Level "WARNING"; return }

    $TotalMailboxes = $MailboxPairs.Count
    $SuccessfulInitiations = 0
    $FailedInitiations = 0
    Write-Log "Found $TotalMailboxes mailbox pairs to process for migration." -Level "INFO"
    $StatusReport = @()

    foreach ($MailboxPair in $MailboxPairs) {
        $RestoreOutcome = Start-MailboxRestoreRequest -MailboxPair $MailboxPair # -BadItemLimit can be passed if needed
        $CurrentTimestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

        if ($RestoreOutcome) {
            $SuccessfulInitiations++
            Write-Log "Successfully processed and initiated migration for $($MailboxPair.OnPremIdentityObject.Identity)." -Level "SUCCESS"
            $StatusReport += [PSCustomObject]@{
                OnPremIdentity = $MailboxPair.OnPremIdentityObject.Identity; CloudIdentity = $MailboxPair.CloudIdentity
                RequestId = $RestoreOutcome.RequestId; RequestName = $RestoreOutcome.RequestName
                Step3Status = "Initiated"; InitiationTime = $CurrentTimestamp
            }
        } else {
            $FailedInitiations++
            Write-Log "Failed to process or initiate migration for $($MailboxPair.OnPremIdentityObject.Identity). Check logs for rollback details." -Level "ERROR"
             $StatusReport += [PSCustomObject]@{
                OnPremIdentity = $MailboxPair.OnPremIdentityObject.Identity; CloudIdentity = $MailboxPair.CloudIdentity
                RequestId = "N/A"; RequestName = "N/A"
                Step3Status = "Failed (KO)"; InitiationTime = $CurrentTimestamp # CA6 on cloud should reflect KO
            }
        }
    }

    if ($StatusReport.Count -gt 0) {
        try { $StatusReport | Export-Csv -Path $StatusReportFile -NoTypeInformation -Force; Write-Log "Exported status report to $StatusReportFile" -Level "SUCCESS" }
        catch { Write-Log "Error exporting status report: $($_.Exception.Message)" -Level "ERROR" }
    }
    Write-Log "Migration initiation process complete. Summary:" -Level "INFO"
    Write-Log "- Total pairs processed: $TotalMailboxes" -Level "INFO"
    Write-Log "- Successful initiations: $SuccessfulInitiations" -Level "SUCCESS"
    Write-Log "- Failed initiations (rollback attempted): $FailedInitiations" -Level "WARNING"
}

#==========================================================================
# Script execution
#==========================================================================
try {
    if (-not (Initialize-Environment)) { Write-Log "Failed to initialize environment. Exiting." -Level "ERROR"; exit 1 }
    if (-not (Connect-ExchangeSessions)) { Write-Log "Failed to set up Exchange sessions. Exiting." -Level "ERROR"; exit 1 }
    Start-MailboxMigration
    Write-Log "Script execution completed." -Level "SUCCESS"
}
catch {
    Write-Log "Unhandled error in script: $($_.Exception.Message)" -Level "ERROR"
    Write-Log "Error details: $($_.Exception.StackTrace)" -Level "ERROR"
    exit 1
}
