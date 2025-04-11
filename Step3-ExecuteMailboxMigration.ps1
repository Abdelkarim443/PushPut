#==========================================================================
# Script: Step3-ExecuteMailboxMigration.ps1
# Author: Manus
# Date: 04/11/2025
# Description: Script to execute mailbox migration via Mailbox Restore Request
#              - Finds mailboxes with CustomAttribute6 containing "STEP2;OK" from Step 2
#              - Disables on-premises mailboxes before migration
#              - Executes New-MailboxRestoreRequest to migrate mailbox content
#              - Updates CustomAttribute6 to track migration initiation
#              - Runs autonomously without user input
#              - Does not verify completion (handled by a separate script)
#==========================================================================

#==========================================================================
# Configuration - Predefined Paths
#==========================================================================
# Base paths - Update these paths to match your environment
$BasePath = "C:\ExchangeMigration"
$LogPath = "$BasePath\Logs"
$CSVPath = "$BasePath\CSV"
$ReportsPath = "$BasePath\Reports"

# Log file
$LogFile = Join-Path -Path $LogPath -ChildPath "Step3-ExecuteMailboxMigration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Status report file
$StatusReportFile = Join-Path -Path $ReportsPath -ChildPath "MigrationInitiated_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Default migration parameters - Change these values to match your requirements
$DefaultAllowLargeItems = $true
$DefaultLargeItemLimit = 50
$DefaultAcceptLargeDataLoss = $true
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
    
    # Create log folder if it doesn't exist
    if (-not (Test-Path -Path $LogPath)) {
        try {
            New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
            Write-Output "Log folder created: $LogPath"
        }
        catch {
            Write-Error "Unable to create log folder: $($_.Exception.Message)"
            return
        }
    }
    
    # Log message format
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$TimeStamp] [$Level] $Message"
    
    # Write to log file
    try {
        Add-Content -Path $LogFile -Value $LogMessage -ErrorAction Stop
    }
    catch {
        Write-Error "Unable to write to log file: $($_.Exception.Message)"
    }
    
    # Also display in console with color coding
    switch ($Level) {
        "INFO" { Write-Host $LogMessage -ForegroundColor Cyan }
        "WARNING" { Write-Host $LogMessage -ForegroundColor Yellow }
        "ERROR" { Write-Host $LogMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $LogMessage -ForegroundColor Green }
        default { Write-Host $LogMessage }
    }
}

#==========================================================================
# Function to establish Exchange sessions
#==========================================================================
function Connect-ExchangeSessions {
    Write-Log "Setting up Exchange sessions" -Level "INFO"
    
    $SessionsCreated = $true
    
    # Check if On-premises Exchange commands are available
    if (-not (Get-Command Get-OnpremMailbox -ErrorAction SilentlyContinue)) {
        Write-Log "On-premises Exchange commands not found. Creating prefix for On-premises Exchange." -Level "INFO"
        
        # Check if Exchange module is loaded
        if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
            Write-Log "Exchange module is not loaded. Attempting to load..." -Level "WARNING"
            
            # Attempt to load Exchange module
            try {
                Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
                Write-Log "Exchange module loaded successfully" -Level "SUCCESS"
            }
            catch {
                Write-Log "Unable to load Exchange module. Make sure you're running this script on an Exchange server or in an Exchange Management Shell session." -Level "ERROR"
                $SessionsCreated = $false
            }
        }
        
        if ($SessionsCreated) {
            # Create prefix for On-premises Exchange commands
            try {
                $OnpremSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
                if ($null -eq $OnpremSession) {
                    Write-Log "No active Exchange session found. Please run this script in an Exchange Management Shell." -Level "ERROR"
                    $SessionsCreated = $false
                }
                else {
                    Import-PSSession $OnpremSession -Prefix "Onprem" -DisableNameChecking -AllowClobber | Out-Null
                    Write-Log "On-premises Exchange commands imported with 'Onprem' prefix" -Level "SUCCESS"
                }
            }
            catch {
                Write-Log "Error setting up On-premises Exchange session: $($_.Exception.Message)" -Level "ERROR"
                $SessionsCreated = $false
            }
        }
    }
    
    # Check if Exchange Online commands are available
    if (-not (Get-Command Get-CloudMailbox -ErrorAction SilentlyContinue)) {
        Write-Log "Exchange Online commands not found. Creating prefix for Exchange Online." -Level "INFO"
        
        # Check if Exchange Online module is installed
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Log "Exchange Online Management module is not installed. Please install it using: Install-Module -Name ExchangeOnlineManagement" -Level "ERROR"
            $SessionsCreated = $false
        }
        else {
            # Import Exchange Online module
            try {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
                Write-Log "Exchange Online Management module imported successfully" -Level "SUCCESS"
                
                # Connect to Exchange Online
                try {
                    # Use stored credentials or certificate-based authentication for automation
                    # This example uses stored credentials - replace with your preferred authentication method
                    $CredentialPath = "$BasePath\ExchangeOnlineCredential.xml"
                    
                    if (Test-Path $CredentialPath) {
                        $CloudCredential = Import-Clixml -Path $CredentialPath
                        Connect-ExchangeOnline -Credential $CloudCredential -ShowBanner:$false -ErrorAction Stop
                    }
                    else {
                        # For first-time setup, create and export credentials
                        Write-Log "Exchange Online credentials not found. Using default connection method." -Level "WARNING"
                        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
                    }
                    
                    # Create prefix for Exchange Online commands
                    $CloudSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" -and $_.ComputerName -like "*.outlook.com" }
                    if ($null -eq $CloudSession) {
                        Write-Log "No active Exchange Online session found." -Level "ERROR"
                        $SessionsCreated = $false
                    }
                    else {
                        Import-PSSession $CloudSession -Prefix "Cloud" -DisableNameChecking -AllowClobber | Out-Null
                        Write-Log "Exchange Online commands imported with 'Cloud' prefix" -Level "SUCCESS"
                    }
                }
                catch {
                    Write-Log "Error connecting to Exchange Online: $($_.Exception.Message)" -Level "ERROR"
                    $SessionsCreated = $false
                }
            }
            catch {
                Write-Log "Error importing Exchange Online Management module: $($_.Exception.Message)" -Level "ERROR"
                $SessionsCreated = $false
            }
        }
    }
    
    return $SessionsCreated
}

#==========================================================================
# Function to find eligible mailboxes
#==========================================================================
function Find-EligibleMailboxes {
    Write-Log "Finding eligible mailboxes with CustomAttribute6 containing 'STEP2;OK'" -Level "INFO"
    
    try {
        # Find on-premises mailboxes with CustomAttribute6 containing "STEP2;OK"
        $OnPremMailboxes = Get-OnpremMailbox -ResultSize Unlimited | 
                           Where-Object { $_.CustomAttribute6 -like "*DEL_MIG;STEP2;OK;*" }
        
        $OnPremCount = $OnPremMailboxes.Count
        
        if ($OnPremCount -eq 0) {
            Write-Log "No eligible on-premises mailboxes found with CustomAttribute6 containing 'STEP2;OK'" -Level "WARNING"
            return $null
        }
        
        Write-Log "Found $OnPremCount eligible on-premises mailboxes" -Level "SUCCESS"
        
        # Create a collection to store mailbox pairs (on-premises and cloud)
        $MailboxPairs = @()
        
        foreach ($OnPremMailbox in $OnPremMailboxes) {
            $OnPremIdentity = $OnPremMailbox.Identity
            $OnPremPrimarySmtpAddress = $OnPremMailbox.PrimarySmtpAddress
            $Username = ($OnPremPrimarySmtpAddress -split '@')[0]
            
            Write-Log "Processing mailbox pair for: $OnPremIdentity" -Level "INFO"
            
            # Find corresponding cloud mailbox
            try {
                # Get all cloud mailboxes and filter by DisplayName
                $CloudMailboxes = Get-CloudMailbox -ResultSize Unlimited | 
                                 Where-Object { $_.DisplayName -eq $OnPremMailbox.DisplayName }
                
                if ($CloudMailboxes.Count -eq 0) {
                    Write-Log "No matching cloud mailbox found for $OnPremIdentity" -Level "WARNING"
                    continue
                }
                
                if ($CloudMailboxes.Count -gt 1) {
                    Write-Log "Multiple cloud mailboxes found with DisplayName: $($OnPremMailbox.DisplayName). Using the first match." -Level "WARNING"
                }
                
                $CloudMailbox = $CloudMailboxes[0]
                
                # Add to collection
                $MailboxPair = [PSCustomObject]@{
                    OnPremIdentity = $OnPremIdentity
                    OnPremPrimarySmtpAddress = $OnPremPrimarySmtpAddress
                    CloudIdentity = $CloudMailbox.Identity
                    CloudPrimarySmtpAddress = $CloudMailbox.PrimarySmtpAddress
                }
                
                $MailboxPairs += $MailboxPair
                Write-Log "Added mailbox pair: $OnPremIdentity -> $($CloudMailbox.Identity)" -Level "SUCCESS"
            }
            catch {
                Write-Log "Error finding cloud mailbox for $OnPremIdentity: $($_.Exception.Message)" -Level "ERROR"
            }
        }
        
        $PairsCount = $MailboxPairs.Count
        Write-Log "Found $PairsCount mailbox pairs ready for migration" -Level "INFO"
        
        return $MailboxPairs
    }
    catch {
        Write-Log "Error finding eligible mailboxes: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

#==========================================================================
# Function to disable on-premises mailbox
#==========================================================================
function Disable-OnPremMailbox {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )
    
    Write-Log "Disabling on-premises mailbox: $Identity" -Level "INFO"
    
    try {
        # Get mailbox information before disabling
        $Mailbox = Get-OnpremMailbox -Identity $Identity -ErrorAction Stop
        
        # Disable the mailbox
        Disable-OnpremMailbox -Identity $Identity -Confirm:$false -ErrorAction Stop
        
        Write-Log "Successfully disabled on-premises mailbox: $Identity" -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Error disabling on-premises mailbox $Identity: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

#==========================================================================
# Function to execute mailbox restore request
#==========================================================================
function Start-MailboxRestoreRequest {
    param (
        [Parameter(Mandatory = $true)]
        [object]$MailboxPair,
        
        [Parameter(Mandatory = $false)]
        [bool]$AllowLargeItems = $DefaultAllowLargeItems,
        
        [Parameter(Mandatory = $false)]
        [int]$LargeItemLimit = $DefaultLargeItemLimit,
        
        [Parameter(Mandatory = $false)]
        [bool]$AcceptLargeDataLoss = $DefaultAcceptLargeDataLoss,
        
        [Parameter(Mandatory = $false)]
        [int]$BadItemLimit = $DefaultBadItemLimit
    )
    
    $OnPremIdentity = $MailboxPair.OnPremIdentity
    $CloudIdentity = $MailboxPair.CloudIdentity
    
    Write-Log "Starting mailbox restore request for: $OnPremIdentity -> $CloudIdentity" -Level "INFO"
    
    try {
        # First, disable the on-premises mailbox
        $DisableResult = Disable-OnPremMailbox -Identity $OnPremIdentity
        
        if (-not $DisableResult) {
            Write-Log "Failed to disable on-premises mailbox. Skipping restore request for $OnPremIdentity" -Level "ERROR"
            return $null
        }
        
        # Create parameters for New-MailboxRestoreRequest
        $RestoreParams = @{
            SourceStoreMailbox = $OnPremIdentity
            TargetMailbox = $CloudIdentity
            AllowLegacyDNMismatch = $true
            BadItemLimit = $BadItemLimit
        }
        
        # Add optional parameters if specified
        if ($AllowLargeItems) {
            $RestoreParams.Add("AllowLargeItems", $true)
            $RestoreParams.Add("LargeItemLimit", $LargeItemLimit)
        }
        
        if ($AcceptLargeDataLoss) {
            $RestoreParams.Add("AcceptLargeDataLoss", $true)
        }
        
        # Generate a unique name for the restore request
        $RequestName = "MIG_$($OnPremIdentity.Replace('@', '_').Replace('.', '_'))_$(Get-Date -Format 'yyyyMMddHHmmss')"
        $RestoreParams.Add("Name", $RequestName)
        
        # Execute the restore request
        $RestoreRequest = New-CloudMailboxRestoreRequest @RestoreParams -ErrorAction Stop
        
        if ($RestoreRequest) {
            Write-Log "Mailbox restore request created successfully. Request ID: $($RestoreRequest.Identity)" -Level "SUCCESS"
            
            # Update CustomAttribute6 to indicate migration has been initiated
            $CurrentDate = Get-Date -Format "yyyy-MM-dd"
            $AttributeValue = "DEL_MIG;STEP3;INITIATED;$CurrentDate"
            
            # Update cloud mailbox
            try {
                Set-CloudMailbox -Identity $CloudIdentity -CustomAttribute6 $AttributeValue -ErrorAction Stop
                Write-Log "Updated cloud mailbox $CloudIdentity CustomAttribute6 to $AttributeValue" -Level "SUCCESS"
            }
            catch {
                Write-Log "Error updating cloud mailbox $CloudIdentity: $($_.Exception.Message)" -Level "ERROR"
            }
            
            # Return the request information
            return [PSCustomObject]@{
                RequestId = $RestoreRequest.Identity
                RequestName = $RequestName
                OnPremIdentity = $OnPremIdentity
                CloudIdentity = $CloudIdentity
                Status = $RestoreRequest.Status
                StartTime = $RestoreRequest.RequestQueue
            }
        }
        else {
            Write-Log "Failed to create mailbox restore request for $OnPremIdentity" -Level "ERROR"
            return $null
        }
    }
    catch {
        Write-Log "Error creating mailbox restore request for $OnPremIdentity: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

#==========================================================================
# Function to verify and create required folders
#==========================================================================
function Initialize-Environment {
    Write-Log "Initializing environment and verifying paths" -Level "INFO"
    
    # Check and create base directory
    if (-not (Test-Path -Path $BasePath)) {
        try {
            New-Item -Path $BasePath -ItemType Directory -Force | Out-Null
            Write-Log "Base directory created: $BasePath" -Level "SUCCESS"
        }
        catch {
            Write-Log "Error creating base directory: $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    }
    
    # Check and create logs directory
    if (-not (Test-Path -Path $LogPath)) {
        try {
            New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
            Write-Log "Logs directory created: $LogPath" -Level "SUCCESS"
        }
        catch {
            Write-Log "Error creating logs directory: $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    }
    
    # Check and create CSV directory
    if (-not (Test-Path -Path $CSVPath)) {
        try {
            New-Item -Path $CSVPath -ItemType Directory -Force | Out-Null
            Write-Log "CSV directory created: $CSVPath" -Level "SUCCESS"
        }
        catch {
            Write-Log "Error creating CSV directory: $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    }
    
    # Check and create reports directory
    if (-not (Test-Path -Path $ReportsPath)) {
        try {
            New-Item -Path $ReportsPath -ItemType Directory -Force | Out-Null
            Write-Log "Reports directory created: $ReportsPath" -Level "SUCCESS"
        }
        catch {
            Write-Log "Error creating reports directory: $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    }
    
    return $true
}

#==========================================================================
# Main function
#==========================================================================
function Start-MailboxMigration {
    Write-Log "Starting mailbox migration process" -Level "INFO"
    Write-Log "Using default migration parameters:" -Level "INFO"
    Write-Log "- Allow large items: $DefaultAllowLargeItems" -Level "INFO"
    Write-Log "- Large item limit: $DefaultLargeItemLimit" -Level "INFO"
    Write-Log "- Accept large data loss: $DefaultAcceptLargeDataLoss" -Level "INFO"
    Write-Log "- Bad item limit: $DefaultBadItemLimit" -Level "INFO"
    
    # Find eligible mailboxes
    $MailboxPairs = Find-EligibleMailboxes
    
    if ($null -eq $MailboxPairs -or $MailboxPairs.Count -eq 0) {
        Write-Log "No eligible mailbox pairs found. Script ending." -Level "WARNING"
        return
    }
    
    # Statistics
    $TotalMailboxes = $MailboxPairs.Count
    $SuccessfulInitiations = 0
    $FailedInitiations = 0
    
    Write-Log "Found $TotalMailboxes mailbox pairs to migrate" -Level "INFO"
    
    # Initialize status report
    $StatusReport = @()
    
    # Process each mailbox pair
    foreach ($MailboxPair in $MailboxPairs) {
        # Start restore request with default parameters
        $RestoreRequest = Start-MailboxRestoreRequest -MailboxPair $MailboxPair
        
        if ($RestoreRequest) {
            $SuccessfulInitiations++
            Write-Log "Successfully initiated migration for $($MailboxPair.OnPremIdentity) -> $($MailboxPair.CloudIdentity)" -Level "SUCCESS"
            
            # Add to status report
            $StatusReport += [PSCustomObject]@{
                OnPremIdentity = $MailboxPair.OnPremIdentity
                CloudIdentity = $MailboxPair.CloudIdentity
                RequestId = $RestoreRequest.RequestId
                RequestName = $RestoreRequest.RequestName
                Status = "Initiated"
                StartTime = Get-Date
            }
        }
        else {
            $FailedInitiations++
            Write-Log "Failed to initiate migration for $($MailboxPair.OnPremIdentity)" -Level "ERROR"
            
            # Add to status report
            $StatusReport += [PSCustomObject]@{
                OnPremIdentity = $MailboxPair.OnPremIdentity
                CloudIdentity = $MailboxPair.CloudIdentity
                RequestId = "N/A"
                RequestName = "N/A"
                Status = "Failed"
                StartTime = Get-Date
            }
        }
    }
    
    # Export status report
    if ($StatusReport.Count -gt 0) {
        try {
            $StatusReport | Export-Csv -Path $StatusReportFile -NoTypeInformation -Force
            Write-Log "Exported migration initiation status report to $StatusReportFile" -Level "SUCCESS"
        }
        catch {
            Write-Log "Error exporting status report: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    # Display final statistics
    Write-Log "Migration initiation process complete. Final summary:" -Level "INFO"
    Write-Log "- Total mailboxes: $TotalMailboxes" -Level "INFO"
    Write-Log "- Successful initiations: $SuccessfulInitiations" -Level "SUCCESS"
    Write-Log "- Failed initiations: $FailedInitiations" -Level "WARNING"
    Write-Log "- Migration verification will be handled by a separate script" -Level "INFO"
}

#==========================================================================
# Script execution
#==========================================================================
try {
    # Initialize environment and verify paths
    $EnvReady = Initialize-Environment
    if (-not $EnvReady) {
        Write-Log "Failed to initialize environment. Script cannot continue." -Level "ERROR"
        exit 1
    }
    
    # Setup Exchange sessions with proper prefixes
    $SessionsReady = Connect-ExchangeSessions
    if (-not $SessionsReady) {
        Write-Log "Failed to set up required Exchange sessions. Script cannot continue." -Level "ERROR"
        exit 1
    }
    
    # Start mailbox migration process with default parameters
    Start-MailboxMigration
    
    Write-Log "Script completed successfully" -Level "SUCCESS"
}
catch {
    Write-Log "Unhandled error in script: $($_.Exception.Message)" -Level "ERROR"
    Write-Log "Error details: $($_.Exception.StackTrace)" -Level "ERROR"
    exit 1
}
