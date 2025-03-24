#==========================================================================
# Script: Step3-ExecuteMailboxMigration.ps1
# Author: Manus
# Date: 03/21/2025
# Description: Script to execute mailbox migration via Mailbox Restore Request
#              - Finds mailboxes with CustomAttribute6 containing "STEP2;OK" from Step 2
#              - Executes New-MailboxRestoreRequest to migrate mailbox content
#              - Monitors restore request status
#              - Updates tracking attributes after successful migration
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
$StatusReportFile = Join-Path -Path $ReportsPath -ChildPath "MigrationStatus_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

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
                    # Prompt for credentials if needed
                    $CloudCredential = Get-Credential -Message "Enter your Exchange Online admin credentials"
                    Connect-ExchangeOnline -Credential $CloudCredential -ShowBanner:$false -ErrorAction Stop
                    
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
# Function to execute mailbox restore request
#==========================================================================
function Start-MailboxRestoreRequest {
    param (
        [Parameter(Mandatory = $true)]
        [object]$MailboxPair,
        
        [Parameter(Mandatory = $false)]
        [switch]$AllowLargeItems,
        
        [Parameter(Mandatory = $false)]
        [int]$LargeItemLimit = 50,
        
        [Parameter(Mandatory = $false)]
        [switch]$AcceptLargeDataLoss,
        
        [Parameter(Mandatory = $false)]
        [int]$BadItemLimit = 10
    )
    
    $OnPremIdentity = $MailboxPair.OnPremIdentity
    $CloudIdentity = $MailboxPair.CloudIdentity
    
    Write-Log "Starting mailbox restore request for: $OnPremIdentity -> $CloudIdentity" -Level "INFO"
    
    try {
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
# Function to monitor restore request status
#==========================================================================
function Get-RestoreRequestStatus {
    param (
        [Parameter(Mandatory = $true)]
        [string]$RequestId,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 5,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryIntervalSeconds = 60
    )
    
    Write-Log "Checking status for restore request: $RequestId" -Level "INFO"
    
    $RetryCount = 0
    $Status = $null
    
    do {
        try {
            $RestoreRequest = Get-CloudMailboxRestoreRequest -Identity $RequestId -ErrorAction Stop
            $Status = $RestoreRequest.Status
            $PercentComplete = $RestoreRequest.PercentComplete
            
            Write-Log "Restore request $RequestId status: $Status ($PercentComplete% complete)" -Level "INFO"
            
            # If the request is completed or failed, return the status
            if ($Status -eq "Completed" -or $Status -eq "Failed" -or $Status -eq "CompletedWithWarning") {
                return [PSCustomObject]@{
                    RequestId = $RequestId
                    Status = $Status
                    PercentComplete = $PercentComplete
                    CompletionTime = $RestoreRequest.CompletionTimestamp
                    FailureCode = $RestoreRequest.FailureCode
                    FailureType = $RestoreRequest.FailureType
                    Message = $RestoreRequest.Message
                }
            }
            
            # If the request is still in progress, wait and retry
            if ($RetryCount -lt $MaxRetries) {
                Write-Log "Restore request still in progress. Waiting $RetryIntervalSeconds seconds before checking again..." -Level "INFO"
                Start-Sleep -Seconds $RetryIntervalSeconds
            }
            
            $RetryCount++
        }
        catch {
            Write-Log "Error checking restore request status: $($_.Exception.Message)" -Level "ERROR"
            $RetryCount++
            
            if ($RetryCount -lt $MaxRetries) {
                Write-Log "Retrying in $RetryIntervalSeconds seconds..." -Level "INFO"
                Start-Sleep -Seconds $RetryIntervalSeconds
            }
        }
    } while ($RetryCount -lt $MaxRetries)
    
    Write-Log "Maximum retries reached. Unable to determine final status for restore request $RequestId" -Level "WARNING"
    
    # Return the last known status
    return [PSCustomObject]@{
        RequestId = $RequestId
        Status = $Status
        PercentComplete = $PercentComplete
        CompletionTime = $null
        FailureCode = "Unknown"
        FailureType = "Unknown"
        Message = "Maximum retries reached. Unable to determine final status."
    }
}

#==========================================================================
# Function to update status after migration
#==========================================================================
function Update-MigrationStatus {
    param (
        [Parameter(Mandatory = $true)]
        [string]$OnPremIdentity,
        
        [Parameter(Mandatory = $true)]
        [string]$CloudIdentity,
        
        [Parameter(Mandatory = $true)]
        [string]$Status,
        
        [Parameter(Mandatory = $false)]
        [string]$Message = ""
    )
    
    $CurrentDate = Get-Date -Format "yyyy-MM-dd"
    
    if ($Status -eq "Completed" -or $Status -eq "CompletedWithWarning") {
        $AttributeValue = "DEL_MIG;STEP3;OK;$CurrentDate"
        $StatusText = "OK"
    }
    else {
        $AttributeValue = "DEL_MIG;STEP3;KO;$CurrentDate"
        $StatusText = "KO"
    }
    
    Write-Log "Updating status for $OnPremIdentity -> $CloudIdentity: $StatusText" -Level "INFO"
    
    # Update on-premises mailbox
    try {
        Set-OnpremMailbox -Identity $OnPremIdentity -CustomAttribute6 $AttributeValue -ErrorAction Stop
        Write-Log "Updated on-premises mailbox $OnPremIdentity CustomAttribute6 to $AttributeValue" -Level "SUCCESS"
    }
    catch {
        Write-Log "Error updating on-premises mailbox $OnPremIdentity: $($_.Exception.Message)" -Level "ERROR"
    }
    
    # Update cloud mailbox
    try {
        Set-CloudMailbox -Identity $CloudIdentity -CustomAttribute6 $AttributeValue -ErrorAction Stop
        Write-Log "Updated cloud mailbox $CloudIdentity CustomAttribute6 to $AttributeValue" -Level "SUCCESS"
    }
    catch {
        Write-Log "Error updating cloud mailbox $CloudIdentity: $($_.Exception.Message)" -Level "ERROR"
    }
    
    # Return status object for reporting
    return [PSCustomObject]@{
        OnPremIdentity = $OnPremIdentity
        CloudIdentity = $CloudIdentity
        Status = $Status
        StatusValue = $StatusText
        CompletionDate = $CurrentDate
        Message = $Message
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
    param (
        [Parameter(Mandatory = $false)]
        [switch]$AllowLargeItems,
        
        [Parameter(Mandatory = $false)]
        [int]$LargeItemLimit = 50,
        
        [Parameter(Mandatory = $false)]
        [switch]$AcceptLargeDataLoss,
        
        [Parameter(Mandatory = $false)]
        [int]$BadItemLimit = 10,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxConcurrentMigrations = 10,
        
        [Parameter(Mandatory = $false)]
        [int]$StatusCheckIntervalMinutes = 15
    )
    
    Write-Log "Starting mailbox migration process" -Level "INFO"
    
    # Find eligible mailboxes
    $MailboxPairs = Find-EligibleMailboxes
    
    if ($null -eq $MailboxPairs -or $MailboxPairs.Count -eq 0) {
        Write-Log "No eligible mailbox pairs found. Script ending." -Level "WARNING"
        return
    }
    
    # Statistics
    $TotalMailboxes = $MailboxPairs.Count
    $SuccessfulMigrations = 0
    $FailedMigrations = 0
    $InProgressMigrations = 0
    
    Write-Log "Found $TotalMailboxes mailbox pairs to migrate" -Level "INFO"
    
    # Create restore requests
    $RestoreRequests = @()
    $BatchCount = [Math]::Min($TotalMailboxes, $MaxConcurrentMigrations)
    
    Write-Log "Creating restore requests for first batch of $BatchCount mailboxes" -Level "INFO"
    
    for ($i = 0; $i -lt $BatchCount; $i++) {
        $MailboxPair = $MailboxPairs[$i]
        
        # Create restore request parameters
        $RestoreParams = @{
            MailboxPair = $MailboxPair
            BadItemLimit = $BadItemLimit
        }
        
        if ($AllowLargeItems) {
            $RestoreParams.Add("AllowLargeItems", $true)
            $RestoreParams.Add("LargeItemLimit", $LargeItemLimit)
        }
        
        if ($AcceptLargeDataLoss) {
            $RestoreParams.Add("AcceptLargeDataLoss", $true)
        }
        
        # Start restore request
        $RestoreRequest = Start-MailboxRestoreRequest @RestoreParams
        
        if ($RestoreRequest) {
            $RestoreRequests += $RestoreRequest
            $InProgressMigrations++
            Write-Log "Started migration for $($MailboxPair.OnPremIdentity) -> $($MailboxPair.CloudIdentity)" -Level "SUCCESS"
        }
        else {
            $FailedMigrations++
            Write-Log "Failed to start migration for $($MailboxPair.OnPremIdentity)" -Level "ERROR"
            
            # Update status for failed migration
            $StatusResult = Update-MigrationStatus -OnPremIdentity $MailboxPair.OnPremIdentity -CloudIdentity $MailboxPair.CloudIdentity -Status "Failed" -Message "Failed to create restore request"
        }
    }
    
    # Initialize status report
    $StatusReport = @()
    
    # Process remaining mailboxes as current ones complete
    $NextMailboxIndex = $BatchCount
    
    while ($RestoreRequests.Count -gt 0) {
        Write-Log "Checking status of $($RestoreRequests.Count) active restore requests" -Level "INFO"
        
        # Wait for the specified interval
        Start-Sleep -Seconds ($StatusCheckIntervalMinutes * 60)
        
        # Check status of each request
        $CompletedRequests = @()
        
        foreach ($Request in $RestoreRequests) {
            $Status = Get-RestoreRequestStatus -RequestId $Request.RequestId
            
            if ($Status.Status -eq "Completed" -or $Status.Status -eq "Failed" -or $Status.Status -eq "CompletedWithWarning") {
                Write-Log "Migration completed for request $($Request.RequestId) with status: $($Status.Status)" -Level "INFO"
                
                # Update status
                $StatusResult = Update-MigrationStatus -OnPremIdentity $Request.OnPremIdentity -CloudIdentity $Request.CloudIdentity -Status $Status.Status -Message $Status.Message
                
                # Add to status report
                $StatusReport += [PSCustomObject]@{
                    OnPremIdentity = $Request.OnPremIdentity
                    CloudIdentity = $Request.CloudIdentity
                    RequestId = $Request.RequestId
                    Status = $Status.Status
                    PercentComplete = $Status.PercentComplete
                    CompletionTime = $Status.CompletionTime
                    FailureCode = $Status.FailureCode
                    FailureType = $Status.FailureType
                    Message = $Status.Message
                }
                
                # Update statistics
                if ($Status.Status -eq "Completed" -or $Status.Status -eq "CompletedWithWarning") {
                    $SuccessfulMigrations++
                }
                else {
                    $FailedMigrations++
                }
                
                $InProgressMigrations--
                
                # Mark request as completed
                $CompletedRequests += $Request
                
                # Start a new migration if there are more mailboxes to process
                if ($NextMailboxIndex -lt $TotalMailboxes) {
                    $MailboxPair = $MailboxPairs[$NextMailboxIndex]
                    
                    # Create restore request parameters
                    $RestoreParams = @{
                        MailboxPair = $MailboxPair
                        BadItemLimit = $BadItemLimit
                    }
                    
                    if ($AllowLargeItems) {
                        $RestoreParams.Add("AllowLargeItems", $true)
                        $RestoreParams.Add("LargeItemLimit", $LargeItemLimit)
                    }
                    
                    if ($AcceptLargeDataLoss) {
                        $RestoreParams.Add("AcceptLargeDataLoss", $true)
                    }
                    
                    # Start restore request
                    $NewRestoreRequest = Start-MailboxRestoreRequest @RestoreParams
                    
                    if ($NewRestoreRequest) {
                        $RestoreRequests += $NewRestoreRequest
                        $InProgressMigrations++
                        Write-Log "Started migration for $($MailboxPair.OnPremIdentity) -> $($MailboxPair.CloudIdentity)" -Level "SUCCESS"
                    }
                    else {
                        $FailedMigrations++
                        Write-Log "Failed to start migration for $($MailboxPair.OnPremIdentity)" -Level "ERROR"
                        
                        # Update status for failed migration
                        $StatusResult = Update-MigrationStatus -OnPremIdentity $MailboxPair.OnPremIdentity -CloudIdentity $MailboxPair.CloudIdentity -Status "Failed" -Message "Failed to create restore request"
                        
                        # Add to status report
                        $StatusReport += [PSCustomObject]@{
                            OnPremIdentity = $MailboxPair.OnPremIdentity
                            CloudIdentity = $MailboxPair.CloudIdentity
                            RequestId = "N/A"
                            Status = "Failed"
                            PercentComplete = 0
                            CompletionTime = Get-Date
                            FailureCode = "RequestCreationFailed"
                            FailureType = "Error"
                            Message = "Failed to create restore request"
                        }
                    }
                    
                    $NextMailboxIndex++
                }
            }
            else {
                Write-Log "Migration in progress for request $($Request.RequestId): $($Status.Status) ($($Status.PercentComplete)% complete)" -Level "INFO"
            }
        }
        
        # Remove completed requests from active list
        foreach ($CompletedRequest in $CompletedRequests) {
            $RestoreRequests = $RestoreRequests | Where-Object { $_.RequestId -ne $CompletedRequest.RequestId }
        }
        
        # Display current statistics
        Write-Log "Current migration status:" -Level "INFO"
        Write-Log "- Total mailboxes: $TotalMailboxes" -Level "INFO"
        Write-Log "- Successful migrations: $SuccessfulMigrations" -Level "SUCCESS"
        Write-Log "- Failed migrations: $FailedMigrations" -Level "WARNING"
        Write-Log "- In progress migrations: $InProgressMigrations" -Level "INFO"
        Write-Log "- Remaining mailboxes: $($TotalMailboxes - $SuccessfulMigrations - $FailedMigrations - $InProgressMigrations)" -Level "INFO"
    }
    
    # Export status report
    if ($StatusReport.Count -gt 0) {
        try {
            $StatusReport | Export-Csv -Path $StatusReportFile -NoTypeInformation -Force
            Write-Log "Exported migration status report to $StatusReportFile" -Level "SUCCESS"
        }
        catch {
            Write-Log "Error exporting status report: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    # Display final statistics
    Write-Log "Migration process complete. Final summary:" -Level "INFO"
    Write-Log "- Total mailboxes: $TotalMailboxes" -Level "INFO"
    Write-Log "- Successful migrations: $SuccessfulMigrations" -Level "SUCCESS"
    Write-Log "- Failed migrations: $FailedMigrations" -Level "WARNING"
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
    
    # Prompt for migration parameters
    $AllowLargeItems = $false
    $AcceptLargeDataLoss = $false
    $BadItemLimit = 10
    $LargeItemLimit = 50
    $MaxConcurrentMigrations = 10
    $StatusCheckIntervalMinutes = 15
    
    $AllowLargeItemsResponse = Read-Host "Allow large items during migration? (y/n, default: n)"
    if ($AllowLargeItemsResponse -eq "y") {
        $AllowLargeItems = $true
        $LargeItemLimitResponse = Read-Host "Large item limit (default: 50)"
        if ($LargeItemLimitResponse -match "^\d+$") {
            $LargeItemLimit = [int]$LargeItemLimitResponse
        }
    }
    
    $AcceptLargeDataLossResponse = Read-Host "Accept large data loss during migration? (y/n, default: n)"
    if ($AcceptLargeDataLossResponse -eq "y") {
        $AcceptLargeDataLoss = $true
    }
    
    $BadItemLimitResponse = Read-Host "Bad item limit (default: 10)"
    if ($BadItemLimitResponse -match "^\d+$") {
        $BadItemLimit = [int]$BadItemLimitResponse
    }
    
    $MaxConcurrentMigrationsResponse = Read-Host "Maximum concurrent migrations (default: 10)"
    if ($MaxConcurrentMigrationsResponse -match "^\d+$") {
        $MaxConcurrentMigrations = [int]$MaxConcurrentMigrationsResponse
    }
    
    $StatusCheckIntervalResponse = Read-Host "Status check interval in minutes (default: 15)"
    if ($StatusCheckIntervalResponse -match "^\d+$") {
        $StatusCheckIntervalMinutes = [int]$StatusCheckIntervalResponse
    }
    
    # Start mailbox migration process
    $MigrationParams = @{
        BadItemLimit = $BadItemLimit
        MaxConcurrentMigrations = $MaxConcurrentMigrations
        StatusCheckIntervalMinutes = $StatusCheckIntervalMinutes
    }
    
    if ($AllowLargeItems) {
        $MigrationParams.Add("AllowLargeItems", $true)
        $MigrationParams.Add("LargeItemLimit", $LargeItemLimit)
    }
    
    if ($AcceptLargeDataLoss) {
        $MigrationParams.Add("AcceptLargeDataLoss", $true)
    }
    
    Start-MailboxMigration @MigrationParams
    
    Write-Log "Script completed successfully" -Level "SUCCESS"
}
catch {
    Write-Log "Unhandled error in script: $($_.Exception.Message)" -Level "ERROR"
    Write-Log "Error details: $($_.Exception.StackTrace)" -Level "ERROR"
    exit 1
}
