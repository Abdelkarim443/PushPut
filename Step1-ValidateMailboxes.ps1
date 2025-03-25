#==========================================================================
# Script: Step1-ValidateMailboxes.ps1
# Author: Manus
# Date: 03/25/2025
# Description: Script to validate mailboxes for migration to Exchange Online
#              - Imports CSV files with today's date
#              - Checks criteria (size < 40 GB, no archive, last logon < 5 years)
#              - Updates custom attribute 6 with migration status
#              - Moves processed CSV files to "Imported CSVs" folder
#==========================================================================

#==========================================================================
# Configuration - Predefined Paths
#==========================================================================
# Base paths - Update these paths to match your environment
$BasePath = "C:\ExchangeMigration"
$LogPath = "$BasePath\Logs"
$CSVPath = "$BasePath\CSV"
$ImportedCSVPath = "$BasePath\Imported CSVs"

# Log file and CSV pattern
$LogFile = Join-Path -Path $LogPath -ChildPath "Step1-ValidateMailboxes_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$Today = Get-Date -Format "yyyyMMdd"
$CSVPattern = "*$Today*.csv"

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
# Function to import CSV files
#==========================================================================
function Import-MigrationCSV {
    param (
        [string]$FolderPath,
        [string]$FilePattern
    )
    
    Write-Log "Searching for CSV files matching pattern '$FilePattern' in folder '$FolderPath'" -Level "INFO"
    
    try {
        # Create CSV folder if it doesn't exist
        if (-not (Test-Path -Path $FolderPath)) {
            try {
                New-Item -Path $FolderPath -ItemType Directory -Force | Out-Null
                Write-Log "CSV folder created: $FolderPath" -Level "INFO"
            }
            catch {
                Write-Log "Unable to create CSV folder: $($_.Exception.Message)" -Level "ERROR"
                return $null
            }
        }
        
        $CSVFiles = Get-ChildItem -Path $FolderPath -Filter $FilePattern -ErrorAction Stop
        
        if ($CSVFiles.Count -eq 0) {
            Write-Log "No CSV files found with today's date ($Today) in $FolderPath" -Level "WARNING"
            return $null
        }
        
        Write-Log "$($CSVFiles.Count) CSV file(s) found" -Level "INFO"
        
        $AllMailboxes = @()
        $ProcessedFiles = @()
        
        foreach ($CSVFile in $CSVFiles) {
            Write-Log "Importing file: $($CSVFile.FullName)" -Level "INFO"
            
            try {
                $Mailboxes = Import-Csv -Path $CSVFile.FullName -ErrorAction Stop
                $AllMailboxes += $Mailboxes
                $ProcessedFiles += $CSVFile
                Write-Log "Import successful: $($Mailboxes.Count) entries found in $($CSVFile.Name)" -Level "SUCCESS"
            }
            catch {
                Write-Log "Error importing file $($CSVFile.Name): $($_.Exception.Message)" -Level "ERROR"
            }
        }
        
        Write-Log "Total mailboxes imported: $($AllMailboxes.Count)" -Level "INFO"
        
        # Return both the mailboxes and the processed files
        return @{
            Mailboxes = $AllMailboxes
            Files = $ProcessedFiles
        }
    }
    catch {
        Write-Log "Error searching for CSV files: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

#==========================================================================
# Function to move processed CSV files
#==========================================================================
function Move-ProcessedCSVFiles {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$Files,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationFolder
    )
    
    Write-Log "Moving processed CSV files to: $DestinationFolder" -Level "INFO"
    
    # Create destination folder if it doesn't exist
    if (-not (Test-Path -Path $DestinationFolder)) {
        try {
            New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
            Write-Log "Imported CSVs folder created: $DestinationFolder" -Level "INFO"
        }
        catch {
            Write-Log "Unable to create Imported CSVs folder: $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    }
    
    $MovedCount = 0
    $ErrorCount = 0
    
    foreach ($File in $Files) {
        $DestinationPath = Join-Path -Path $DestinationFolder -ChildPath $File.Name
        
        # If file already exists in destination, add timestamp to filename
        if (Test-Path -Path $DestinationPath) {
            $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $NewFileName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name) + "_" + $Timestamp + [System.IO.Path]::GetExtension($File.Name)
            $DestinationPath = Join-Path -Path $DestinationFolder -ChildPath $NewFileName
        }
        
        try {
            Move-Item -Path $File.FullName -Destination $DestinationPath -Force -ErrorAction Stop
            Write-Log "Moved file: $($File.Name) to $DestinationPath" -Level "SUCCESS"
            $MovedCount++
        }
        catch {
            Write-Log "Error moving file $($File.Name): $($_.Exception.Message)" -Level "ERROR"
            $ErrorCount++
        }
    }
    
    Write-Log "File move operation complete: $MovedCount files moved, $ErrorCount errors" -Level "INFO"
    
    return ($ErrorCount -eq 0)
}

#==========================================================================
# Function to validate mailboxes
#==========================================================================
function Test-MailboxCriteria {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )
    
    Write-Log "Checking criteria for mailbox: $Identity" -Level "INFO"
    
    try {
        # Get mailbox information from On-premises Exchange
        $Mailbox = Get-OnpremMailbox -Identity $Identity -ErrorAction Stop
        $MailboxStats = Get-OnpremMailboxStatistics -Identity $Identity -ErrorAction Stop
        
        # Check mailbox size (< 40 GB)
        $MailboxSizeGB = [math]::Round(($MailboxStats.TotalItemSize.Value.ToBytes() / 1GB), 2)
        Write-Log "Mailbox size for $Identity: $MailboxSizeGB GB" -Level "INFO"
        
        if ($MailboxSizeGB -ge 40) {
            Write-Log "Mailbox $Identity exceeds 40 GB limit (Current size: $MailboxSizeGB GB)" -Level "WARNING"
            return @{
                Result = $false
                Reason = "Size exceeds 40 GB ($MailboxSizeGB GB)"
            }
        }
        
        # Check if archive is enabled
        if ($Mailbox.ArchiveStatus -eq "Active") {
            Write-Log "Mailbox $Identity has an active archive" -Level "WARNING"
            return @{
                Result = $false
                Reason = "Archive enabled"
            }
        }
        
        # Check last logon time (< 5 years)
        $LastLogonTime = $MailboxStats.LastLogonTime
        if ($null -eq $LastLogonTime) {
            Write-Log "Mailbox $Identity has no last logon time" -Level "WARNING"
            return @{
                Result = $false
                Reason = "No last logon time"
            }
        }
        
        $FiveYearsAgo = (Get-Date).AddYears(-5)
        if ($LastLogonTime -lt $FiveYearsAgo) {
            Write-Log "Last logon for mailbox $Identity is more than 5 years ago (Last logon: $LastLogonTime)" -Level "WARNING"
            return @{
                Result = $false
                Reason = "Last logon > 5 years ($LastLogonTime)"
            }
        }
        
        # All criteria met
        Write-Log "Mailbox $Identity meets all migration criteria" -Level "SUCCESS"
        return @{
            Result = $true
            Reason = "All criteria met"
        }
    }
    catch {
        Write-Log "Error checking criteria for $Identity: $($_.Exception.Message)" -Level "ERROR"
        return @{
            Result = $false
            Reason = "Error: $($_.Exception.Message)"
        }
    }
}

#==========================================================================
# Function to update custom attribute
#==========================================================================
function Update-CustomAttribute {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $true)]
        [bool]$IsEligible,
        
        [Parameter(Mandatory = $false)]
        [string]$Reason = ""
    )
    
    $CurrentDate = Get-Date -Format "yyyy-MM-dd"
    
    if ($IsEligible) {
        $AttributeValue = "DEL_MIG;STEP1;OK;$CurrentDate"
        $StatusText = "OK"
    }
    else {
        $AttributeValue = "DEL_MIG;STEP1;KO;$CurrentDate"
        $StatusText = "KO"
    }
    
    Write-Log "Updating custom attribute 6 for $Identity: $AttributeValue" -Level "INFO"
    
    try {
        Set-OnpremMailbox -Identity $Identity -CustomAttribute6 $AttributeValue -ErrorAction Stop
        Write-Log "Custom attribute 6 successfully updated for $Identity (Status: $StatusText)" -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Error updating custom attribute 6 for $Identity: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

#==========================================================================
# Main function
#==========================================================================
function Start-MailboxValidation {
    Write-Log "Starting mailbox validation process" -Level "INFO"
    
    # Import CSV files
    $ImportResult = Import-MigrationCSV -FolderPath $CSVPath -FilePattern $CSVPattern
    
    if ($null -eq $ImportResult -or $null -eq $ImportResult.Mailboxes -or $ImportResult.Mailboxes.Count -eq 0) {
        Write-Log "No mailboxes to process. Script ending." -Level "WARNING"
        return
    }
    
    $Mailboxes = $ImportResult.Mailboxes
    $ProcessedFiles = $ImportResult.Files
    
    # Statistics
    $TotalMailboxes = $Mailboxes.Count
    $EligibleMailboxes = 0
    $NonEligibleMailboxes = 0
    $ErrorMailboxes = 0
    
    # Process each mailbox
    foreach ($Mailbox in $Mailboxes) {
        # Check that identity property exists in CSV
        if (-not $Mailbox.Identity -and -not $Mailbox.UserPrincipalName -and -not $Mailbox.PrimarySmtpAddress -and -not $Mailbox.Alias) {
            Write-Log "CSV file does not contain a valid identity column (Identity, UserPrincipalName, PrimarySmtpAddress, or Alias)" -Level "ERROR"
            continue
        }
        
        # Determine identity to use
        $Identity = $Mailbox.Identity
        if ([string]::IsNullOrEmpty($Identity)) { $Identity = $Mailbox.UserPrincipalName }
        if ([string]::IsNullOrEmpty($Identity)) { $Identity = $Mailbox.PrimarySmtpAddress }
        if ([string]::IsNullOrEmpty($Identity)) { $Identity = $Mailbox.Alias }
        
        Write-Log "Processing mailbox: $Identity" -Level "INFO"
        
        # Validate criteria
        $ValidationResult = Test-MailboxCriteria -Identity $Identity
        
        # Update custom attribute
        $UpdateResult = Update-CustomAttribute -Identity $Identity -IsEligible $ValidationResult.Result -Reason $ValidationResult.Reason
        
        # Update statistics
        if ($ValidationResult.Result) {
            $EligibleMailboxes++
        }
        else {
            $NonEligibleMailboxes++
        }
        
        if (-not $UpdateResult) {
            $ErrorMailboxes++
        }
    }
    
    # Move processed CSV files to Imported CSVs folder
    if ($ProcessedFiles.Count -gt 0) {
        $MoveResult = Move-ProcessedCSVFiles -Files $ProcessedFiles -DestinationFolder $ImportedCSVPath
        if ($MoveResult) {
            Write-Log "Successfully moved all processed CSV files to the Imported CSVs folder" -Level "SUCCESS"
        }
        else {
            Write-Log "There were errors moving some CSV files to the Imported CSVs folder" -Level "WARNING"
        }
    }
    
    # Display final statistics
    Write-Log "Processing complete. Summary:" -Level "INFO"
    Write-Log "- Total mailboxes processed: $TotalMailboxes" -Level "INFO"
    Write-Log "- Eligible mailboxes: $EligibleMailboxes" -Level "SUCCESS"
    Write-Log "- Non-eligible mailboxes: $NonEligibleMailboxes" -Level "WARNING"
    Write-Log "- Update errors: $ErrorMailboxes" -Level "ERROR"
}

#==========================================================================
# Function to establish Exchange sessions
#==========================================================================
function Connect-ExchangeSessions {
    Write-Log "Setting up Exchange sessions" -Level "INFO"
    
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
                return $false
            }
        }
        
        # Create prefix for On-premises Exchange commands
        try {
            $OnpremSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
            if ($null -eq $OnpremSession) {
                Write-Log "No active Exchange session found. Please run this script in an Exchange Management Shell." -Level "ERROR"
                return $false
            }
            
            Import-PSSession $OnpremSession -Prefix "Onprem" -DisableNameChecking -AllowClobber | Out-Null
            Write-Log "On-premises Exchange commands imported with 'Onprem' prefix" -Level "SUCCESS"
        }
        catch {
            Write-Log "Error setting up On-premises Exchange session: $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    }
    
    return $true
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
    
    # Check and create Imported CSVs directory
    if (-not (Test-Path -Path $ImportedCSVPath)) {
        try {
            New-Item -Path $ImportedCSVPath -ItemType Directory -Force | Out-Null
            Write-Log "Imported CSVs directory created: $ImportedCSVPath" -Level "SUCCESS"
        }
        catch {
            Write-Log "Error creating Imported CSVs directory: $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    }
    
    return $true
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
    
    # Start validation process
    Start-MailboxValidation
    
    Write-Log "Script completed successfully" -Level "SUCCESS"
}
catch {
    Write-Log "Unhandled error in script: $($_.Exception.Message)" -Level "ERROR"
    Write-Log "Error details: $($_.Exception.StackTrace)" -Level "ERROR"
    exit 1
}
