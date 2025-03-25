#==========================================================================
# Script: Step2-CreateCloudSharedMailboxes.ps1
# Author: Manus
# Date: 03/25/2025
# Description: Script to create shared mailboxes in Exchange Online for migration
#              - Finds mailboxes with CustomAttribute6 containing "OK" from Step 1
#              - Collects all attributes, especially custom attributes
#              - Creates shared mailboxes in Exchange Online with onmicrosoft.com domain
#==========================================================================

#==========================================================================
# Configuration - Predefined Paths
#==========================================================================
# Base paths - Update these paths to match your environment
$BasePath = "C:\ExchangeMigration"
$LogPath = "$BasePath\Logs"
$CSVPath = "$BasePath\CSV"

# Log file
$LogFile = Join-Path -Path $LogPath -ChildPath "Step2-CreateCloudSharedMailboxes_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Export path for collected attributes
$AttributesExportPath = Join-Path -Path $BasePath -ChildPath "AttributesExport"

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
    Write-Log "Finding eligible mailboxes with CustomAttribute6 containing 'OK'" -Level "INFO"
    
    try {
        # Find mailboxes with CustomAttribute6 containing "OK"
        $EligibleMailboxes = Get-OnpremMailbox -ResultSize Unlimited | 
                            Where-Object { $_.CustomAttribute6 -like "*DEL_MIG;STEP1;OK;*" }
        
        $Count = $EligibleMailboxes.Count
        
        if ($Count -eq 0) {
            Write-Log "No eligible mailboxes found with CustomAttribute6 containing 'OK'" -Level "WARNING"
            return $null
        }
        
        Write-Log "Found $Count eligible mailboxes" -Level "SUCCESS"
        return $EligibleMailboxes
    }
    catch {
        Write-Log "Error finding eligible mailboxes: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

#==========================================================================
# Function to collect mailbox attributes
#==========================================================================
function Get-MailboxAttributes {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Mailbox
    )
    
    $Identity = $Mailbox.Identity
    Write-Log "Collecting attributes for mailbox: $Identity" -Level "INFO"
    
    try {
        # Get detailed mailbox information
        $DetailedMailbox = Get-OnpremMailbox -Identity $Identity -ErrorAction Stop
        
        # Create attributes object
        $Attributes = [PSCustomObject]@{
            Identity = $Identity
            DisplayName = $DetailedMailbox.DisplayName
            Alias = $DetailedMailbox.Alias
            PrimarySmtpAddress = $DetailedMailbox.PrimarySmtpAddress
            EmailAddresses = $DetailedMailbox.EmailAddresses
            CustomAttribute1 = $DetailedMailbox.CustomAttribute1
            CustomAttribute2 = $DetailedMailbox.CustomAttribute2
            CustomAttribute3 = $DetailedMailbox.CustomAttribute3
            CustomAttribute4 = $DetailedMailbox.CustomAttribute4
            CustomAttribute5 = $DetailedMailbox.CustomAttribute5
            CustomAttribute6 = $DetailedMailbox.CustomAttribute6
            CustomAttribute7 = $DetailedMailbox.CustomAttribute7
            CustomAttribute8 = $DetailedMailbox.CustomAttribute8
            CustomAttribute9 = $DetailedMailbox.CustomAttribute9
            CustomAttribute10 = $DetailedMailbox.CustomAttribute10
            CustomAttribute11 = $DetailedMailbox.CustomAttribute11
            CustomAttribute12 = $DetailedMailbox.CustomAttribute12
            CustomAttribute13 = $DetailedMailbox.CustomAttribute13
            CustomAttribute14 = $DetailedMailbox.CustomAttribute14
            CustomAttribute15 = $DetailedMailbox.CustomAttribute15
            ExtensionCustomAttribute1 = $DetailedMailbox.ExtensionCustomAttribute1
            ExtensionCustomAttribute2 = $DetailedMailbox.ExtensionCustomAttribute2
            ExtensionCustomAttribute3 = $DetailedMailbox.ExtensionCustomAttribute3
            ExtensionCustomAttribute4 = $DetailedMailbox.ExtensionCustomAttribute4
            ExtensionCustomAttribute5 = $DetailedMailbox.ExtensionCustomAttribute5
            HiddenFromAddressListsEnabled = $DetailedMailbox.HiddenFromAddressListsEnabled
            RecipientTypeDetails = $DetailedMailbox.RecipientTypeDetails
            UMDtmfMap = $DetailedMailbox.UMDtmfMap
            ExchangeGuid = $DetailedMailbox.ExchangeGuid
            LegacyExchangeDN = $DetailedMailbox.LegacyExchangeDN
            MaxSendSize = $DetailedMailbox.MaxSendSize
            MaxReceiveSize = $DetailedMailbox.MaxReceiveSize
            OfflineAddressBook = $DetailedMailbox.OfflineAddressBook
            AddressBookPolicy = $DetailedMailbox.AddressBookPolicy
            RetentionPolicy = $DetailedMailbox.RetentionPolicy
            SharingPolicy = $DetailedMailbox.SharingPolicy
            EmailAddressPolicyEnabled = $DetailedMailbox.EmailAddressPolicyEnabled
        }
        
        # Export attributes to CSV for reference
        if (-not (Test-Path -Path $AttributesExportPath)) {
            New-Item -Path $AttributesExportPath -ItemType Directory -Force | Out-Null
            Write-Log "Created attributes export directory: $AttributesExportPath" -Level "INFO"
        }
        
        $ExportFile = Join-Path -Path $AttributesExportPath -ChildPath "$($Identity.Replace('@', '_').Replace('.', '_'))_attributes.csv"
        $Attributes | Export-Csv -Path $ExportFile -NoTypeInformation -Force
        Write-Log "Exported attributes for $Identity to $ExportFile" -Level "INFO"
        
        return $Attributes
    }
    catch {
        Write-Log "Error collecting attributes for $Identity: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

#==========================================================================
# Function to create cloud shared mailbox
#==========================================================================
function New-CloudSharedMailbox {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Attributes,
        
        [Parameter(Mandatory = $true)]
        [string]$OnMicrosoftDomain
    )
    
    $Identity = $Attributes.Identity
    Write-Log "Creating cloud shared mailbox for: $Identity" -Level "INFO"
    
    try {
        # Extract username from primary SMTP address
        $Username = ($Attributes.PrimarySmtpAddress -split '@')[0]
        
        # Create new primary SMTP address with onmicrosoft.com domain
        $NewPrimarySmtpAddress = "$Username@$OnMicrosoftDomain"
        
        Write-Log "New primary SMTP address will be: $NewPrimarySmtpAddress" -Level "INFO"
        
        # Check if mailbox already exists in Exchange Online
        $ExistingMailbox = Get-CloudMailbox -Identity $NewPrimarySmtpAddress -ErrorAction SilentlyContinue
        
        if ($ExistingMailbox) {
            Write-Log "Mailbox $NewPrimarySmtpAddress already exists in Exchange Online" -Level "WARNING"
            return $false
        }
        
        # Create new shared mailbox in Exchange Online
        $NewMailbox = New-CloudMailbox -Shared -Name $Attributes.DisplayName -DisplayName $Attributes.DisplayName -Alias $Attributes.Alias -PrimarySmtpAddress $NewPrimarySmtpAddress
        
        if (-not $NewMailbox) {
            Write-Log "Failed to create shared mailbox $NewPrimarySmtpAddress" -Level "ERROR"
            return $false
        }
        
        Write-Log "Created shared mailbox: $NewPrimarySmtpAddress" -Level "SUCCESS"
        
        # Configure the mailbox to not receive emails
        Set-CloudMailbox -Identity $NewPrimarySmtpAddress -HiddenFromAddressListsEnabled $true -ErrorAction Stop
        Write-Log "Configured mailbox to be hidden from address lists" -Level "SUCCESS"
        
        # Set custom attributes
        $SetMailboxParams = @{
            Identity = $NewPrimarySmtpAddress
            CustomAttribute1 = $Attributes.CustomAttribute1
            CustomAttribute2 = $Attributes.CustomAttribute2
            CustomAttribute3 = $Attributes.CustomAttribute3
            CustomAttribute4 = $Attributes.CustomAttribute4
            CustomAttribute5 = $Attributes.CustomAttribute5
            CustomAttribute6 = "DEL_MIG;STEP2;OK;$(Get-Date -Format 'yyyy-MM-dd')"
            CustomAttribute7 = $Attributes.CustomAttribute7
            CustomAttribute8 = $Attributes.CustomAttribute8
            CustomAttribute9 = $Attributes.CustomAttribute9
            CustomAttribute10 = $Attributes.CustomAttribute10
            CustomAttribute11 = $Attributes.CustomAttribute11
            CustomAttribute12 = $Attributes.CustomAttribute12
            CustomAttribute13 = $Attributes.CustomAttribute13
            CustomAttribute14 = $Attributes.CustomAttribute14
            CustomAttribute15 = $Attributes.CustomAttribute15
        }
        
        Set-CloudMailbox @SetMailboxParams -ErrorAction Stop
        Write-Log "Set custom attributes for $NewPrimarySmtpAddress" -Level "SUCCESS"
        
        # Set extension custom attributes if they exist
        if ($Attributes.ExtensionCustomAttribute1 -or 
            $Attributes.ExtensionCustomAttribute2 -or 
            $Attributes.ExtensionCustomAttribute3 -or 
            $Attributes.ExtensionCustomAttribute4 -or 
            $Attributes.ExtensionCustomAttribute5) {
            
            $ExtensionParams = @{
                Identity = $NewPrimarySmtpAddress
            }
            
            if ($Attributes.ExtensionCustomAttribute1) { $ExtensionParams.Add("ExtensionCustomAttribute1", $Attributes.ExtensionCustomAttribute1) }
            if ($Attributes.ExtensionCustomAttribute2) { $ExtensionParams.Add("ExtensionCustomAttribute2", $Attributes.ExtensionCustomAttribute2) }
            if ($Attributes.ExtensionCustomAttribute3) { $ExtensionParams.Add("ExtensionCustomAttribute3", $Attributes.ExtensionCustomAttribute3) }
            if ($Attributes.ExtensionCustomAttribute4) { $ExtensionParams.Add("ExtensionCustomAttribute4", $Attributes.ExtensionCustomAttribute4) }
            if ($Attributes.ExtensionCustomAttribute5) { $ExtensionParams.Add("ExtensionCustomAttribute5", $Attributes.ExtensionCustomAttribute5) }
            
            Set-CloudMailbox @ExtensionParams -ErrorAction Stop
            Write-Log "Set extension custom attributes for $NewPrimarySmtpAddress" -Level "SUCCESS"
        }
        
        # Update on-premises mailbox to indicate Step 2 is complete
        Set-OnpremMailbox -Identity $Identity -CustomAttribute6 "DEL_MIG;STEP2;OK;$(Get-Date -Format 'yyyy-MM-dd')" -ErrorAction Stop
        Write-Log "Updated on-premises mailbox $Identity CustomAttribute6 to indicate Step 2 is complete" -Level "SUCCESS"
        
        return $true
    }
    catch {
        Write-Log "Error creating cloud shared mailbox for $Identity: $($_.Exception.Message)" -Level "ERROR"
        return $false
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
    
    # Check and create attributes export directory
    if (-not (Test-Path -Path $AttributesExportPath)) {
        try {
            New-Item -Path $AttributesExportPath -ItemType Directory -Force | Out-Null
            Write-Log "Attributes export directory created: $AttributesExportPath" -Level "SUCCESS"
        }
        catch {
            Write-Log "Error creating attributes export directory: $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    }
    
    return $true
}

#==========================================================================
# Main function
#==========================================================================
function Start-CloudMailboxCreation {
    param (
        [Parameter(Mandatory = $true)]
        [string]$OnMicrosoftDomain
    )
    
    Write-Log "Starting cloud shared mailbox creation process" -Level "INFO"
    
    # Find eligible mailboxes
    $EligibleMailboxes = Find-EligibleMailboxes
    
    if ($null -eq $EligibleMailboxes -or $EligibleMailboxes.Count -eq 0) {
        Write-Log "No eligible mailboxes found. Script ending." -Level "WARNING"
        return
    }
    
    # Statistics
    $TotalMailboxes = $EligibleMailboxes.Count
    $SuccessfulMailboxes = 0
    $FailedMailboxes = 0
    
    # Process each mailbox
    foreach ($Mailbox in $EligibleMailboxes) {
        $Identity = $Mailbox.Identity
        Write-Log "Processing mailbox: $Identity" -Level "INFO"
        
        # Collect attributes
        $Attributes = Get-MailboxAttributes -Mailbox $Mailbox
        
        if ($null -eq $Attributes) {
            Write-Log "Failed to collect attributes for $Identity. Skipping." -Level "ERROR"
            $FailedMailboxes++
            continue
        }
        
        # Create cloud shared mailbox
        $Result = New-CloudSharedMailbox -Attributes $Attributes -OnMicrosoftDomain $OnMicrosoftDomain
        
        if ($Result) {
            $SuccessfulMailboxes++
        }
        else {
            $FailedMailboxes++
        }
    }
    
    # Display final statistics
    Write-Log "Processing complete. Summary:" -Level "INFO"
    Write-Log "- Total mailboxes processed: $TotalMailboxes" -Level "INFO"
    Write-Log "- Successful mailbox creations: $SuccessfulMailboxes" -Level "SUCCESS"
    Write-Log "- Failed mailbox creations: $FailedMailboxes" -Level "WARNING"
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
    
    # Prompt for onmicrosoft.com domain
    $OnMicrosoftDomain = Read-Host "Enter your onmicrosoft.com domain (e.g., contoso.onmicrosoft.com)"
    
    if ([string]::IsNullOrEmpty($OnMicrosoftDomain) -or -not $OnMicrosoftDomain.EndsWith("onmicrosoft.com")) {
        Write-Log "Invalid onmicrosoft.com domain provided. Domain must end with 'onmicrosoft.com'" -Level "ERROR"
        exit 1
    }
    
    # Start cloud mailbox creation process
    Start-CloudMailboxCreation -OnMicrosoftDomain $OnMicrosoftDomain
    
    Write-Log "Script completed successfully" -Level "SUCCESS"
}
catch {
    Write-Log "Unhandled error in script: $($_.Exception.Message)" -Level "ERROR"
    Write-Log "Error details: $($_.Exception.StackTrace)" -Level "ERROR"
    exit 1
}
