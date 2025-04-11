s imported with 'Onprem' prefix" -Level "SUCCESS"
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
