# Exchange Migration to Exchange Online Scripts

## Description
This project contains a series of PowerShell scripts to automate the migration of mailboxes from on-premises Exchange to Exchange Online. The scripts are designed to work autonomously using CSV files as data sources.

## Script 1: Mailbox Validation (Step1-ValidateMailboxes.ps1)

### Features
- Automatically searches for and imports CSV files containing today's date in their name
- Checks the following criteria for each mailbox:
  - Size less than 40 GB
  - No archive enabled
  - Last logon less than 5 years ago
- Updates custom attribute 6 (CustomAttribute6) with the format "DEL_MIG;STEP1;OK;DATE" for eligible mailboxes
- Updates custom attribute 6 (CustomAttribute6) with the format "DEL_MIG;STEP1;KO;DATE" for non-eligible mailboxes
- Moves processed CSV files to the "Imported CSVs" folder after processing
- Generates detailed log files in the Logs folder

## Script 2: Create Cloud Shared Mailboxes (Step2-CreateCloudSharedMailboxes.ps1)

### Features
- Finds all mailboxes with CustomAttribute6 containing "OK" from Step 1
- Collects all attributes, especially custom attributes and extension custom attributes
- Creates shared mailboxes in Exchange Online with:
  - The same attributes as the on-premises mailbox
  - A primary SMTP address using the same local part but with the onmicrosoft.com domain
  - Configuration to not receive emails (hidden from address lists)
- Updates CustomAttribute6 on both on-premises and cloud mailboxes to track migration progress
- Exports collected attributes to CSV files for reference

## Script 3: Execute Mailbox Migration (Step3-ExecuteMailboxMigration.ps1)

### Features
- Finds mailboxes with CustomAttribute6 containing "STEP2;OK" from Step 2
- Executes New-MailboxRestoreRequest to migrate mailbox content from on-premises to cloud
- Updates CustomAttribute6 to "DEL_MIG;STEP3;INITIATED;DATE" to track migration initiation
- Does not verify completion (this will be handled by a separate script)
- Generates a report of initiated migrations
- Handles errors and provides comprehensive logging

### Migration Parameters
The script allows customization of several migration parameters:
- Allow large items (with configurable large item limit)
- Accept large data loss
- Bad item limit

### Prerequisites
- PowerShell 5.1 or higher
- Exchange Management Shell module for on-premises Exchange
- Exchange Online Management module for Exchange Online
- Exchange administrator rights for both environments
- Access to the mailboxes to be migrated

### Command Prefixes
The scripts use command prefixes to distinguish between Exchange platforms:
- "Onprem" prefix for Exchange On-premises commands (e.g., Get-OnpremMailbox)
- "Cloud" prefix for Exchange Online commands (e.g., Get-CloudMailbox)

### Directory Structure
The scripts use predefined paths that will be created automatically if they don't exist:
- Base directory: C:\ExchangeMigration
- Logs directory: C:\ExchangeMigration\Logs
- CSV directory: C:\ExchangeMigration\CSV
- Imported CSVs directory: C:\ExchangeMigration\Imported CSVs
- Attributes export directory: C:\ExchangeMigration\AttributesExport
- Reports directory: C:\ExchangeMigration\Reports

You can modify these paths at the beginning of each script to match your environment.

### Usage
1. First run Step1-ValidateMailboxes.ps1 to validate mailboxes
   - CSV files with today's date will be processed and moved to the "Imported CSVs" folder
2. Then run Step2-CreateCloudSharedMailboxes.ps1 to create cloud shared mailboxes
3. Run Step3-ExecuteMailboxMigration.ps1 to initiate the mailbox migration:
   ```powershell
   .\Step3-ExecuteMailboxMigration.ps1
   ```
4. When prompted, enter the migration parameters or accept the defaults
5. Use a separate script (to be created) to verify migration completion

### Logging
Logs are created in the Logs directory (C:\ExchangeMigration\Logs by default) with the format:
```
Step1-ValidateMailboxes_YYYYMMDD_HHMMSS.log
Step2-CreateCloudSharedMailboxes_YYYYMMDD_HHMMSS.log
Step3-ExecuteMailboxMigration_YYYYMMDD_HHMMSS.log
```

### Migration Reports
Migration initiation reports are created in the Reports directory (C:\ExchangeMigration\Reports by default) with the format:
```
MigrationInitiated_YYYYMMDD_HHMMSS.csv
```
