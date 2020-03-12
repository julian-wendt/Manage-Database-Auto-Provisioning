#Requires -Version 4.0

<#  
.SYNOPSIS
    Exclude/include mailbox databases from Exchange Server Database auto provisioning.

.DESCRIPTION
    Script to automatically exclude/include Exchange Server Mailbox Databases based on the
    available disk space and the database whitespace from the mailbox provisioning load balancer
    that distributes new mailboxes randomly and evenly across the available databases,
    based on the available disk space and database whitespace.

.NOTES
    Author     : Julian Wendt
    Version    : 1.0.0
#>

[CmdletBinding(DefaultParametersetName='None')]
param (
    # Threshold value in percent, when a database is excluded or resumed
    [Parameter(Mandatory)]
    [int]$Treshold,

    # List of databases that should not be processed by the script
    [array]$ExcludedDatabases,
    
    # Send a report by mail after script execution
    [Parameter(ParameterSetName = 'SendReport')]
    [switch]$SendReport,

    # Directory in which the report is saved before it is sent as a mail
    [Parameter(ParameterSetName = 'SendReport', Mandatory)]
    [ValidateScript({Test-Path -Path $_ -PathType Container})]
    [string]$ReportPath,

    # List of recipients who will receive the report
    [Parameter(ParameterSetName = 'SendReport', Mandatory)]
    [array]$ReportRecipients,

    # The sender's address
    [Parameter(ParameterSetName = 'SendReport', Mandatory)]
    [string]$ReportSender,

    # Subject of the mail
    [Parameter(ParameterSetName = 'SendReport')]
    [string]$ReportSubject = 'Database Suspension Report',

    # SMTP server that is used to send the mail
    [Parameter(ParameterSetName = 'SendReport', Mandatory)]
    [string]$SmtpServer,

    # Port over which the SMTP server accepts the mail
    [Parameter(ParameterSetName = 'SendReport')]
    [int]$SmtpPort = 25,

    # Credentials of the sender. Not tested yet!
    [Parameter(ParameterSetName = 'SendReport')]
    [int]$SmtpCredential
)

function Convert-Size {
    [cmdletbinding()]
    param(
        [Parameter(ParameterSetName = 'From')]
        [ValidateSet("Bytes", "KB", "MB", "GB", "TB")]
        [string]$From,
        
        [Parameter(ParameterSetName = 'To')]
        [ValidateSet("Bytes", "KB", "MB", "GB", "TB")]
        [string]$To,

        [Parameter(Mandatory)]
        [double]$Value,

        [int]$Precision = 1
    )

    switch ($From) {
        "KB" { $value = $Value * 1024 }
        "MB" { $value = $Value * 1024 * 1024 }
        "GB" { $value = $Value * 1024 * 1024 * 1024 }
        "TB" { $value = $Value * 1024 * 1024 * 1024 * 1024 }
        default { $Value = $Value }
    }

    switch ($To) {
        "KB" { $Value = $Value / 1KB }
        "MB" { $Value = $Value / 1MB }
        "GB" { $Value = $Value / 1GB }
        "TB" { $Value = $Value / 1TB }
        default { return $value }
    }

    return [Math]::Round($value, $Precision, [MidPointRounding]::AwayFromZero)
}

# ------------------------------------------------------------------------------------
# Prepare variables
# ------------------------------------------------------------------------------------

# Required properties
$DatabaseProperties = (
    'Name',
    'Server',
    'EdbFilePath',
    'DatabaseSize',
    'AvailableNewMailboxSpace',
    'IsExcludedFromProvisioning'
)

# Ordered result properties
$OrderedResults = (
    'Database',
    'Suspended',
    'SizeGb',
    'DiskSpaceGb',
    'DiskSpacePct',
    'WhiteSpaceGb',
    'WhiteSpacePct',
    'TotalSpaceGb',
    'TotalSpacePct'
)

$Resume = @{
    IsExcludedFromProvisioning       = $false
    IsExcludedFromProvisioningReason = $null
}

$Suspend = @{
    IsExcludedFromProvisioning       = $true
    IsExcludedFromProvisioningReason = 'Too low disk space.'
}

$Export = @{
    Delimiter         = ';'
    Encoding          = 'UTF8'
    NoTypeInformation = $true
}

if ($PSCmdlet.ParameterSetName -eq 'SendReport') {
    $MailSettings = @{
        To         = $ReportRecipients
        From       = $ReportSender
        Subject    = $ReportSubject
        Body       = 'Find the database suspension report attached.'
        SmtpServer = $SmtpServer
        Port       = $SmtpPort
    }

    if ($SmtpCredential) {
        # Add sender credentials
        $MailSettings.Add('Credential', $SmtpCredential)
    }
}

# ------------------------------------------------------------------------------------
# Create a list of all required databases
# ------------------------------------------------------------------------------------
Write-Verbose -Message 'Search databases...'

try {
    # Search all databases
    $Databases = Get-MailboxDatabase -Status -ErrorAction 'Stop' |
        Select-Object -Property $DatabaseProperties
}
catch {
    Write-Error -Message "Failed to search databases. $PSItem"
}

# Remove excluded databases and sort them by name
$Databases = $Databases | Where-Object { $_.Name -notin $ExcludedDatabases } | Sort-Object -Property Name

# ------------------------------------------------------------------------------------
# Create a list of all databases and their free space
# ------------------------------------------------------------------------------------
Write-Verbose -Message 'Calculate available database space...'

$DBSpace = foreach ($DB in $Databases) {

    $Disk, $Arguments, $WhiteSpace = $null

    Write-Verbose "Current database: $($DB.Name)"

    # --------------------------------------------------------------------------------
    # Database on local host
    # --------------------------------------------------------------------------------
    if ($DB.Server -eq $env:COMPUTERNAME) {
        try {
            # Get disk space params for database volume
            $Disk = (Get-Volume -FilePath $DB.EdbFilePath -ErrorAction 'Stop' |
                Select-Object -Property Size, SizeRemaining)[0]
        }
        catch {
            Write-Error -Message "Failed to get disk size for $($DB.Name). $PSItem"
            continue
        }
    }

    # --------------------------------------------------------------------------------
    # Database on remote server
    # --------------------------------------------------------------------------------
    if ($DB.Server -ne $env:COMPUTERNAME) {
        # Setup a list with vars
        $Arguments = $DB.EdbFilePath
        
        try {
            # Invoke command on remote server
            $Disk = Invoke-Command -ComputerName $DB.Server -ArgumentList $Arguments -ScriptBlock {
                # Get disk space params for database volume
                (Get-Volume -FilePath $args[0] -ErrorAction 'Stop' |
                    Select-Object -Property Size, SizeRemaining)[0]
            }
        }
        catch {
            Write-Error -Message "Failed to get disk size for $($DB.Name). $PSItem"
            continue
        }
    }

    # --------------------------------------------------------------------------------
    # Disk values present
    # --------------------------------------------------------------------------------
    if ($null -ne $Disk) {
        
        # Save the white space to recude future command length
        $WhiteSpace = $DB.AvailableNewMailboxSpace.ToBytes()
    
        # Create an object with available space details
        New-Object -TypeName PSObject -Property @{
            Database      = $DB.Name
            Suspended     = $DB.IsExcludedFromProvisioning
            SizeGb        = Convert-Size -To GB -Value $DB.DatabaseSize.ToBytes()
            DiskSpaceGb   = Convert-Size -To GB -Value $Disk.SizeRemaining
            DiskSpacePct  = [math]::Round(($Disk.SizeRemaining / $Disk.Size * 100))
            WhiteSpaceGb  = Convert-Size -To GB -Value $WhiteSpace
            WhiteSpacePct = [math]::Round(($WhiteSpace / $DB.DatabaseSize.ToBytes() * 100))
            TotalSpaceGb  = Convert-Size -To GB -Value ($Disk.SizeRemaining + $WhiteSpace)
            TotalSpacePct = [math]::Round((($Disk.SizeRemaining + $WhiteSpace) / $Disk.Size * 100))
        }
    }
}

# If verbose is active, return the results
if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {
    $DBSpace | Format-Table -Property $OrderedResults -AutoSize
}

# ------------------------------------------------------------------------------------
# Resume and suspend auto database provisioning
# ------------------------------------------------------------------------------------
Write-Verbose -Message 'Update databases provisioning settings...'

foreach ($DB in $DBSpace) {

    Write-Verbose "Current database: $($DB.Database)"

    # --------------------------------------------------------------------------------
    # Resume suspended database with enough available space
    # --------------------------------------------------------------------------------
    if ($DB.Suspended -eq $true -and $DB.TotalSpacePct -gt $Treshold) {
        try {
            # Set value to false
            ($DBSpace | Where-Object { $_.Database -eq $DB.Database }).Suspended = $false
            
            # Resume database for auto provisioning
            Set-MailboxDatabase @Resume -Identity $DB.Database -ErrorAction 'Stop'
            Write-Verbose 'Database resumed.'
            continue
        }
        catch {
            Write-Error -Message "Failed to resume database $($DB.Database). $PSItem"
            continue
        }
    }

    # --------------------------------------------------------------------------------
    # Suspend database with less available space
    # --------------------------------------------------------------------------------
    if ($DB.Suspended -eq $false -and $DB.TotalSpacePct -le $Treshold) {
        try {
            # Set value to true
            ($DBSpace | Where-Object { $_.Database -eq $DB.Database }).Suspended = $true

            # Suspend database from auto provisioning
            Set-MailboxDatabase @Suspend -Identity $DB.Database -ErrorAction 'Stop'
            Write-Verbose 'Database suspended.'
            continue
        }
        catch {
            Write-Error -Message "Failed to suspend database $($DB.Database). $PSItem"
            continue
        }
    }
}

# ------------------------------------------------------------------------------------
# Export report to disk
# ------------------------------------------------------------------------------------
if ($PSCmdlet.ParameterSetName -eq 'SendReport') {

    Write-Verbose -Message 'Export report to disk...'
    
    # Setup the export file name
    $FileName = (Get-Date -Format 'yyyy-MM-dd') + ' DBSuspensionReport.csv'

    try {
        # Export the report to disk
        $DBSpace | Select-Object -Property $OrderedResults |
            Export-Csv @Export -Path "$ReportPath\$FileName" -ErrorAction 'Stop'
    }
    catch {
        Write-Error -Message "Failed to export the report. $PSItem"
        break
    }
}

# ------------------------------------------------------------------------------------
# Send report by mail
# ------------------------------------------------------------------------------------
if ($PSCmdlet.ParameterSetName -eq 'SendReport') {

    Write-Verbose -Message 'Send report by mail...'

    try {
        # Send mail report
        Send-MailMessage @MailSettings -Attachments "$ReportPath\$FileName"
    }
    catch {
        Write-Error -Message "Failed to send the report. $PSItem"
        break    
    }
}