[CmdletBinding(DefaultParametersetName='None')]
param (
    # Treshold in % for the total free space
    [Parameter(Mandatory)]
    [int]$Treshold,

    # Send a report by mail
    [Parameter(ParameterSetName = 'SendReport')]
    [switch]$SendReport,

    # Path to export suspension reports
    [Parameter(ParameterSetName = 'SendReport', Mandatory)]
    [ValidateScript({Test-Path -Path $_ -PathType Container})]
    [string]$ReportPath,

    # List or report recipients
    [Parameter(ParameterSetName = 'SendReport', Mandatory)]
    [array]$MailRecipients,

    # Sender mail address
    [Parameter(ParameterSetName = 'SendReport', Mandatory)]
    [string]$MailSender,

    # Mail server to send the mail
    [Parameter(ParameterSetName = 'SendReport', Mandatory)]
    [string]$SmtpServer,

    # Listing smtp port
    [Parameter(ParameterSetName = 'SendReport')]
    [int]$SmtpPort = 25,

    # Mail subject
    [Parameter(ParameterSetName = 'SendReport')]
    [string]$MailSubject = 'Database Suspension Report',

    # Basic sender credentials
    [Parameter(ParameterSetName = 'SendReport')]
    [int]$SenderCredential,

    # List of databases to exclude from management
    [array]$ExcludedDatabases
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
        To         = $MailRecipients
        From       = $MailSender
        Subject    = $MailSubject
        Body       = 'Find the database suspension report attached.'
        SmtpServer = $SmtpServer
        Port       = $SmtpPort
    }

    if ($SenderCredential) {
        # Add sender credentials
        $MailSettings.Add('Credential', $SenderCredential)
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

    # Save the white space to recude future command length
    $WhiteSpace = $DB.AvailableNewMailboxSpace.ToBytes()

    # Disk values present
    if ($null -ne $Disk) {
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

    # Suspended but enough available space present
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

    # Not suspended but to less available total space
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