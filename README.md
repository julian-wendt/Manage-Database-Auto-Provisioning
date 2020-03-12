# Manage-Database-Auto-Provisioning

PowerShell Script to automatically exclude/include Exchange Server Mailbox Databases based on the available disk space and the database whitespace from the mailbox provisioning load balancer that distributes new mailboxes randomly and evenly across the available databases, based on the available disk space and database whitespace.

## How the script works

The script creates a list of all Mailbox Databases in the Organization and checks the available disk space where the EDB-File is homed. For each database, an object is created and added to a list, containing the following information:

* Database exclusion state
* Database size in GB
* Database white space in GB and percent
* Available disk space in GB and percent
* Total free space in GB and percent, consisting of available disk space and database whitespace

Based on the the total free space and a treshold, databases will be automatically excluded from provisioning or resumed.

## Using the script

This chapter provides all necessary information to execute the script.

### Parameters

Param | Type | Mandatory | Description
--- | --- | --- | ---
`Treshold` | `int` | `true` | Specifies the threshold value in percent, when a database is excluded or resumed.
`ExcludedDatabases` | `array` |`false` | List of databases that should not be processed by the script.
`SendReport` | `switch` | `false` | Use the switch to send a report by mail after an execution. The report will contain a csv file with the collected information from the chapter above. To send the report, further parameters are required.
`ReportPath` | `string` | `true` | Directory in which the report is saved before it is sent as a mail.
`ReportRecipients` | `array` | `true` | List of recipients who will receive the report.
`ReportSender` | `string` | `true` | The sender's address.
`ReportSubject` | `string` |`false` | Subject of the mail. Default value: *Database Suspension Report*
`SmtpServer` | `string` | `true` | SMTP server that is used to send the mail.
`SmtpPort` | `int` | `true` | Port over which the SMTP server accepts the mail. Default value: *25*
`SmtpCredential` | `PSCredential` |`false` | Credentials of the sender. Not tested yet!

**Please note that the sending of mails with the `SmtpCredentials` has not been tested by me and was only implemented for completeness.**

### Examples

Use this command to manage the suspension for all databases except DB1 and DB2 that have less than 20 percent total storage space:
`.\ManageDatabaseAutoProvisioning.ps1 -Treshold 20 -ExcludedDatabases 'DB1', 'DB2'`

Use this command to manage the suspension for all databases that have less than 20 percent total storage space. Afterwards send a report to two recipients:
`.\ManageDatabaseAutoProvisioning.ps1 -Treshold 20 -SendReport -ReportPath 'C:\Temp\' -ReportRecipients 'julian@example.com', 'exchangeadmins@example.com' -ReportSender -suspensionreport@example.com -SmtpServer 'mail.example.com'`
