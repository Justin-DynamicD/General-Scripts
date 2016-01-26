#Get Complete Databaselist to crawl
$DBList = get-mailboxdatabase

#Loop Each DB
foreach ($DB in $DBList) {

    #Grab Current DB Limits and DB name
    $DBIdentity = $DB.Identity
    IF ($DB.ProhibitSendReceiveQuota.IsUnlimited) {$DBReceiveQuota = "Unlimited"} Else {$DBReceiveQuota = $DB.ProhibitSendReceiveQuota.Value}
    IF ($DB.ProhibitSendQuota.IsUnlimited) {$DBSendQuota = "Unlimited"} Else {$DBSendQuota = $DB.ProhibitSendQuota.Value}
    IF ($DB.IssueWarningQuota.IsUnlimited) {$DBWarning = "Unlimited"} Else {$DBWarning = $DB.IssueWarningQuota.Value}

    #Grab All Users in the Current DB
    $MBList = get-mailbox -database $DBIdentity

    #Check each mailbox for custom limits, apply storage limits is inherited.
    foreach ($MB in $MBList) {
        IF ($MB.UseDatabaseQuotaDefaults) {
            set-mailbox $MB -ProhibitSendQuota $DBSendQuota -ProhibitSendReceiveQuota $DBReceiveQuota -IssueWarningQuota $DBWarning -UseDatabaseQuotaDefaults:$false
            } #Update Limits
        } #MB Loop
    } #DB Loop
