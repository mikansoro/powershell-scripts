<#
.NOTES
    This was originally designed to be copy/pasted into a running intermedia powershell session. They do not provide a module to integrate with their "HostPilot" platform. 
    Intermedia only provides an interactive shell. See https://kb.intermedia.net/Article/23442 for details on how to set it up.
#>
$filter = (Read-Host -Prompt "Enter a filter to search for mailboxes. Use * for a wildcard")
$destinationEmailDomain = (Read-Host -Prompt "Enter a domain for each account to forward to (in the format '@contoso.com')")
$accounts = Get-User -count 1000 | Where-Object UserPrincipalName -like $filter
$processed = 0
$accounts | foreach-object {
    Write-Progress -Activity "Setup Mailbox Forwarding" -Status "Progress:" -PercentComplete ($processed/$accounts.count*100)
    $NewAddress = $_.UserPrincipalName.split("@")[0] + $destinationEmailDomain
    $contact = New-MailContact -DisplayName $_.UserPrincipalName.split("@")[0] -EmailAddress $NewAddress
    Set-ExchangeMailbox -Identity $_.DistinguishedName -ForwardingAddress $contact.DistinguishedName -DeliverAndForwardMail $true
    $processed += 1
}