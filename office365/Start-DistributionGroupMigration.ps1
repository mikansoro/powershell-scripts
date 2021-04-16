<#
.SYNOPSIS
    Migrates AD Distribution Groups that are synced with Office 365 to native Office 365 Distribution Groups
.DESCRIPTION
    Searches for any Distribution group which name matches $GroupNameFilter. Aggregates all group information
    and exports it to JSON files in %windir%\temp. Deletes the groups from AD, and forces a sync of that change
    to Office 365. Once groups are removed, creates new distribution groups in Office 365 natively, and resets settings/members/owners/etc. 
.PARAMETER GroupNameFilter
    Required: False
    Default: "*"
    Name filter used to select groups from Active Directory using Get-ADGroup -Filter "Name -like $GroupNameFilter"
    Filters to select multiple groups should be postfixed with/prefixed with/include one or more '*' (asterisk).
.PARAMETER ExchangeUserPrincipalName
    Required: True
    The User Principal Name (usually email address) of the Office 365 Administrator account that will be used 
    to commit the new group changes to Office 365.
.PARAMETER SyncServerFQDN
    Required: True
    The Fully Qualified Domain Name of the on-premesis Active Directory Sync Server, that syncs AD information 
    to Office 365. This is used to initiate a manual sync to Office 365 after changes are committed. 
.PARAMETER DefaultManagerDN
    Required: True
    The Office 365 Distinguished Name of the User/Group used as the default Group manager. 
    Only applies to groups that do not already have a manager. 
    To find DN: (Get-<User/Group> -Identity <email>).DistinguishedName 
    This is NOT the Active Directory Distinguished Name of the Group.
.PARAMETER Limit
    Required: False
    Default: 100000
    A limiter on how many AD groups matching GroupNameFilter will be processed by the script
    Default is sufficiently large enough that in 99% of cases, all groups returned will be included.
.EXAMPLE
    C:\PS> .\Start-GroupMigration.ps1 -Verbose -GroupNameFilter "TEST_O365*" -ExchangeUserPrincipalName "admin@contoso.com" -SyncServerFQDN "syncserver.contoso.com" -Limit 20
.NOTES
    Requires: AD Powershell Snap-In installed, ExchangeOnline module installed, WinRM/PS Remoting enabled on Azure AD Connect Server, Credentials with permission to 
    Delete Distribution Groups from Active Directory and manage Syncing on Azure AD Connect server.

    Supports: -Verbose, -Debug, -WhatIf, -Confirm

    Author: Michael Rowland
#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
param(
    [Parameter(Mandatory=$false)]
    [pscredential]$credential = $(Get-Credential),
    [Parameter(Mandatory=$false)]
    [string]$GroupNameFilter = "*",
    [Parameter(Mandatory)]
    [string]$ExchangeUserPrincipalName,
    [Parameter(Mandatory)]
    [string]$SyncServerFQDN,
    [Parameter(Mandatory)]
    [string]$DefaultManagerDN,
    [Parameter(Mandatory=$false)]
    [int]$Limit = 100000 # any stupidly high number will do here, but will need to be larger than your maximum number of migrations in a single attempt
)

function Get-GroupsAndMembers() {
    Write-Verbose -Message "Get all selected Distribution Groups from AD"
    $adgroups = Get-ADGroup -Credential $credential -Filter "(GroupCategory -eq 'Distribution') -and (Name -like '$GroupNameFilter')" -Properties mail,legacyExchangeDN | Where-Object {
        ($null -ne $_.mail) # skip those Groups that don't have a mail address
    }
    # limit groups to size $Limit
    $adgroups = $adgroups | Select-Object -First $Limit
    if ($Limit -ne 5000) {
        Write-Verbose -Message "Selected $($adgroups.length) groups to migrate based on limit: $Limit."
    }

    $adgroups | ConvertTo-Json | Set-Content -Path "$($Env:windir)\temp\o365migrate-adgroupinfo-$(Get-Timestamp).json"

    $Groups = @()
    $Skipped = @()
    Write-Verbose -Message "Collecting group members & managers for Groups"
    $i = 0
    foreach ($group in $adgroups) {
        $err = $false
        try {
            #with select-object, returns array of DN for group members from o365
            $members = Get-DistributionGroupMember -Identity $group.mail -ErrorAction Stop | Select-Object -ExpandProperty DistinguishedName
        }
        catch { 
            Write-Host "Failed to collect Group Membership information. Skipping Group $($group.mail)." -BackgroundColor yellow -ForegroundColor black
            Write-Host "Reason: $_" -ForegroundColor yellow
            Write-Host
            $err = $true
        }

        # group management is messy. thanks exchange powershell team.
        try {
            #with select-object, returns array of Names for group owners from o365
            $managers_tmp = Get-DistributionGroup -Identity $group.mail -ErrorAction Stop | Select-Object -ExpandProperty ManagedBy
            #get a list of owners. then pick either a user, usermailbox, mailboxuser, or whaterver else type of account to _actually_ be the manager, and add to $managers
            $managers = @()
            foreach($manager in $managers_tmp) { #$managers_temp.getType() = Object[] [String]
                if ($manager -ne "Organization Management" -and $managers_tmp -ne "") {
                    $mgrusers = Get-User -Identity $managers_tmp
                    foreach($mgr in $mgrusers) {
                        if ($mgr.RecipientType -ne "User") {
                            $managers += $mgr.DistinguishedName
                        }
                    }
                }
            }
            if($managers.length -eq 0) {
                $managers += $DefaultManagerDN
            }
        }
        catch { 
            Write-Host "Failed to collect Group Ownership information. Skipping Group $($group.mail)." -BackgroundColor yellow -ForegroundColor black
            Write-Host "Reason: $_" -ForegroundColor yellow
            Write-Host
            $err = $true
        }

        try {
            $properties = Get-DistributionGroup -Identity $group.mail -ErrorAction Stop | Select-Object -Property EmailAddresses,HiddenFromAddressListsEnabled,legacyExchangeDN,ReportToOriginatorEnabled,RequireSenderAuthenticationEnabled
        }
        catch { 
            Write-Host "Failed to collect Group Properties information. Skipping Group $($group.mail)." -BackgroundColor yellow -ForegroundColor black
            Write-Host "Reason: $_" -ForegroundColor yellow
            Write-Host
            $err = $true
        }

        if (-not $err) {
            $Groups += [PSCustomObject]@{
                Name = $group.name
                mail = $group.mail
                members = $members # array of ADUser(s)
                managers = $managers # will be a [string] of names
                objectGUID = $group.objectGUID.ToString()
                EmailAddresses = $properties.EmailAddresses
                HiddenFromAddressListsEnabled = $properties.HiddenFromAddressListsEnabled
                legacyExchangeDN = $properties.legacyExchangeDN
                ReportToOriginatorEnabled = $properties.ReportToOriginatorEnabled
                RequireSenderAuthenticationEnabled = $properties.RequireSenderAuthenticationEnabled
            }
        } else {
            $Skipped += $group
        }
        Write-Progress -Activity "Collect AD Groups" -Status "Progress: " -PercentComplete ($i/$adgroups.count*100)
        $i++
    }
    $Groups | ConvertTo-Json | Set-Content -Path "$($Env:windir)\temp\o365migrate-scriptgroupinfo-$(Get-Timestamp).json"
    $Skipped | ConvertTo-Json | Set-Content -Path "$($Env:windir)\temp\o365migrate-skippedgroupinfo-$(Get-Timestamp).json"

    Write-Verbose -Message "Finished collecting group members for Groups"
    $Groups | ConvertTo-Json | Write-Debug

    return $Groups
}

function Get-Timestamp() {
    return Get-Date -Format o | ForEach-Object { $_ -replace ":", "." }
}

function Remove-OriginalGroups($Groups) {
    $i = 0
    foreach($group in $Groups){
        Write-Verbose -Message "Removing group $($group.name) from AD"
        Remove-ADGroup -Identity $group.objectGUID -Credential $credential # Going to require confirms here as it's potentially quite destructive
        Write-Progress -Activity "Remove AD Groups" -Status "Progress: " -PercentComplete ($i/$Groups.count*100)
        $i++
    }
}

function New-CloudDistributionGroups($Groups) {
    Write-Verbose -Message "Beginning Creation of Groups in Exchange Online"
    $i = 0
    foreach ($group in $Groups) {
        Write-Verbose -Message "Creation of group $($group.Name)"
        New-DistributionGroup -Name $group.Name -PrimarySmtpAddress $group.mail -Type Distribution -RequireSenderAuthenticationEnabled $group.RequireSenderAuthenticationEnabled 
        Set-DistributionGroup -Identity $group.mail -EmailAddresses $group.EmailAddresses -HiddenFromAddressListsEnabled $group.HiddenFromAddressListsEnabled -ReportToOriginatorEnabled $group.ReportToOriginatorEnabled
        Write-Verbose -Message "Group $($group.Name) created."

        $i++
        Write-Progress -Activity "Create Cloud Groups" -Status "Progress: " -PercentComplete (($i/2)/$Groups.count*100)
    } 
    
    Write-Verbose -Message "Begin population of Group Members in Exchange Online"
    foreach ($group in $Groups) {
        Write-Verbose -Message "Adding Members to group $($group.Name)"
        Set-DistributionGroup -Identity $group.mail -ManagedBy $group.managers
        Update-DistributionGroupMember -Identity $group.mail -Members $group.members -Confirm:$false
        Write-Verbose -Message "Finish Adding Members to group $($group.Name)"
        $i++
        Write-Progress -Activity "Create Cloud Groups" -Status "Progress: " -PercentComplete (($i/2)/$Groups.count*100)
    }
}

function Start-O365Sync($Session) {
    Get-O365SyncStatus -SyncServerSession $Session

    Write-Verbose -Message "Initialize sync of AD to O365."
    Invoke-Command -Session $Session -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
    for ($i = 15; $i -gt 0; $i--) {
        Write-Progress -Activity "Initializing Sync To Office 365" -Status "Seconds Remaining: " -SecondsRemaining $i
        Start-Sleep -Seconds 1
    }
    Write-Host "Initialized Sync of AD to O365" -ForegroundColor Green

    Get-O365SyncStatus -SyncServerSession $Session

    Write-Host "Sync Complete. Script will now continue." -ForegroundColor Green 
}

function Get-O365SyncStatus($SyncServerSession) {

    Write-Verbose -Message "Checking for active AD Connector Sync"
    $SyncStatus = Invoke-Command -Session $SyncServerSession -ScriptBlock {Get-ADSyncConnectorRunStatus} #returns null when no sync active

    if ($SyncStatus) {
        Write-Warning -Message "AD Sync currently in progress. Waiting for sync to complete before continuing."
        while($null -ne $SyncStatus) {
            for ($i = 10; $i -gt 0; $i--) {
                Write-Progress -Activity "Time to Next Sync Completion Check" -Status "Seconds Remaining: " -SecondsRemaining $i
                Start-Sleep -Seconds 1
            }
            $SyncStatus = Invoke-Command -Session $SyncServerSession -ScriptBlock {Get-ADSyncConnectorRunStatus}
        }
    }
    else {
        Write-Host "No Active Sync Found. Safe to continue." -ForegroundColor Green
    }
    Write-Verbose -Message "Finished checking for Active AD Connector Sync."
}

function Start-O365ChangePropigationCheck($StartGroup, $EndGroup) {
    $ChangesPropigated = [psobject]@{
        firstgroup = $false
        lastgroup = $false
    }
    $i = 0
    Write-Verbose "Checking Changes Propagated"
    while ($ChangesPropigated.firstgroup -eq $false -or $ChangesPropigated.lastgroup -eq $false) {
        if ($i % 30 -eq 0) {
            if ($ChangesPropigated.firstgroup -eq $false) {
                $check = Get-DistributionGroup -Identity $StartGroup.mail
                if (-not $check) {
                    $ChangesPropigated.firstgroup = $true
                }
            }
            if ($ChangesPropigated.lastgroup -eq $false) {
                $check = Get-DistributionGroup -Identity $EndGroup.mail
                if (-not $check) {
                    $ChangesPropigated.lastgroup = $true
                }
            }
        }
        
        Write-Progress -Activity "Waiting for Changes" -Status "Time Elapsed: " -SecondsRemaining $i
        $i++
        Start-Sleep -Seconds 1
    }

    # wait an additional 5 minutes after detecting changes, just in case more changes from middle of large group not processed yet
    # horrible hacky idea, should change this later
    $target = $i + 180
    for($i; $i -lt $target; $i++) {
        Write-Progress -Activity "Waiting for Changes" -Status "Time Elapsed: " -SecondsRemaining $i
        $i++
        Start-Sleep -Seconds 1
    }
}

# ------------------

function Start-GroupMigration() {
    begin {
        if (-not (Test-Connection $SyncServerFQDN)) {
            Write-Error "No server found with name $SyncServerFQDN"
        }

        Write-Verbose -Message "Begin Connection to Exchange Online"
        Connect-ExchangeOnline -UserPrincipalName $ExchangeUserPrincipalName -ShowProgress $true -ErrorAction Stop
        Write-Verbose -Message "Connection Established to Exchange Online"
    }
    process {
        # this is a mess
        if ($GroupNameFilter -eq "*") {
            if ($PSCmdlet.ShouldProcess(
                "All distribution groups with registered mail addresses would be removed from AD, and rebuilt in Office 365.",
                "You have selected a group selection filter of * (any group). Are you sure you would like to proceed selecting ALL AD Distribution Groups for Migration?",
                "No Filter Selection Check"
            )) {
                $SyncServerSession = New-PSSession -ComputerName $SyncServerFQDN -Credential $credential -ErrorAction Stop
                $Groups = Get-GroupsAndMembers
            }
            else {    
                $Groups = Get-GroupsAndMembers
            }
        }
        else {
            Write-Debug -Message "GroupNameFilter = $GroupNameFilter"
            $SyncServerSession = New-PSSession -ComputerName $SyncServerFQDN -Credential $credential -ErrorAction Stop
            $Groups = Get-GroupsAndMembers
        }

        if ($PSCmdlet.ShouldProcess(
            "$($Groups.length) Distribution Groups would be removed from AD, and rebuilt in Office 365 with the set filter \'$($GroupNameFilter)\'",
            "Would you like to migrate $($Groups.length) groups?",
            "Migrate Groups"
        )) {
            Remove-OriginalGroups -Groups $Groups
            Start-O365Sync -Session $SyncServerSession
            Start-O365ChangePropigationCheck -StartGroup $Groups[0] -EndGroup $Groups[1]
            New-CloudDistributionGroups -Groups $Groups
        }          
    }
    end {
        Write-Host "Successfully Updated all Groups in Exchange Online" -ForegroundColor Green
        Write-Verbose -Message "Begin Disconnection from Exchange Online. Ready for use."
        Disconnect-ExchangeOnline
        Write-Verbose -Message "Disconnected from Exchange Online"
    } 
}

Start-GroupMigration