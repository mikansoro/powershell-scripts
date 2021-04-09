<#
.SYNOPSIS
    Reads a folder for subfolders, creates a DFS Folder mount for each one. Creates 2 Domain Local Active Directory Groups for each folder, one for "READ" and one for "MODIFY". Does not set NTFS permissions (yet).
.NOTES
    Requires: ADDS Powershell Snap-In installed, DFS-N Server Role Enabled with Powershell Module and dfsutil.exe accessible

    Originally designed to be run from the fileserver hosting DFS
    Still needs to set NTFS permissions on the local folder for proper DFS sharing. 
    Maybe could build folders optionally from a list, inside $LiteralPath ? 
    Only supports -Verbose and -Debug from CmdletBinding() so far.

    Author:         Michael Rowland
    Date Created:   2021-04-08
#>
[CmdletBinding()]
param (
    [string]$LiteralPath = (Get-Location),
    [string]$DFSRootFolder = $env:USERDOMAIN,
    [string]$DFSDomain = $env:UserDnsDomain,
    [pscredential]$Credential = (Get-Credential),
    [Parameter(Mandatory)]
    [string]$DLGroupPath,
    [Parameter(Mandatory=$false)]
    [switch]$ModifyACL
)

function Update-AclRules ($groups, $folder) {
    $folderACL = Get-Acl -Path $folder.FullName
    $folderACL.SetAccessRuleProtection($true,$false) # Disable inheritance and delete inherited ACLs
    Write-Verbose "Setting NTFS Permissions"
    foreach ($group in $groups) {
        $newrule = New-Object System.Security.AccessControl.FileSystemAccessRule($group.name, $group.permission, "Allow")
        $folderACL.SetAccessRule($newrule)
    }
    # add local administrators group as full-control on each folder
    $newrule = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "Allow")
    $folderACL.SetAccessRule($newrule)
    Write-Debug $folderACL.Access
    $folderACL | Set-Acl -Path $folder.FullName
}

# if templates array initialized empty, dfs folder access will be default, i.e. $ADGroupNameTemplates = @()
$ADGroupNameTemplates = @{Name = "DL-FS-DFS-(<FOLDER>)-READ"; Permission = "ReadAndExecute"}, @{Name = "DL-FS-DFS-(<FOLDER>)-MODIFY"; Permission = "Modify"}

if (-not (Test-Path -LiteralPath $LiteralPath)) {
    Write-Warning "Path is not a Literal Path valid on this server. Exiting..."
    exit
}
Write-Verbose "Path Check Success"

if (-not (Get-DfsnRoot -Path "\\$DFSDomain\$DFSRootFolder" -ErrorAction SilentlyContinue)) {
    Write-Warning "DFS Root Folder \\$DFSDomain\$DFSRootFolder does not exist. Exiting..."
    exit
}
Write-Verbose "DFS-N Root Check Success"

$Folders = Get-ChildItem -Directory -Depth 0 -LiteralPath $LiteralPath 
foreach ($folder in $Folders) {

    Write-Verbose "Current Folder: $folder"
    $DFSFolderPath = "\\$DFSDomain\$DFSRootFolder\$($folder.Name)"
    
    #if already in DFS, skip
    if (Get-DfsnFolder -Path "$DFSFolderPath" -ErrorAction SilentlyContinue) {
        Write-Verbose "$($folder.name) in DFS already. Skipped."
        continue
    }
    
    # Create server-local smb share if folder is not shared already
    $WMIFolderPath = $folder.FullName -replace '\\','\\'
    $WMIShare = Get-CimInstance -Query "SELECT * FROM Win32_Share WHERE Path='$WMIFolderPath'"
    Write-Debug "WMI Share missing: $((-not ([bool]$WMIShare)))"

    if (-not ([bool]$WMIShare)) {
        Write-Verbose "Creating new SMB Share for $($folder.FullName)"
        New-SMBShare -Name $folder.Name -Path $folder.FullName -FullAccess "Everyone"
        $SMBSharePath = "\\$($env:computername)\$($folder.Name)"
    } else {
        Write-Verbose "Share exists for $($folder.fullname) already"
        $SMBSharePath = "\\$($env:computername)\$($WMIShare.Name)"
    }
    Write-Debug "SharePath: $SMBSharePath"

    Write-Verbose "Creating DFS Folder: $DFSFolderPath"
    # Create/Register DFS folders
    New-DfsnFolder -Path $DFSFolderPath -TargetPath $SMBSharePath 

    # generate group information from templates for this folder
    $groups = @()
    foreach ($template in $ADGroupNameTemplates) {
        Write-Debug "folder.name = $($folder.Name), template.name = $($template.name)"
        $group = @{Name = $template.Name -replace '<FOLDER>',"$($folder.Name)"; Permission = $template.Permission}
        $group.NetbiosName = "$($env:UserDomain)\$($group.Name)"
        Write-Verbose "Access Group: $($group.Name)"
        $groups += $group

        try {
            Get-ADGroup -Credential $Credential -Identity $group.Name | Out-Null
        } catch {
            Write-Verbose "Group $($Group.Name) not found. Creating new group."
            New-ADGroup -Credential $Credential -DisplayName $group.Name -SAMAccountName $group.Name -Name $group.Name -GroupCategory Security -GroupScope DomainLocal -Path $DLGroupPath
            start-sleep 20
        }

        # janky fix to wait for New-ADGroup to finish making the group and replicate it
        Do {
            If($Idx -gt 0) {Start-sleep -s 5}
            $r = Get-ADGroup -Identity $group.name
            Write-Verbose "."
            $Idx = $Idx + 1
        } Until($r)
    }

    foreach ($group in $groups) {
        Write-Verbose "Granting DFS Access to folder $DFSFolderPath : $($group.NetbiosName)"
        # use dfsutil.exe instead of Grant-DfsnAccess due to bug in Grant-DfsnAccess
        # https://docs.microsoft.com/en-us/troubleshoot/windows-client/system-management-components/grant-dfsnaccess-not-change-inheritance
        $params = "property","sd","grant",$DFSFolderPath,"$($group.NetbiosName):RX","protect"
        & dfsutil @params
        # Grant-DfsnAccess -Path $DFSFolderPath -AccountName "$($group.NetbiosName)"
    }

    if ($ModifyACL) {
        Update-AclRules -groups $groups -folder $folder
    }
}