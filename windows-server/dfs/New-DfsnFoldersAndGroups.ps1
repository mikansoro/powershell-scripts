<#
.SYNOPSIS
    Reads a folder for subfolders, creates a DFS Folder mount for each one. Creates 2 Domain Local Active Directory Groups for each folder, one for "READ" and one for "MODIFY". Does not set NTFS permissions (yet).
.NOTES
    Requires AD Powershell Snap-In installed on the local machine.
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
    [string]$DLGroupPath
)

# if templates array initialized empty, dfs folder access will be default, i.e. $ADGroupNameTemplates = @()
$ADGroupNameTemplates = "DL-FS-DFS-(<FOLDER>)-READ", "DL-FS-DFS-(<FOLDER>)-MODIFY"

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
    
    #if already in DFS, skip
    if (Get-DfsnFolder -Path "\\$DFSDomain\$DFSRootFolder\$($folder.Name)" -ErrorAction SilentlyContinue) {
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
        $SharePath = "\\$($env:computername)\$($folder.Name)"
    } else {
        Write-Verbose "Share exists for $($folder.fullname) already"
        $SharePath = "\\$($env:computername)\$($WMIShare.Name)"
    }
    Write-Debug "SharePath: $SharePath"

    Write-Verbose "Creating DFS Folder: \\$DFSDomain\$DFSRootFolder\$($folder.Name)"
    # Create/Register DFS folders
    New-DfsnFolder -Path "\\$DFSDomain\$DFSRootFolder\$($folder.Name)" -TargetPath $SharePath 
    foreach ($template in $ADGroupNameTemplates) {
        $group = $template -replace '<FOLDER>',"$($folder.Name)"
        Write-Verbose "Access Group: $group"

        try {
            Get-ADGroup -Credential $Credential -Identity $group | Out-Null
        } catch {
            Write-Verbose "Group $Group not found. Creating new group."
            New-ADGroup -Credential $Credential -DisplayName $group -SAMAccountName $group -Name $group -GroupCategory Security -GroupScope DomainLocal -Path $DLGroupPath
        }

        Write-Verbose "Granting DFS Access to folder \\$DFSDomain\$DFSRootFolder\$($folder.Name) : $($env:UserDomain)\$group"

        # use dfsutil.exe instead of Grant-DfsnAccess due to bug in Grant-DfsnAccess
        # https://docs.microsoft.com/en-us/troubleshoot/windows-client/system-management-components/grant-dfsnaccess-not-change-inheritance
        $params = "property","sd","grant","\\$DFSDomain\$DFSRootFolder\$($folder.Name)","$($env:UserDomain)\$($group):RX","protect"
        & dfsutil @params
        # Grant-DfsnAccess -Path "\\$DFSDomain\$DFSRootFolder\$($folder.Name)" -AccountName "$($env:UserDomain)\$group"
    }
}