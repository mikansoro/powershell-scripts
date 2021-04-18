<#
.SYNOPSIS
    Prints key mkv track information for all mkv files in a folder, for use with mkvmerge or Remove-ExtraMkvTracks.ps1
.PARAMETER Folder
    Type: String
    Path where MKV files exist. Required to be an absolute path. 
.NOTES
    Requires: mkvmerge.exe, mkvpropedit.exe both on your user or system path.
    
    Example Usage: .\Get-MkvTrackInformation.ps1 -Folder "C:\path\to\media"
    
    Currently, the only output format is JSON. It's easy enough to read for this purpose, and didn't
        require building a custom formatter. Eventually, an optional YAML output (easier to read for humans)
        will be added, once YAML support is added to mainline powershell. In versions before YAML support (when added), 
        JSON will still most likely be the default output.
        See this bug on YAML status if that's your thing. https://github.com/PowerShell/PowerShell/issues/3607

    Author: Michael Rowland
    Date:   2021-04-17
#>
param (
    [string]$Folder
)

# accepts two sets of tracks as psobject[], checks several key metrics to determine if they're indeed equal
function tracksAreEqual($trackset1, $trackset2) {
    # If the tracksets aren't the same length, they're not equal. Plain and simple.
    if ($trackset1.count -ne $trackset2.count) {return $false}

    # Check key attributes. If key attributes for every track in $trackset1 match the corresponding key attributes 
    # in the $trackset2 track index, they're equal. Otherwise, not equal.
    for ($i = 0; $i -lt $trackset1.count; $i++) {
        if ($trackset1[$i].id -ne $trackset2[$i].id) {
            #this should never happen. but just in case, for a corrupted/misformatted mkv edge case. 
            return $false
        } elseif ($trackset1[$i].type -ne $trackset2[$i].type) {
            return $false
        } elseif ($trackset1[$i].codec -ne $trackset2[$i].codec) {
            return $false
        } elseif ($trackset1[$i].language -ne $trackset2[$i].language) {
            return $false
        } elseif ($trackset1[$i].name -ne $trackset2[$i].name) {
            return $false
        } elseif ($trackset1[$i].default_track -ne $trackset2[$i].default_track) {
            return $false
        }
    }
    return $true
}

function selectKeyTrackInfo ($trackset) {
    $output = @()
    foreach ($track in $trackset.tracks) {
        $output += @{id = $track.id; type = $track.type; language = $track.properties.language; codec = $track.codec; name = $track.properties.track_name; default_track = $track.properties.default_track}
    }
    return $output
}

if (-not (Split-Path -Path $folder -IsAbsolute)) {
    Write-Warning '$folder should be a literal path. Exiting.'
    exit
}

try {
    $files = Get-ChildItem -File -Filter "*.mkv" -Path $folder 
} catch {
    # need to also check -LiteralPath, for some edge case folder names that include a sub release group in brackets, i.e. C:\Downloads\[SUBGroup] show [info]\[SUBGroup] release [info].mkv
    # about half the time, Powershell will think that it is regex without using -LiteralPath
    $files = Get-ChildItem -File -Filter "*.mkv" -LiteralPath $folder 
}
if ($files.count -lt 1) {
    Write-Warning "No MKV files found. Exiting."
    exit
}

$UniqueTrackSets = @()

foreach ($file in $files) {
    $mkvparams = "-J",$file.FullName
    $trackset = (& mkvmerge @mkvparams) | ConvertFrom-Json | Select-Object -Property tracks
    $trackset = selectKeyTrackInfo -trackset $trackset
    if ($UniqueTrackSets.Count -eq 0) {
        $UniqueTrackSets += @{files = @($file.Name); tracks = $trackset}
    } else {
        $match = $false
        for ($i = 0; $i -lt $UniqueTrackSets.Count; $i++) {
            if (tracksAreEqual -trackset1 $trackset -trackset2 $UniqueTrackSets[$i].tracks) {
                $UniqueTrackSets[$i].files += $file.Name
                $match = $true
            }
        }
        if (-not $match) {
            $UniqueTrackSets += @{files = @($file.Name); tracks = $trackset}
        }
    }
}

$UniqueTrackSets | ConvertTo-Json -Depth 10