<#
.SYNOPSIS
    Allows a user to define a subset of video, audio, and subtitle tracks to keep for mkv files that exist with a directory, and remove all others via mkvmerge remuxing. 
    Users can set options for each track or use existing data. mkv "title" attribute is removed for each file. 
.PARAMETER Folder
    Type: String
    Path where MKV files exist. Required to be an absolute path. 
.PARAMETER Video
    Type: int[]
    Array of video track indicies to keep in the output file
.PARAMETER Audio
    Type: object[]
    Array of audio track information to keep in the output file. Each element can either be a track index, or a partial or full hashmap of information, defined below.
    ID is the only required field in the hashmap.
    @{
        id = 3
        name = "English"
        lang = "eng" # should be an iso 3 letter language code, per mkvmerge
        default = "true" # or "false"
    }
.PARAMETER Subtitle
    Type: object[]
    Array of subtitle track information to keep in the output file. Each element can either be a track index, or a partial or full hashmap of information, defined below.
    ID is the only required field in the hashmap.
    @{
        id = 3
        name = "English"
        lang = "eng" # should be an iso 3 letter language code, per mkvmerge
        default = "true" # or "false"
    }
.NOTES
    Author: Michael Rowland
    Date:   2021-04-08
#>
[CmdletBinding()]
param (
    [string]$Folder,
    [int[]]$Video,
    [object[]]$Audio,
    [object[]]$Subtitle
)

if (-not (Split-Path -Path $folder -IsAbsolute)) {
    Write-Warning '$folder should be a literal path. Exiting.'
    exit
}

$files = Get-ChildItem -File -Filter "*.mkv" -Path $folder -Depth 0 -Exclude "temp.mkv"
if ($files.count -lt 1) {
    # need to also check -LiteralPath, for some edge case folder names that include a sub release group in brackets, i.e. C:\Downloads\[SUBGroup] show [info]\[SUBGroup] release [info].mkv
    # about half the time, Powershell will think that it is regex without using -LiteralPath
    $files = Get-ChildItem -File -Filter "*.mkv" -LiteralPath $folder -Depth 0 -Exclude "temp.mkv"
    if ($files.count -lt 1) {
        Write-Warning "No MKV files found. Exiting."
        exit
    }
}

$folder = Split-Path -Path $files[0].FullName
$propeditparams = "--delete","title"

foreach ($file in $files) {

    $tempfile = "$folder\temp.mkv"
    $params = "-o",$tempfile
    $filedata = & mkvmerge "-J" $file.FullName | ConvertFrom-Json

    $videotracks = ""
    foreach($track in $video) {
        $trackid = $(if ($track.id) {$track.id} else {$track})
        $videotracks += "$trackid,"
    }
    $params += "-d",$videotracks.substring(0,$videotracks.Length-1)
    
    # horribly janky way of building "-a x,y,z", but it all has to be defined at once. cannot define "-a x -a y -a z" as only "z" will be committed to the final file
    $audiotracks = ""
    foreach ($track in $audio) {
        $trackid = $(if ($track.id) {$track.id} else {$track})
        $audiotracks += "$trackid,"
    }
    $params += "-a",$audiotracks.substring(0,$audiotracks.Length-1)

    foreach($track in $audio) {
        $trackid = $(if ($track.id) {$track.id} else {$track})
        $params += "--track-name","""$($trackid):$(if ($track.name) {$track.name} else {$filedata.tracks[$trackid].properties.track_name})"""
        $params += "--language","""$($trackid):$(if ($track.lang) {$track.lang} else {$filedata.tracks[$trackid].properties.language})"""
        $params += "--default-track","""$($trackid):$(if ($track.default) {$track.default} else {$filedata.tracks[$trackid].properties.default_track})"""
    }

    # same issue as above with "-a"
    $subtracks = ""
    foreach ($track in $subtitle) {
        $trackid = $(if ($track.id) {$track.id} else {$track})
        $subtracks += "$trackid,"
    }
    $params += "-s",$subtracks.substring(0,$subtracks.Length-1)

    foreach($track in $subtitle) {
        $trackid = $(if ($track.id) {$track.id} else {$track})
        $params += "--track-name","""$($trackid):$(if ($track.name) {$track.name} else {$filedata.tracks[$trackid].properties.track_name})"""
        $params += "--language","""$($trackid):$(if ($track.lang) {$track.lang} else {$filedata.tracks[$trackid].properties.language})"""
        $params += "--default-track","""$($trackid):$(if ($track.default) {$track.default} else {$filedata.tracks[$trackid].properties.default_track})"""
    }

    Write-Debug "params for mkvmerge: \n $params"
    & mkvmerge @params $file.FullName
    # need to get folderpath of file to splice onto tempfile and destination, to make sure everything goes to the right place
    Remove-Item -LiteralPath $file.FullName
    Move-Item -LiteralPath $tempfile -Destination $file.FullName

    & mkvpropedit $file.FullName @propeditparams
}