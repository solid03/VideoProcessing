#region Archive
Class Archive {
    [string]$Name
    [string]$FullName
    [string]$ArchiveDir
    [string]$ParentDir
    
    Archive([System.IO.FileInfo]$File) {
        $this.Name = $File.Name
        $this.FullName = $File.FullName
        $this.ArchiveDir = $File.Directory
        $this.ParentDir = $this.ArchiveDir | Split-Path
    }

    # Extract()
    #
    # Extract .rar files
    # Requires 7zip
    [void] Extract() {
        $7z = "C:\Program Files\7-Zip\7z.exe"
        try {
            Write-Host -ForegroundColor Cyan "$($this.Name)"
            & $7z e $this.FullName -o"$($this.ParentDir)" | Out-Host
        }
        catch {throw $_}
    }
}
## This function name conflicts with actual Expand-Archive commandlet
<#
Function Expand-Archive {
param(
    [parameter(Mandatory,ValueFromPipeline)]
    $Path
)
Process {
    switch ((Get-Item $Path).PSIsContainer) {
        $true {
            $archive = Get-ChildItem $Path *.rar
        }
        default {
            if ($Path -notmatch '.rar') {Write-Error "'$Path' is not an archive" ; return}
            $archive = Get-Item $Path *.rar
        }
    }
    if (!($path)) {return}
    $7z = "C:\Program Files\7-Zip\7z.exe"
    & $7z e $archive.FullName -o"$($archive.directory | split-path)"
    #[archive]::new($archive).Extract()    Calling 7z.exe from method doesn't show percent complete
}
}
#>

#endregion
#region Video
Class Video {
[string]$Basename
[string]$Name
[string]$FullName
[string]$Extension
[string]$Directory
[string]$Title
[string]$Genre
[string]$Studio
[string]$Actor
[string]$Date
[string]$Codec
[string]$Height
[string]$Bitrate
[string]$Ratio
[string]$FrameRate
[string]$Duration
[string]$Comment

Video([System.IO.FileInfo]$File) {
    $foo = ffprobe.exe -v error -hide_banner -show_streams -show_format -select_streams v:0 -print_format json $File | ConvertFrom-JSON
    if (!$foo) {
        #throw new ArgumentException("Color must be 'green' or 'yellow'", "color")
        throw [System.ArgumentException]::new("Input must be a video file","File")
    }
    $this.Basename = $File.BaseName
    $this.Name = $File.Name
    $this.FullName = $File.FullName
    $this.Extension = $File.Extension
    $this.Directory = $File.Directory
    $this.Title = $foo.format.tags.title
    $this.Genre = $foo.format.tags.genre
    #$this.Studio = $foo.format.tags.album
    $this.Studio = $foo.format.tags.copyright
    $this.Actor = $foo.format.tags.artist
    $this.Date = $foo.format.tags.date
    $this.Codec = $foo.streams.codec_name
    $this.Height = $foo.streams.height
    $this.Bitrate = $foo.streams.bit_rate
    $this.Ratio = [system.math]::round($foo.streams.bit_rate/($foo.streams.height*$foo.streams.width),2)
    $this.FrameRate = [system.math]::Round([System.Data.DataTable]::new().Compute($foo.streams.avg_frame_rate,""),2)
    $this.Duration = $foo.format.duration
    $this.Comment = $foo.format.tags.comment
}
Video([String]$Path) {
    $File = Get-Item $Path
    $foo = ffprobe.exe -v error -hide_banner -show_streams -show_format -select_streams v:0 -print_format json $File | ConvertFrom-JSON
    if (!$foo) {
        #throw new ArgumentException("Color must be 'green' or 'yellow'", "color")
        throw [System.ArgumentException]::new("Input must be a video file","File")
    }
    $this.Basename = $File.BaseName
    $this.Name = $File.Name
    $this.FullName = $File.FullName
    $this.Extension = $File.Extension
    $this.Directory = $File.Directory
    $this.Title = $foo.format.tags.title
    $this.Genre = $foo.format.tags.genre
    #$this.Studio = $foo.format.tags.album
    $this.Studio = $foo.format.tags.copyright
    $this.Actor = $foo.format.tags.artist
    $this.Date = $foo.format.tags.date
    $this.Codec = $foo.streams.codec_name
    $this.Height = $foo.streams.height
    $this.Bitrate = $foo.streams.bit_rate
    $this.FrameRate = [system.math]::Round([System.Data.DataTable]::new().Compute($foo.streams.avg_frame_rate,""),2)
    $this.Duration = $foo.format.duration
    $this.Comment = $foo.format.tags.comment
}

[void] EncodeH265(){
    ## Settled on hevc_nvenc for speed/quality/compression.  Test sample at q27 had vamf score of 98 and compression ratio of ~.5.

    if ($this.Codec -match 'hevc') {Write-Warning "`'$($this.name)`' is already encoded in H.265" ; return}
    if (!(Test-Path .\encode)) {[void](mkdir encode)}
    #if ($this.Height -gt 2160){  ## Tune option for 4K video
    #    ffmpeg -i $this.FullName -c:v libx265 -preset slow -tune fastdecode ".\encode\$($this.name)" | Out-Host
    #}
    #else {
        #ffmpeg -i $this.FullName -c:v hevc_nvenc -preset p7 ".\encode\$($this.name)" | Out-Host
        #ffmpeg -i $this.FullName -c:v libx265 -preset slow ".\encode\$($this.name)" | Out-Host
        ffmpeg -i $this.FullName -c:v hevc_nvenc -rc vbr -cq 27 -qmin 27 -qmax 27 -profile:v main -pix_fmt p010le -b:v 0K ".\encode\$($this.name)" | Out-Host
    #}    
    
    Remove-Item $this.FullName
    Move-Item ".\encode\$($this.Name)" -Destination $this.Directory
}

# ParseName()
#
# Parse video information/metadata from file name
#
[Hashtable] ParseName() {
    $narcosString = 'narcos-(?<studio>\w+)-(?<date>(?<year>\d\d)-(?<month>\d\d)-(?<day>\d\d))-(?<title>.*)-\d+p'
    $withEpisodeString = '(?<studio>\w+)\W(?<episode>\w{5})\W(?<title>.*)'
    $maybeDateString = '(?<studio>\w+)\W(?<date>(?<year>\d\d)\W(?<month>\d\d)\W(?<day>\d\d))?\.?(?<title>.*)'
    
    $this.Basename -match $withEpisodeString
    $this.Basename -match $maybeDateString
    $this.Basename -match $narcosString

    ## Using Regex switch changes $matches
<#  
    switch -Wildcard ($matches.studio) {
        Add cases to this statement to change the studio name before it is written to the file
    }
#>
        
    return $Matches
}

[string] ToString() {
    return $this.FullName
}

#####Clear metadata
#ffmpeg -y -i "test.mkv" -c copy -map_metadata -1 -metadata title="My Title" -metadata creation_time=2016-09-20T21:30:00 -map_chapters -1 "test.mkv"


# WriteMetadata()
#
# Write video metadata to object. Requires FFMPEG.
# Todo: encode folder appears at current path, not in parent folder

[void] WriteMetadata() {
    ## test for object equality
    $referenceObject = Get-Video $this.FullName
    if (!(Compare-Object $this.psobject.properties.value $referenceObject.psobject.properties.value)) {
        Write-Warning "The command completed successfully, but no changes were made to `'$($this.Name)`'"
        return
    }

    if (!(Test-Path .\encode)) {[void](mkdir encode)}
    ffmpeg -hide_banner -i $this.FullName -metadata comment=$($this.Comment) -metadata genre=$($this.Genre) -metadata artist=$($this.Actor) -metadata album="" -metadata title=$($this.Title) -metadata copyright=$($this.Studio) -metadata date=$($this.Date) -c copy ".\encode\$($this.name)" | Out-Host
    Remove-Item $this.FullName
    Move-Item ".\encode\$($this.Name)" -Destination $this.Directory
}
[void] WriteMetadataAndEncode() {
    ## test for object equality
    $referenceObject = Get-Video $this.FullName
    if (!(Compare-Object $this.psobject.properties.value $referenceObject.psobject.properties.value)) {
        Write-Warning "The command completed successfully, but no changes were made to `'$($this.Name)`'"
        $this.EncodeH265()
        return
    }
    if ($this.Codec -match 'hevc') {
        Write-Warning "`'$($this.name)`' is already encoded in H.265"
        $this.WriteMetadata()
        return
    }

    if (!(Test-Path .\encode)) {[void](mkdir encode)}
    ffmpeg -hide_banner -i $this.FullName -metadata comment=$($this.Comment) -metadata genre=$($this.Genre) -metadata artist=$($this.Actor) -metadata album="" -metadata title=$($this.Title) -metadata copyright=$($this.Studio) -metadata date=$($this.Date) -c:v hevc_nvenc -rc vbr -cq 27 -qmin 27 -qmax 27 -profile:v main -pix_fmt p010le -b:v 0K ".\encode\$($this.name)" | Out-Host
    Remove-Item $this.FullName
    Move-Item ".\encode\$($this.Name)" -Destination $this.Directory
}
}
Function New-VideoSample {
param(
    [parameter(Mandatory,ValueFromPipeline)]
    $Path
)
    $video = [Video]::new((Get-ChildItem $Path))
    ffmpeg.exe -ss $([System.Random]::new().Next($($video.duration - 60))) -i $video.FullName -t 60 -c copy "$($video.Directory)\sample.mp4"

}

Function ConvertTo-H265 {
param(
    [parameter(Mandatory,ValueFromPipeline)]
    $Path
)
process{
    [Video]::new((Get-ChildItem $Path)).EncodeH265()
}
}

Function Search-Video {
param(
    #[parameter(Mandatory)]
    [string]$Filter,
    [switch]$Recurse
)
    switch ($Recurse.IsPresent) {
        True {$File = Get-ChildItem -Recurse *.mp4 | where {$_.BaseName -match ($Filter -split ' ' -join '.')}}
        False {$File = Get-ChildItem *.mp4 | where {$_.BaseName -match ($Filter -split ' ' -join '.')}}
    }
    if ($File) {
        $File | Get-Video    
    }
}

Function Get-Video {
param(
    [parameter(Mandatory,ValueFromPipeline)]
    $Path
)
process {
    [Video]::new((Get-ChildItem $Path))
}
}
Function Invoke-ActorNameWizard {
param (
    [switch]$Override,
    [switch]$Recurse
)
foreach ($File in (Get-ChildItem *.mp4 -Recurse:$Recurse)) {
    $video = Get-Video $File
    if ($video.actor -and (!$Override.IsPresent)) {Continue}
    $video
    $Actor = Read-Host "Actor Name(s)"

    if ([string]$Actor -eq "skip") {Continue}
    add-videometadata $video -actor $Actor
}
}
<#
Function Invoke-TitleWizard {
param ([switch]$Override)
foreach ($File in (Get-ChildItem -Recurse -File *.mp4)) {
    $video = Get-Video $File
    $video
    if ($video.Title -and (!$Override.IsPresent)) {Continue}
    $Title = $video.basename

    if ([string]$Title -eq "skip") {Continue}
    add-videometadata $video -Title $Title
}
}
#>

Function Invoke-VideoChapterWizard {
param(
    [parameter(Mandatory)] $Path
)
    $video = [Video]::new((Get-ChildItem $Path))
    $oldMetadata = ffmpeg -hide_banner -v error -i $video.FullName -f ffmetadata -
    $moreChapters = $true

    do {
        $startTimeStamp = Read-Host "Chapter start timestamp"
        if ($startTimeStamp -ne "0") {
            $startTimeStamp -match '(?<hours>\d+)\W(?<minutes>\d+)\W(?<seconds>)\d+' | Out-Null
            $startTimeStamp = (New-TimeSpan -Hours $Matches.hours -Minutes $Matches.minutes -Seconds $Matches.seconds).TotalSeconds
        }  
        $endTimeStamp = Read-Host "Chapter end timestamp"
        if ($endTimeStamp -match 'end') {
            $endTimeStamp = [System.Math]::Floor($video.Duration)
            $moreChapters = $false
        } else {
            $endTimeStamp -match '(?<hours>\d+)\W(?<minutes>\d+)\W(?<seconds>)\d+' | Out-Null
            $endTimeStamp = (New-TimeSpan -Hours $Matches.hours -Minutes $Matches.minutes -Seconds $Matches.seconds).TotalSeconds
        }
        $chapterTitle = Read-Host "Chapter Title"


        $oldMetadata += `
"[CHAPTER]
TIMEBASE=1/1
START=$startTimeStamp
END=$endTimeStamp
title=$chapterTitle"
    } while ($moreChapters)

    Set-Content $env:TEMP\chapterwiz.txt -Value $oldMetadata

    if (!(Test-Path .\encode)) {[void](mkdir encode)}
    ffmpeg -i $video.FullName -i $env:TEMP\chapterwiz.txt -map_metadata 1 -codec copy ".\encode\$($video.Name)"
    Remove-Item $video.FullName
    Move-Item ".\encode\$($video.Name)" -Destination ".\"
}

Function Set-VideoTitle {
param(
    [parameter(Mandatory,ValueFromPipeline)]
    $Path,
    [switch]$EncodeToH265
)
Process {
    $video = [Video]::new((Get-ChildItem $Path))
    $newName = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($($video.ParseName().title -replace '\W',' '))
    $video.Title = $newName
    #$video.studio = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($video.parsename().studio)
    $video.studio = $video.parsename().studio
    if ($video.parsename().date) {
        $video.date = "$($video.parsename().month).$($video.parsename().day).$($video.parsename().year)"
    }
    if ($EncodeToH265) {$video.WriteMetadataAndEncode()}
    else {$video.WriteMetadata()}
}
}
Function Set-VideoMetadata {
param(
    [parameter(Mandatory,ValueFromPipeline)]
    $Path,
    [string]$Genre,
    [string]$Title,
    [string]$Actor,
    [string]$Studio,
    [string]$Comment
)
Process {
    $video = [Video]::new((Get-ChildItem $Path))
    $video.Genre = $Genre
    $video.Title = $Title
    $video.Actor = $Actor
    $video.Studio = $Studio
    $video.Comment = $Comment
    $video.WriteMetadata()
}
}
Function Add-VideoMetadata {
param(
    [parameter(Mandatory,ValueFromPipeline)]
    $Path,
    [string]$Genre,
    [string]$Title,
    [string]$Actor
)
process {
    if (!$Genre -and !$Title -and !$Actor) {return}    ## check for empty parameters
    $video = [Video]::new((Get-ChildItem $Path))

    [array]$oldGenre = $video.Genre.Trim() -split ','
    [array]$newGenre = $Genre.Trim() -split ','
    $newGenre = ([System.Linq.Enumerable]::Distinct([string[]]$($newGenre + $oldGenre)) -join ',') -replace '(^,)?(,$)?','' ## -replace prevents extra commas
    $video.Genre = $newGenre

    [array]$oldActor = $video.Actor.Trim() -split ','
    [array]$newActor = $Actor.Trim() -split ','
    $newActor = ([System.Linq.Enumerable]::Distinct([string[]]$($newActor + $oldActor)) -join ',') -replace '(^,)?(,$)?','' ## -replace prevents extra commas
    $video.Actor = $newActor

    if ($Title){$video.Title = $Title}
    $video.WriteMetadata()
}
}
Function Get-VideoScreenshots {
param(
    [parameter(Mandatory,ValueFromPipeline)]
    $Path,
    $Destination = "X:\Thumbs"
)
process {
    $video = [Video]::new((Get-ChildItem $Path))
    ffmpeg -i $video.FullName -f image2 -vf fps=1/$([System.Math]::Round($video.Duration/16)),scale=500:-1,tile=4x4 -frames:v 1 "$Destination\$($video.Basename).jpg"
}
}

<#
Function Unfuck-VideoMetadata {
param (    
    [parameter(Mandatory,ValueFromPipeline)]
    $Path
    #[parameter(Mandatory)]
    #[System.String] $Studio
)
process {
    $video = Get-Video $Path
    #$video.Studio = $Studio
    if ($video.studio) {
        Write-Warning "The command completed successfully, but no changes were made to `'$($video.Name)`'"
        return
    }
    $foo = ffprobe.exe -v error -hide_banner -show_streams -show_format -select_streams v:0 -print_format json $Path | ConvertFrom-JSON    
    $video.studio = $foo.format.tags.album
    $video.writemetadata()
}
}
#>

#endregion


#region Classless functions
Function Invoke-DVDWizard {
param(
    [parameter(Mandatory,ValueFromPipeline)] $File,
    [parameter(Mandatory)] [string] $URI,
    [parameter()] [string] $CoverPhotoURI
)
    Function Parse-DVDMetadata {
    param (
        [parameter(Mandatory)] [string] $URI
    )
        $webResponse = Invoke-WebRequest $URI

        $hash = [ordered] @{
            title = ($webResponse.AllElements | where {$_.class -eq 'spacing-bottom'}).innerText
            studio = ($webResponse.AllElements | where {$_.class -eq 'studio'}).innerText
            actor = (($webResponse.AllElements | where {$_.class -eq 'boxcover girl'}).innerText -replace "`r`n"," ").trim() -join ","
            comment = $URI
        }

        New-Object psobject -Property $hash
    }
    $File = Get-ChildItem $File

    if (!(Test-Path $file.BaseName)) {
        $newHome = New-Item -ItemType Directory -Name $File.BaseName
    } else {
        $newHome = Get-ChildItem -Directory $File.BaseName
    }

    if ($CoverPhotoURI) {
        Invoke-WebRequest -Uri $CoverPhotoURI -OutFile "$newHome\$($file.BaseName).jpg"
    }

    $metadata = Parse-DVDMetadata -URI $URI
    Set-VideoMetadata $File -Title $metadata.Title -Genre DVD -Studio $metadata.studio -Actor $metadata.Actor -Comment $metadata.comment
    Move-Item $File.FullName -Destination $newHome
}



Function Import-Video {
param(
    [parameter(Mandatory)] [string] $URI,
    [string] $Archive
)
    if ($Archive){
        if (!(Test-Path $Archive)) {New-Item -Type File $Archive}
        $archiveFile = Get-ChildItem $Archive
        yt-dlp.exe --download-archive $archiveFile.FullName -o "%(title)s.%(ext)s" $URI
    }
    else {yt-dlp.exe -o "%(title)s.%(ext)s" $URI}
}
Function Import-Song {
param(
    [parameter(Mandatory)] [string] $URI
)
    yt-dlp.exe -x --audio-format mp3 -o "%(title)s.%(ext)s" $URI
}
Function Split-Photo {
param(
    [parameter(Mandatory,Position=0)][string]$Path,
    [switch]$Left,
    [switch]$DVD
)
    if (!(Test-Path .\encode)) {[void](mkdir encode)}
    try{$photo = Get-ChildItem $Path -ErrorAction Stop}
    catch {$_}
    if ($photo.Extension -notmatch '(jpg|png)') {Write-Error "Unable to crop $($photo.Name).  It is not an image file." ; return}
    
    if (!$DVD){
        if ($Left) {$filter = "crop=(iw/2):(ih):0:0"}
        else {$filter = "crop=(iw/2):(ih):(iw/2):0"}
    }
    else {
        if ($Left) {$filter = "crop=(iw*.47445):(ih):0:0"}
        else {$filter = "crop=(iw*.47445):(ih):(iw*(1-.47445)):0"}
    }

    ffmpeg -i $photo.FullName -vf $filter ".\encode\$($photo.name)" | Out-Host
    
    Remove-Item $photo.FullName
    Move-Item ".\encode\$($photo.Name)" -Destination $photo.Directory


}
#endregion
