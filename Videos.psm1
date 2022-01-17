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

Video([System.IO.FileInfo]$File) {
    $foo = ffprobe -v error -hide_banner -of default=noprint_wrappers=0 -print_format json -select_streams v:0 -show_format $File | ConvertFrom-Json
    $this.Basename = $File.BaseName
    $this.Name = $File.Name
    $this.FullName = $File.FullName
    $this.Extension = $File.Extension
    $this.Directory = $File.Directory
    $this.Title = $foo.format.tags.title
    $this.Genre = $foo.format.tags.genre
    $this.Studio = $foo.format.tags.album
    $this.Actor = $foo.format.tags.artist
    $this.Date = $foo.format.tags.date
}
Video([String]$Path) {
    $File = Get-Item $Path
    $foo = ffprobe -v error -hide_banner -of default=noprint_wrappers=0 -print_format json -select_streams v:0 -show_format $File | ConvertFrom-Json
    $this.Basename = $File.BaseName
    $this.Name = $File.Name
    $this.FullName = $File.FullName
    $this.Extension = $File.Extension
    $this.Directory = $File.Directory
    $this.Title = $foo.format.tags.title
    $this.Genre = $foo.format.tags.genre
    $this.Studio = $foo.format.tags.album
    $this.Actor = $foo.format.tags.artist
    $this.Date = $foo.format.tags.date
}

[void] EncodeH265(){
    if (!(Test-Path .\encode)) {[void](mkdir encode)}
    ffmpeg -i $this.FullName -c:v hevc -preset veryslow ".\encode\$($this.name)" | Out-Host
    Remove-Item $this.FullName
    Move-Item ".\encode\$($this.Name)" -Destination $this.Directory
}

# ParseName()
#
# Parse video information/metadata from file name
#
[Hashtable] ParseName() {
    $withEpisodeString = '(?<studio>\w+)\.(?<episode>\w{5})\.(?<title>.*)'
    $maybeDateString = '(?<studio>\w+)\.(?<date>(?<year>\d\d)\.(?<month>\d\d)\.(?<day>\d\d))?\.?(?<title>.*)'
    $narcosString = 'narcos-(?<studio>gx)-(?<date>(?<year>\d\d)-(?<month>\d\d)-(?<day>\d\d))-(?<title>.*)-\d+p'
            
    $this.Basename -match $withEpisodeString
    $this.Basename -match $maybeDateString
    $this.Basename -match $narcosString
        
    return $Matches
}

[string] ToString() {
    return $this.FullName
}

# WriteMetadata()
#
# Write video metadata to object. Requires FFMPEG.
# Todo: create check to prevent writing metadata even if nothing has changed

[void] WriteMetadata() {
    if (!(Test-Path .\encode)) {[void](mkdir encode)}
    #$video = [Video]::new((Get-ChildItem $Path))

    ffmpeg -hide_banner -i $this.FullName -metadata genre=$($this.Genre) -metadata artist=$($this.Actor) -metadata title=$($this.Title) -metadata album=$($this.Studio) -metadata date=$($this.Date) -c copy ".\encode\$($this.name)" | Out-Host
    Remove-Item $this.FullName
    Move-Item ".\encode\$($this.Name)" -Destination $this.Directory
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
param ([switch]$Override)
foreach ($File in (Get-ChildItem -File)) {
    $video = Get-Video $File
    $video
    if ($video.actor -and (!$Override.IsPresent)) {Continue}
    $Actor = Read-Host "Actor Name(s)"

    if ([string]$Actor -eq "skip") {Continue}
    add-videometadata $video -actor $Actor
}
}

Function Set-VideoTitle {
param(
    [parameter(Mandatory,ValueFromPipeline)]
    $Path
)
Process {
    $video = [Video]::new((Get-ChildItem $Path))
    $newName = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($($video.ParseName().title -replace '\W',' '))
    $video.Title = $newName
    $video.studio = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($video.parsename().studio)
    if ($video.parsename().date) {
        $video.date = "$($video.parsename().month).$($video.parsename().day).$($video.parsename().year)"
    }
    $video.WriteMetadata()
}
}
Function Set-VideoMetadata {
param(
    [parameter(Mandatory,ValueFromPipeline)]
    $Path,
    [string]$Genre,
    [string]$Title,
    [string]$Actor
)
Process {
    $video = [Video]::new((Get-ChildItem $Path))
    $video.Genre = $Genre
    $video.Title = $Title
    $video.Actor = $Actor
    if ($Title) {$video.Title = $Title}
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
    #if ($newGenre -eq $($oldGenre -join ',')) {
    #    Write-Warning "The command completed successfully, but no changes were made to $($video.Name)"
    #    return
    #}
    $video.Genre = $newGenre

    [array]$oldActor = $video.Actor.Trim() -split ','
    [array]$newActor = $Actor.Trim() -split ','
    $newActor = ([System.Linq.Enumerable]::Distinct([string[]]$($newActor + $oldActor)) -join ',') -replace '(^,)?(,$)?','' ## -replace prevents extra commas

    ## this check needs to be moved to WriteMetadata() method

    #if ($newActor -eq $($oldActor -join ',')) {
    #    Write-Warning "The command completed successfully, but no changes were made to $($video.Name)"
    #    return
    #}
    $video.Actor = $newActor
    $video.WriteMetadata()
}
}
Function Get-VideoScreenshots {
param(
    [parameter(Mandatory,ValueFromPipeline)]
    $Path,
    $Destination = "X:\Thumbs",
    [string] $SceneFilter = "0.4",
    [string] $Grid = "4x4",
    [string] $SS = "5"
)
    $video = [Video]::new((Get-ChildItem $Path))
    #ffmpeg -ss 3 -i $video.FullName -vf "select=gt(scene\,0.5)" -frames:v 5 -vsync vfr "$Destination\$($video.Basename)%02d.jpg"
    #ffmpeg -ss 3 -i $video.FullName -vf "select=gt(scene\,0.4),fps=1/600" -frames:v 5 -vsync vfr "$Destination\$($video.Basename)%02d.jpg"
    #ffmpeg -i $video.FullName -vf "select=gt(scene\,0.25),scale=100:-1,tile=5x5" -frames:v 1 -qscale:v 3 "$Destination\$($video.Basename).jpg"
    ffmpeg -ss $SS -i $video.FullName -vf "select=gt(scene\,$SceneFilter),scale=500:-1,tile=$Grid" -frames:v 1 -qscale:v 3 "$Destination\$($video.Basename).jpg"
}
#endregion



#region Classless functions
Function Import-Video {
param(
    [parameter(Mandatory)]
    [string]$URI,
    $Archive
)
    if ($Archive){
        $archiveFile = Get-ChildItem $Archive
        yt-dlp.exe --download-archive $archiveFile.FullName -o "%(title)s.%(ext)s" $URI
    }
    else {yt-dlp.exe -o "%(title)s.%(ext)s" $URI}
}
#endregion
