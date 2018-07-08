#processes media files so they are organized into a directory structure by date.  Photo files are automatically rotated to the correct orientation 
#and Video files are optionally reencoded to fix their orientation
#set-executionpolicy remotesigned
#.\fixmedia.ps1 -DestinationDirectory 'M:\output' -SourceDirectory 'D:\Phone\Video' -processVideoRotation 0
#.\fixmedia.ps1 -DestinationDirectory 'D:\output' -SourceDirectory 'E:\imported' -processVideoRotation 0 -suppressVideoRenaming 1

param (
		[parameter(Mandatory=$true)]
		[ValidateScript({Test-Path $_ -PathType 'Container'})] 
        [string]$DestinationDirectory,      #TODO: add support for optional just destination directory so images arent copied
        [bool]$processVideoRotation= $true,
   		[parameter(Mandatory=$true)]
   		[ValidateScript({
            If(Test-Path $_){$true}else{Throw "Invalid path given: $_"}
            If($_.CompareTo($DestinationDirectory)){$true}else{Throw "Source path cannot be same as destination: $_"}
        })] 
        [string]$SourceDirectory,
        [string]$ffmpegpath="M:\Os Restore\ffmpeg-20160105-git-fc703f5-win64-static\bin\ffmpeg.exe",
        [string]$LogFilePath="C:\fixmedia.log",
        [bool]$suppressVideoRenaming= 0
 )

Add-Type -Assembly 'System.Drawing'

$script:ErrorActionPreference = "Stop"

#Add cwd to env so load library searches through it
$CurrentValue = [Environment]::GetEnvironmentVariable("PSModulePath", "Machine")
[Environment]::SetEnvironmentVariable("PSModulePath", $CurrentValue + ";" + $(Get-Location).Path, "Machine")


#http://blogs.technet.com/b/jamesone/archive/2010/07/05/exploring-the-image-powershell-module.aspx
Import-Module Image

<# 
.Synopsis 
   Write-Log writes a message to a specified log file with the current time stamp. 
.DESCRIPTION 
   The Write-Log function is designed to add logging capability to other scripts. 
   In addition to writing output and/or verbose you can write to a log file for 
   later debugging. 
.NOTES 
   Created by: Jason Wasser @wasserja 
   Modified: 11/24/2015 09:30:19 AM   
 
   Changelog: 
    * Code simplification and clarification - thanks to @juneb_get_help 
    * Added documentation. 
    * Renamed LogPath parameter to Path to keep it standard - thanks to @JeffHicks 
    * Revised the Force switch to work as it should - thanks to @JeffHicks 
 
   To Do: 
    * Add error handling if trying to create a log file in a inaccessible location. 
    * Add ability to write $Message to $Verbose or $Error pipelines to eliminate 
      duplicates. 
.PARAMETER Message 
   Message is the content that you wish to add to the log file.  
.PARAMETER Path 
   The path to the log file to which you would like to write. By default the function will  
   create the path and file if it does not exist.  
.PARAMETER Level 
   Specify the criticality of the log information being written to the log (i.e. Error, Warning, Informational) 
.PARAMETER NoClobber 
   Use NoClobber if you do not wish to overwrite an existing file. 
.EXAMPLE 
   Write-Log -Message 'Log message'  
   Writes the message to c:\Logs\PowerShellLog.log. 
.EXAMPLE 
   Write-Log -Message 'Restarting Server.' -Path c:\Logs\Scriptoutput.log 
   Writes the content to the specified log file and creates the path and file specified.  
.EXAMPLE 
   Write-Log -Message 'Folder does not exist.' -Path c:\Logs\Script.log -Level Error 
   Writes the message to the specified log file as an error message, and writes the message to the error pipeline. 
.LINK 
   https://gallery.technet.microsoft.com/scriptcenter/Write-Log-PowerShell-999c32d0 
#> 
function Write-Log 
{ 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
 
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path='C:\Logs\PowerShellLog.log', 
         
        [Parameter(Mandatory=$false)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info", 
         
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 
 
    Begin 
    { 
        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        $VerbosePreference = 'Continue' 
    } 
    Process 
    { 
         
        # If the file already exists and NoClobber was specified, do not write to the log. 
        if ((Test-Path $Path) -AND $NoClobber) { 
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
            } 
 
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        elseif (!(Test-Path $Path)) { 
            Write-Verbose "Creating $Path." 
            $NewLogFile = New-Item $Path -Force -ItemType File 
            } 
 
        else { 
            # Nothing to see here yet. 
            } 
 
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
 
         # Write log entry to $Path 
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
        
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) { 
            'Error' { 
                Write-Error $Message 
                $LevelText = 'ERROR:' 
                } 
            'Warn' { 
                Write-Warning $Message 
                $LevelText = 'WARNING:' 
                } 
            'Info' { 
                Write-Verbose $Message 
                $LevelText = 'INFO:' 
                } 
            } 
    } 
    End 
    { 
    } 
}
  
function PSUsing
{
    param
    (
        [IDisposable] $disposable,
        [ScriptBlock] $scriptBlock
    )
 
    try
    {
        & $scriptBlock
    }
    finally
    {
        if ($disposable -ne $null)
        {
            $disposable.Dispose()
        }
    }
}
 
function Set-ExifRotation
{
    param
    (
        [int16] $RotationValue,
        [string] $ImagePath
    )

    $fullPath = (Resolve-Path $ImagePath).Path
 	$tempFile = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($ImagePath), [System.IO.Path]::GetRandomFileName())

    if (Test-Path -Path $tempFile)
    {
        Remove-Item $tempFile 
    }

    try
    {
        PSUsing ($fs = [System.IO.File]::OpenRead($fullPath)) `
        {
         
            PSUsing ($image = [System.Drawing.Image]::FromStream($fs, $false, $false)) `
            {
                #fix rotation
                #http://superuser.com/questions/424474/standard-application-to-automatically-rotate-pictures-based-on-exif

                #horizontal
                if ( $RotationValue -ne 1)
                {

                    #180
                    if ($RotationValue -eq 3)
                    {
                         $image.rotateflip("Rotate180FlipNone")
                    }
                    #90
                    elseif ($RotationValue -eq 6)
                    {
                         $image.rotateflip("Rotate90FlipNone")
                    }
                    #270
                    elseif ($RotationValue -eq 8)
                    {
                         $image.rotateflip("Rotate270FlipNone")
                    }
                    $image.Save($tempFile)
                }
            }
        }

        if ( $RotationValue -ne 1)
        {
            Remove-Item $fullPath 
            Rename-Item $tempFile $fullPath 
        }
    }   
    finally
    {


    }

    return $null
}

function Get-VideoRotation
{
    param
    (
        [string] $ImagePath
    )
        
    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $ffmpegpath 
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = "-i `"" + $ImagePath + "`""
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()
        
    #parse out "rotate          : XX"
    $rotate = $null;
    if ($stderr-match "rotate *: (.*?)\n") { 
      $rotate = $matches[1]
    }
      
    return $rotate;
}

function Get-VideoCreatedDate
{
    param
    (
        [string] $ImagePath
    )
        
    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $ffmpegpath 
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = "-i `"" + $ImagePath + "`""
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()
        
    #parse out "creation_time          : XX"
    $date = $null;
    if ($stderr-match "creation_time *: (.*?)\n") { 
      $date = $matches[1]
    }

    if ($date -eq $null)
    {
        return $null
    }

    $dateTime = [DateTime]::MinValue
    if ([DateTime]::TryParseExact($date.Trim(), "yyyy-MM-dd HH:mm:ss", $null, [System.Globalization.DateTimeStyles]::None, [ref] $dateTime))
    {
        return $dateTime
    }

    #fall back to file modified time
    $msg =  [string]::Format("Unable to get created date defaulting to modified timestamp for: {0}",  $ImagePath)
    Write-Log -Message $msg -Path $LogFilePath -Level WARN
    return (Get-Item $ImagePath).LastWriteTime
}

function Set-VideoRotation
{
    param
    (
        [int16] $RotationValue,
        [string] $ImagePath
    )
    
    $fullPath = (Resolve-Path $ImagePath).Path
    $tempFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($fullPath), [System.IO.Path]::GetFileNameWithoutExtension([System.IO.Path]::GetRandomFileName()) + ".mp4")
    $newFileName = [System.IO.Path]::GetFileNameWithoutExtension($fullPath) + ".mp4"
    $newFilePath  =  [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($fullPath), $newFileName)
    $transposeValue = $null 
            
    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $ffmpegpath 
    $pinfo.RedirectStandardError = $false
    $pinfo.RedirectStandardOutput = $false
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = "-i `"" + $fullPath + "`" -c:a copy `"" + $tempFilePath + "`""
    $p = New-Object System.Diagnostics.Process
    $pinfo.CreateNoWindow = $true;
    
    $p.StartInfo = $pinfo
    $result = $p.Start() 
    $result = $p.WaitForExit()

    if (Test-Path -Path $newFilePath)
    {
        Remove-Item $newFilePath 
    }
    rename-item -path  $tempFilePath -newname $newFileName 

    #remove original if its not the same name as the newly created file
    if ($newFilePath -ne $ImagePath) 
    {
        if (Test-Path -Path $ImagePath)
        {
            Remove-Item $ImagePath
        }
    }

    return $newFilePath;
}

function CopyAllFilesFromSourceToDestinationFolder
{
    param
    (
        [string] $DstDir,
        [string] $SrcDir
    )
	$srcDirFilter = [io.path]::combine($SrcDir, "*.*")

    $copy = Copy-Item -Path $srcDirFilter -Destination $DstDir -recurse -Force
    return $copy
}

function GetExifData
{
    param
    (
        [string] $Path
    )
    
    $Image  =  Get-Image -Path $Path 
    $ExifData = Get-Exif -image $Image 
    $DateTaken = $null
    $Orientation = $null
    $GPS = $null
    
    $dateTime = [DateTime]::MinValue
    if ($ExifData.DateTaken -ne $null)
    {
         $DateTaken = $ExifData.DateTaken
    }
    else
    {
         #fall back to file modified time
        $msg =  [string]::Format("Unable to get created date defaulting to modified timestamp for: {0}",  $Path)
        Write-Log -Message $msg -Path $LogFilePath -Level WARN
        $DateTaken = (Get-Item $Path).LastWriteTime
    }
   
    $Orientation = $ExifData.Orientation
    
    if ($ExifData.GPS -ne $null)
    {
        
        $GPS = ConvertDMSToDecimalDegrees -Dmslocation  $ExifData.GPS #chicago 41.8369° N, 87.6847° 
    }

    New-Object PSObject -Property @{
        DateTaken = $DateTaken
        Orientation = $Orientation
        Latitude = $GPS.Latitude
        Longitude = $GPS.Longitude
        Altitude = $GPS.Altitude
    }
}

function ConvertDMSToDecimalDegrees
{
    param
    (
        [string] $Dmslocation
    )
    $DecimalNotaion = $null
    $LatLong = $null
    $Altitude = $null
    $Latitude = $null
    $Longitude = $null
    
    #parse out long, lat, alt from string in format 41°38'16.76"N  88°14'44.12"W, 215.78612716763M above Sea Level
        
    $Dmslocation.Split(",") | ForEach `
    {
        if ($_.Contains("°"))
        {
            $LatLong = $_.Trim() 
        }
        elseif ($_.Contains("M"))
        {
            $trim = $_.Trim()
            if ($trim -match "^-?\d+(?:\.\d+)?") 
            { 
                $Altitude= $matches[0].Trim()
            }
        }
    }

    if ($LatLong -ne $null)
    {
        $Dmslocation.Split(" ") | ForEach `
        {
            if ($_.Contains("°"))
            {
                if ( $Latitude -eq $null)
                {
                    $split = $_.Split("°'""")
                    #S and W need to be negative

                    $Latitude = ([decimal]$split[0])+([decimal]($split[1]/60))+([decimal]($split[2]/3600))

                    if ($_.Contains("S") -or $_.Contains("W"))
                    {
                        $Latitude = $Latitude * -1
                    }
                }
                elseif ( $Longitude -eq $null)
                {
                    $split = $_.Split("°'""")
                    $Longitude = ([decimal]$split[0])+([decimal]($split[1]/60))+([decimal]($split[2]/3600))
                    
                    if ($_.Contains("S") -or $_.Contains("W"))
                    {
                        $Longitude = $Longitude * -1
                    }
                }
            }
        }
    }

    New-Object PSObject -Property @{
        Latitude = $Latitude
        Longitude = $Longitude
        Altitude = [decimal]($Altitude)
    }
}

function GetImageLocation
{
    param
    (
        [string] $Longitude,
        [string] $Latitude
    )
   
#TODO: add caching to avoid rpc calls if we already have the information for the set of coordinates   
       
    $url = "https://maps.googleapis.com/maps/api/geocode/json?latlng=$Latitude,$Longitude&amp;sensor=true&amp;language=en"
    $json = Invoke-WebRequest -Uri $url | ConvertFrom-JSON
    if ($json.status -ne "OK")
    {
        if ($json.status -ne "ZERO_RESULTS")
        { 
            $msg =  [string]::Format("Reverse geocoding failed with {0} for url {1}", $json.status, $url)
            Write-Log -Message $msg -Path $LogFilePath -Level ERROR
            return $null
        }
        else
        {
            #TODO: handle ZERO_RESULTS properly
            $msg =  [string]::Format("Reverse geocoding failed with {0} for url {1}", $json.status, $url)
            Write-Log -Message $msg -Path $LogFilePath -Level WARN
        }
    }
    else
    {
        #use locality, political attrbute to describe location
        $shortLocation = $json.results[0].address_components | where { $_.types.Contains("locality") } | Select -First 1
        return $shortLocation.short_name
    }
}

function ProcessImageFiles
{
    param
    (
        [string] $DstDir
    )

    $imagelist = @("*.jpeg","*.gif","*.jpg")

    $dstDirFilter = [io.path]::combine($DstDir, "*.*")
    $dstFiles = Get-ChildItem $dstDirFilter -recurse -include $imagelist  
    $curIndex = 1
    foreach ($File in ($dstFiles))
    {
        $msg =  [string]::Format("Processing image {0} of {1} {2}", $curIndex, $dstFiles.Count, $File.FullName)
        Write-Log -Message $msg -Path $LogFilePath -Level INFO

        $curIndex = $curIndex + 1
        $FileName = $File.Name
        $ExifData = GetExifData -Path $File.FullName  
        $ImageLocation = $null

        if ($ExifData.Longitude -ne $null -and $ExifData.Latitude -ne $null)
        {
            $ImageLocation = GetImageLocation -Longitude $ExifData.Longitude -Latitude $ExifData.Latitude
        }

        if ($ExifData.Orientation -ne $null -and $ExifData.Orientation  -ne 1)
        {
            Set-ExifRotation -ImagePath $File.FullName -RotationValue $ExifData.Orientation 
        }
        else
        {
            if (![string]::IsNullOrEmpty($ExifData.Orientation))
            {
                $msg =  [string]::Format("Skipped rotating {0} RotateValue: {1}", $File.FullName, $ExifData.Orientation)
                Write-Log -Message $msg -Path $LogFilePath -Level WARN
            }
        }

        #move image files in subdirectory eg. root/year/month day - imagelocation/filename
        if ($ImageLocation -eq $null)
        {
            $ImageLocation = " - " + "No Location"
            $msg =  [string]::Format("Unable to get location for: {0}", $File.FullName)
            Write-Log -Message $msg -Path $LogFilePath -Level WARN
        }
        else
        {
            $ImageLocation = " - " + $ImageLocation 
        }
        
        $newDirPath = [io.path]::combine($DestinationDirectory, "Image", $ExifData.DateTaken.Year, [string]::Format("{0} {1} {2}", $ExifData.DateTaken.Month, $ExifData.DateTaken.Day, $ImageLocation))
        if (!(Test-Path -Path $newDirPath))
        {
            $new = New-Item $newDirPath -type directory
        }
    
        $origfilePath = [io.path]::combine($DestinationDirectory, $FileName )
        $move = Move-Item $origfilePath $newDirPath
   
    }
    return $null
}

function ProcessVideoFiles
{
    param
    (
        [string] $DstDir
    )

    $videolist = @("*.mov", "*.mp4", "*.avi")
    
    $dstDirFilter = [io.path]::combine($DstDir, "*.*")
    $dstFiles = Get-ChildItem $dstDirFilter -recurse -include $videolist  
    $curIndex = 1
    foreach ($File in ($dstFiles))
    {
        $msg =  [string]::Format("Processing video {0} of {1} {2}", $curIndex, $dstFiles.Count, $File.FullName)
        Write-Log -Message $msg -Path $LogFilePath -Level INFO
        $curIndex = $curIndex + 1

        $RotateValue = Get-VideoRotation -ImagePath $File.FullName 
        $CreatedDate = Get-VideoCreatedDate -ImagePath $File.FullName 
        $FilePath  = $File.FullName
        $FileName = $File.Name 
    
        if ($RotateValue -ne $null -and $RotateValue -ne "0" -and $processVideoRotation)
        {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $FilePath  = Set-VideoRotation -ImagePath $File.FullName -RotationValue $RotateValue 
            $sw.Stop()
            $elapsed = $sw.Elapsed
            $msg =  [string]::Format("Took {0} sec to rotate {1}", $elapsed.TotalSeconds, $FilePath)
            Write-Log -Message $msg -Path $LogFilePath -Level INFO 
        }
        else
        {
            if (![string]::IsNullOrEmpty($RotateValue))
            {
                $msg =  [string]::Format("Skipped rotating {0} RotateValue: {1}", $File.FullName, $RotateValue)
                Write-Log -Message $msg -Path $LogFilePath -Level WARN 
            }
        }

        if ($CreatedDate -ne $null -and $suppressVideoRenaming -eq 0)
        {
            $NewDatePrefix = $CreatedDate.ToString("MM dd ")
            $newFileName = $NewDatePrefix + [System.IO.Path]::GetFileName($FilePath)
            $newFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($File.FullName), $newFileName)
            if (Test-Path -Path $newFilePath)
            {
                Remove-Item $newFilePath 
            }
            rename-item -path $filePath -newname $newFileName 
            $FileName = $newFileName
        }
        else
        {
			if ($CreatedDate -eq $null)
						{
            	$msg =  [string]::Format("Failed renaming {0} CreatedDate: {1}", $File.FullName, $CreatedDate)
            	Write-Log -Message $msg -Path $LogFilePath -Level ERROR
            }
            if ($suppressVideoRenaming -eq 1)
						{
            	$msg =  [string]::Format("Not renaming {0}", $File.FullName)
            	Write-Log -Message $msg -Path $LogFilePath -Level INFO
            }
        }
        
        #move image files in subdirectory eg. root/year/filename
        $newDirPath = [io.path]::combine($DestinationDirectory, "Video", $CreatedDate.Year)
        if (!(Test-Path -Path $newDirPath))
        {
            $new = New-Item $newDirPath -type directory
        }
        $origfilePath = [io.path]::combine($DestinationDirectory, $FileName)
        $move = Move-Item $origfilePath $newDirPath
    }  
    return $null
}

    If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
        [Security.Principal.WindowsBuiltInRole] "Administrator"))
    {
        $msg = "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
        Write-Log -Message $msg -Path $LogFilePath -Level ERROR
         exit 1
    }
     
    trap { 
        $msg =  Out-String -InputObject (format-list -force -InputObject $Error[0])
        #TODO: to be an info not an error since error outputing is disabled in a trap
        Write-Log -Message $msg -Path $LogFilePath -Level INFO
        exit 1
    }
        
    CopyAllFilesFromSourceToDestinationFolder -DstDir $DestinationDirectory -SrcDir $SourceDirectory

    ProcessImageFiles -DstDir $DestinationDirectory

    ProcessVideoFiles -DstDir $DestinationDirectory
    

 


    




 
