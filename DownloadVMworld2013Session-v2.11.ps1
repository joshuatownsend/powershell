# VMworld Session Downloader
#
# v1.0 - Initial release 2010 by stu@vinternals.com
# v1.5 - Updated for 2011 Videos by Alan Renouf
# v2.0 - Updated for 2012 Videos by Alan Renouf
# v2.1 - added native download option and only 1 script instead of multiple.
# v2.2 - added new session IDS and number lengths
# v2.3 - added ability to specify session types for specific downloads
# v2.4 - added ability to download based on content type by Damian Karlson
# v2.5 - added ability to folder session contents together
#        Top level folder by three letter code and then
#        session number for content folder
#        Added trap on error if person signed in already
#        by Brent Quick
# v2.6 - Hopefully fixed HTML parsing/DOM differences between IE8 & IE9 by Damian Karlson
# v2.7 - Fixed a bug with determining IE version; thank you Andr?? Pett! -- Damian Karlson
# v2.8 - Modified the script to download all PDF and MP3 files
#        Added an option to skip already existing files (don't overwrite)
#        Added an option to interactively enter username and password
#        Added some verbose script output
#        $outputfolder can be an absolute path or relative to this script
#        Some minor fixes and optimizations
#        by Andr?? Pett - 11/11/2012
# v2.9 - replace illegal characters in mp4 session name - a.p. - 11/15/2012
# v2.10 - changes for 2013 content.
#		- Modified before EU show, in case EU is broked
#		- Added IE10 support
#		- Damian Karlson @sixfootdad
# v2.11 - some fixes for Windows 7 & IE10, PowerShell 2 & IE10 is still hit and miss. Upgrade to PowerShell 3 if you run into errors.
#		- Damian Karlson @sixfootdad

$ScriptVersion = "v2.11"

# TROUBLESHOOTING 
# If the script fails, make sure that you're logged out of vmworld.com and try again.
# On Windows 7 w/IE 10, vmworld.com MUST be in IE's Compatibility View list.
# If Windows 7 w/IE 10 are still broken after that, upgrade to PowerShell 3.

# Edit the following information before starting the script
# Warning these sessions take a lot of space!

$user = 'U$ERN@ME' # Your User ID for VMworld.Com
$pass = 'P@$$W0RD' # Your password for VMworld.com
$download = $false # Change to $false to just list the session titles and MP4 paths.
$skipexisting = $true # Change to $false to overwrite already existing files
$OutputFolder = 'C:\VMworld2013' # e.g. 'C:\Temp\' (absolute) or '.\SubFolder\' (relative)
$sessionRegex = "all" # See below for session types.
$downloadRegex = "all" # See below for download types
$CDO = $true # See below for impact

# The following session regex patterns can be used to download content by session types, 
# these are the first three or four chars in front of the VMworld session id.
#
# ALL = All Sessions
# BCO = Business Continuity
# EUC = End User Computing
# NET = Networking
# OPT = Operations
# PHC = Public Hybrid Cloud
# SEC = Security & Compliance
# TEX = Technology Exchange for Alliance Partner (You'll be unable to retrieve if login isn't tagged for TAP)
# VAPP = Virtualizing Applications
# VCM = Virtualization and Cloud Management
# VSVC = vSphere & vCloud Suite

# The following download regex patterns can be used to download specific content types
#
# MP4 = MP4 video file
# PDF = PDF file
# MP3 = MP3 audio file
# All = all content types

# The following used to enable CDO mode (not OCD since that is not alphabetical) which will build
# folder structure and place a single sessions files together.
# CDO = $true - unique folder for each session with all DL'ed contents together
# CDO = $false - single download folder

#
# DO NOT edit anything past this point.
#

function Get-VMworldMP4 {
	if ($mp4URL) {
			$tmpURL = $mp4URL
			$tmpFilename = $mp4Filename
			Get-VMworldFile
	} else {
		Write-Output "$sessionURL doesn't have a video listed."
	}
}

function Get-VMworldMP3 {
	# ensure $mp3URL is an array
	if (!($mp3URL -is [system.array])) {$mp3URL = @($mp3URL)}
	if ($mp3URL.Length -gt 0) {
		for ($i = 1; $i -le $mp3URL.Length; $i++) {
			$tmpURL = $mp3URL[$i - 1]
			$tmpFilename = $tmpURL.Substring($tmpURL.LastIndexOf("/") + 1)
			$tmpFilename = $tmpFilename.Replace(" ","_")
			Get-VMworldFile
		}
	} else {
		Write-Output "$sessionURL doesn't have an MP3 listed."
	}
}

function Get-VMworldPDF {
	# ensure $pdfURL is an array
	if (!($pdfURL -is [system.array])) {$pdfURL = @($pdfURL)}
	if ($pdfURL.Length -gt 0) {
		for ($i = 1; $i -le $pdfURL.Length; $i++) {
			$tmpURL = $pdfURL[$i - 1]
			$tmpFilename = $tmpURL.Substring($tmpURL.LastIndexOf("/") + 1)
			$tmpFilename = $tmpFilename.Replace(" ","_")
			Get-VMworldFile
		}
	} else {
		Write-Output "$sessionURL doesn't have a PDF listed."
	}
}

function Get-VMworldFile {
	if ($download) {
			Write-Output "Downloading: $tmpFilename from $tmpURL"
			$tmpFilename = $outputFolder + $tmpFilename
		if ((Test-Path ($tmpFilename)) -and ($skipexisting)) {
			Write-Output "File already exists."
		} else {
			$client = New-Object System.Net.WebClient
			$client.DownloadFile($tmpURL, $tmpFilename)
		}
	} else {
		Write-Output $tmpFilename
		Write-Output $tmpURL
	}
}

function makeCDOFolder {
	# Add a trailing slash in case it wasn't entered correctly
	if ($CDOoutputFolder -notmatch "[\\]$") {$CDOoutputFolder += "\"}    
	# Create the outputFolder if it doesn't exist
	if (!(Test-Path $CDOoutputFolder)){md $CDOoutputFolder | out-null}
}

### Start Main
$Host.UI.RawUI.WindowTitle = "VMworld Session Downloader [$ScriptVersion]"

# Enter Credentials unless already specified
if ($user -eq 'U$ERN@ME') {
	$user = Read-Host 'What is your VMworld.com username?'
}
if (! $user) {exit}
if ($pass -eq 'P@$$W0RD') {
	$pass = Read-Host 'What is your VMworld.com password?' -AsSecureString
	$pass = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))
}
if (! $pass) {exit}

# Create the outputFolder if it doesn't exist
if ($outputFolder.Substring(0,2) -eq ".\") {
	$outputFolder = ($MyInvocation.MyCommand.Path | Split-Path) + $OutputFolder.Substring(1)
}
write-host "Output folder: $OutputFolder" -foregroundcolor "magenta"
if (!(Test-Path $outputFolder)){md $outputFolder}
# Add a trailing slash in case it wasn't entered correctly
if ($outputFolder -notmatch "[\\]$") {$outputFolder += "\"}

write-host "Starting Internet Explorer" -foregroundcolor "magenta"
$loginURL = "http://www.vmworld.com/login.jspa"
$ie = New-Object -com InternetExplorer.Application
$ieVersion = ($ie.FullName | Get-ChildItem).VersionInfo.ProductVersion
write-host "Detected IE Version $ieVersion" -foregroundcolor "magenta"

# login via IE because using System.Net.WebClient with cookie auth was just too hard for me
write-host "Opening $loginURL" -foregroundcolor "magenta"
$ie.visible=$true
$ie.navigate($loginURL)
while($ie.ReadyState -ne 4) {start-sleep -m 100}
$checkIEUser = $ie.document.getElementByID("username01")
$checkIEPass = $ie.document.getElementByID("password01")
$checkIEForm = $ie.document.getElementByID("loginform01")
if ($checkIEUser -eq $null -or $checkIEPass -eq $null -or $checkIEform -eq $null) {
	#Do nothing logged in
	Write-Host "User already logged in." -ForegroundColor "magenta"
} else {
	$ie.document.getElementById("username01").value=$user
	$ie.document.getElementById("password01").value=$pass
	$ie.document.getElementById("loginform01").submit()
}
sleep -Seconds 10
# pull session listing page into a string
while($ie.ReadyState -ne 4) {start-sleep -m 100}
write-host "Reading Session list" -foregroundcolor "magenta"
$sessionList = $ie.Navigate("http://www.vmworld.com/community/sessions/2013/")
while($ie.ReadyState -ne 4) {start-sleep -m 100}
if ($ieVersion -like "8.*") { #it's IE8
	$sessionList = [string]$ie.Document.body.OuterHtml
} elseif (($ieVersion -like "9.*") -or ($ieVersion -like "10.*")) { #it's IE9 or 10
	$sessionList = [string]$ie.Document.documentElement.outerHTML
}

# regex to pull out session number
$sessionRegex = $sessionRegex.ToUpper()
if ($sessionRegex -eq "ALL") {
	$sessionRegex = [regex] "(?s)(BCO|EUC|NET|OPT|PHC|SEC|STO|TEX|VAPP|VCM|VSVC)\d{4}.*?/docs/DOC-\d{4}"
} else {
	$sessionRegex = [regex] "(?s)($sessionRegex)\d{4}.*?/docs/DOC-\d{4}"
}
$sessionMatch = $sessionRegex.matches($sessionList)
# loop through matches
if ($sessionMatch.Success) {
	foreach ($match in $sessionMatch.value) {
		# get the doc ID
		$docID = [regex]::Match($match,'[DOC]{3}-[0-9]{4}').Value
		# session number will be first 7 or 8 characters of a match, depending on the session id abbreviation
		$sessionID = [regex]::Match($match,'^[a-zA-Z0-9]+').Value
		# Remove bogus characters for sessions with only 3 chars in the session ID
		$sessionID = $sessionID.Replace("<","")
		# build session page URL
		$sessionURL = "http://www.vmworld.com/docs/" + $docID
		# open session page
		$ie.navigate($sessionURL)
		while($ie.ReadyState -ne 4) {start-sleep -m 100}
		# grab necessary details from the session page
		$sessionTitle = $ie.Document.Title
		write-host "$sessionTitle" -foregroundcolor "magenta"
		if ($ieVersion -like "8.*") { #it's IE8
			$sessionURLSource = [string]$ie.Document.body.OuterHtml
		} elseif (($ieVersion -like "9.*") -or ($ieVersion -like "10.*")) { #it's IE9 or 10
			$sessionURLSource = [string]$ie.Document.documentElement.outerHTML
		}
		# regex to pull the MyLearn Class ID
		$classIDRegex = [regex] "classID=\d{6}"
		$classMatch = $classIDRegex.match($sessionURLSource)
		if ($classMatch.Success)
		{        
			# ClassID will be last 5 chars of a match
			$classID = [regex]::Match($classMatch.value,'[0-9]+').Value
			if ($outputFolder -notmatch "[\\]$") {$outputFolder += "\"}
			$CDOoutputFolder = $outputFolder + [regex]::Match($sessionid,'[a-zA-Z]+').Value + "\" + $sessionID +"\"
			if ($CDO) {
				$tmpOutputFolder = $outputFolder
				makeCDOFolder
				$outputFolder = $CDOoutputFolder
			}        
			# Construct text showing session name and direct MP4 download link
			$mp4URL = "http://sessions.vmworld.com/lcms/mL_course/courseware/" + $classID + "/" + $sessionID + ".mp4"
			$mp4Filename = $sessionTitle.substring(13) + " - " + $sessionID + ".mp4"
			$mp4Filename = [regex]::Replace($mp4Filename, '[\\/:*?><|"]', ".")

			# we have to go to the session theater page since the PDF & MP3 URLs aren't predictable
			$ie.Navigate("http://vmworld.com/mylearn?classID=" + $classID)
			while($ie.ReadyState -ne 4) {start-sleep -m 100}
			if ($ieVersion -like "8.*") { #it's IE8
				$theaterSource = [string]$ie.Document.body.OuterHtml
			} elseif (($ieVersion -like "9.*") -or ($ieVersion -like "10.*")) { #it's IE9 or 10
				$theaterSource = [string]$ie.Document.documentElement.outerHTML
			}
			# Extract URLs for all PDF and MP3 files
			$pdfRegex = '(?<=href=\").+\.pdf(?=\")'
			$pdfURL = $theaterSource | select-string -Pattern $pdfRegex -AllMatches | % { $_.Matches } | % { $_.Value }
			$mp3Regex = '(?<=href=\").+\.mp3(?=\")'
			$mp3URL = $theaterSource | select-string -Pattern $mp3Regex -AllMatches | % { $_.Matches } | % { $_.Value }
			
			# Navigate to SessionURL to avoid automatic MP4 download
			$ie.navigate($sessionURL)
			
			switch ($downloadRegex) {
				MP4 {
					Get-VMworldMP4
				}
				MP3 {
					Get-VMworldMP3
				}
				PDF {
					Get-VMworldPDF
				}
				All {
					Get-VMworldMP4
					Get-VMworldMP3
					Get-VMworldPDF
				}
			}
			if ($CDO){
				$outputFolder = $tmpOutputFolder
			}
		}
	}
}