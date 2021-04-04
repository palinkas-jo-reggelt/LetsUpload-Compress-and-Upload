<#

.SYNOPSIS
	LetsUpload Backup Utility

.DESCRIPTION
	Compresses and uploads folder contents to LetsUpload.io

.FUNCTIONALITY
	Compresses and uploads folder contents to LetsUpload.io

.PARAMETER UploadFolder
	Specifies the folder on local filesystem to compress and upload
	
.PARAMETER UploadName
	Specifies the name (description) of the archive to be created as well as letsupload folder name
	
.NOTES
	Create account and get API keys from https://www.letsupload.io then fill in $APIKey variables under USER VARIABLES
	Run from task scheduler daily
	Windows only
	API: https://letsupload.io/api.html
	Install latest 7-zip and put into system path
	
.EXAMPLE
	PS C:\Users\username> C:\scripts\letsupload.ps1 "C:\Path\To\Folder\To\Backup" "Backup Description (email, work, etc)"

#>

Param(
	[Parameter(Mandatory=$True)]
	[ValidatePattern("^[A-Z]\:\\")]
	[String]$UploadFolder,

	[Parameter(Mandatory=$False)]
	[String]$UploadName
)

<###   CONFIG   ###>
Try {
	.("$PSScriptRoot\LetsUploadConfig.ps1")
}
Catch {
	Write-Output "$(Get-Date) -f G) : ERROR : Unable to load supporting PowerShell Scripts : $query $(Error[0])" | out-file "$PSScriptRoot\PSError.log" -append
}

<###   FUNCTIONS   ###>
Function Debug ($DebugOutput) {
	If ($VerboseFile) {Write-Output "$(Get-Date -f G) : $DebugOutput" | Out-File $DebugLog -Encoding ASCII -Append}
	If ($VerboseConsole) {Write-Host "$(Get-Date -f G) : $DebugOutput"}
}

Function Email ($EmailOutput) {
	If ($UseHTML){
		If ($EmailOutput -match "\[OK\]") {$EmailOutput = $EmailOutput -Replace "\[OK\]","<span style=`"background-color:green;color:white;font-weight:bold;font-family:Courier New;`">[OK]</span>"}
		If ($EmailOutput -match "\[INFO\]") {$EmailOutput = $EmailOutput -Replace "\[INFO\]","<span style=`"background-color:yellow;font-weight:bold;font-family:Courier New;`">[INFO]</span>"}
		If ($EmailOutput -match "\[ERROR\]") {$EmailOutput = $EmailOutput -Replace "\[ERROR\]","<span style=`"background-color:red;color:white;font-weight:bold;font-family:Courier New;`">[ERROR]</span>"}
		If ($EmailOutput -match "^\s$") {$EmailOutput = $EmailOutput -Replace "\s","&nbsp;"}
		Write-Output "<tr><td>$EmailOutput</td></tr>" | Out-File $EmailBody -Encoding ASCII -Append
	} Else {
		Write-Output $EmailOutput | Out-File $EmailBody -Encoding ASCII -Append
	}	
}

Function EmailResults {
	Try {
		$Body = (Get-Content -Path $EmailBody | Out-String )
		If (($AttachDebugLog) -and (Test-Path $DebugLog) -and (((Get-Item $DebugLog).length/1MB) -lt $MaxAttachmentSize)){$Attachment = New-Object System.Net.Mail.Attachment $DebugLog}
		$Message = New-Object System.Net.Mail.Mailmessage $EmailFrom, $EmailTo, $Subject, $Body
		$Message.IsBodyHTML = $UseHTML
		If (($AttachDebugLog) -and (Test-Path $DebugLog) -and (((Get-Item $DebugLog).length/1MB) -lt $MaxAttachmentSize)){$Message.Attachments.Add($DebugLog)}
		$SMTP = New-Object System.Net.Mail.SMTPClient $SMTPServer,$SMTPPort
		$SMTP.EnableSsl = $SSL
		$SMTP.Credentials = New-Object System.Net.NetworkCredential($SMTPAuthUser, $SMTPAuthPass); 
		$SMTP.Send($Message)
	}
	Catch {
		Debug "Email ERROR : $($Error[0])"
	}
}

Function Plural ($Integer) {
	If ($Integer -eq 1) {$S = ""} Else {$S = "s"}
	Return $S
}

Function ElapsedTime ($EndTime) {
	$TimeSpan = New-Timespan $EndTime
	If (([int]($TimeSpan).Hours) -eq 0) {$Hours = ""} ElseIf (([int]($TimeSpan).Hours) -eq 1) {$Hours = "1 hour "} Else {$Hours = "$([int]($TimeSpan).Hours) hours "}
	If (([int]($TimeSpan).Minutes) -eq 0) {$Minutes = ""} ElseIf (([int]($TimeSpan).Minutes) -eq 1) {$Minutes = "1 minute "} Else {$Minutes = "$([int]($TimeSpan).Minutes) minutes "}
	If (([int]($TimeSpan).Seconds) -eq 1) {$Seconds = "1 second"} Else {$Seconds = "$([int]($TimeSpan).Seconds) seconds"}
	
	If (($TimeSpan).TotalSeconds -lt 1) {
		$Return = "less than 1 second"
	} Else {
		$Return = "$Hours$Minutes$Seconds"
	}
	Return $Return
}

Function MakeArchive {
	$StartArchive = Get-Date
	Debug "----------------------------"
	Debug "Create archive : $BackupName"
	Debug "Archive folder : $UF"
	$VolumeSwitch = "-v$VolumeSize"
	$PWSwitch = "-p$ArchivePassword"
	Try {
		$SevenZip = & cmd /c 7z a $VolumeSwitch -t7z -m0=lzma2 -mx=9 -mfb=64 -md=32m -ms=on -mhe=on $PWSwitch "$BackupLocation\$BackupName\$BackupName.7z" "$UF\*" | Out-String
		Debug $SevenZip
		Debug "Archive creation finished in $(ElapsedTime $StartArchive)"
		Debug "Wait a few seconds to make sure archive is finished"
		Email "[OK] 7z archive created"
		Start-Sleep -Seconds 3
	}
	Catch {
		Debug "[ERROR] Archive Creation : $($Error[0])"
		Email "[ERROR] Archive Creation : Check Debug Log"
		Email "[ERROR] Archive Creation : $($Error[0])"
		EmailResults
		Exit
	}
}

Function OffsiteUpload {

	$BeginOffsiteUpload = Get-Date
	Debug "----------------------------"
	Debug "Begin offsite upload process"

	<#  Authorize and get access token  #>
	Debug "Getting access token from LetsUpload"
	$URIAuth = "https://letsupload.io/api/v2/authorize"
	$AuthBody = @{
		'key1' = $APIKey1;
		'key2' = $APIKey2;
	}
	$GetAccessTokenTries = 1
	Do {
		Try{
			$Auth = Invoke-RestMethod -Method GET $URIAuth -Body $AuthBody -ContentType 'application/json; charset=utf-8' 
			$AccessToken = $Auth.data.access_token
			$AccountID = $Auth.data.account_id
			Debug "Access Token : $AccessToken"
			Debug "Account ID   : $AccountID"
		}
		Catch {
			Debug "LetsUpload Authentication ERROR : Try $GetAccessTokenTries : $($Error[0])"
		}
		$GetAccessTokenTries++
	} Until (($GetAccessTokenTries -gt $MaxUploadTries) -or ($AccessToken -match "^\w{128}$"))
	If ($GetAccessTokenTries -gt $MaxUploadTries) {
		Debug "LetsUpload Authentication ERROR : Giving up"
		Email "[ERROR] LetsUpload Authentication : Check Debug Log"
		EmailResults
		Exit
	}

	<#  Create Folder  #>
	Debug "----------------------------"
	Debug "Creating Folder $BackupName at LetsUpload"
	$URICF = "https://letsupload.io/api/v2/folder/create"
	$CFBody = @{
		'access_token' = $AccessToken;
		'account_id' = $AccountID;
		'folder_name' = $BackupName;
		'is_public' = $IsPublic;
	}
	$CreateFolderTries = 1
	Do {
		Try {
			$CreateFolder = Invoke-RestMethod -Method GET $URICF -Body $CFBody -ContentType 'application/json; charset=utf-8' 
			$CFResponse = $CreateFolder.response
			$FolderID = $CreateFolder.data.id
			$FolderURL = $CreateFolder.data.url_folder
			Debug "Response   : $CFResponse"
			Debug "Folder ID  : $FolderID"
			Debug "Folder URL : $FolderURL"
		}
		Catch {
			Debug "LetsUpload Folder Creation ERROR : Try $CreateFolderTries : $($Error[0])"
		}
		$CreateFolderTries++
	} Until (($CreateFolderTries -gt $MaxUploadTries) -or ($FolderID -match "^\d+$"))
	If ($CreateFolderTries -gt $MaxUploadTries) {
		Debug "LetsUpload Folder Creation ERROR : Giving up"
		Email "[ERROR] LetsUpload Folder Creation Error : Check Debug Log"
		EmailResults
		Exit
	}

	<#  Upload Files  #>
	$StartUpload = Get-Date
	Debug "----------------------------"
	Debug "Begin uploading files to LetsUpload"
	$CountArchVol = (Get-ChildItem "$BackupLocation\$BackupName").Count
	Debug "There are $CountArchVol files to upload"
	$UploadCounter = 1
	$TotalUploadErrors = 0

	Get-ChildItem "$BackupLocation\$BackupName" | ForEach {

		$FileName = $_.Name;
		$FilePath = $_.FullName;
		$FileSize = $_.Length;
		
		$UploadURI = "https://letsupload.io/api/v2/file/upload";
		Debug "----------------------------"
		Debug "Encoding file $FileName"
		$BeginEnc = Get-Date
		Try {
			$FileBytes = [System.IO.File]::ReadAllBytes($FilePath);
			$FileEnc = [System.Text.Encoding]::GetEncoding('ISO-8859-1').GetString($FileBytes);
		}
		Catch {
			Debug "Error in encoding file $UploadCounter."
			Debug "$($Error[0])"
		}
		Debug "Finished encoding file in $(ElapsedTime $BeginEnc)";
		$Boundary = [System.Guid]::NewGuid().ToString(); 
		$LF = "`r`n";

		$BodyLines = (
			"--$Boundary",
			"Content-Disposition: form-data; name=`"access_token`"",
			'',
			$AccessToken,
			"--$Boundary",
			"Content-Disposition: form-data; name=`"account_id`"",
			'',
			$AccountID,
			"--$Boundary",
			"Content-Disposition: form-data; name=`"folder_id`"",
			'',
			$FolderID,
			"--$Boundary",
			"Content-Disposition: form-data; name=`"upload_file`"; filename=`"$FileName`"",
			"Content-Type: application/json",
			'',
			$FileEnc,
			"--$Boundary--"
		) -join $LF
			
		Debug "Uploading $FileName - $UploadCounter of $CountArchVol"
		$UploadTries = 1
		$BeginUpload = Get-Date
		Do {
			$Error.Clear()
			$Upload = $UResponse = $UURL = $USize = $UStatus = $NULL
			Try {
				$Upload = Invoke-RestMethod -Uri $UploadURI -Method POST -ContentType "multipart/form-data; boundary=`"$boundary`"" -Body $BodyLines
				$UResponse = $Upload.response
				$UURL = $Upload.data.url
				$USize = $Upload.data.size
				$USizeFormatted = "{0:N2}" -f (($USize)/1MB)
				$UStatus = $Upload._status
				$UFileID = $upload.data.file_id
				If ($USize -ne $FileSize) {Throw "Local and remote filesizes do not match! Local file: $Filesize ::: Remote file: $USize"}
				Debug "Upload try $UploadTries"
				Debug "Response : $UResponse"
				Debug "File ID  : $UFileID"
				Debug "URL      : $UURL"
				Debug "Size     : $USizeFormatted mb"
				Debug "Status   : $UStatus"
				Debug "Finished uploading file in $(ElapsedTime $BeginUpload)"
			} 
			Catch {
				Debug "Upload try $UploadTries"
				Debug "[ERROR]  : $($Error[0])"
				If (($USize -gt 0) -and ($UFileID -match '\d+')) {
					Debug "Deleting file due to size mismatch"
					$URIDF = "https://letsupload.io/api/v2/file/delete"
					$DFBody = @{
						'access_token' = $AccessToken;
						'account_id' = $AccountID;
						'file_id' = $UFileID;
					}
					Try {
						$DeleteFile = Invoke-RestMethod -Method GET $URIDF -Body $DFBody -ContentType 'application/json; charset=utf-8' 
						Debug "Mismatched upload deleted. Trying again."
					}
					Catch {
						Debug "Mismatched upload file delete ERROR : $($Error[0])"
						Email "[ERROR] Un-Repairable Upload Mismatch! See debug log. Quitting script."
						EmailResults
						Exit
					}
				}
				$TotalUploadErrors++
			}
			$UploadTries++
		} Until (($UploadTries -gt $MaxUploadTries) -or ($UStatus -match "success"))

		If (-not($UStatus -Match "success")) {
			Debug "Error in uploading file number $UploadCounter. Check the log for errors."
			Email "[ERROR] in uploading file number $UploadCounter. Check the log for errors."
			EmailResults
			Exit
		}
		$UploadCounter++
	}
	
	<#  Count remote files  #>
	Debug "----------------------------"
	Debug "Counting uploaded files at LetsUpload"
	$URIFL = "https://letsupload.io/api/v2/folder/listing"
	$FLBody = @{
		'access_token' = $AccessToken;
		'account_id' = $AccountID;
		'parent_folder_id' = $FolderID;
	}
	Try {
		$FolderListing = Invoke-RestMethod -Method GET $URIFL -Body $FLBody -ContentType 'application/json; charset=utf-8' 
	}
	Catch {
		Debug "LetsUpload Folder Listing ERROR : $($Error[0])"
		Email "[ERROR] LetsUpload Folder Listing : Check Debug Log"
		Email "[ERROR] LetsUpload Folder Listing : $($Error[0])"
	}
	$FolderListingStatus = $FolderListing._status
	$RemoteFileCount = ($FolderListing.data.files.id).Count
	
	<#  Report results  #>
	If ($FolderListingStatus -match "success") {
		Debug "There are $RemoteFileCount file$(Plural $RemoteFileCount) in the remote folder"
		If ($RemoteFileCount -eq $CountArchVol) {
			Debug "----------------------------"
			Debug "Finished uploading $CountArchVol file$(Plural $CountArchVol) in $(ElapsedTime $BeginOffsiteUpload)"
			If ($TotalUploadErrors -gt 0){Debug "$TotalUploadErrors upload error$(Plural $TotalUploadErrors) successfully resolved"}
			Debug "Upload sucessful. $CountArchVol file$(Plural $CountArchVol) uploaded to $FolderURL"
			If ($UseHTML) {
				Email "[OK] $CountArchVol file$(Plural $CountArchVol) uploaded to <a href=`"$FolderURL`">letsupload.io</a>"
			} Else {
				Email "[OK] $CountArchVol file$(Plural $CountArchVol) uploaded to $FolderURL"
			}
		} Else {
			Debug "----------------------------"
			Debug "Finished uploading in $(ElapsedTime $StartUpload)"
			Debug "[ERROR] Number of archive files uploaded does not match count in remote folder"
			Debug "[ERROR] Archive volumes   : $CountArchVol"
			Debug "[ERROR] Remote file count : $RemoteFileCount"
			Email "[ERROR] Number of archive files uploaded does not match count in remote folder - see debug log"
		}
	} Else {
		Debug "----------------------------"
		Debug "Error : Unable to obtain file count from remote folder"
		Email "[ERROR] Unable to obtain uploaded file count from remote folder - see debug log"
	}

}

Function CheckForUpdates {
	Debug "----------------------------"
	Debug "Checking for script update at GitHub"
	$GitHubVersion = $LocalVersion = $NULL
	$GetGitHubVersion = $GetLocalVersion = $False
	$GitHubVersionTries = 1
	Do {
		Try {
			$GitHubVersion = [decimal](Invoke-WebRequest -UseBasicParsing -Method GET -URI https://raw.githubusercontent.com/palinkas-jo-reggelt/LetsUpload-Compress-and-Upload/main/version.txt).Content
			$GetGitHubVersion = $True
		}
		Catch {
			Debug "[ERROR] Obtaining GitHub version : Try $GitHubVersionTries : Obtaining version number: $($Error[0])"
		}
		$GitHubVersionTries++
	} Until (($GitHubVersion -gt 0) -or ($GitHubVersionTries -gt $MaxUploadTries))
	If (Test-Path "$PSScriptRoot\version.txt") {
		$LocalVersion = [decimal](Get-Content "$PSScriptRoot\version.txt")
		$GetLocalVersion = $True
	}
	If (($GetGitHubVersion) -and ($GetLocalVersion)) {
		If ($LocalVersion -lt $GitHubVersion) {
			Debug "[INFO] Upgrade to version $GitHubVersion available at https://github.com/palinkas-jo-reggelt/LetsUpload-Compress-and-Upload"
			If ($UseHTML) {
				Email "[INFO] Upgrade to version $GitHubVersion available at <a href=`"https://github.com/palinkas-jo-reggelt/LetsUpload-Compress-and-Upload`">GitHub</a>"
			} Else {
				Email "[INFO] Upgrade to version $GitHubVersion available at https://github.com/palinkas-jo-reggelt/LetsUpload-Compress-and-Upload"
			}
		} Else {
			Debug "Backup & Upload script is latest version: $GitHubVersion"
		}
	} Else {
		If ((-not($GetGitHubVersion)) -and (-not($GetLocalVersion))) {
			Debug "[ERROR] Version test failed : Could not obtain either GitHub nor local version information"
			Email "[ERROR] Version check failed"
		} ElseIf (-not($GetGitHubVersion)) {
			Debug "[ERROR] Version test failed : Could not obtain version information from GitHub"
			Email "[ERROR] Version check failed"
		} ElseIf (-not($GetLocalVersion)) {
			Debug "[ERROR] Version test failed : Could not obtain local install version information"
			Email "[ERROR] Version check failed"
		} Else {
			Debug "[ERROR] Version test failed : Unknown reason - file issue at GitHub"
			If ($UseHTML) {
				Email "[ERROR] Version test failed : Unknown reason - <a href=`"https://github.com/palinkas-jo-reggelt/LetsUpload-Compress-and-Upload`">file issue at GitHub</a>"
			} Else {
				Email "[ERROR] Version check failed"
			}
		}
	}
}

<###   BEGIN SCRIPT   ###>
$StartScript = Get-Date

<#  Clear out error variable  #>
$Error.Clear()

<#  Use UploadName (or not)  #>
$UploadName = $UploadName -Replace '\s','_'
$UploadName = $UploadName -Replace '[^a-zA-Z0-9-]',''
If ($UploadName) {
	$BackupName = "$((Get-Date).ToString('yyyy-MM-dd'))_$UploadName"
} Else {
	$BackupName = "$((Get-Date).ToString('yyyy-MM-dd'))_Backup"
}

<#  Delete old debug files and create new  #>
$EmailBody = "$PSScriptRoot\EmailBody.log"
If (Test-Path $EmailBody) {Remove-Item -Force -Path $EmailBody}
New-Item $EmailBody
$DebugLog = "$BackupLocation\$($BackupName)_Debug.log"
If (Test-Path $DebugLog) {Remove-Item -Force -Path $DebugLog}
New-Item $DebugLog
Write-Output "::: hMailServer Backup Routine $(Get-Date -f D) :::" | Out-File $DebugLog -Encoding ASCII -Append
Write-Output " " | Out-File $DebugLog -Encoding ASCII -Append
If ($UseHTML) {
	Write-Output "
		<!DOCTYPE html><html>
		<head><meta name=`"viewport`" content=`"width=device-width, initial-scale=1.0 `" /></head>
		<body style=`"font-family:Arial Narrow`"><table>
	" | Out-File $EmailBody -Encoding ASCII -Append
}

<#  Validate backup folder  #>
$UF = $UploadFolder -Replace('\\$','')
If (Test-Path $UF) {
	Debug "The folder to be backed up is $UF"
} Else {
	Debug "Error : The folder to be backed up could not be found. Quitting Script"
	Debug "$UploadFolder does not exist"
	Email "Error : The folder to be backed up could not be found. Quitting Script"
	Email "$UploadFolder does not exist"
	EmailResults
	Exit
}

<#  Set Email Header  #>
If ($UseHTML) {
	Email "<center>:::&nbsp;&nbsp;&nbsp;$UploadName Backup Routine&nbsp;&nbsp;&nbsp;:::</center>"
	Email "<center>$(Get-Date -f D)</center>"
	Email " "
} Else {
	Email ":::   $UploadName Backup Routine   :::"
	Email "       $(Get-Date -f D)"
	Email " "
}

<#  Compress backup into 7z archives  #>
MakeArchive

<#  Upload archive to LetsUpload.io  #>
OffsiteUpload

<#  Check for updates  #>
CheckForUpdates

<#  Finish up and send email  #>
Debug "----------------------------"
Debug "$UploadName Backup & Upload routine completed in $(ElapsedTime $StartScript)"
Email " "
Email "$UploadName Backup & Upload routine completed in $(ElapsedTime $StartScript)"
If ($UseHTML) {Write-Output "</table></body></html>" | Out-File $EmailBody -Encoding ASCII -Append}
EmailResults