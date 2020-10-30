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
	Write-Output "$(Get-Date) -f G) : ERROR : Unable to load supporting PowerShell Scripts : $query $Error" | out-file "$PSScriptRoot\PSError.log" -append
}

<###   FUNCTIONS   ###>
Function Debug ($DebugOutput) {
	If ($VerboseFile) {Write-Output "$(Get-Date -f G) : $DebugOutput" | Out-File $DebugLog -Encoding ASCII -Append}
	If ($VerboseConsole) {Write-Host "$(Get-Date -f G) : $DebugOutput"}
}

Function Email ($EmailOutput) {
	Write-Output $EmailOutput | Out-File $VerboseEmail -Encoding ASCII -Append
}

Function EmailResults {
	Try {
		$Body = (Get-Content -Path $VerboseEmail | Out-String )
		If (($AttachDebugLog) -and (Test-Path $DebugLog) -and (((Get-Item $DebugLog).length/1MB) -lt $MaxAttachmentSize)){$Attachment = New-Object System.Net.Mail.Attachment $DebugLog}
		$Message = New-Object System.Net.Mail.Mailmessage $EmailFrom, $EmailTo, $Subject, $Body
		$Message.IsBodyHTML = $HTML
		If (($AttachDebugLog) -and (Test-Path $DebugLog) -and (((Get-Item $DebugLog).length/1MB) -lt $MaxAttachmentSize)){$Message.Attachments.Add($DebugLog)}
		$SMTP = New-Object System.Net.Mail.SMTPClient $SMTPServer,$SMTPPort
		$SMTP.EnableSsl = $SSL
		$SMTP.Credentials = New-Object System.Net.NetworkCredential($SMTPAuthUser, $SMTPAuthPass); 
		$SMTP.Send($Message)
	}
	Catch {
		Debug "Email ERROR : $Error"
	}
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

<###   BEGIN SCRIPT   ###>
$StartScript = Get-Date

<#  Clear out error variable  #>
$Error.Clear()

<#  Delete old debug file and create new  #>
$VerboseEmail = "$PSScriptRoot\VerboseEmail.log"
If (Test-Path $VerboseEmail) {Remove-Item -Force -Path $VerboseEmail}
New-Item $VerboseEmail
$DebugLog = "$PSScriptRoot\LetsUploadDebug.log"
If (Test-Path $DebugLog) {Remove-Item -Force -Path $DebugLog}
New-Item $DebugLog

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

<#  Use UploadName (or not)  #>
$UploadName = $UploadName -Replace '\s','-'
$UploadName = $UploadName -Replace '[^a-zA-Z0-9-]',''
If ($UploadName) {
	$BackupName = "$((Get-Date).ToString('yyyy-MM-dd'))-$UploadName"
} Else {
	$BackupName = "$((Get-Date).ToString('yyyy-MM-dd'))-Backup"
}

<#  Create archive  #>
$StartArchive = Get-Date
Debug "----------------------------"
Debug "Create archive : $BackupName"
Debug "Archive folder : $UF"
$VolumeSwitch = "-v$VolumeSize"
$PWSwitch = "-p$ArchivePassword"
Try {
	$SevenZip = & cmd /c 7z a $VolumeSwitch -t7z -m0=lzma2 -mx=9 -mfb=64 -md=32m -ms=on -mhe=on $PWSwitch "$BackupLocation\$BackupName\$BackupName.7z" "$UF\*" | Out-String
	Debug $SevenZip
}
Catch {
	Debug "Archive Creation ERROR : $Error"
	Email "Archive Creation ERROR : Check Debug Log"
	Email "Archive Creation ERROR : $Error"
	EmailResults
	Exit
}
Debug "Archive creation finished in $(ElapsedTime $StartArchive)"
Debug "Wait a few seconds to make sure archive is finished"
Start-Sleep -Seconds 3

<#  Authorize and get access token  #>
Debug "----------------------------"
Debug "Getting access token from LetsUpload"
$URIAuth = "https://letsupload.io/api/v2/authorize"
$AuthBody = @{
	'key1' = $APIKey1;
	'key2' = $APIKey2;
}
Try{
	$Auth = Invoke-RestMethod -Method GET $URIAuth -Body $AuthBody -ContentType 'application/json; charset=utf-8' 
	$AccessToken = $Auth.data.access_token
	$AccountID = $Auth.data.account_id
	Debug "Access Token : $AccessToken"
	Debug "Account ID   : $AccountID"
}
Catch {
	Debug "LetsUpload Authentication ERROR : $Error"
	Email "LetsUpload Authentication ERROR : Check Debug Log"
	Email "LetsUpload Authentication ERROR : $Error"
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
	Debug "LetsUpload Folder Creation ERROR : $Error"
	Email "[ERROR] LetsUpload Folder Creation : Check Debug Log"
	Email "[ERROR] LetsUpload Folder Creation : $Error"
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
			Debug "$Error"
			Debug " "
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
			If ($USize -ne $FileSize) {Throw "Local and remote filesizes do not match!"}
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
			Debug "[ERROR]  : $Error"
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
				}
				Catch {
					Debug "File delete ERROR : $Error"
				}
			}
		}
		$UploadTries++
	} Until (($UploadTries -eq ($MaxUploadTries + 1)) -or ($UStatus -match "success"))

	If (-not($UStatus -Match "success")) {
		Debug "Error in uploading file number $UploadCounter. Check the log for errors."
		Email "[ERROR] in uploading file number $UploadCounter. Check the log for errors."
		EmailResults
		Exit
	}
	$UploadCounter++
}
Debug "----------------------------"
Debug "Finished offsite upload process in $(ElapsedTime $BeginOffsiteUpload)"

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
	Debug "LetsUpload Folder Listing ERROR : $Error"
	Email "[ERROR] LetsUpload Folder Listing : Check Debug Log"
	Email "[ERROR] LetsUpload Folder Listing : $Error"
}
$FolderListingStatus = $FolderListing._status
$RemoteFileCount = ($FolderListing.data.files.id).Count

<#  Report results  #>
If ($FolderListingStatus -match "success") {
	Debug "There are $RemoteFileCount files in the remote folder"
	If ($RemoteFileCount -eq $CountArchVol) {
		Debug "----------------------------"
		Debug "Finished uploading $CountArchVol files in $(ElapsedTime $StartUpload)"
		Debug "Upload sucessful. $CountArchVol files uploaded to $FolderURL"
		Email "* Offsite upload of backup archive completed successfully:"
		Email "* $CountArchVol files uploaded to $FolderURL"
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

<#  Delete old backups  #>
If ($DeleteOldBackups) {
	$FilesToDel = Get-ChildItem -Path $BackupLocation  | Where-Object {$_.LastWriteTime -lt ((Get-Date).AddDays(-$DaysToKeep))}
	$CountDel = $FilesToDel.Count
	If ($CountDel -gt 0) {
		Debug "Deleting $CountDel items older than $DaysToKeep days"
		Debug "----------------------------"
	}
	$FilesToDel | ForEach {
		$FullName = $_.FullName
		$Name = $_.Name
		If (Test-Path $_.FullName -PathType Container) {
			Remove-Item -Force -Recurse -Path $FullName
			Debug "Deleting folder: $Name"
		}
		If (Test-Path $_.FullName -PathType Leaf) {
			Remove-Item -Force -Path $FullName
			Debug "Deleting file  : $Name"
		}
	}
}

<#  Finish up and email results  #>
Debug "Sending Email"
If (($AttachDebugLog) -and (Test-Path $DebugLog) -and (((Get-Item $DebugLog).length/1MB) -gt $MaxAttachmentSize)){
	Email "Debug log size exceeds maximum attachment size. Please see log file in script folder"
}
EmailResults