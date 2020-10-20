<#

.SYNOPSIS
	LetsUpload Backup Utility

.DESCRIPTION
	Compresses and uploads folder contents to LetsUpload.io

.FUNCTIONALITY
	Compresses and uploads folder contents to LetsUpload.io

.PARAMETER UploadFolder
	Specifies the folder on local filesystem to compress and upload
	DO NOT include trailing slash "\"
	
.PARAMETER UploadName
	Specifies the name (description) of the archive to be created as well as letsupload folder name
	
.NOTES
	Create account and get API keys from https://www.letsupload.io then fill in $APIKey variables under USER VARIABLES.
	Run from task scheduler daily.
	Windows only.
	API: https://letsupload.io/api.html
	Install latest 7-zip and put into system path.
	
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
	Write-Output "$(Get-Date) -f G) : ERROR : Unable to load supporting PowerShell Scripts : $query `n$Error[0]" | out-file "$PSScriptRoot\PSError.log" -append
}

<###   FUNCTIONS   ###>
Function Debug ($DebugOutput) {
	If ($VerboseFile) {Write-Output "$(Get-Date -f G) $DebugOutput" | Out-File $DebugLog -Encoding ASCII -Append}
	If ($VerboseConsole) {Write-Host "$(Get-Date -f G) $DebugOutput"}
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
		Debug "Email ERROR : `n$Error[0]"
	}
}

<###   BEGIN SCRIPT   ###>
$StartScript = Get-Date

<#  Clear out error variable  #>
$Error.Clear()

<#  Use UploadName (or not)  #>
$UploadName = $UploadName -Replace '\s','-'
$UploadName = $UploadName -Replace '[^a-zA-Z0-9-]',''
If ($UploadName) {
	$BackupName = "$UploadName-$((Get-Date).ToString('yyyy-MM-dd'))"
} Else {
	$BackupName = "Backup-$((Get-Date).ToString('yyyy-MM-dd'))"
}

<#  Delete old debug file and create new  #>
$VerboseEmail = "$PSScriptRoot\VerboseEmail.log"
If (Test-Path $VerboseEmail) {Remove-Item -Force -Path $VerboseEmail}
New-Item $VerboseEmail
$DebugLog = "$PSScriptRoot\LetsUploadDebug.log"
If (Test-Path $DebugLog) {Remove-Item -Force -Path $DebugLog}
New-Item $DebugLog

<#  Create archive  #>
$StartArchive = Get-Date
Debug "Create archive : $BackupName"
Debug "Archive folder : $UploadFolder"
$VolumeSwitch = "-v$VolumeSize"
$PWSwitch = "-p$ArchivePassword"
Try {
	& cmd /c 7z a $VolumeSwitch -t7z -m0=lzma2 -mx=9 -mfb=64 -md=32m -ms=on -mhe=on $PWSwitch "$BackupLocation\$BackupName\$BackupName.7z" "$UploadFolder\*"
}
Catch {
	Debug "Archive Creation ERROR : `n$Error[0]"
	Email "Archive Creation ERROR : Check Debug Log"
	Email "Archive Creation ERROR : `n$Error[0]"
	EmailResults
	Exit
}
Debug "Archive creation finished in $([int]((New-Timespan $StartArchive).TotalMinutes)) minutes."

<#  Authorize and get access token  #>
Debug "Getting access token from LetsUpload"
$URIAuth = "https://letsupload.io/api/v2/authorize"
$AuthBody = @{
	'key1' = $APIKey1;
	'key2' = $APIKey2;
}
Try{
	$Auth = Invoke-RestMethod -Method GET $URIAuth -Body $AuthBody -ContentType 'application/json; charset=utf-8' 
}
Catch {
	Debug "LetsUpload Authentication ERROR : `n$Error[0]"
	Email "LetsUpload Authentication ERROR : Check Debug Log"
	Email "LetsUpload Authentication ERROR : `n$Error[0]"
	EmailResults
	Exit
}
$AccessToken = $Auth.data.access_token
$AccountID = $Auth.data.account_id
Debug "Access Token : $AccessToken"
Debug "Account ID   : $AccountID"

<#  Create Folder  #>
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
}
Catch {
	Debug "LetsUpload Folder Creation ERROR : `n$Error[0]"
	Email "LetsUpload Folder Creation ERROR : Check Debug Log"
	Email "LetsUpload Folder Creation ERROR : `n$Error[0]"
	EmailResults
	Exit
}
$CreateFolder.response
$FolderID = $CreateFolder.data.id
$FolderURL = $CreateFolder.data.url_folder
Debug "Folder ID  : $FolderID"
Debug "Folder URL : $FolderURL"

<#  Upload Files  #>
$StartUpload = Get-Date
Debug "Begin uploading files to LetsUpload"
$Count = (Get-ChildItem "$BackupLocation\$BackupName").Count
Debug "There are $Count files to upload"
Email "There are $Count files to upload"
$N = 1

Try {
	Get-ChildItem "$BackupLocation\$BackupName" | ForEach {

		$FileName = $_.Name;
		$FilePath = "$BackupLocation\$BackupName\$FileName";
		
		$UploadURI = "https://letsupload.io/api/v2/file/upload";
		Debug "----------------------------"
		Debug "Encoding file $FileName"
		$FileBytes = [System.IO.File]::ReadAllBytes($FilePath);
		$FileEnc = [System.Text.Encoding]::GetEncoding('UTF-8').GetString($FileBytes);
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
			
		Debug "Uploading $FileName - $N of $Count"
		$Upload = Invoke-RestMethod -Uri $UploadURI -Method POST -ContentType "multipart/form-data; boundary=`"$boundary`"" -Body $BodyLines

		$UResponse = $Upload.response
		$UURL = $Upload.data.url
		$USize = $Upload.data.size
		$UStatus = $Upload._status

		Debug "Response : $UResponse"
		Debug "URL      : $UURL"
		Debug "Size     : $USize"
		Debug "Status   : $UStatus"

		If ($UResponse -NotMatch "File uploaded") {
			Debug "Error in uploading file number $N. Check the log for errors."
			Email "Error in uploading file number $N. Check the log for errors."
			EmailResults
			Exit
		}

		$N++
	}
}
Catch {
		Debug "Upload ERROR : `n$Error[0]"
		Email "Upload ERROR : Check Debug Log"
		Email "Upload ERROR : `n$Error[0]"
		EmailResults
		Exit
}

Debug "Upload finished in $([int]((New-Timespan $StartUpload).TotalMinutes)) minutes."

<#  Email results  #>
Debug "Upload sucessful. $Count files uploaded to $FolderURL"
Email "Upload sucessful. $Count files uploaded to $FolderURL"
Debug "Script completed in $([int]((New-Timespan $StartScript).TotalMinutes)) minutes."
Email "Script completed in $([int]((New-Timespan $StartScript).TotalMinutes)) minutes."
Debug "Sending Email"
EmailResults