<#

.SYNOPSIS
	LetsUpload Download Utility

.DESCRIPTION
	Download backup folder contents from LetsUpload.io

.FUNCTIONALITY
	Download backup folder contents from LetsUpload.io

.PARAMETER DownloadFolder
	Specifies the folder on local filesystem where the files will be saved
	
.PARAMETER BackupName
	Specifies the name (description) of the archive previously uploaded to letsupload
	Script uses this parameter to find the most recent backup
	
.NOTES
	Create account and get API keys from https://www.letsupload.io then fill in $APIKey variables under USER VARIABLES
	Run to obtain previously uploaded backups
	Windows only
	API: https://letsupload.io/api.html
	
.EXAMPLE
	PS C:\Users\username> C:\scripts\letsupload.ps1 "C:\Path\To\Folder\To\Backup" "Backup Description (email, work, etc)"

#>

Param(
	[Parameter(Mandatory=$True)]
	[ValidatePattern("^[A-Z]\:\\")]
	[String]$DownloadFolder,

	[Parameter(Mandatory=$True)]
	[String]$BackupName
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
		Debug "Email ERROR : `n$Error[0]"
	}
}

Function ElapsedTime ($TimeSpan) {
	If (([int]($TimeSpan).TotalHours) -eq 0) {$Hours = ""} ElseIf (([int]($TimeSpan).TotalHours) -eq 1) {$Hours = "1 hour "} Else {$Hours = "$([int]($TimeSpan).TotalHours) hours "}
	If (([int]($TimeSpan).Minutes) -eq 0) {$Minutes = ""} ElseIf (([int]($TimeSpan).Minutes) -eq 1) {$Minutes = "1 minute "} Else {$Minutes = "$([int]($TimeSpan).Minutes) minutes "}
	If (([int]($TimeSpan).Seconds) -eq 0) {$Seconds = ""} ElseIf (([int]($TimeSpan).Seconds) -eq 1) {$Seconds = "1 second"} Else {$Seconds = "$([int]($TimeSpan).Seconds) seconds"}
	Return "$Hours$Minutes$Seconds"
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
$DF = $DownloadFolder -Replace('\\$','')
If (Test-Path $DF) {
	Debug "The folder to store downloaded files is $DF"
} Else {
	Debug "Error : The folder to be backed up could not be found. Quitting Script"
	Debug "$DownloadFolder does not exist"
	Email "Error : The folder to be backed up could not be found. Quitting Script"
	Email "$DownloadFolder does not exist"
	EmailResults
	Exit
}

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
	Debug "LetsUpload Authentication ERROR : $Error"
	Email "LetsUpload Authentication ERROR : Check Debug Log"
	Email "LetsUpload Authentication ERROR : $Error"
	EmailResults
	Exit
}
$AccessToken = $Auth.data.access_token
$AccountID = $Auth.data.account_id
Debug "Access Token : $AccessToken"
Debug "Account ID   : $AccountID"

<#  Get folder_id of last upload  #>
$URIFolderListing = "https://letsupload.io/api/v2/folder/listing"
$FLBody = @{
	'access_token' = $AccessToken;
	'account_id' = $AccountID;
}
Try{
	$FolderListing = Invoke-RestMethod -Method GET $URIFolderListing -Body $FLBody -ContentType 'application/json; charset=utf-8'
}
Catch {
	Debug "ERROR obtaining backup folder ID : $Error"
	Email "ERROR obtaining backup folder ID : Check Debug Log"
	Email "ERROR obtaining backup folder ID : $Error"
	EmailResults
	Exit
}
$NewestBackup = $FolderListing.data.folders | Sort-Object date_added -Descending | Where {$_.folderName -match $BackupName} | Select -First 1
$FolderID = $NewestBackup.id
$FolderName = $NewestBackup.folderName
Debug "Folder Name: $FolderName"
Debug "Folder ID  : $FolderID"


<#  Get file listing within latest backup folder  #>
$URIFolderListing = "https://letsupload.io/api/v2/folder/listing"
$FLBody = @{
	'access_token' = $AccessToken;
	'account_id' = $AccountID;
	'parent_folder_id' = $FolderID;
}
Try{
	$FileListing = Invoke-RestMethod -Method GET $URIFolderListing -Body $FLBody -ContentType 'application/json; charset=utf-8'
}
Catch {
	Debug "ERROR obtaining backup file listing : $Error"
	Email "ERROR obtaining backup file listing : Check Debug Log"
	Email "ERROR obtaining backup file listing : $Error"
	EmailResults
	Exit
}

$Count = ($FileListing.data.files).Count
Debug "File count: $Count"
$N = 1
$DLSuccess = 0
Debug "Starting file download"

<#  Loop through results and download files  #>
$FileListing.data.files | ForEach {
	$FileID = $_.id
	$FileName = $_.filename
	$FileURL = $_.url_file
	Debug "----------------------------"
	Debug "File $N of $Count"
	Debug "File ID     : $FileID"
	Debug "File Name   : $FileName"

	$URIDownload = "https://letsupload.io/api/v2/file/download"
	$DLBody = @{
		'access_token' = $AccessToken;
		'account_id' = $AccountID;
		'file_id' = $FileID;
	}
	Try{
		$FileDownload = Invoke-RestMethod -Method GET $URIDownload -Body $DLBody -ContentType 'application/json; charset=utf-8'
	}
	Catch {
		Debug "ERROR obtaining download URL : $Error"
		Email "ERROR obtaining download URL : Check Debug Log"
		Email "ERROR obtaining download URL : $Error"
		EmailResults
		Exit
	}

	$FDStatus = $FileDownload._status
	If ($FDStatus -notmatch "success") {
		Debug "Error : Could not obtain download URL"
	} Else {
		$DownloadURL = $FileDownload.data.download_url
		Debug "Download URL: $DownloadURL"
		<#  Download file using BITS  #>
		Try {
			$BeginDL = Get-Date
			Import-Module BitsTransfer
			Start-BitsTransfer -Source $DownloadURL -Destination "$DF\$FileName"
			Debug "File $N downloaded in $(ElapsedTime (New-TimeSpan $BeginDL))"
			If (Test-Path "$DF\$FileName") {$DLSuccess++}
		}
		Catch {
			Debug "BITS ERROR downloading file $N of $FileCount : $Error"
		}
	}
	
	$N++
}

<#  Finish up and email results  #>
Debug "----------------------------"
If ($DLSuccess -eq $Count) {
	Debug "Download sucessful. $Count files downloaded to $DF"
	Email "Download sucessful. $Count files downloaded to $DF"
} Else {
	Debug "Download FAILED to download $($Count - $DLSuccess) files - check debug log"
	Email "Download FAILED to download $($Count - $DLSuccess) files - check debug log"
}

Email " "
Debug "Script completed in $(ElapsedTime (New-TimeSpan $StartScript))"
Email "Script completed in $(ElapsedTime (New-TimeSpan $StartScript))"
Debug "Sending Email"
If (($AttachDebugLog) -and (Test-Path $DebugLog) -and (((Get-Item $DebugLog).length/1MB) -gt $MaxAttachmentSize)){
	Email "Debug log size exceeds maximum attachment size. Please see log file in script folder"
}
EmailResults