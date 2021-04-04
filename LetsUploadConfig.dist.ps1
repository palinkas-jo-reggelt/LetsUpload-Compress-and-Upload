<#

.SYNOPSIS
	LetsUpload Backup Utility Config File

.DESCRIPTION
	LetsUpload Backup Utility Config File

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

<###   USER VARIABLES   ###>
$APIKey1           = "64_character_API_string"
$APIKey2           = "64_character_API_string"
$ArchivePassword   = "supersecretpassword"  # Password to 7z archive
$BackupLocation    = "C:\BACKUP"            # Location archive files will be stored
$VolumeSize        = "100m"                 # Size of archive volume parts - maximum 200m recommended - valid suffixes for size units are (b|k|m|g)
$IsPublic          = 0                      # 0 = Private, 1 = Unlisted, 2 = Public in site search
$VerboseConsole    = $True                  # If true, will output debug to console
$VerboseFile       = $True                  # If true, will output debug to file
$DeleteOldBackups  = $False                 # If true, will delete local backups older than N days
$DaysToKeep        = 5                      # Number of days to keep backups - all others will be deleted at end of script
$MaxUploadTries    = 5                      # If file upload error, number of times to retry before giving up

<###   EMAIL VARIABLES   ###>
$EmailFrom         = "notify@mydomain.net"
$EmailTo           = "admin@mydomain.net"
$Subject           = "Offsite Backup: $UploadName"
$SMTPServer        = "mydomain.net"
$SMTPAuthUser      = "notify@mydomain.net"
$SMTPAuthPass      = "supersecretpassword"
$SMTPPort          =  587
$SSL               = $True
$UseHTML           = $True
$AttachDebugLog    = $True                  # If true, will attach debug log to email report - must also select $VerboseFile
$MaxAttachmentSize = 1                      # Size in MB