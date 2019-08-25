<# 
 
.SYNOPSIS 
Rebuild The Quick Access Menu based on the listings in a CSV File
 
.DESCRIPTION 
Rebuild The Quick Access Menu based on the listings in a CSV File

Script Process
-Clears all the Quick Access and Recent Files
-Reads entries from CSV
-Creates a Quick Access Pins directory in your User Profile to store SymLinks
-Creates Symlinks based on the entries in the CSV File
-Audits Symlinks / Deletes Orphans in the Quick Access Pins
-Puts Pins on the Quick Access Menu with your custom names (from the SymDirName in the CSV file) 
 
.EXAMPLE 
 
.NOTES 
 
.LINK 
 
#>

#Get 
Function Find-CRSPath {
    IF ($host.name -eq 'Windows Powershell ISE Host') { #If the script is being run in ISE
        $Global:CRSPath = $psISE.CurrentFile.FullPath | Split-Path}  #Get the Directory where this script is saved, and then chop off the file name
        

    IF ($host.name -eq 'ConsoleHost') { #If the Script is being run normally
        $Global:CRSPath = Split-Path $PSCommandPath}
            
    #Error Checking
    IF ($CRSPath -eq "" -or !(Test-Path variable:global:CRSPath)) {Messages -Msgtype Error -MsgLevel Fatal -Location Function-CRSPath  -Message "Path of current running script not found" -Impact "Script will NOT be able to find Modules, cannot continue."}
}  

Find-CRSPath

Import-Module $CRSPath\Set-QuickAccess.psm1

Remove-Item "$env:APPDATA\Microsoft\Windows\Recent\*" -Include *.* -Force -Recurse
#Remove-Item "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations"
#Remove-Item "$env:APPDATA\Microsoft\Windows\Recent\CustomDestinations"

$Pins = Import-Csv -Path "$CRSPath\QuickAccess.csv"
If (!(Test-Path "$env:USERPROFILE\Quick Access Pins")) {New-Item -ItemType Directory -Path $env:USERPROFILE -Name Quick Access Pins }
$PinsDir = "$env:USERPROFILE\Quick Access Pins"

$CurrentPinSymDirs = Get-ChildItem -Path $PinsDir | Where-Object { ($_.Attributes -match "ReparsePoint")}

#Clear Out Existing SymLinks in the $PinsDir that are NOT in the CSV List (Orphans)
foreach ($Dir in $CurrentPinSymDirs) {
$DirName = $Dir.Name
If (!($Pins.SymDirName -contains $DirName)) {
    Write-Host "Found Orphan SymLink: $PinsDir\$Dir.Name"
    Get-Item $PinsDir\$DirName | %{$_.Delete()}
    }
}

#Create New SymLinks in the $PinsDir that are in the CSV List and Create PIN (In order of the CSV file)
foreach ($Pin in $Pins) {
$PinSymDirName = $Pin.SymDirName
$PinReferencedPath = $Pin.ReferencedPath
If (!(Get-ChildItem -Path $PinsDir\$PinSymDirName)) {New-Item -ItemType Junction -Path "$PinsDir\$PinSymDirName" -Target $PinReferencedPath}
Set-QuickAccess -Action Pin -Path $PinsDir\$PinSymDirName
}