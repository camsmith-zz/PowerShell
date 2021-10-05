<#
  .SYNOPSIS
    Sets up N: on QFES PC's
  .DESCRIPTION
    Removes old common file short cut that points to t: and re-creates one that points to N:
    Deletes the Backup.bat file from the common startup 
    Creates a new mapdrives.bat to map n: as a backup if Group Policy dosen't work
    Deletes the sharedfiles.lnk shortcut from the all users desktop
    Writes a log of PC's that cannot be contacted.  (errorlog.txt)
  .PARAMATER FPN
    This is the FPN the script will create a drive map to
  .PARAMATER FPNAREA
    This is the FPNAREA that N: will connect to on the DFS
  .PARAMATER PCList
    This is a list of PC's that the script will run against
  .PARAMATER Log
    This is log file where the files have been successfully updated
  .PARAMATER ErrorLog
    This is errorlog where the pc was uncontactable or the files failed to update
  .EXAMPLE
    Set-LocalShares
  .EXAMPLE
    Set-LocalShares -FPNAREA 'EMDFPN06' -PClist 'c:\temp\pclist.txt'
  .EXAMPLE
    Set-LocalShares -FPNAREA 'EMDFPN06' -PClist 'c:\temp\pclist.txt' -ErrorLog 'c:\temp\errorlog.txt' -Log 'c:\temp\log.txt
  .NOTES
    General notes
    Created By: Cam Smith
    Created On: July 2021  
    Last Modified : 05 October 2021
  #>
  param (
    $FPN = 'EMDFPN02',
    $FPNAREA = 'EMDAREA',
    $PClist = '\\ROKFPN06\n$\Scripts\QFES\pclist.txt',
    $Log = '\\ROKFPN06\n$\Scripts\QFES\log.txt',
    $ErrorLog = '\\ROKFPN06\n$\Scripts\QFES\Errorlog.txt',
    $deleteFiles = '\\ROKFPN06\n$\Scripts\QFES\removefiles.txt'
)
clear-host
$RemoveFiles = (Get-Content "$deletefiles")
$InstallSoftware = (Get-Content "$SoftwareToInstall")
while ($FPNAREA -eq '' ) {
    $FPNAREA = Read-Host -Prompt "Enter FPN Area 'eg ROKAREA'"
}
if (Test-Path -Path "$PClist") {
    $PCs = (Get-Content "$PClist")
}
else {   
    Write-Host -ForegroundColor Red "`n$pclist does not exist`n"
    $PClist = (Read-host -Prompt "Enter path of PCLIST")
    $PCs = (Get-Content "$PClist")
}
Write-Host -ForegroundColor Red "This List is going to remove the existing drive mappings,
the common shortcut on \Users\desktop, and will create
a new mapdrive.bat file and store it in the common startup 
folder.  It will also create a new common short cut on 
all users desktop. It will also fix the QFES add 
printer group policy issue with users not being 
able to add printers`n"
#CMD /c PAUSE
 foreach ($PC in $PCs) {
        Write-Host -ForegroundColor Cyan "Checking $PC for files to delete"
        foreach ($RemoveFile in $RemoveFiles){
            $IsItTrue = Test-Path "\\$PC\c$\$RemoveFile"
            if ($IsItTrue -eq $True) {
                Remove-Item -Path "\\$PC\c$\$RemoveFile" -Force -Recurse
                Write-Host -ForegroundColor Gray "\\$PC\c$\$RemoveFile"
                $Connected = $true
            }
        }
        if ($connected -eq $True){
        Write-Host -ForegroundColor Cyan "Applying new files to $PC"
        #$SourceFilePath = "N:\"
        $SourceFilePath = "\\$FPN\$FPNAREA\"
        $ShortcutPath = "\\$PC\c$\Users\Public\Desktop\Common.lnk"
        $WScriptObj = New-Object -ComObject ("WScript.Shell")
        $shortcut = $WscriptObj.CreateShortcut($ShortcutPath)
        $shortcut.TargetPath = $SourceFilePath
        $shortcut.Save()
        [void](New-item -Path "\\$PC\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp" -Name "mapdrive.bat"`
        -ItemType "file"`
        -Value "@echo off `nif exist t:\ (net use t: /delete /y >nul 2>nul)`nif exist n:\ (net use n: /delete /y >nul 2>nul)`
net use n: \\DESQLD.INTERNAL\FPNSHARE\$FPNAREA >nul 2>nul")
        if (Test-Path "\\$PC\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\mapdrive.bat") {
            Write-Host -ForegroundColor Gray "\\$PC\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\mapdrive.bat"
        }
        else {
            Write-Host -ForegroundColor RED "Failed to add \\$PC\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\mapdrive.bat"
        }
            
        if (Test-Path "\\$PC\c$\Users\Public\Desktop\Common.lnk") {
            Write-Host -ForegroundColor Gray "\\$PC\c$\Users\Public\Desktop\Common.lnk"
        }
        else {
            Write-Host -ForegroundColor RED "Failed to add \\$PC\c$\Users\Public\Desktop\Common.lnk"
        }   
        
        Write "$PC" | Out-File $Log -Append
        Write-Host -foreground DarkCyan "`n$log written to successfully`n" 
        $connected = $false
        }
        else {
        Write-Host -ForegroundColor Red "`n$PC - Fail" 
        Write "$PC" | Out-File $ErrorLog -Append
        Write-Host -foreground DarkCyan "`n$ErrorLog written to successfully`n" 
        } 
    }

Read-Host -Prompt "Press any key to continue or CTRL+C to quit" 
