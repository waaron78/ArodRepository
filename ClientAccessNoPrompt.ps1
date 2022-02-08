


#Log file created on Local drive in Temp Folder
Start-Transcript -Path "C:\Temp\AS400log.txt"  -IncludeInvocationHeader -Force

$ErrorActionPreference = "Stop"

#Check PC for Client Access Program Private Folder
$AppPath = Test-Path 'C:\Program Files (x86)\IBM\Client Access\Emulator\Private' -PathType Container
If ($AppPath -eq $True) {Echo "Path exists on PC."}
Else {Echo "Path NOT exist on PC."
exit
}
#Give Users Group Full Control of IBM Folder and Public Desktop
icacls "C:\Program Files (x86)\IBM" /grant '"Users":(OI)(CI)(F)'

#Delete Contens from Client Access Private folder on local PC
gci "C:\Program Files (x86)\IBM\Client Access\Emulator\Private" | ri -force 

#Copy contents from private folder on server to local private folder
cp "\\tcp004b\Desktop Central\Deploy\Client Access\(blank sessions)\for 6.1 (win7)\private\*" -Destination "C:\Program Files (x86)\IBM\Client Access\Emulator\Private" -Recurse

#Get Computer Name First 5 Characters
$name = Get-WMIObject Win32_ComputerSystem | Select-Object -ExpandProperty name
$name = $name.substring(0,5)
$name = "TCP" + $name
$name

#Edit .ws file with Computer name for Session1
$S1 = $name + "S1"
((gc -path "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #1.ws" -Raw) -replace 'TCPS1',$S1 ) | sc -Path "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #1.ws"

#Edit .ws file with Computer name for Session2
$S2 = $name + "S2"
((gc -path "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #2.ws" -Raw) -replace 'TCPS2',$S2 ) | sc -Path "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #2.ws"


#Edit .ws file with Computer name for Session3
$S3 = $name + "S3"
((gc -path "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #3.ws" -Raw) -replace 'TCPS3',$S3 ) | sc -Path "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #3.ws"


#Edit .ws file with Computer name for Session4
$S4 = $name + "S4"
((gc -path "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #4.ws" -Raw) -replace 'TCPS4',$S4 ) | sc -Path "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #4.ws"


#Edit .ws file with Computer name for Staging
$ST = $name + "ST"
((gc -path "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\STAGING.ws" -Raw) -replace 'TCPST',$ST ) | sc -Path "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\STAGING.ws"

#Give Domain Admins Full Control Public Desktop
icacls "C:\Users\Public\Desktop" /grant '"Users":(OI)(CI)(F)'

#Create Shortcuts for Session1 on Public Desktop
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\AS400S1.lnk")
$Shortcut.TargetPath = "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #1.ws"
$Shortcut.Save()

#Create Shortcuts for Session2 on Public Desktop
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\AS400S2.lnk")
$Shortcut.TargetPath = "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #2.ws"
$Shortcut.Save()

#Create Shortcuts for Session3 on Public Desktop
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\AS400S3.lnk")
$Shortcut.TargetPath = "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #3.ws"
$Shortcut.Save()

#Create Shortcuts for Session4 on Public Desktop
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\AS400S4.lnk")
$Shortcut.TargetPath = "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\AS400 Session #4.ws"
$Shortcut.Save()

#Create Shortcuts for Staging on Public Desktop
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\STAGING.lnk")
$Shortcut.TargetPath = "C:\Program Files (x86)\IBM\Client Access\Emulator\Private\STAGING.WS"
$Shortcut.Save()


#Stop Transcript and Delete Log File
Stop-Transcript



#---------------End of Script----------------