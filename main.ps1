<#
.Synopsis
   GUI that will launch PowerShell scripts upon button clicks.
.DESCRIPTION
   This GUI was designed to make launching scripts easier and faster.
   Read the description above each button to know how the script will perform.
   Upon button click, a seperate PowerShell window will appear and begin the script.
   The window must be manually closed by the user when finished.
   This gives the user the ability to fully read their input and output before ending.
   After the PowerShell window is closed, the user can click another button or close the GUI.
.OUTPUTS
   Various output depending which script is selected and run.
.NOTES
   Author : Devon Bardoul
   version: 1.0
   Date   : 24-03-202
#>

#Necessary assembly added to the script for extended functionality
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#Form creation with custom attributes
$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '360,610'
$Form.text                       = "DB PowerShell Scripts"
$Form.BackColor                  = "#101010"
$Form.TopMost                    = $true
$Form.FormBorderStyle            = 'Fixed3D'
$Form.MaximizeBox                = $false
$Form.StartPosition              = 'CenterScreen'
$Form.TopMost                    = $false

#Panels (numbered from top to bottom of GUI) creation with custom attributes
$Panel1                          = New-Object system.Windows.Forms.Panel
$Panel1.height                   = 65
$Panel1.width                    = 340
$Panel1.BackColor                = "#1d1d1d"
$Panel1.location                 = New-Object System.Drawing.Point(10,10)

$Panel2                          = New-Object system.Windows.Forms.Panel
$Panel2.height                   = 65
$Panel2.width                    = 340
$Panel2.BackColor                = "#1d1d1d"
$Panel2.location                 = New-Object System.Drawing.Point(10,85)

$Panel3                          = New-Object system.Windows.Forms.Panel
$Panel3.height                   = 65
$Panel3.width                    = 340
$Panel3.BackColor                = "#1d1d1d"
$Panel3.location                 = New-Object System.Drawing.Point(10,160)

$Panel4                          = New-Object system.Windows.Forms.Panel
$Panel4.height                   = 65
$Panel4.width                    = 340
$Panel4.BackColor                = "#1d1d1d"
$Panel4.location                 = New-Object System.Drawing.Point(10,235)

$Panel5                          = New-Object system.Windows.Forms.Panel
$Panel5.height                   = 65
$Panel5.width                    = 340
$Panel5.BackColor                = "#1d1d1d"
$Panel5.location                 = New-Object System.Drawing.Point(10,310)

$Panel6                          = New-Object system.Windows.Forms.Panel
$Panel6.height                   = 65
$Panel6.width                    = 340
$Panel6.BackColor                = "#1d1d1d"
$Panel6.location                 = New-Object System.Drawing.Point(10,385)

$Panel7                          = New-Object system.Windows.Forms.Panel
$Panel7.height                   = 65
$Panel7.width                    = 340
$Panel7.BackColor                = "#1d1d1d"
$Panel7.location                 = New-Object System.Drawing.Point(10,460)

$Panel8                          = New-Object system.Windows.Forms.Panel
$Panel8.height                   = 65
$Panel8.width                    = 340
$Panel8.BackColor                = "#1d1d1d"
$Panel8.location                 = New-Object System.Drawing.Point(10,535)

#Labels (numbered from top to bottom of GUI) creation with custom attributes
$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Copies file(s) to hostnames"
$Label1.BackColor                = "#1d1d1d"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(20,15)
$Label1.Font                     = 'Microsoft Sans Serif,9'
$Label1.ForeColor                = "#e0e0e0"

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Returns OU, ping status, and AD groups for hostnames"
$Label2.BackColor                = "#1d1d1d"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(20,90)
$Label2.Font                     = 'Microsoft Sans Serif,9'
$Label2.ForeColor                = "#e0e0e0"

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "Adds list of hostnames to a specified AD group"
$Label3.BackColor                = "#1d1d1d"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(20,165)
$Label3.Font                     = 'Microsoft Sans Serif,9'
$Label3.ForeColor                = "#e0e0e0"

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Compares KB list versus installed KBs on hostnames"
$Label4.BackColor                = "#1d1d1d"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(20,240)
$Label4.Font                     = 'Microsoft Sans Serif,9'
$Label4.ForeColor                = "#e0e0e0"

$Label5                          = New-Object system.Windows.Forms.Label
$Label5.text                     = "Moves hostnames to specified OU"
$Label5.BackColor                = "#1d1d1d"
$Label5.AutoSize                 = $true
$Label5.width                    = 25
$Label5.height                   = 10
$Label5.location                 = New-Object System.Drawing.Point(20,315)
$Label5.Font                     = 'Microsoft Sans Serif,9'
$Label5.ForeColor                = "#e0e0e0"

$Label6                          = New-Object system.Windows.Forms.Label
$Label6.text                     = "Removes list of hostnames from AD"
$Label6.BackColor                = "#1d1d1d"
$Label6.AutoSize                 = $true
$Label6.width                    = 25
$Label6.height                   = 10
$Label6.location                 = New-Object System.Drawing.Point(20,390)
$Label6.Font                     = 'Microsoft Sans Serif,9'
$Label6.ForeColor                = "#e0e0e0"

$Label7                          = New-Object system.Windows.Forms.Label
$Label7.text                     = "Renames hostnames to specified naming scheme"
$Label7.BackColor                = "#1d1d1d"
$Label7.AutoSize                 = $true
$Label7.width                    = 25
$Label7.height                   = 10
$Label7.location                 = New-Object System.Drawing.Point(20,465)
$Label7.Font                     = 'Microsoft Sans Serif,9'
$Label7.ForeColor                = "#e0e0e0"

$Label8                          = New-Object system.Windows.Forms.Label
$Label8.text                     = "Adds NSHData5 and NSHDataRS to hostnames"
$Label8.BackColor                = "#1d1d1d"
$Label8.AutoSize                 = $true
$Label8.width                    = 25
$Label8.height                   = 10
$Label8.location                 = New-Object System.Drawing.Point(20,540)
$Label8.Font                     = 'Microsoft Sans Serif,9'
$Label8.ForeColor                = "#e0e0e0"

#Buttons (numbered from top to bottom of GUI) creation with custom attributes
$Button1                         = New-Object system.Windows.Forms.Button
$Button1.BackColor               = "#0099bc"
$Button1.text                    = "Copy-Files"
$Button1.width                   = 110
$Button1.height                  = 32
$Button1.location                = New-Object System.Drawing.Point(20,35)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.BackColor               = "#2d7d9a"
$Button2.text                    = "Hostname-Report"
$Button2.width                   = 130
$Button2.height                  = 32
$Button2.location                = New-Object System.Drawing.Point(20,110)
$Button2.Font                    = 'Microsoft Sans Serif,10'

$Button3                         = New-Object system.Windows.Forms.Button
$Button3.BackColor               = "#00b7c3"
$Button3.text                    = "Add-ADGroup"
$Button3.width                   = 110
$Button3.height                  = 32
$Button3.location                = New-Object System.Drawing.Point(20,185)
$Button3.Font                    = 'Microsoft Sans Serif,10'

$Button4                         = New-Object system.Windows.Forms.Button
$Button4.BackColor               = "#038387"
$Button4.text                    = "Compare-KBs"
$Button4.width                   = 110
$Button4.height                  = 32
$Button4.location                = New-Object System.Drawing.Point(20,260)
$Button4.Font                    = 'Microsoft Sans Serif,10'

$Button5                         = New-Object system.Windows.Forms.Button
$Button5.BackColor               = "#00b294"
$Button5.text                    = "Move-ADObject"
$Button5.width                   = 110
$Button5.height                  = 32
$Button5.location                = New-Object System.Drawing.Point(20,335)
$Button5.Font                    = 'Microsoft Sans Serif,10'

$Button6                         = New-Object system.Windows.Forms.Button
$Button6.BackColor               = "#018574"
$Button6.text                    = "Remove-ADComputer"
$Button6.width                   = 150
$Button6.height                  = 32
$Button6.location                = New-Object System.Drawing.Point(20,410)
$Button6.Font                    = 'Microsoft Sans Serif,10'

$Button7                         = New-Object system.Windows.Forms.Button
$Button7.BackColor               = "#00cc6a"
$Button7.text                    = "Rename-Computer"
$Button7.width                   = 130
$Button7.height                  = 32
$Button7.location                = New-Object System.Drawing.Point(20,485)
$Button7.Font                    = 'Microsoft Sans Serif,10'

$Button8                         = New-Object system.Windows.Forms.Button
$Button8.BackColor               = "#10893e"
$Button8.text                    = "WiFi-Profiles"
$Button8.width                   = 110
$Button8.height                  = 32
$Button8.location                = New-Object System.Drawing.Point(20,560)
$Button8.Font                    = 'Microsoft Sans Serif,10'

#Panels are sent to the back of the form
$Panel1.SendToBack();
$Panel2.SendToBack();
$Panel3.SendToBack();
$Panel4.SendToBack();
$Panel5.SendToBack();
$Panel6.SendToBack();
$Panel7.SendToBack();
$Panel8.SendToBack();

#Style is applied to each button
$Button1.FlatStyle = "Popup";
$Button2.FlatStyle = "Popup";
$Button3.FlatStyle = "Popup";
$Button4.FlatStyle = "Popup";
$Button5.FlatStyle = "Popup";
$Button6.FlatStyle = "Popup";
$Button7.FlatStyle = "Popup";
$Button8.FlatStyle = "Popup";

#Panels are added to the Form
$Form.controls.AddRange(@($Label1,$Label2,$Label3,$Label4,$Label5,$Label6,$Label7,$Label8,$Button1,$Button2,$Button3,$Button4,$Button5,$Button6,$Button7,$Button8,$Panel1,$Panel2,$Panel3,$Panel4,$Panel5,$Panel6,$Panel7,$Panel8))

#Button click events are created
$Button1.Add_Click({ Button1_Click })
$Button2.Add_Click({ Button2_Click })
$Button3.Add_Click({ Button3_Click })
$Button4.Add_Click({ Button4_Click })
$Button5.Add_Click({ Button5_Click })
$Button6.Add_Click({ Button6_Click })
$Button7.Add_Click({ Button7_Click })
$Button8.Add_Click({ Button8_Click })

#Button 1 click event runs Copy-Files.ps1
function Button1_Click {
    Start-Process Powershell.exe -ArgumentList "-NoExit -ExecutionPolicy Bypass -File .\Copy-Files.ps1"
}

#Button 2 click event runs Hostname-Report.ps1
function Button2_Click {
    Start-Process Powershell.exe -ArgumentList "-NoExit -ExecutionPolicy Bypass -File .\Hostname-Report.ps1"
}

#Button 3 click event runs Add-ADPrincipalGroupMember.ps1
function Button3_Click {
    Start-Process Powershell.exe -ArgumentList "-NoExit -ExecutionPolicy Bypass -File .\Add-ADGroup.ps1"
}

#Button 4 click event runs List-KBs.ps1
function Button4_Click {
    Start-Process Powershell.exe -ArgumentList "-NoExit -ExecutionPolicy Bypass -File .\Compare-KBs.ps1"
}

#Button 5 click event runs Move-ADObject_withSelection.ps1
function Button5_Click {
    Start-Process Powershell.exe -ArgumentList "-NoExit -ExecutionPolicy Bypass -File .\Move-ADObject.ps1"
}

#Button 6 click event runs Remove-ADComputer.ps1
function Button6_Click {
    Start-Process Powershell.exe -ArgumentList "-NoExit -ExecutionPolicy Bypass -File .\Remove-ADComputer.ps1"
}

#Button 7 click event runs Rename-Computer.ps1
function Button7_Click {
    Start-Process Powershell.exe -ArgumentList "-NoExit -ExecutionPolicy Bypass -File .\Rename-Computer.ps1"
}

#Button 8 click event runs Network-TSing.ps1
function Button8_Click {
    Start-Process Powershell.exe -ArgumentList "-NoExit -ExecutionPolicy Bypass -File .\WiFi-Profiles.ps1"
}

[void]$Form.ShowDialog()