Clear-Host

Write-Host "Please select text or csv file with list of hostnames." -f Yellow

#Begin Function file explorer pop-up selection
Function Get-OpenFile($initialDirectory) { 
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = "c:\\"
    $OpenFileDialog.filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv"
    $OpenFileDialog.title = "### Please select text or csv file with list of hostnames ###"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
    $OpenFileDialog.ShowHelp = $true
    }
#End Funtion file explorer pop-up selection

try {    
    $InputFile = Get-OpenFile
    $computers = Get-Content $InputFile
    } catch {
        Write-Host "Text or csv file must be selected. Script ended." -f Red
        exit
}

Write-Host "File selected." -f Cyan

Write-Host "`nPlease enter your AD credentials." -f Yellow

try {
    $creds = Get-Credential -Credential $null
    } catch {
        Write-Host "Credentials must be entered. Script ended." -f Red
        exit
}

Write-Host "Credentials securely stored." -f Cyan

Function Add-ADGroup($groupName) {
    Add-ADPrincipalGroupMembership -Identity $c$ -MemberOf $groupName -Credential $creds
}

Write-Host "`nSelect the AD group that will be added to each hostname:" -f Yellow 
Write-Host "(1) " -f Yellow -NoNewline
Write-Host "Office_2016_installer"
Write-Host "(2) " -f Yellow -NoNewline
Write-Host "Greenway-Computer"
Write-Host "(3) " -f Yellow -NoNewline
Write-Host "MBAM_Installer"
Write-Host "(4) " -f Yellow -NoNewline
Write-Host "JavaRuleSet_Bypass_Computers"

do {
    $choice = Read-Host -Prompt "`nPlease choose an option"
} while(($choice -lt 1) -or ($choice -gt 4))

ForEach ($c in $computers) {
    switch($choice) {
        1 {Add-ADGroup("Office_2016_installer")
           $group = "Office_2016_installer"}
	    2 {Add-ADGroup("Greenway-Computers")
           $group = "Greenway-Computers"}
        3 {Add-ADGroup("MBAM_Installer")
           $group = "MBAM_Installer"}
        4 {Add-ADGroup("JavaRuleSet_Bypass_Computers")
           $group = "JavaRuleSet_Bypass_Computers"}
        default {"Please choose an option"}
    }

    Write-Host "Added $c to " -NoNewline
    Write-Host "$group" -f Cyan
}