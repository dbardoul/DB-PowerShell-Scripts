Clear-Host

Write-Host "Please select text or csv file with list of hostnames." -f Yellow

#Begin Function file explorer pop-up selection
Function Get-OpenFile { 
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

Write-Host "File selected.`n" -f Cyan
    
$sourcefile = ".\WiFi-Profiles\*.xml"
    
ForEach ($c in $computers) {
    try {
        $destinationFolder = "\\$c\C$\temp\WiFi-Profiles"
    
        if (!(Test-Path -path $destinationFolder)) {
            New-Item $destinationFolder -Type Directory | Out-Null
        }

        Copy-Item -Path $sourcefile -Destination $destinationFolder -Recurse
    } catch {
            Write-Host "Could not copy to $c"
    }
    
    
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $c) 
    $regKey = $reg.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WinRM\\Service\\WinRS",$true)
    
    $winRMStatus = Get-Service -Name winRM -ComputerName $c
    
    if (($regKey.GetValue("AllowRemoteShellAccess") -eq 0) -and ($winRMStatus.Status -eq "Stopped")) {
        $regKey.SetValue("AllowRemoteShellAccess","1",[Microsoft.Win32.RegistryValueKind]::DWORD)
        $winRMStatus | Set-Service -Status Running
    }
    if (($regKey.GetValue("AllowRemoteShellAccess") -eq 0) -and ($winRMStatus.Status -eq "Running")) {
        $winRMStatus | Stop-Service -Force
        $regKey.SetValue("AllowRemoteShellAccess","1",[Microsoft.Win32.RegistryValueKind]::DWORD)
        $winRMStatus | Set-Service -Status Running
    }
    if (($regKey.GetValue("AllowRemoteShellAccess") -eq 1) -and ($winRMStatus.Status -eq "Stopped")) {
        $winRMStatus | Set-Service -Status Running
    }
    
    #Begin remote script
    $s = New-PSSession -ComputerName $c
    Invoke-Command -Session $s -ScriptBlock `
    {
        netsh wlan add profile filename="C:\temp\WiFi-Profiles\Wi-Fi-NSHDataRS.xml"
        netsh wlan add profile filename="C:\temp\WiFi-Profiles\Wi-Fi-NSHData5.xml"
    } | Out-Null
    Remove-PSSession $s
    #End remote script

    Write-Host "Added NSHData5 and NSHDataRS to " -NoNewline
    Write-Host $c -f Yellow
    
    $regKey.SetValue("AllowRemoteShellAccess","0",[Microsoft.Win32.RegistryValueKind]::DWORD)
    
    $winRMStatus | Stop-Service -Force
    
    try {
        Remove-Item –Path "\\$c\c$\temp\WiFi-Profiles\*" –recurse
    } catch {
        Write-Host "Could not remove from $c"
    }
}