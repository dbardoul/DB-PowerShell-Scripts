Clear-Host

Write-Host "Please select text or csv file with list of hostnames." -f Yellow

#Begin Function file explorer pop-up selection
Function Get-OpenFile($initialDirectory) { 
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
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

Write-Host "Credentials securely stored.`n" -f Cyan

ForEach ($c in $computers) {
    Remove-ADComputer -Identity $c -Credential $creds

    Write-Host "Removed " -NoNewline
    Write-Host "$c " -f Cyan -NoNewline
    Write-Host "from AD"
}