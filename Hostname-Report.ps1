Clear-Host

$ErrorActionPreference = "Continue"

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

Write-Host "`nPlease select location to store report." -f Yellow

#Begin Function folder explorer pop-up selection
Function Get-OpenFolder {
    [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') |
        Out-Null

    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.RootFolder = 'MyComputer'
    $FolderBrowserDialog.Description = "### Please select location to store report ###"
    $FolderBrowserDialog.ShowDialog() | Out-Null
    return $FolderBrowserDialog.SelectedPath
}
#End Funtion folder explorer pop-up selection

$OutputFile = Get-OpenFolder
if (!$OutputFile) {
    Write-Host "Location must be selected. Script ended." -f Red
    exit
}

Write-Host "Location selected." -f Cyan

Write-Host "`nCreating Hostname Report spreadsheet on your desktop... Please wait..." -ForegroundColor Yellow

<# Credit for Convert-LHSADName goes to PLantell from a TechNet Gallery
   Link - https://gallery.technet.microsoft.com/scriptcenter/Translating-Active-5c80dd67/view/Discussions
#>

Function Convert-LHSADName 
{ 
    
[cmdletbinding()]   
 
[OutputType('System.String')]  
 
Param( 
 
    [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True, 
        HelpMessage='Active Directory Object Name you want to translate.')] 
    [string]$Identity, 
 
    [Parameter(Position=1,Mandatory=$True, 
        HelpMessage='Specify the format used for representing distinguished names.')] 
    [ValidateSet("DN","Canonical","NT4","Display","DomainSimple","EnterpriseSimple","GUID","UPN","CanonicalEx","SPN","SID")]   
    [string]$OutputType, 
 
    [Parameter(Position=2,HelpMessage='Type of binding to perform on a Name Translate.')]  
    [ValidateSet("GC","Domain","Server")]    
    [String]$InitType="GC", 
 
    [Parameter(Position=3,HelpMessage='Translation is performed on objects that do not belong to this directory and are the referrals returned from referral chasing')] 
    [Switch]$ChaseReferrals, 
 
    [Parameter(Position=4)] 
    [Alias('RunAs')] 
    [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty 
 
   ) 
 
DynamicParam { 
    if ( ($PSBoundParameters['InitType']) -and ($InitType -ne "GC") ) 
    { 
        <# 
        dynamically add a new parameter called -InitName when -InitType is not 'GC' (Global Catalog Server)  
        #> 
         
        #create a new ParameterAttribute Object 
        $Attribute = New-Object -TypeName System.Management.Automation.ParameterAttribute 
        $Attribute.Position = 3 
        $Attribute.Mandatory = $true 
        $Attribute.HelpMessage = "Supply the domain name within a directory forest or a machine name of a directory server." 
         
        #create an attributecollection object for the attribute we just created. 
        $attributeCollection = New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute] 
         
        #add our custom attribute 
        $attributeCollection.Add($Attribute) 
         
        #add our paramater specifying the attribute collection 
        $ParameterName = 'InitName' 
        $InitName_Param = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $attributeCollection) 
         
        #expose the name of our parameter 
        $paramDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary 
        $paramDictionary.Add($ParameterName, $InitName_Param) 
        return $paramDictionary 
    } 
} 
 
 
BEGIN  
{ 
    Set-StrictMode -Version 2.0 
    ${CmdletName} = $Pscmdlet.MyInvocation.MyCommand.Name 
 
 
    If (-not($PSBoundParameters['InitName'])) 
    { 
        $InitName = $null 
    } 
    else 
    { 
        $InitName = $PSBoundParameters.InitName 
    } 
    Write-Debug ("Parameters:`n{0}" -f ($PSBoundParameters | Out-String)) 
 
 
    #region hash tables 
 
    # https://msdn.microsoft.com/en-us/library/aa772267%28v=vs.85%29.aspx 
    $ADS_NAME_TYPE_ENUM = @{  
        ADS_NAME_TYPE_1779                    = 1;  # Name format as specified in RFC 1779. For example, "CN=Jeff Smith,CN=users,DC=Fabrikam,DC=com". 
        ADS_NAME_TYPE_CANONICAL               = 2;  # Canonical name format. For example, "Fabrikam.com/Users/Jeff Smith". 
        ADS_NAME_TYPE_NT4                     = 3;  # Account name format used in Windows. For example, "Fabrikam\JeffSmith". 
        ADS_NAME_TYPE_DISPLAY                 = 4;  # Display name format. For example, "Jeff Smith". 
        ADS_NAME_TYPE_DOMAIN_SIMPLE           = 5;  # Simple domain name format. For example, "JeffSmith@Fabrikam.com". 
        ADS_NAME_TYPE_ENTERPRISE_SIMPLE       = 6;  # Simple enterprise name format. For example, "JeffSmith@Fabrikam.com". 
        ADS_NAME_TYPE_GUID                    = 7;  # Global Unique Identifier format. For example, "{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}". 
        ADS_NAME_TYPE_UNKNOWN                 = 8;  <# Unknown name type. The system will estimate the format. This element is a meaningful option  
                                                       only with the IADsNameTranslate.Set or the IADsNameTranslate.SetEx method,  
                                                       but not with the IADsNameTranslate.Get or IADsNameTranslate.GetEx method.#> 
        ADS_NAME_TYPE_USER_PRINCIPAL_NAME     = 9;  # User principal name format. For example, "JeffSmith@Fabrikam.com". 
        ADS_NAME_TYPE_CANONICAL_EX            = 10; # Extended canonical name format. For example, "Fabrikam.com/Users Jeff Smith". 
        ADS_NAME_TYPE_SERVICE_PRINCIPAL_NAME  = 11; # Service principal name format. For example, "www/www.fabrikam.com@fabrikam.com". 
        ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME = 12; <# A SID string, as defined in the Security Descriptor Definition Language (SDDL),  
                                                       for either the SID of the current object or one from the object SID history.  
                                                       For example, "O:AOG:DAD:(A;;RPWPCCDCLCSWRCWDWOGA;;;S-1-0-0)" For more information,  
                                                       see Security Descriptor String Format.#> 
    } 
 
    # https://msdn.microsoft.com/en-us/library/aa772266%28v=vs.85%29.aspx 
    $ADS_NAME_INITTYPE_ENUM = @{  
        ADS_NAME_INITTYPE_DOMAIN = 1; # Initializes a NameTranslate object by setting the domain that the object binds to. 
        ADS_NAME_INITTYPE_SERVER = 2; # Initializes a NameTranslate object by setting the server that the object binds to. 
        ADS_NAME_INITTYPE_GC     = 3; # Initializes a NameTranslate object by locating the global catalog that the object binds to. 
    }  
     
    # https://msdn.microsoft.com/en-us/library/aa772250.aspx 
    $ADS_CHASE_REFERRALS_ENUM = @{  
        ADS_CHASE_REFERRALS_NEVER       = (0x00); #The client should never chase the referred-to server. Setting this option prevents a client from contacting other servers in a referral process. 
        ADS_CHASE_REFERRALS_SUBORDINATE = (0x20); #The client chases only subordinate referrals which are a subordinate naming context in a directory tree. For example, if the base search is requested for "DC=Fabrikam,DC=Com", and the server returns a result set and a referral of "DC=Sales,DC=Fabrikam,DC=Com" on the AdbSales server, the client can contact the AdbSales server to continue the search. The ADSI LDAP provider always turns off this flag for paged searches. 
        ADS_CHASE_REFERRALS_EXTERNAL    = (0x40); #The client chases external referrals. For example, a client requests server A to perform a search for "DC=Fabrikam,DC=Com". However, server A does not contain the object, but knows that an independent server, B, owns it. It then refers the client to server B. 
        ADS_CHASE_REFERRALS_ALWAYS      = (0x60); #Referrals are chased for either the subordinate or external type. 
    }    
     
    #endregion hash tables 
 
 
    $ADS_InitType = switch($InitType) 
        { 
            'Domain' {$ADS_NAME_INITTYPE_ENUM.ADS_NAME_INITTYPE_DOMAIN} 
            'Server' {$ADS_NAME_INITTYPE_ENUM.ADS_NAME_INITTYPE_SERVER} 
            'GC'     {$ADS_NAME_INITTYPE_ENUM.ADS_NAME_INITTYPE_GC} 
            default  {$ADS_NAME_INITTYPE_ENUM.ADS_NAME_INITTYPE_GC} 
        }     
 
    $ADS_OutputType = switch($OutputType) 
        { 
            "DN"               {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_1779} 
            "Canonical"        {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_CANONICAL} 
            "NT4"              {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_NT4} 
            "Display"          {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_DISPLAY} 
            "DomainSimple"     {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_DOMAIN_SIMPLE} 
            "EnterpriseSimple" {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_ENTERPRISE_SIMPLE} 
            "GUID"             {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_GUID} 
            "UPN"              {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_USER_PRINCIPAL_NAME} 
            "CanonicalEx"      {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_CANONICAL_EX} 
            "SPN"              {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_SERVICE_PRINCIPAL_NAME} 
            "SID"              {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME} 
            "Unkonwn"          {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_UNKNOWN} 
            default            {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_UNKNOWN} 
        } 
 
 
    #region Functions 
 
    # Accessor functions from Bill Stewart to simplify calls to NameTranslate 
    function Invoke-Method([__ComObject] $object, [String] $method, $parameters)  
    { 
        $output = $Null 
        $output = $object.GetType().InvokeMember($method, "InvokeMethod", $NULL, $object, $parameters) 
        Write-Output $output 
    } 
 
    function Get-Property([__ComObject] $object, [String] $property)  
    { 
        $object.GetType().InvokeMember($property, "GetProperty", $NULL, $object, $NULL) 
    } 
 
    function Set-Property([__ComObject] $object, [String] $property, $parameters)  
    { 
        [Void] $object.GetType().InvokeMember($property, "SetProperty", $NULL, $object, $parameters) 
    } 
 
    #endregion Functions 
 
 
} # end BEGIN 
 
PROCESS  
{ 
    #region Initialize IADsNameTranslate 
    $NameTranslate = New-Object -ComObject NameTranslate 
 
    If ($PSBoundParameters['Credential']) 
    { 
        Try 
        { 
            $Cred = $Credential.GetNetworkCredential() 
  
            Invoke-Method $NameTranslate "InitEx" ( 
                $ADS_InitType, 
                $InitName, 
                $Cred.UserName, 
                $Cred.Domain, 
                $Cred.Password 
            ) 
        } 
        Catch [System.Management.Automation.MethodInvocationException]  
        { 
            Write-Error $_ 
            break 
        } 
        Finally  
        { 
            Remove-Variable Cred 
        } 
    } 
    Else 
    { 
        Try  
        { 
            Invoke-Method $NameTranslate "Init" ( 
                $ADS_InitType, 
                $InitName 
            ) 
        } 
        Catch [System.Management.Automation.MethodInvocationException]  
        { 
            Write-Error $_ 
            break 
        } 
    } 
    #endregion Initialize IADsNameTranslate 
 
 
    If ($PSBoundParameters['ChaseReferrals'])  
    { 
        Set-Property $NameTranslate "ChaseReferral" ($ADS_CHASE_REFERRALS_ENUM.ADS_CHASE_REFERRALS_ALWAYS) 
    } 
 
 
    Try 
    { 
        Invoke-Method $NameTranslate "Set" ($ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_UNKNOWN,$Identity) 
        Invoke-Method $NameTranslate "Get" ($ADS_OutputType) 
    } 
    Catch [System.Management.Automation.MethodInvocationException]  
    { 
        Write-Error "'$Identity' - $($_.Exception.InnerException.Message)" 
    } 
 
} # end PROCESS 
 
END { Write-Verbose "Function ${CmdletName} finished." } 
 
} # end Function Convert-LHSADName              
 
#Export-ModuleMember -Function Convert-LHSADName 

<# Excel spreadsheet creation credit goes to Boe Prox from https://learn-powershell.net/
   Link - https://learn-powershell.net/2012/12/20/powershell-and-excel-adding-some-formatting-to-your-report/
#>

#Create excel COM object
$excel = New-Object -ComObject excel.application

#Make Visible
$excel.Visible = $False

#Add a workbook
$workbook = $excel.Workbooks.Add()

#Connect to first worksheet to rename and make active
$serverInfoSheet = $workbook.Worksheets.Item(1)
$serverInfoSheet.Name = 'Hostname Report'
$serverInfoSheet.Activate() | Out-Null

#Create a Title for the first worksheet and adjust the font
$row = 1
$Column = 1
$serverInfoSheet.Cells.Item($row,$column)= 'Hostname Report'

$serverInfoSheet.Cells.Item($row,$column).Font.Size = 18
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$serverInfoSheet.Cells.Item($row,$column).Font.Name = "Cambria"
$serverInfoSheet.Cells.Item($row,$column).Font.ThemeFont = 1
$serverInfoSheet.Cells.Item($row,$column).Font.ThemeColor = 4
$serverInfoSheet.Cells.Item($row,$column).Font.ColorIndex = 55
$serverInfoSheet.Cells.Item($row,$column).Font.Color = 8210719

$range = $serverInfoSheet.Range("a1","d2")
$range.Merge() | Out-Null
$range.VerticalAlignment = -4160

#Increment row for next set of data
$row++;$row++

#Save the initial row so it can be used later to create a border
$initalRow = $row

#Create a header for Hostname Report; set each cell to Bold and add a background color
$serverInfoSheet.Cells.Item($row,$column)= 'Hostname'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'Organizational Unit'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'Ping Status'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'AD groups'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True

#Increment Row and reset Column back to first column
$row++
$Column = 1

#Loop through list of computers and gather data for each row
ForEach ($c in $computers) {

    #Hostname
    $serverInfoSheet.Cells.Item($row,$column) = $c.ToUpper()
    $Column++

    #OU
    $orgUnit = (Get-ADComputer $c).DistinguishedName
    $orgConversion = $orgUnit | Convert-LHSADName -OutputType Canonical
    $orgFormatted = (($orgConversion | Out-String).Trim() -replace '/[^/]*$')
    $serverInfoSheet.Cells.Item($row,$column) = $orgFormatted
    $Column++

    #Ping results
    if(Test-Connection -ComputerName $c -Count 3 -ErrorAction SilentlyContinue) {
        $serverInfoSheet.Cells.Item($row,$column) = "Online"
        } else {
        $serverInfoSheet.Cells.Item($row,$column) = "Offline"
        }
    $Column++

    #AD groups
    $groups = (([adsisearcher]"(&(objectCategory=computer)(cn=$c))").FindOne().Properties.memberof -replace '^CN=([^,]+).+$','$1')
    $groupsFormatted = ($groups | Out-String).Trim()
    $serverInfoSheet.Cells.Item($row,$column) = $groupsFormatted

    #Increment to next row and reset Column to 1
    $Column = 1
    $row++
    }

$row--
$dataRange = $serverInfoSheet.Range(("A{0}"  -f $initalRow),("D{0}"  -f $row))
7..12 | ForEach {
    $dataRange.Borders.Item($_).LineStyle = 1
    $dataRange.Borders.Item($_).Weight = 2
}

#Auto fit everything so it looks better
$usedRange = $serverInfoSheet.UsedRange	
$serverInfoSheet.columns.item('d').columnWidth = 60
$usedRange.EntireColumn.AutoFit() | Out-Null
$usedRange.EntireRow.AutoFit() | Out-Null

#Timestamp variable for saved file name
$timestamp = $(((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmss"))

#Save the file
$workbook.SaveAs($OutputFile + "\Hostname_Report_$timestamp.xlsx")

#Quit the application
$excel.Quit()

#Release COM Object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null

Write-Host "Hostname Report complete." -f Cyan