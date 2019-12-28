#region variables
$eporeport = @()
$filepath = "E:\SSS-win-scripts\output\"
$filename ="epo_data_" + (Get-Date).ToString().Replace(" ","T").Replace(":",".") + "_" + (whoami).replace("\",".")
#endregion

[System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") | Out-Null 
$username = Read-Host -Prompt "Please enter your username : "
$domainfqdn = "fqdn.com"
$adcred = New-Object –TypeName "System.Management.Automation.PSCredential" –ArgumentList ($domainfqdn + "\"+$shsdirusername), $(Read-Host -Prompt "Please enter your password : " -AsSecureString)
$adcred | %{"password correct ? " + (New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain, $($_.GetNetworkCredential().Domain ) ) ).ValidateCredentials($_.GetNetworkCredential().UserName, $_.GetNetworkCredential().Password).ToString()}

$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($adcred.Password))
$epopwd = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
$epocred = new-object System.Net.NetworkCredential($username, ($epopwd | ConvertTo-SecureString -AsPlainText -force))
$eposerver = "eposerver"
$epoport = 9005


#region ePo
# Code specific for ePo API
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
$netAssembly = [Reflection.Assembly]::GetAssembly([System.Net.Configuration.SettingsSection])
if($netAssembly)
{
    $bindingFlags = [Reflection.BindingFlags] "Static,GetProperty,NonPublic"
    $settingsType = $netAssembly.GetType("System.Net.Configuration.SettingsSectionInternal")

    $instance = $settingsType.InvokeMember("Section", $bindingFlags, $null, $null, @())

    if($instance)
    {
        $bindingFlags = "NonPublic","Instance"
        $useUnsafeHeaderParsingField = $settingsType.GetField("useUnsafeHeaderParsing", $bindingFlags)

        if($useUnsafeHeaderParsingField)
        {
          $useUnsafeHeaderParsingField.SetValue($instance, $true)
        }
    }
}
[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
$url=  "https://" + $eposerver + ":" + $epoport + "/remote/core.executeQuery?target=EPOLeafNode&select=(select+EPOComputerProperties.ComputerName+EPOComputerProperties.DomainName+EPOComputerProperties.IPAddress+EPOComputerProperties.Description+EPOLeafNode.Tags+EPOLeafNode.os+EPOLeafNode.LastUpdate+EPOProdPropsView_VIRUSCAN.productversion+EPOProdPropsView_VIRUSCAN.datver)"
$wc = new-object System.net.WebClient 
$wc.Credentials = $epocred
$webpage = $wc.DownloadString($url)

$nl = [System.Environment]::NewLine
$systemnames = ($webpage -split "$nl$nl").split($nl) | %{if($_ -like "*System Name:*"){$_ -replace "System Name: "}} 

foreach($systemname in $systemnames)
{

$domainname = (($webpage -split "$nl$nl") -match $systemname).split($nl)  | %{if($_ -like "*Domain Name*"){$_.Split(":")[1].trimstart()}}
$IPaddress = (($webpage -split "$nl$nl") -match $systemname).split($nl)  | %{if($_ -like "*IP address*"){$_.Split(":")[1].trimstart()}}
$Tags = (($webpage -split "$nl$nl") -match $systemname).split($nl)  | %{if($_ -like "*Tags*"){$_.Split(":")[1].trimstart()}}
$OperatingSystem = (($webpage -split "$nl$nl") -match $systemname).split($nl)  | %{if($_ -like "*Operating System*"){$_.Split(":")[1].trimstart()}}
$LastCommunication = (($webpage -split "$nl$nl") -match $systemname).split($nl)  | %{if($_ -like "*Last Communication*"){$_.Split(":")[1].trimstart()}}
$ProductVersion = (($webpage -split "$nl$nl") -match $systemname).split($nl)  | %{if($_ -like "*Product Version*"){$_.Split(":")[1].trimstart()}}
$DATVersion =(($webpage -split "$nl$nl") -match $systemname).split($nl)  | %{if($_ -like "*DAT Version*"){$_.Split(":")[1].trimstart()}}
$Description =(($webpage -split "$nl$nl") -match $systemname).split($nl)  | %{if($_ -like "*Description*"){$_.Split(":")[1].trimstart()}}
$props = [ordered]@{
         systemname = $systemname
         domainname = $domainname
         IPaddress = $IPaddress
         Tags = $Tags
         Description = $Description
         OperatingSystem = $OperatingSystem
         LastCommunication = $LastCommunication
         ProductVersion = $ProductVersion
         DatVersion = $DATVersion
         }
$obj = New-Object -TypeName PSObject -Property $props
$EPOReport += $obj
}
$eporeport | Export-excel ($filepath + $filename + ".xlsx") -NoNumberConversion $true
$eporeport | Export-Csv ($filepath + $filename + ".csv") -NoClobber -NoTypeInformation -Delimiter ";"
#endregion
