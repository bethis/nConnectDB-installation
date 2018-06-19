Param
(
[string] $ApplicationName, #Name Displayed in Control Panel or in Registry  
[string] $ProjectName,  #Hammertime or Callisto or RoyalDream
[string] $ProjectPhase,   #Dev/Release/Hotfix 
[string] $BuildVersion,
[string] $BuildNumber 
)



$BuildNumberSplit = $BuildNumber.Split('.')
$BuildRev = $BuildNumberSplit[$BuildNumberSplit.Length - 1]
$ProductVersion="$BuildVersion.$BuildRev"

Write-Verbose "Product Version to be installed: $ProductVersion"
    
Write-Verbose "$ApplicationName $ProjectName $ProjectPhase"
            
$AppNameXML=$ApplicationName -replace '\s',''     #Product Name Without WhiteSpace, Same as to be written in XML Configuration File  

$CommandVariables=" /s /v/qn"                     #To Install Silently  
#XML Configuration At Target Server
$xmlFileLocation="C:\ABDT\$ProjectName\$ProjectPhase\ApplicationConfiguration.xml"
[xml]$xmlDocument=Get-Content $xmlFileLocation

#To Pick exe file from Target Server
$InstallerName=$xmlDocument.Settings.$AppNameXML | Select InstallerName
$EXEPath="C:\ABDT\"+$ProjectName+"\"+"$ProjectPhase"+"\"+"$ApplicationName"
$ExeName=$InstallerName.InstallerName+"*.exe"
$InstallerPath=Get-ChildItem $EXEPath -Filter $ExeName

#Getting Information About Installed Application
Write-Verbose "Installed Application Detail for 32 bit:"
$InstalledProductDetails=Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |  Select-Object DisplayName, DisplayVersion, InstallLocation, Publisher, InstallDate | Where {$_.DisplayName -like "$ApplicationName*"}
Write-Verbose " $InstalledProductDetails"
Write-Verbose "Installed Application Detail for 64 bit:"
$InstalledProductDetails=Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |  Select-Object  DisplayName, DisplayVersion,InstallLocation, Publisher, InstallDate | Where {$_.DisplayName -like "$ApplicationName*"}
Write-Verbose "$InstalledProductDetails "


##Uninstallation of Application
$uninstallCommand="/c wmic product where name=""$ApplicationName"" call uninstall /nointeractive"
Write-Verbose "Running Command to Uninstall: $uninstallCommand"
$uninstallProc=start-process cmd  $uninstallCommand -Wait

## Read XML
$xmlData=$xmlDocument.SelectNodes("/Settings/$AppNameXML/$ProjectName/$ProjectPhase")[0]
foreach ($Data in $xmlData.ChildNodes) 
{
    if($Data.Name -ne "#comment")
    {
        $Field=$Data.Name
        Write-Verbose "Field: $Field"
        $Value=$Data.InnerText
        Write-Verbose "Value: $Value"
        if($Value -match "^\d+$")
        {
            $CommandVariables=$CommandVariables + " /v""$Field=$Value"""
        } 
        else
        {
            $CommandVariables=$CommandVariables + " /v""$Field=\""$Value\"""""
        }
    }

}

##Installation
Write-Verbose "Running Command to Install: $($InstallerPath.FullName) -ArgumentList $CommandVariables"
$Installationproc=Start-Process $InstallerPath.FullName  -ArgumentList "$CommandVariables"  -Wait

##Verifying
$ProductInstalled="False"
Write-Verbose "Installed Application Detail for 32 bit:"
$InstalledProductDetails=Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |  Select-Object DisplayName, DisplayVersion, InstallLocation, Publisher, InstallDate | Where {$_.DisplayName -like "$ApplicationName*"}
Write-Verbose " $InstalledProductDetails"
if($InstalledProductDetails -ne $null)
{
    if($ProductVersion -match $($InstalledProductDetails[0].DisplayVersion))
    {
        $ProductInstalled="True"
    }
}
Write-Verbose "Installed Application Detail for 64 bit:"
$InstalledProductDetails=Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |  Select-Object  DisplayName, DisplayVersion,InstallLocation, Publisher, InstallDate | Where {$_.DisplayName -like "$ApplicationName*"}
if($InstalledProductDetails -ne $null)
{
    if($ProductVersion -match $($InstalledProductDetails[0].DisplayVersion))
    {
        $ProductInstalled="True"
    }
}
Write-Verbose "$InstalledProductDetails"
Write-Verbose "Product Installed: $ProductInstalled"
if($ProductInstalled -match "False")
{
    Write-Error "Installed version is not matching the product version"
    exit 1
 }    
    