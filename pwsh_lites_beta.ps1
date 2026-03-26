<################################################
 POWERCLI : "LITES" LOOK INSIDE THE ESXI SERVER
 ---
 Get NIC/HBA Driver's Detailed Information
 ================================================
 Version                  :      Beta Version 1.0
 Author                   :        Yuri P. Bodrov         
 Company                  :   PLC Alfastrahovanie             
 E-mail                   : bodrovyp@alfastrah.ru   
 Phone №                  :          +79259929596
 Telegram_ID              :           @YuriBodrov           
################################################>

<### TODO : #####################################
$hbaList = $esxcli.Storage.san.fc.list.Invoke()

    foreach ($hba in $hbaList) {
        $hbaDetails = [PSCustomObject]@{
            HostName        = $esx.Name
            Adapter         = $hba.Adapter
            DriverName      = $hba.DriverName
            DriverVersion   = $hba.DriverVersion
            FirmwareVersion = $hba.FirmwareVersion
            Model           = $hba.Model
        }
        $report += $hbaDetails
    }
################################################>

$vcsaAddress = "vmc-01-new.vesta.ru"

$credentialsForAll = Get-Credential -UserName "administrator@vsphere.local"
Connect-VIServer -Server $vcsaAddress -Credential $credentialsForAll

#$vmhost = "vmware26.vesta.ru"
$vmhost      = "vmware139-sap.vesta.ru"
$esxcli_info = Get-EsxCli -VMHost $vmhost -V2

$esxi_nics = $esxcli_info.network.nic.list.Invoke()    # Get ESXi NICs Information
$esxi_hbas = $esxcli_info.storage.san.fc.list.Invoke() # Get ESXi HBAs Information

ForEach ($nic in $esxi_nics) 
{  
  # Get details for each ESXi NIC
  $nic_details = $esxcli_info.network.nic.get.Invoke(@{nicname = $nic.Name})
        
  $VMHost          = $vmhost
  $Device          = $nic_details.Name
  $Driver          = $nic_details.DriverInfo.Driver
  $DriverVersion   = $nic_details.DriverInfo.Version
  $FirmwareVersion = $nic_details.DriverInfo.FirmwareVersion
  $BusInfo         = $nic_details.DriverInfo.BusInfo
  $Description     = $nic.Description
  $Speed           = $nic.Speed

  # TODO! PSCustomObject - Must Have! 
  [PSCustomObject]@{
    Name        = $Device
    Driver      = $Driver
    Description = $Description
  } | Format-Table -AutoSize
  
  #Get-ChildItem $nic_details | Format-Table -AutoSize
  
  <#
  $report += [PSCustomObject]@{
    VMHost          = $vmhost
    Device          = $nic_details.Name
    Driver          = $nic_details.DriverInfo.Driver
    DriverVersion   = $nic_details.DriverInfo.Version
    FirmwareVersion = $nic_details.DriverInfo.FirmwareVersion
    BusInfo         = $nic_details.DriverInfo.BusInfo
  }
  #>
  Write-Host ("---")
  Write-Host ($Device + " | " + $Driver + " | " + $DriverVersion + " | " + $FirmwareVersion + " | " + $BusInfo + " | " + `
  $Description + " | " + $Speed)
  Write-Host ("")
}

ForEach ($hba in $esxi_hbas)
{
  #HostName        = $esx.Name
  $hbaAdapter         = $hba.Adapter
  $hbaDriverName      = $hba.DriverName
  $hbaDriverVersion   = $hba.DriverVersion
  $hbaFirmwareVersion = $hba.FirmwareVersion
  $hbaModel           = $hba.ModelDescription
  $hbaPortState       = $hba.PortState

  Write-Host ($hbaAdapter + " | " + $hbaDriverName + " | " + $hbaDriverVersion + " | " + $hbaFirmwareVersion + " | " `
  + $hbaModel + " | " + $hbaPortState)
}

# Output to console
#$report | Format-Table -AutoSize
<#
Write-Host ("---")
Write-Host ($Device + " | " + $Driver + " | " + $DriverVersion + " | " + $FirmwareVersion + " | " + $Description)
#>
#Write-Host ($Driver)

Disconnect-VIServer $vcsaAddress -Confirm:$false

