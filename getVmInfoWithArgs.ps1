Clear-Host

Write-Host ("")

Write-Host ("0 : Search in all registered vCenter Servers") -ForegroundColor White
Write-Host ("1 : Search in specified vCenter Server") -ForegroundColor White
Write-Host ("--------------------------------------------") -ForegroundColor DarkGray

### Case Operator for initial operation type: ####
$caseVar = Read-Host "Select an operation's type "

switch ($caseVar) 
{
  0 { Write-Host "You entered 0" }
  1 { Write-Host "You entered 1" }
}


### Read credentials from the keyboard: ##############################
$vCenterInstance = Read-Host "vCenter Server Address "
Write-Host ("-------------------------------------------------------")
$credentials = Get-Credential

### Create the connection to vCSA with a specified creds #########
Connect-VIServer -Server $vCenterInstance -Credential $credentials


### Check if connection was successful: ###############
if (!$DefaultVIServer)
  {
    Write-Host ("Unable to connect to vCenter Server!")
  }

else
  {
    Write-Host ("")
    Write-Host ("Connection successfully initiated! :-)") -ForegroundColor Green

    ### Add new vCenter Server record to a TXT-file: ####
    $vCenterInstance | Out-File "vcenter_server_list.txt"

    Write-Host ("-------------------------------------------------------")

    $vmToSearch = Read-Host ("Enter VM Name to find ")

    ### Get a vDatacenters from vCenter Server:
    $getDataCenters = Get-Datacenter

    ### Get a vClusters from vCenter Server:
    $getClusters = Get-Cluster

    ### Get the DateTime parameter's value before Start
    ### next blocks of code:
    $taskStartTime = Get-Date -Format yyyy-MM-dd-HHmm
    
    ####### Initialization of the Delimiter function ###################
    function DelimiterPrint 
    {
      param ($strValue)
      
      $tmpDelValue = ""

      for ($i = 1; $i -le $strValue.Length; $i++)
      {
        $tmpDelValue = $tmpDelValue + "-"
      }

      Write-Host $tmpDelValue
    }
    ####################################################################
    
    foreach ($dc in $getDataCenters)
      {
        Write-Host ("")
        
        $header01 = "Datacenter : " + $dc
        #$delimiter01 = ""

        Write-Host ($header01)
        
        # Call the Delimiter Print function:
        DelimiterPrint -strValue $header01

        <#
        for ($i = 1; $i -le $header01.Length; $i++)
        {
          $delimiter01 = $delimiter01 + "-"
        }

        Write-Host ($delimiter01)
        #>
        $vms = Get-VM -Location $dc | Get-View
        
        foreach ($vm in $vms) 
        {
          if ($vm.Name -match $vmToSearch)
          {
            $coresPerSocket = $vm.Config.Hardware.NumCoresPerSocket
            $vcpuSockets = $vm.Config.Hardware.NumCPU / $coresPerSocket
            $memInGB = $vm.Config.Hardware.MemoryMB / 1024

            Write-Host ("vCPU Sockets in " + $vm.Name + "          : " + $vcpuSockets)
            Write-Host ("Cores per vCPU Socket in " + $vm.Name + " : " +$coresPerSocket)
            Write-Host ("Total vCPUs count : " + $coresPerSocket * $vcpuSockets)
            Write-Host ("Memory size in " + $vm.Name + "           : " + $memInGB)

            foreach ($dev in $vm.Config.Hardware.Device)
            {
              If (($dev.gettype()).Name -eq "VirtualDisk")
              {
                Write-Host ("===========================================")
                Write-Host ($dev.DeviceInfo.Label + " type is " + "Standard VMFS-based disk.")
                Write-Host ("===========================================")

                #Write-Host ("Datastore Name: " + ($dev.Backing.Datastore).Value)
                Write-Host ("Datastore Name: " + ($dev.Backing.FileName).Split("]")[0].TrimStart("["))
                Write-Host ("Path to VMDK: " + $dev.Backing.FileName)
                Write-Host ("Disk volume size (GB): " + ($dev.CapacityInKB)/1048576)
              }
            }

          }
        }

      }

    # Get the DateTime parameter's value after Stop
    # that block of code:
    $taskStopTime = Get-Date -Format yyyy-MM-dd-HHmm

    Write-Host ("-------------------------------------------------------")
    
    Write-Host ("Started at :" + $taskStartTime)
    Write-Host ("Stopped at :" + $taskStopTime)

    #Write-Host ($getDataCenter)

    # Disconnect with vCenter Server:
    Disconnect-VIServer $vCenterInstance -Confirm:$false
  }

 
