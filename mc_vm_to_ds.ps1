<############################################################
 POWERCLI :     MC (Mission Critical) VM to Datastore Mapping
 ============================================================
 Version                              :           Release 1.0
 Author                               :        Yuri P. Bodrov         
 Company                              :   PLC Alfastrahovanie             
 E-mail                               : bodrovyp@alfastrah.ru   
 Phone №                              :          +79259929596
 Telegram_ID                          :           @YuriBodrov           
############################################################>

Clear-Host # Clear the Console Output before all Operations

# GLOBAL VARIABLES. START ##########################################################################################

$getStartDateTimeValue = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff")   # Script's Start Datetime
[string]$vcAddress     = "vmc-01-new.vesta.ru"                            # vCSA Address to Connect with
$dsHashTable           = @{}                                              # Initialize an Empty Datastores HashTable
$mcVmNameList          = [System.Collections.Generic.List[string]]::New() # Initialize MC VM Names List

# Initialize the New DataSet for 'Datastore <> Disk <> VM' Mapping +
$vmToDsDataSet = New-Object System.Data.DataSet("vmToDsDataSet")

# Create a New DataTable to place the Data inside the DataSet 'vmToDsDataSet' +
$dataSetTable = New-Object System.Data.DataTable("vmToDsDataTable")
$dataSetTable.Columns.Add("vmname", [string]) | Out-Null
$dataSetTable.Columns.Add("diskid", [string]) | Out-Null
$dataSetTable.Columns.Add("vmdkpath", [string]) | Out-Null
$dataSetTable.Columns.Add("datastore", [string]) | Out-Null
$dataSetTable.Columns.Add("capacity_GB", [double]) | Out-Null

# Initialize the New DataSet for 'Tag <> Datastore <> VM Disk <> VM Name' +
$resultDataSet = New-Object System.Data.DataSet("resultDataSet")

# Create a New DataTable to place the Data inside the DataSet 'resultDataSet' +
$resultDsTable = New-Object System.Data.DataTable("resultDsTable")
$resultDsTable.Columns.Add("vmname", [string]) | Out-Null
$resultDsTable.Columns.Add("diskid", [string]) | Out-Null
$resultDsTable.Columns.Add("vmdkpath", [string]) | Out-Null
$resultDsTable.Columns.Add("datastore", [string]) | Out-Null
$resultDsTable.Columns.Add("capacityGB", [double]) | Out-Null
$resultDsTable.Columns.Add("dstagvalue", [string]) | Out-Null

########################################################################################### GLOBAL VARIABLES. STOP #
Function FillTheDsAndTagsHashTable ### Function 'FillTheDsAndTagsHashTable' Starts Here...
{
  $tagCategoryToFilterDs = "datastoreClass" # Declare Tag Category to Optimize 
                                            # the Get-View Datastore Query

  # Filter and Count More Quickly with Get-View Output.
  $dsView = Get-View -ViewType Datastore -Property Name | Sort-Object -Property Name
  
  [int]$dsViewCount  = $dsView.Count # Progress Bar Maximum Value

  Write-Host ("Filling the Datastores HashTable...")
  
  For ($i = 0; $i -lt $dsViewCount; $i++)
  {
    # Calculate the Progress Bar Percentage
    $percent = (($i + 1) / $dsViewCount) * 100

    $viObject          = Get-VIObjectByVIView -VIView $dsView[$i] # This cmdlet converts a 
                                                                  # vSphere View object to a VIObject.
    # Get Tag Category and its Value of each Datastore
    $tagAssignCategory = Get-TagAssignment -Entity $viObject -Category $tagCategoryToFilterDs
    $tagNameVal        = $tagAssignCategory.Tag.Name
    
    # Display the Progress Bar
    Write-Progress -Activity "Processing Datastores" -Status "Item '$($viObject.Name)' with ID : $($i + 1) of $dsViewCount" `
    -PercentComplete $percent -CurrentOperation "Analyzing data..."
    
    # Delay of the Progress Bar Status Changing  
    Start-Sleep -Milliseconds 100
    
    # Strictly Checks If Tag Value is Not Null/Empty or WhiteSpace and 
    # If Datastore's Item Does Not Already Exist in HashTable
    If ( (![string]::IsNullOrWhiteSpace($tagNameVal)) -and (!$dsHashTable.ContainsKey($viObject.Name)) )
    {

      $dsHashTable.Add($viObject.Name, $tagNameVal) # Add Key/Value Pair to HashTable!
      $dsCount = $dsCount + 1                       # Increase Datastores Count by 1

    }

  }

  Write-Host ("Done.")
  Write-Host ("---")
  Write-Host ("Datastores Count : " + $dsCount)
  Write-Host ("")
  Start-Sleep -Milliseconds 2000
  
} ### ...Function 'FillTheDsAndTagsHashTable' Ends Here.

Function FillTheVmDisksDataSet ### Function 'FillTheVmDisksDataSet' Starts Here...
{
  Write-Host ("Getting All MC VMs and their Disks...") -NoNewline
  
  # Filter and Count More Quickly with Get-View Output. Get only MC VMs
  $bcCustomAttrName = Get-CustomAttribute -Name "Business Category"
  $vmView           = Get-View -ViewType VirtualMachine -Property Name, CustomValue, Config.Hardware.Device | `
  Where-Object { $_.CustomValue.Key -eq $bcCustomAttrName.Key -and $_.CustomValue.Value -eq "MC" } | Sort-Object -Property Name
  
  [int]$mcVMsCount = $vmView.Count # Progress Bar Maximum Value
  
  For ($j = 0; $j -lt $mcVMsCount; $j++)
  {
    # Calculate the Percentage
    $percent = (($j + 1) / $mcVMsCount) * 100

    $vmName  = $vmView[$j]
    $mcVmNameList.Add($vmName.Name)
    
    # Isolate MC VM's Hard Disks from the Hardware Device List
    $vmNameDisks = $vmName.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualDisk] }

    ForEach ($vmNameDisk in $vmNameDisks)
    {
      $diskLabel     = $vmNameDisk.DeviceInfo.Label                              # Disk Serial Number ('Hard Disk [N]')
      $datastoreName = $vmNameDisk.Backing.FileName.Split(']')[0].TrimStart('[') # VMFS Datastore Name
      $vmdkPath      = $vmNameDisk.Backing.FileName                              # Path to VMDK File
      $capacityInGB  = [Math]::Round(($vmNameDisk.CapacityInKB/1024)/1024, 2)    # VMDK Capacity in GB

      # Add Data Rows into the DataTable
      $dataSetTableRow                = $dataSetTable.NewRow()
      $dataSetTableRow["vmname"]      = $vmName.Name   # DataTable Row Index : 0
      $dataSetTableRow["diskid"]      = $diskLabel     # DataTable Row Index : 1
      $dataSetTableRow["vmdkpath"]    = $vmdkPath      # DataTable Row Index : 2
      $dataSetTableRow["datastore"]   = $datastoreName # DataTable Row Index : 3
      $dataSetTableRow["capacity_GB"] = $capacityInGB  # DataTable Row Index : 4
      $dataSetTable.Rows.Add($dataSetTableRow)
    }

    # Display the Progress Bar
    Write-Progress -Activity "Processing MC VMs" -Status "Item '$($vmName.Name)' $($j + 1) of $mcVMsCount" `
    -PercentComplete $percent -CurrentOperation "Analyzing data..."
    
    # Delay of the Progress Bar Status Changing  
    Start-Sleep -Milliseconds 100

  }

  # Bind the DataTable to DataSet
  $vmToDsDataSet.Tables.Add($dataSetTable)

  Write-Host ("Done.")
  Write-Host ("---")
  Write-Host ("MC VMs. Total   : " + $mcVMsCount)
  Write-Host ("VM Disks. Total : " + $vmToDsDataSet.Tables[0].Rows.Count)
  Write-Host ("")
  Start-Sleep -Milliseconds 2000

} ### ...Function 'FillTheVmDisksDataSet' Ends Here.

Function FinalizeMcVmStat ### Function 'FinalizeMcVmStat' Starts Here...
{
  
  ForEach ($htItem in $dsHashTable.GetEnumerator())
  {
    For ($i = 0; $i -lt $vmToDsDataSet.Tables[0].Rows.Count; $i++)
    {
      If ($htItem.Key -eq $vmToDsDataSet.Tables[0].Rows[$i][3])
      {
        Write-Host ("vmname     : " + $vmToDsDataSet.Tables[0].Rows[$i][0])
        Write-Host ("diskid     : " + $vmToDsDataSet.Tables[0].Rows[$i][1])
        Write-Host ("vmdkpath   : " + $vmToDsDataSet.Tables[0].Rows[$i][2])
        Write-Host ("datastore  : " + $vmToDsDataSet.Tables[0].Rows[$i][3])
        Write-Host ("capacityGB : " + $vmToDsDataSet.Tables[0].Rows[$i][4])
        Write-Host ("dstagvalue : " + $htItem.Value)
        Write-Host ("---")

        # Add Data Rows into the DataTable 'resultDsTable'
        $resultDsTableRow               = $resultDsTable.NewRow()
        $resultDsTableRow["vmname"]     = $vmToDsDataSet.Tables[0].Rows[$i][0] # DataTable Row Index : 0
        $resultDsTableRow["diskid"]     = $vmToDsDataSet.Tables[0].Rows[$i][1] # DataTable Row Index : 1
        $resultDsTableRow["vmdkpath"]   = $vmToDsDataSet.Tables[0].Rows[$i][2] # DataTable Row Index : 2
        $resultDsTableRow["datastore"]  = $vmToDsDataSet.Tables[0].Rows[$i][3] # DataTable Row Index : 3
        $resultDsTableRow["capacityGB"] = $vmToDsDataSet.Tables[0].Rows[$i][4] # DataTable Row Index : 4
        $resultDsTableRow["dstagvalue"] = $htItem.Value                        # DataTable Row Index : 5
        $resultDsTable.Rows.Add($resultDsTableRow)
      }
    }
  }

  # Bind the DataTable to DataSet
  $resultDataSet.Tables.Add($resultDsTable)
  
  # Export 'resultDataSet' DataTable to MS Excel File
  # $PSScriptRoot - Directory where this Script is located
  $resultDsTable | Select-Object * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | Sort-Object vmname -Descending | `
  Export-Excel -Path "$PSScriptRoot\mc_vms_to_ds_tmp.xlsx" -WorksheetName "mc_vms_to_ds" -AutoSize -TableStyle Medium11

} ### ...Function 'FinalizeMcVmStat' Ends Here.

Try
{
  $vcsaAddress = "vmc-01-new.vesta.ru"
  $credentialsForAll = Get-Credential -UserName "administrator@vsphere.local"
  Connect-VIServer -Server $vcsaAddress -Credential $credentialsForAll

  If ($global:DefaultVIServers.Count -gt 0)
  {
    Write-Host ("")
    Write-Host ("Status : Connection successful. Server : " + $vcAddress)    
    Write-Host ("")

    FillTheDsAndTagsHashTable # Function's Call
    FillTheVmDisksDataSet     # Function's Call
    FinalizeMcVmStat          # Function's Call
  }

  Disconnect-VIServer vmc-01-new.vesta.ru -Confirm:$false
  Start-Sleep -Milliseconds 2000
}

Catch
{
  Write-Host ("Status : No connection with vCSA or Another Error! Check this out!")
  Write-Host ("---")
  Write-Host ($_.Exception.Message)
  Write-Host ("")
  Disconnect-VIServer vmc-01-new.vesta.ru -Confirm:$false
  Exit # Exit the Program
}

$getStopDateTimeValue = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff") # Script's Stop Datetime

Write-Host ("")
Write-Host ("---")
Write-Host ("Start Time : " + $getStartDateTimeValue)
Write-Host ("Stop Time  : " + $getStopDateTimeValue)
Write-Host ("")
Start-Sleep -Milliseconds 2000

Exit
