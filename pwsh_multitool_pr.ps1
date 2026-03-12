<##############################################################
 POWERSHELL MULTITOOL : CMDB SYNCHRONIZER + VM/CLUSTER MAPPER
 ITIL_CMDB Attributes -> vCenter Server Synchronization Program
 ==============================================================
 Version                              :       Pre-Release 1.1
 Author                               :        Yuri P. Bodrov         
 Company                              :   PLC Alfastrahovanie             
 E-mail                               : bodrovyp@alfastrah.ru   
 Phone №                              :          +79259929596
 Telegram_ID                          :           @YuriBodrov           
###############################################################>

<# TODO! 
0. USEFUL! You must specify a Full Path to the Log Files for TaskScheduler Correct Work! +
1. Log File Size <= 10MB! If >= 10MB -> ZIP Archive?
2. Log like Everybody is Watching! +
3. Modify 'CreateLogFile_Func' Function! +
#>

<# FIGURE IT OUT 
(Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.PrefixOrigin -ne 'WellKnown' -and $_.IPAddress -notlike "127.0.0.1" }).IPAddress
Get-NetIPConfiguration
#>

### PRIMARY SCRIPT'S ARGUMENTS AND ITS VALUES : ############
param ([Parameter(Mandatory=$true, Position=0)][string]$Arg)

### SCRIPT'S GLOBAL VARIABLES : BLOCK STARTS ########################################################

# String Variables 
[string]$logFilePath = "$PSScriptRoot\cmdb-sync-log-file.log"
[string]$logFilePathCsv = "$PSScriptRoot\cmdb-sync-log-file.csv"
[string]$vcAddress = "vmc-01-new.vesta.ru"
[string]$cmdbAddress = "10.213.20.29"
[string]$smtpServer = "relay.alfastrah.ru"
[string]$emailSender = "vmware_automation@alfastrah.ru"
#[string]$emailAddr_01 = "oksenuk@alfastrah.ru"
#[string]$emailAddr_02 = "bodrovyp@alfastrah.ru"
[string]$pwshServerIpAddr = (Get-NetIPAddress -InterfaceIndex 8).IPAddress
[string]$pwshServerFqdn = (Resolve-DnsName -Name $pwshServerIpAddr).NameHost

# Array Lists 
$emailAddr_List = "bodrovyp@alfastrah.ru", "oksenuk@alfastrah.ru", "SolovevSV1@alfastrah.ru"
$vmNamesArray = [System.Collections.ArrayList]::new()    # Gets 'VM Name' Values from JSON!
$itServicesArray = [System.Collections.ArrayList]::new() # Gets 'IT Service' Values from JSON!
$envTypeArray = [System.Collections.ArrayList]::new()    # Gets 'EnvironmentType' Values from JSON!
$bsnsscatArray = [System.Collections.ArrayList]::new()   # Gets 'Business Category' Values from JSON!
$ownrArray = [System.Collections.ArrayList]::new()       # Gets 'Owner' Values from JSON!
$spprtArray = [System.Collections.ArrayList]::new()      # Gets 'Support' Values from JSON!
$hststtArray = [System.Collections.ArrayList]::new()     # Gets 'HostState' Values from JSON!

# Credentials from PowerCLI Credentials Store 
$viCreds = Get-VICredentialStoreItem -Host $vcAddress

######################################################### SCRIPT'S GLOBAL VARIABLES : BLOCK STOPS ###

### FUNCTION : CreateLogFileFunc. Declaring. START ########################################################
### Note! This Function creates a New Log File in the Current Script's Directory and verifies its existing.
### Note! After creation It appends a 1st String into the Log File : Log File Creation Date. 
Function CreateLogFile_Func 
{   
  If (Test-Path $logFilePath) # If The Log File already Exists... 
  {
    Write-Host "The Log File '$logFilePath' already Exists."
    Write-Host "- - -"
  } 
  Else # If Does Not Exist... 
  {
    Write-Host "The file '$logFilePath' does not exist."
    Write-Host ""
    Write-Host "Let's create it..."
    Start-Sleep -Milliseconds 1000

    # Log File Creation Process : -----------------------------------
    New-Item -Path $logFilePath -ItemType File
    $logFileCreationDate = Get-Date -Format "dd-MM-yyyy HH:mm:ss.fff"
    # ---------------------------------------------------------------

    If (Test-Path $logFilePath) # If Log File is Successfully Created...
    {
      # Add 1st Record to a Log File -------------------------------------------------------
      "This Log File Created at : " + $logFileCreationDate | Out-File -FilePath $logFilePath
      "- - -" | Out-File -FilePath $logFilePath -Append -NoClobber
      # ------------------------------------------------------------------------------------
    }

    Write-Host "Exiting..."
    Start-Sleep -Milliseconds 1000
    Exit # Is it neccessary? TODO! Check this out!
  }

  Return
}
######################################################### FUNCTION : CreateLogFileFunc. Declaring. STOP ###

### FUNCTION : AppendToLogFileFunc. Declaring. START ###########################################################################################
### Note! This Function appends the Transaction Records to a Different Log File Formats : .LOG/.CSV
Function AppendToLogFile_Func
{
  param (
    [Parameter(Mandatory)]$tr_time,     # Transaction/Log Record Creation Datetime!
    [Parameter(Mandatory)]$object_name, # VI Object Name!
    [Parameter(Mandatory)]$tr_type,     # Transaction/Log Record Type : Information/ValueChanging/Warning/Error!
    [Parameter(Mandatory)]$tr_rslt,     # Transaction Metadata or Exception Messages!
    $cattr_key                          # Custom Attribute's Name
  )

  # Append to Log File. Logic ----------------------------------------------------------------------------------------------------------------
  # global_variable : $logFilePath = "$PSScriptRoot\cmdb-sync-log-file.log"
  # global_variable : $logFilePathCsv = "$PSScriptRoot\cmdb-sync-log-file.csv"
  If ($tr_type -eq "Error") 
  {
    # To Native .LOG File
    "TransactionTime    : " + $tr_time | Out-File -FilePath $logFilePath -Append -NoClobber
    "ObjectName         : " + $object_name | Out-File -FilePath $logFilePath -Append -NoClobber
    "TransactionType    : " + $tr_type | Out-File -FilePath $logFilePath -Append -NoClobber
    "TransactionResult  : " + $tr_rslt | Out-File -FilePath $logFilePath -Append -NoClobber
    "---" | Out-File -FilePath $logFilePath -Append -NoClobber

    # To .CSV Log File
    $tr_time + "|" + $object_name + "|" + $tr_type + "|" + $tr_rslt | Out-File -FilePath $logFilePathCsv -Append -NoClobber
  }
  
  Else 
  {
    # To Native .LOG File
    "TransactionTime    : " + $tr_time | Out-File -FilePath $logFilePath -Append -NoClobber
    "ObjectName         : " + $object_name | Out-File -FilePath $logFilePath -Append -NoClobber
    "TransactionType    : " + $tr_type | Out-File -FilePath $logFilePath -Append -NoClobber
    "CustomAttributeKey : " + $cattr_key | Out-File -FilePath $logFilePath -Append -NoClobber
    "TransactionResult  : " + $tr_rslt | Out-File -FilePath $logFilePath -Append -NoClobber
    "---" | Out-File -FilePath $logFilePath -Append -NoClobber

    # To .CSV Log File
    $tr_time + "|" + $object_name + "|" + $tr_type + "|" + $cattr_key + "|" + $tr_rslt | Out-File -FilePath $logFilePathCsv -Append -NoClobber
  }
  # ------------------------------------------------------------------------------------------------------------------------------------------
  Return
}
############################################################################################## FUNCTION : AppendLogFileFunc. Declaring. STOP ###

### FUNCTION : EmailSender_Func. Declaring. START ##################################################################################################################
### Note! This Function sends Success, Statistics or Error Messages through Email Sender Subsystem.
### Note! It gets Parameter as Email Message Text or HTML Page Formats.
Function EmailSender_Func
{
  # Use these Global Vars
  # [string]$smtpServer = "relay.alfastrah.ru"
  # [string]$emailSender = "vmware_automation@alfastrah.ru"
  # [string]$emailAddr_01 = "oksenuk@alfastrah.ru"
  # [string]$emailAddr_02 = "bodrovyp@alfastrah.ru"
  
  param (
    [Parameter(Mandatory)][string]$emailMessageContent,
    [Parameter(Mandatory)][string]$emailMessageSubject
  )
  
  <#
  $emailMessageContent = "Hello, Yuri!" + [Environment]::NewLine + "This is a Test Message. Do not reply on it. Have a Nice Day! :-)" + 
  #[Environment]::NewLine + "`n" + "---" + [Environment]::NewLine + "Powershell Multitool. osis@alfastrah.ru"
  [Environment]::NewLine + [Environment]::NewLine + "---" + [Environment]::NewLine + "Powershell CMDB Sync MultiTool" + 
  [Environment]::NewLine + "Email to : osis@alfastrah.ru"
  #>
  
  Send-MailMessage -To $emailAddr_List -From $emailSender -Subject $emailMessageSubject -Body $emailMessageContent -SmtpServer $smtpServer
  
  Return
}
#################################################################################################################### FUNCTION : EmailSenderFunc. Declaring. STOP ###

### FUNCTION : GetAllJsonInfo_Func. Declaring. START #########################################################################
### Note! Interacts with a downloaded CMDB JSON File.
Function GetAllJsonInfo_Func
{   
  Try
  {
    $json_cmdb_doc_url = "http://" + $cmdbAddress + "/as_itil_cmdb_conf/hs/api/getListVM" # TODO! This Function Must Run with Try..Catch..Finally! 
    $cmdbCreds = Get-VICredentialStoreItem -Host 1C
    $cmdbAuth = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($cmdbCreds.User+':'+$cmdbCreds.Password))
    $cmdbAuthHeaders = @{Authorization = "Basic $cmdbAuth"}
    $spec_symbols = "?","<>","null","<?>"

    $cmdbObjects = ConvertFrom-Json (Invoke-WebRequest -Uri $json_cmdb_doc_url -Method Get -Headers $cmdbAuthHeaders).Content
    
    If ($cmdbObjects) 
    {
      ForEach ($testCmdbObj in $cmdbObjects) # Get All Required Values of JSON's Notation Keys
      {
        [String]$tmp_tr_rslt_name = $testCmdbObj.Name
        [String]$tmp_tr_rslt_clnt = $testCmdbObj.Client
        [String]$tmp_tr_rslt_ownr = $testCmdbObj.Owner
        [String]$tmp_tr_rslt_spprt = $testCmdbObj.Support
        [String]$tmp_tr_rslt_itsrvc = $testCmdbObj.ITService.ITservice
        [String]$tmp_tr_rslt_clsctr = $testCmdbObj.Classificator
        [String]$tmp_tr_rslt_hststt = $testCmdbObj.HostState
        [String]$tmp_tr_rslt_pwrstt = $testCmdbObj.PowerState
        [String]$tmp_tr_rslt_envtype = $testCmdbObj.EnvironmentType
        [String]$tmp_tr_rslt_bsnsscat = $testCmdbObj.BusinessCategory
          
        If ($tmp_tr_rslt_clsctr -match "vsphere") 
        {
          # Append 'VM Name' Value into the Special String Array List
          $vmNamesArray.Add($tmp_tr_rslt_name) | Out-Null
                  
          # Append 'IT Service' Value into the Special String Array List
          $itServicesArray.Add($tmp_tr_rslt_itsrvc) | Out-Null
                  
          # Append 'HostState' Value into the Special String Array List
          $hststtArray.Add($tmp_tr_rslt_hststt) | Out-Null

          # Append 'Environment Type' Value into the Special String Array List
          $envTypeArray.Add($tmp_tr_rslt_envtype) | Out-Null

          # Append 'Owner' Value into the Special String Array List
          $ownrArray.Add($tmp_tr_rslt_ownr) | Out-Null

          # Append 'Support' Value into the Special String Array List
          $spprtArray.Add($tmp_tr_rslt_spprt) | Out-Null

          # Append 'Business Category' Value into the Special String Array List
          $bsnsscatArray.Add($tmp_tr_rslt_bsnsscat) | Out-Null
        }
      }
    }
  }
     
  Catch 
  {
    $tmp_err_time = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff")
    $tmp_err_obj_name = [Environment]::MachineName
    $tmp_err_type = "Error"
    $tmp_err_desc_excep = $_.Exception.Message
    $tmp_err_desc = "Unable to get JSON File from " + $json_cmdb_doc_url + ". " + $tmp_err_desc_excep
    AppendToLogFile_Func -tr_time $tmp_err_time -object_name $tmp_err_obj_name -tr_type $tmp_err_type -tr_rslt $tmp_err_desc

    # Send an Email Error Message
    $subjectText = "ERROR! Powershell MultiTool Software. Unable to Get CMDB JSON File!"
    $messageText = "ERROR!" + [Environment]::NewLine + $tmp_err_desc + [Environment]::NewLine + [Environment]::NewLine + "---" + 
    [Environment]::NewLine + "Powershell MultiTool Software 1.0" + [Environment]::NewLine + "Server Address : " + $pwshServerFqdn + `
    "[" + $pwshServerIpAddr + "]." + [Environment]::NewLine + "Log File Path : " + $logFilePath + "." + [Environment]::NewLine + "Email to : osis@alfastrah.ru"
    EmailSender_Func -emailMessageSubject $subjectText -emailMessageContent $messageText

    Start-Sleep -Milliseconds 2000
    Exit
  }

  Return
  
}
############################################################################# FUNCTION MainOperationsFunc. Declaring. STOP ###

### FUNCTION : AppendJsonInfoToVc_Func. Declaring. START ###########################################################################################################
### Note! Interacts with a downloaded CMDB JSON File and with Pwsh-Script Function 'AppendToLogFileFunc' 01.10.2025 -> TODO! Change this Description ASAP!
Function AppendJsonInfoToVc_Func
{
  Try
  {
    Connect-VIServer $vcAddress -User $viCreds.User -Password $viCreds.Password -ErrorAction Stop | Out-Null
        
    If ($global:DefaultVIServers.Count -gt 0)
    {
      # VM SELECTION METHOD'S NEW LOGIC. BLOCK STARTS : ----------------------------------------------------------------------------------------------------------------------------------------
      ForEach ($datacenter in Get-Datacenter) 
      {
        ForEach ($cluster in (Get-Cluster -Location $datacenter)) 
        {
          ForEach ($vmhost in (Get-VMHost -Location $cluster))
          {
            $VMsOnHost = Get-VM -Location $vmhost
            ForEach ($VM in $VMsOnHost) 
            {
              $vm_name_to_check = Get-VM -Name $VM | Select-Object Name
              $it_srvc_from_vc = $VM | Get-Annotation -CustomAttribute "ITService"                                     # Gets 'IT Service' VM's Custom Attribute Value from vCenter Server
              $ownr_from_vc = $VM | Get-Annotation -CustomAttribute "Owner"                                            # Gets 'Owner' VM's Custom Attribute Value from vCenter Server
              $spprt_from_vc = $VM | Get-Annotation -CustomAttribute "Support"                                         # Gets 'Support' VM's Custom Attribute Value from vCenter Server
              $bsnsscat_from_vc = $VM | Get-Annotation -CustomAttribute "Business Category"                            # Gets 'Business Category' VM's Custom Attribute Value from vCenter Server
              $envType_from_vc = $VM | Get-Annotation -CustomAttribute "EnvironmentType" -ErrorAction SilentlyContinue # Gets 'EnvironmentType' VM's Custom Attribute Value from vCenter Server
                            
              If ($vm_name_to_check.Name -in $vmNamesArray)
              {
                $itemIndexToFind = $vm_name_to_check.Name
                $indexToFind = [System.Array]::IndexOf($vmNamesArray, $itemIndexToFind)
                                
                ### Comparing JSON and vCenter Server Values. Custom Attribute : 'ITService' #####################################################################
                If ($itServicesArray[$indexToFind] -ne $it_srvc_from_vc.Value) # Changes are Present!
                {
                  Set-Annotation -Entity $VM -CustomAttribute "ITService" -Value $itServicesArray[$indexToFind] | Out-Null
                                    
                  # Append this Transaction Record to the Log File
                  $tmp_tr_time = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff")
                  $tmp_object_name = $vm_name_to_check.Name
                  $tmp_tr_type = "ValueChanging"
                  $tmp_cattr_key = "ITService"
                  $tmp_tr_rslt = $it_srvc_from_vc.Value + " -> " + $itServicesArray[$indexToFind]
                  AppendToLogFile_Func -tr_time $tmp_tr_time -object_name $tmp_object_name -tr_type $tmp_tr_type -cattr_key $tmp_cattr_key -tr_rslt $tmp_tr_rslt
                }
                                
                ### Comparing JSON and vCenter Server Values. Custom Attribute : 'Owner'##########################################################################
                If ($ownrArray[$indexToFind] -ne $ownr_from_vc.Value)
                {
                  Set-Annotation -Entity $VM -CustomAttribute "Owner" -Value $ownrArray[$indexToFind] | Out-Null

                  # Append this Transaction Record to the Log File:
                  $tmp_tr_time = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff")
                  $tmp_object_name = $vm_name_to_check.Name
                  $tmp_tr_type = "ValueChanging"
                  $tmp_cattr_key = "Owner"
                  $tmp_tr_rslt = $ownr_from_vc.Value + " -> " + $ownrArray[$indexToFind]
                  AppendToLogFile_Func -tr_time $tmp_tr_time -object_name $tmp_object_name -tr_type $tmp_tr_type -cattr_key $tmp_cattr_key -tr_rslt $tmp_tr_rslt
                }

                ### Comparing JSON and vCenter Server Values. Custom Attribute : 'Support' #######################################################################
                If ($spprtArray[$indexToFind] -ne $spprt_from_vc.Value)
                {
                  Set-Annotation -Entity $VM -CustomAttribute "Support" -Value $spprtArray[$indexToFind] | Out-Null

                  # Append this Transaction Record to the Log File:
                  $tmp_tr_time = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff")
                  $tmp_object_name = $vm_name_to_check.Name
                  $tmp_tr_type = "ValueChanging"
                  $tmp_cattr_key = "Support"
                  $tmp_tr_rslt = $spprt_from_vc.Value + " -> " + $spprtArray[$indexToFind]
                  AppendToLogFile_Func -tr_time $tmp_tr_time -object_name $tmp_object_name -tr_type $tmp_tr_type -cattr_key $tmp_cattr_key -tr_rslt $tmp_tr_rslt
                }

                ### Comparing JSON and vCenter Server Values. Custom Attribute : 'EnvironmentType' ###############################################################
                If ($envTypeArray[$indexToFind] -ne $envType_from_vc.Value)
                {
                  Set-Annotation -Entity $VM -CustomAttribute "EnvironmentType" -Value $envTypeArray[$indexToFind] | Out-Null

                  # Append this Transaction Record to the Log File:
                  $tmp_tr_time = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff")
                  $tmp_object_name = $vm_name_to_check.Name
                  $tmp_tr_type = "ValueChanging"
                  $tmp_cattr_key = "EnvironmentType"
                  $tmp_tr_rslt = $envType_from_vc.Value + " -> " + $envTypeArray[$indexToFind]
                  AppendToLogFile_Func -tr_time $tmp_tr_time -object_name $tmp_object_name -tr_type $tmp_tr_type -cattr_key $tmp_cattr_key -tr_rslt $tmp_tr_rslt
                }

                ### Comparing JSON and vCenter Server Values. Custom Attribute : 'Business Category' #############################################################
                If ($bsnsscatArray[$indexToFind] -ne $bsnsscat_from_vc.Value)
                {
                  Set-Annotation -Entity $VM -CustomAttribute "Business Category" -Value $bsnsscatArray[$indexToFind] | Out-Null

                  # Append this Transaction Record to the Log File:
                  $tmp_tr_time = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff")
                  $tmp_object_name = $vm_name_to_check.Name
                  $tmp_tr_type = "ValueChanging"
                  $tmp_cattr_key = "Business Category"
                  $tmp_tr_rslt = $bsnsscat_from_vc.Value + " -> " + $bsnsscatArray[$indexToFind]
                  AppendToLogFile_Func -tr_time $tmp_tr_time -object_name $tmp_object_name -tr_type $tmp_tr_type -cattr_key $tmp_cattr_key -tr_rslt $tmp_tr_rslt
                }
              }
                            
              Else # If VM Object Exists in vCenter Server Inventory and missing from CMDB
              {
                # Append this Transaction Record to the Log File:
                $tmp_tr_time = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff")
                $tmp_object_name = $vm_name_to_check.Name
                $tmp_tr_type = "Warning"
                $tmp_cattr_key = "n/a"
                $tmp_tr_rslt = "This VM Item is missing from CMDB Data!"
                AppendToLogFile_Func -tr_time $tmp_tr_time -object_name $tmp_object_name -tr_type $tmp_tr_type -cattr_key $tmp_cattr_key -tr_rslt $tmp_tr_rslt
              }

            }
          }
        }
      }
      # ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
    }
  }
    
  Catch
  {
    $tmp_err_time = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff")
    $tmp_err_obj_name = [Environment]::MachineName
    $tmp_err_type = "Error"
    $tmp_err_desc = $_.Exception.Message
    AppendToLogFile_Func -tr_time $tmp_err_time -object_name $tmp_err_obj_name -tr_type $tmp_err_type -tr_rslt $tmp_err_desc

    # Send an Email Error Message
    $subjectText = "ERROR! Powershell MultiTool Software. Unable to Connect to vCenter Server " + $vcAddress +"!"
    $messageText = "ERROR!" + [Environment]::NewLine + $tmp_err_desc + [Environment]::NewLine + [Environment]::NewLine + "---" + 
    [Environment]::NewLine + "Powershell MultiTool Software 1.0" + [Environment]::NewLine + "Server Address : " + $pwshServerFqdn + `
    "[" + $pwshServerIpAddr + "]." + [Environment]::NewLine + "Log File Path : " + $logFilePath + "." + [Environment]::NewLine + "Email to : osis@alfastrah.ru"
    EmailSender_Func -emailMessageSubject $subjectText -emailMessageContent $messageText
    
    Disconnect-VIServer $vcAddress -Confirm:$false
    Exit
  }

  Finally
  {
    Disconnect-VIServer $vcAddress -Confirm:$false
  }

  Return
}
############################################################################################################## FUNCTION AppendJsonInfoToVc_Func. Declaring. STOP ###

$delimiter = "===================================================================================================="

If ($Arg -eq "sync") # CMDB Synchronization Starts Here
{
  Clear-Host
  Write-Host ("---")
  Write-Host ("You have just selected : " + $Arg)
    
  #EmailSender_Func
  #GetAllJsonInfo_Func
  
  $tmp_local_hostname = [Environment]::MachineName
    
  $getStartDateTimeValue = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff")
  AppendToLogFile_Func -tr_time $getStartDateTimeValue -object_name $tmp_local_hostname -tr_type "MULTITOOL_OPS_STARTED" -cattr_key "n/a" -tr_rslt $delimiter

  GetAllJsonInfo_Func
  AppendJsonInfoToVc_Func

  $getStopDateTimeValue = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff")
  AppendToLogFile_Func -tr_time $getStopDateTimeValue -object_name $tmp_local_hostname -tr_type "MULTITOOL_OPS_STOPPED" -cattr_key "n/a" -tr_rslt $delimiter

  "" | Out-File -FilePath $logFilePath -Append -NoClobber
  "" | Out-File -FilePath $logFilePathCsv -Append -NoClobber

  # Send an Email Information Message
  [string]$tmp_info_desc_start = "Operation_Start_Time : " + $getStartDateTimeValue + ". "
  [string]$tmp_info_desc_stop = "Operation_Stop_Time : " + $getStopDateTimeValue + "."
  $subjectText = "INFO. Powershell MultiTool Software. Operation Status."
  $messageText = "CMDB Data Synchronization successfully completed!" + [Environment]::NewLine + "---" + [Environment]::NewLine + `
  $tmp_info_desc_start + [Environment]::NewLine + $tmp_info_desc_stop + [Environment]::NewLine + [Environment]::NewLine + "---" + `
  [Environment]::NewLine + "Powershell MultiTool Software 1.0" + [Environment]::NewLine + "Server Address : " + $pwshServerFqdn + `
  "[" + $pwshServerIpAddr + "]." + [Environment]::NewLine + "Log File Path : " + $logFilePath + "." + [Environment]::NewLine + "Email to : osis@alfastrah.ru"
  EmailSender_Func -emailMessageSubject $subjectText -emailMessageContent $messageText

}

ElseIf ($Arg -eq "map") # VM <> ESXi Server Mapping Starts Here
{
  #Clear-Host
  #Write-Host ("---")
  #Write-Host ("You have just selected : " + $Arg)
}

Else
{
  #Clear-Host
  #Write-Host ("---")
  #Write-Host ("You must Enter one of these Arguments : Sync/Map")
  #Write-Host ("Goodbye!") -ForegroundColor Magenta
  Exit
}
