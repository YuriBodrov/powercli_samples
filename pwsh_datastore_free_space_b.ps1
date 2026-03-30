<############################################################
 POWERCLI : VMWARE DATASTORE FREE SPACE ANALYZER AND REPORTER
 ============================================================
 Version                              :              Beta 1.0
 Author                               :        Yuri P. Bodrov         
 Company                              :   PLC Alfastrahovanie             
 E-mail                               :  bodrovyp@hotmail.com   
 Phone №                              :          +79259929596
 Telegram_ID                          :           @YuriBodrov           
############################################################>

Clear-Host

$getStartDateTimeValue = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff") # Script's Start Datetime

# Global Variables 
[string]$vcAddress = "[vcsa_addr]"
[string]$smtpServer = "[smtp_server_addr]"
[string]$emailSender = "[email_sender_addr]"
$emailAddr_List = "[email_receip_addr_01]", "email_receip_addr_02", "email_receip_addr_03"
[string]$pwshServerIpAddr = (Get-NetIPAddress -InterfaceIndex 8).IPAddress
[string]$pwshServerFqdn = (Resolve-DnsName -Name $pwshServerIpAddr).NameHost

# Output Table Class Initialization 
$datastoreCapTable = New-Object system.Data.DataTable “Datastores_Capacity_Report”                       # XLSX 1st DataSheet's Name
$datastoreCapTable.Columns.Add((New-Object system.Data.DataColumn Datastore_Name, ([string])))           # DataSheet's Column's Header #01
$datastoreCapTable.Columns.Add((New-Object system.Data.DataColumn Datastore_TotalCap_GBs, ([double])))   # DataSheet's Column's Header #02
$datastoreCapTable.Columns.Add((New-Object system.Data.DataColumn Datastore_FreeCap_GBs, ([double])))    # DataSheet's Column's Header #03
$datastoreCapTable.Columns.Add((New-Object system.Data.DataColumn Datastore_Util_In_Prcnts, ([double]))) # DataSheet's Column's Header #04
$datastoreCapTable.Columns.Add((New-Object system.Data.DataColumn Datastore_Util_State, ([string])))     # DataSheet's Column's Header #05

Function EmailSender_Func
{
  # Uses these Global Vars:
  # [string]$smtpServer
  # [string]$emailSender
  # $emailAddr_List
  
  param (
    [Parameter(Mandatory)][string]$emailMessageContent,
    [Parameter(Mandatory)][string]$emailMessageSubject,
		[Parameter(Mandatory)]$emailMessageAttach
  )
  
  Send-MailMessage -To $emailAddr_List -From $emailSender -Subject $emailMessageSubject -Body $emailMessageContent `
	-Attachments $emailMessageAttach -SmtpServer $smtpServer
  
  Return
}

# Connection to vCSA. Credentials from PowerCLI Credentials Store 
$viCreds = Get-VICredentialStoreItem -Host $vcAddress
Connect-VIServer $vcAddress -User $viCreds.User -Password $viCreds.Password -ErrorAction Stop | Out-Null

# Get All Required Parameter Values from Datastore VI (Virtual Infrastructure) Objects
$datastoreList = Get-Datastore | Sort-Object Name # Sort this List by Datastore's Name
ForEach ($datastoreName in $datastoreList)
{
	$datastoreName_Cap = [math]::Round($datastoreName.CapacityGB ,2)
	$datastoreName_FreeCap = [math]::Round($datastoreName.FreeSpaceGB ,2)
	$datastoreName_Util_In_Prcnts = [math]::Round((($datastoreName.CapacityGB - $datastoreName.FreeSpaceGB) / $datastoreName.CapacityGB) * 100, 2)
	Write-Host ("Datastore_Name            : " + $datastoreName.Name)           # Datastore's Name
	Write-Host ("Datastore_Capacity_GB     : " + $datastoreName_Cap)            # Total Datastore's Capacity in GBs
	Write-Host ("Datastore_FreeCapacity_GB : " + $datastoreName_FreeCap)        # Free Datastore's Capacity in GBs
	Write-Host ("Utilization_In_%          : " + $datastoreName_Util_In_Prcnts) # Used Datastore's Capacity in %

	# Check : 
	# IF '80% <= Utilization_In_% < 90%' -> Warning 
	# IF 'Utilization_In_% >= 90%'       -> Error 
	[string]$datastoreName_Util_State = ""
	If (($datastoreName_Util_In_Prcnts -ge 80) -and ($datastoreName_Util_In_Prcnts -lt 90))
	{
		$datastoreName_Util_State = "Warning"
		Write-Host ("State                     : " + $datastoreName_Util_State)   # Datastore's Utilization State = Warning 
	}

	ElseIf ($datastoreName_Util_In_Prcnts -ge 90)
	{
		$datastoreName_Util_State = "Error"
		Write-Host ("State                     : " + $datastoreName_Util_State)   # Datastore's Utilization State = Error
	}

	Else
	{
		$datastoreName_Util_State = "Normal"
		Write-Host ("State                     : " + $datastoreName_Util_State)   # Datastore's Utilization State = Normal
	}

	$dataRow = $datastoreCapTable.NewRow()
	$dataRow.Datastore_Name = $datastoreName.Name
	$dataRow.Datastore_TotalCap_GBs = $datastoreName_Cap
	$dataRow.Datastore_FreeCap_GBs = $datastoreName_FreeCap
	$dataRow.Datastore_Util_In_Prcnts = $datastoreName_Util_In_Prcnts
	$dataRow.Datastore_Util_State = $datastoreName_Util_State

	Write-Host ("---")

	$datastoreCapTable.Rows.Add($dataRow)
}

$attXlsxFile = "[path_to_directory]_$(Get-Date -Format "dd-MM-yyyy_HH-mm-ss_fff").xlsx"
$datastoreCapTable | Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | sort Datastore_Util_In_Prcnts -Descending | `
Export-Excel $attXlsxFile -AutoSize -FreezeTopRow -TableStyle Medium11 -WorksheetName "Datastores_Capacity_Report"

Disconnect-VIServer $vcAddress -Confirm:$false

$getStopDateTimeValue = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss.fff") # Script's Stop Datetime

# Send an Email Information Message
[string]$tmp_info_desc_start = "Operation_Start_Time : " + $getStartDateTimeValue + "."
[string]$tmp_info_desc_stop = "Operation_Stop_Time : " + $getStopDateTimeValue + "."
$subjectText = "INFO. Powershell VMware Datastores Utilization Analyzer. Operation Status."
$messageText = "Datastores Utilization Analysis successfully completed." + [Environment]::NewLine + "See attachment..." + `
[Environment]::NewLine + "---" + [Environment]::NewLine + $tmp_info_desc_start + [Environment]::NewLine + $tmp_info_desc_stop + `
[Environment]::NewLine + [Environment]::NewLine + "---" + [Environment]::NewLine + "Powershell VMware Datastores Utilization Analyzer Beta 1.0." + `
[Environment]::NewLine + "Email to : [email_address]"
EmailSender_Func -emailMessageSubject $subjectText -emailMessageContent $messageText -emailMessageAttach $attXlsxFile


