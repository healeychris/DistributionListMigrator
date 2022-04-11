

### Change As Needed ###########################################################################################################
$ContactGroupOU             =       'MovedDistributionGroups'   # OU name for Contacts to be created, change as required
$DCServer                   =       'BGBSADC1001'               # DC ServerName
$OnPremExchange             =       'BGBSAXH1001'               # On Prem Exchange Server
$DefaultManagedBy           =       "MTS-HybridAdmin"           # Change as required, used for default owner if not set already
$AllowGroupswithNestedDL    =       $false                      # Allow DL's to be created if contaning nested groups
$DBName                     =       'DistributionListDatabase'  # Default Database used to store all values   
$CSVDataFile                =       '.\ADGroups.csv'            # CSV File used to store the groups to be processed 
$version                    =       '0.3'                       # Script Version
################################################################################################################################

Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client -Name AllowBasic -Value 1 


# FUNCTION - Request to continue the operation
function ShowMenu {
    
    Write-host 
    Write-Host "(0) - Load Single Group"
    Write-Host "(1) - Load CSV Of Groups to Process"
    Write-Host "(2) - Pre-Check Groups for Nested/Security Groups"
    Write-Host "(3) - Create PlaceHolder Groups in (365)"
    Write-Host "(4) - Create Contact / Hide Group (OnPremise)"
    Write-Host "(5) - Finalize Cloud groups to Original (365)"
    Write-Host
    Write-host "(10) - Connect to Onprem Exchange Server"
    Write-host "(11) - Connect to Exchange Online"
    Write-host "(12) - Close all Open Exchange Sessions"

    Write-host 


    $input = read-host " *** Please select a number to perform the operation *** "

     switch ($input) `
    {

    '0' {
        # Load Single Group
        CheckforPSliteDBModule 
        ImportPSliteDBModule
        DatabaseConnection
        CheckOnline
        SingleADGroup
        ShowMenu 
    
        }

    '1' {
        # Load CSV Of Groups to Process
        CheckforPSliteDBModule 
        ImportPSliteDBModule
        DatabaseConnection
        CheckOnline
        ImportCSVData
        ShowMenu

    }
    
    '2' {
        # Pre-Check Groups for Nested/Security Groups
        CheckDistributionGroupDataLoaded
        [int]$script:ProcessGroups = '0'
        Foreach ($DLGroup in $script:DistributionListsDetails){
            # Build String
            $DLPrimarySmtpAddress = $DLGroup.PrimarySmtpAddress
            CheckGroupForSecOrNested


        }
        ShowMenu

    }

    '3' {
        # Create PlaceHolder Groups in (365) - None Nested
        CheckDistributionGroupDataLoaded
        CheckOnline
        BatchInformation
        [int]$script:ProcessGroups = '0'
        Foreach ($DLGroup in $script:DistributionListsDetails){
            CreatePlaceHolder 

        }
        ShowMenu

	}

    '4' {
        # Create Contact / Hide Group (OnPremise)
        CheckOnPrem
        $script:CreatePlaceHolderMarker = $true
        GetBatchIDForExistingData
        [int]$script:ProcessGroups = '0'
        Foreach ($DLGroup in $script:DistributionListsDetails){
            CreateContact


        }
        ShowMenu

	}

    '5' {
        # Finalize Cloud groups to Original (365)
        CheckOnline
        $script:FinalisationMarker = $true
        GetBatchIDForExistingData
        [int]$script:ProcessGroups = '0'
        Foreach ($DLGroup in $script:DistributionListsDetails){
            FinaliseGroup

        }
        ShowMenu

	}
    
    '10' {
        # Connect to Onprem Exchange Server
        ConnectOnPremExchange
        ShowMenu 
    
    }

    '11' {
        # Connect to Exchange Online
        ConnectOnlineExchange
        ShowMenu 
    
    }

    '12' {
        # Connect to Exchange Online
        CloseExchangeConnection 
        ShowMenu 
    
    }
   

    default {
        write-host 'You may select one of the options'
        ShowMenu
    }   
    }
}

# FUNCTION - Display Banner information
function DisplayExtendedInfo () {

    # Display to notify the operator before running
    Clear-Host
    Write-Host 
    Write-Host 
    Write-Host  '-------------------------------------------------------------------------------'	
	Write-Host  '                   Distribution Group Re-creator Tool Kit                      '   -ForegroundColor Green
	Write-Host  '-------------------------------------------------------------------------------'
    Write-Host  '                                                                               '
    Write-Host  '  This tool is used to help identify DLs that can be moved/re-created in 365   '   -ForegroundColor YELLOW
    Write-Host  '                                                                               '   -ForegroundColor YELLOW
    Write-Host  "                                                              version: $version"   -ForegroundColor YELLOW
    Write-Host  '-------------------------------------------------------------------------------'
    Write-Host 
}

# FUNCTION - WriteTransaction Log function    
function WriteTransactionsLogs  {

    #WriteTransactionsLogs -Task 'Creating folder' -Result information  -ScreenMessage true -ShowScreenMessage true exit #Writes to file and screen, basic display
          
    #WriteTransactionsLogs -Task task -Result Error -ErrorMessage errormessage -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError true #Writes to file and screen and system "error[0]" is recorded
         
    #WriteTransactionsLogs -Task task -Result Error -ErrorMessage errormessage -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError false  #Writes to file and screen but no system "error[0]" is recorded
         


    [CmdletBinding()]
    param(
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]$Task,
 
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateSet('Information','Warning','Error','Completed','Processing')]
        [string]$Result,
 
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [string]$ErrorMessage,
    
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateSet('True','False')]
        [string]$ShowScreenMessage,
 
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [string]$ScreenMessageColour,
 
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$IncludeSysError,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$ExportData
)
 
    process {
 
        # Stores Variables
        #$LogsFolder           = 'Logs'
 
        # Date
        $DateNow = Get-Date -f g    
        
        # Error Message
        $script:SysErrorMessage = $error[0].Exception.message
 
  
 
        $TransactionLogScreen = [pscustomobject][ordered]@{}
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Date"-Value $DateNow 
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Task" -Value $Task
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Result" -Value $Result
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Error" -Value $ErrorMessage
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "SystemError" -Value $script:SysErrorMessage
        
       
        # Output to screen
       
        if  ($Result -match "Information|Warning" -and $ShowScreenMessage -eq "$true"){
 
        Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
        Write-host " | " -NoNewline
        Write-Host $TransactionLogScreen.Task  -NoNewline
        Write-host " | " -NoNewline
        Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour 
        }
 
       if  ($Result -eq "Error" -and $ShowScreenMessage -eq "$true" -and $IncludeSysError -eq "$false"){
       Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
       Write-host " | " -NoNewline
       Write-Host $TransactionLogScreen.Task  -NoNewline
       Write-host " | " -NoNewline
       Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour -NoNewline 
       Write-host " | " -NoNewline
       Write-Host $ErrorMessage  -ForegroundColor $ScreenMessageColour
       }
 
       if  ($Result -eq "Error" -and $ShowScreenMessage -eq "$true" -and $IncludeSysError -eq "$true"){
       Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
       Write-host " | " -NoNewline
       Write-Host $TransactionLogScreen.Task  -NoNewline
       Write-host " | " -NoNewline
       Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour -NoNewline 
       Write-host " | " -NoNewline
       Write-Host $ErrorMessage -NoNewline -ForegroundColor $ScreenMessageColour
       if (!$SysErrorMessage -eq $null) {Write-Host " | " -NoNewline}
       Write-Host $script:SysErrorMessage -ForegroundColor $ScreenMessageColour
       Write-Host
       }
   
        # Build PScustomObject
        $TransactionLogFile = [pscustomobject][ordered]@{}
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Date"-Value "$datenow"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Task"-Value "$task"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Result"-Value "$result"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Error"-Value "$ErrorMessage"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "SystemError"-Value "$script:SysErrorMessage"
 
        # Connect to Database
        if ($script:DatabaseConnected -eq $true){Open-LiteDBConnection $DBName -Mode shared | Out-Null}

        # Export data if NOT specified
        if(!($ExportData)){$TransactionLogFile |  ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection Transactions}
        
        
 
 
        # Clear Error Messages
        $error.clear()
    }   
 
}
# FUNCTION - Check for NoSQL Database Module
function CheckforPSliteDBModule () {

    # Find is PSliteDB module is installed
    if (Get-Module -ListAvailable -Name PSLiteDB) {WriteTransactionsLogs -Task "Found PSliteDB Module via ListAvailable" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False -ExportData False
        $Global:PSliteDBModuleLocation = 'ListAvailable'}

    # Check of the PSliteDB module is located in the script directory
    Elseif (Test-Path .\PSliteDB\module\PSLiteDB.psd1) {WriteTransactionsLogs -Task "PSLiteDB Module found in script directory" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False -ExportData False
        $Global:PSliteDBModuleLocation = 'Directory'}

    Else {WriteTransactionsLogs -Task "No Database Module found, install the PSliteDB from 'https://github.com/v2kiran/PSLiteDB', application will now close" -Result Information -ErrorMessage "Missing Database Module" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False -ExportData False
    Exit}
}
# FUNCTION - Import PSlite Module 
function ImportPSliteDBModule () {

    # Import module from List Available
    if ($PSliteDBModuleLocation -eq "ListAvailable") {
        WriteTransactionsLogs -Task "Importing PSliteDB Module via ListAvailable" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False -ExportData False
        try {Import-Module PSliteDB -ErrorAction Stop; $Global:PSliteModuleImported = $true}
        Catch {WriteTransactionsLogs -Task "Error loading module from ListAvailable" -Result Information -ErrorMessage "Import Failed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False -ExportData False}
    }

    # import module from Directory
    if ($PSliteDBModuleLocation -eq "Directory") {
        WriteTransactionsLogs -Task "Importing PSliteDB Module via Directory" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False -ExportData False
        try {Import-Module .\PSliteDB\module\PSLiteDB.psd1 -ErrorAction Stop ; $Global:PSliteModuleImported = $true}
        Catch {WriteTransactionsLogs -Task "Error Importing module from Directory" -Result Information -ErrorMessage "Import Failed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError True -ExportData False}
    }

    if ($null -eq $PSliteModuleImported ){TerminateScript}
}

# FUNCTION - Setup PSlteDB Database
function DatabaseConnection () {

    if ($PSliteModuleImported -eq $true){

        # Test if database exists
        $TestDBExists = Test-Path $DBName

        # Checks if database exists and then creates if not found
        if ($TestDBExists){ 
            try {Open-LiteDBConnection $DBname -Mode shared | Out-Null ; $script:DatabaseConnected = $true
            WriteTransactionsLogs -Task "Connected to Database $DBName" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False -ExportData False}
            Catch {WriteTransactionsLogs -Task "Connection to database Failed" -Result Error -ErrorMessage "Connection Error:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError True -ExportData False}
        }
        Else {Try {New-LiteDBDatabase -Path $DBname | Out-Null
            WriteTransactionsLogs -Task "Creating Database $DBname" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False -ExportData False
            Open-LiteDBConnection $DBName -Mode shared | Out-Null ; $script:DatabaseConnected = $true} 
            catch {WriteTransactionsLogs -Task "Failed to Create Database $DBname" -Result Information -ErrorMessage "Error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError True -ExportData False}
        }
    
        if ($script:DatabaseConnected -eq $true){
    
            # Create Collections in Database
            WriteTransactionsLogs -Task "Checking for Database Collections" -Result Information -ErrorMessage "None" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False -ExportData False
            New-LiteDBCollection Transactions -ErrorAction SilentlyContinue -WarningAction SilentlyContinue 
            New-LiteDBCollection DLGroups -ErrorAction SilentlyContinue -WarningAction SilentlyContinue 
            New-LiteDBCollection DLMembers -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            New-LiteDBCollection DLFailedTasks -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            New-LiteDBCollection DLCompleted -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            New-LiteDBCollection BatchInformation -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            
            

        }
    }

}

# FUNCTION - Get CSV Data from File
function ImportCSVData () {

    WriteTransactionsLogs -Task "Importing Data file................$CSVDataFile"   -Result Information none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
    try {$script:DistributionLists = Import-Csv ".\$CSVDataFile" -Delimiter "," -ea stop
        WriteTransactionsLogs -Task "Loaded Groups Data"   -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    } 
    catch {WriteTransactionsLogs -Task "Error loading Users data File" -Result Error -ErrorMessage "An error happened importing the data file, Please Check File" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
         Exit
    }

    if ($script:DistributionLists){
        
            if ($script:DistributionListsDetails){Remove-Variable -Scope Global -Name DistributionListsDetails}
            
            # Build PScustomObject
            $script:DistributionListsDetails = @()
            $script:ProcessedGroups = '0'
            Foreach ($DLGroup in $script:DistributionLists) {

                # Build String
                $DLPrimarySmtpAddress = $DLGroup.PrimarySmtpAddress
                $CloudDLPrimarySmtpAddress = "cloud-$DLGroup.PrimarySmtpAddress"

                # See if the existing cloud DL has been created to skip

                Try {Get-DistributionGroup -identity $CloudDLPrimarySmtpAddress -EA Stop
                    WriteTransactionsLogs -Task "Exising Cloud Group found $CloudDLPrimarySmtpAddress - Group wil be skipped" -Result Information -ErrorMessage "ExisingGroupFound" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false}
                Catch {$ExistingCloudDLFound = 'NO'}
            
                IF ($ExistingCloudDLFound -eq 'NO'){
                    Try {$script:DistributionListsDetails += Get-DistributionGroup -identity $DLPrimarySmtpAddress -EA Stop}
                    Catch {WriteTransactionsLogs -Task "NOT Found Group $DLPrimarySmtpAddress" -Result ERROR -ErrorMessage "Group Not Found" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
                    $GroupNotFound ++
                }
            }
        }

    }
    if (!($GroupNotFound)){$GroupNotFoundCount = '0'}
    Else {$GroupNotFoundCount = $GroupNotFound }
    $DistributionListsDetailsCount = $script:DistributionListsDetails | Measure-Object | Select-Object -ExpandProperty Count
    WriteTransactionsLogs -Task "Active Directory Groups Found $DistributionListsDetailsCount from CSV, Not Found $GroupNotFoundCount" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False
    ShowMenu
}

# FUNCTION - Create Batch Information to track jobs
function BatchInformation () {
    
    # Get-date
    $datenow = Get-Date -f g

    # Open Database Connection
    Open-LiteDBConnection $DBName -Mode shared | Out-Null

    # Clean up variable
    Remove-Variable nextbatch -Scope Global -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
 
    # Find if any batches have been created before
    $AnyExistingBatches = Find-LiteDBDocument -Collection 'BatchInformation'

    if ($AnyExistingBatches){
     
        # Find last batch ID
        [int]$lastbatch = Find-LiteDBDocument -Collection 'BatchInformation' | Select-Object -ExpandProperty  _id | Measure-Object -Maximum | Select-Object -ExpandProperty Count

        # New batch ID
        [int]$Global:nextbatch = $lastbatch +1

    }

    # Build PScustomObject
    $BatchReport = @()
    $BatchReport = [pscustomobject][ordered]@{}

    # Ask For information about the batch... free text
    $BatchNotes = Read-Host -Prompt "Enter notes about this batch"

    if ($Global:nextbatch) {$BatchReport | Add-Member -MemberType NoteProperty -Name "_id" -Value "$nextbatch" -force}
    if (!($Global:nextbatch)){$BatchReport | Add-Member -MemberType NoteProperty -Name "_id" -Value "1" -force ;[int]$global:nextbatch = '1' }
       
    $BatchReport | Add-Member -MemberType NoteProperty -Name "DateJobCreation" -Value "$datenow" -force
    $BatchReport | Add-Member -MemberType NoteProperty -Name "ScriptVersion" -Value $Version -force
    $BatchReport | Add-Member -MemberType NoteProperty -Name "BatchNotes" -Value $BatchNotes -force


   
    # Add Batch Details to Database 
    $BatchReport | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection BatchInformation
    Close-liteDBConnection

    # Update Status bar
    $host.ui.RawUI.WindowTitle = "Distribution Group Toolkit | $script:Connected2 | BatchID $nextbatch" 

        

}

# FUNCTION - Check Powershell version 
function CheckPSVersion () {

    if ($PSVersionTable.PSVersion.Major -ge '7') {}
    Else {
    
    Write-host
    Write-host
    Write-host
    Write-host "Powershell 7 and above is required to run this script. This script will now close." -ForegroundColor YELLOW 
    Write-host
    Write-host

    Exit  
 }
}

# FUNCTION - Check for Nested or Security Groups as members
function CheckGroupForSecOrNested () {

    Try{
    $DistributionGroupMembers = Get-DistributionGroupMember -Identity $DLPrimarySmtpAddress -ResultSize Unlimited
    $NestedGroupsFound = $DistributionGroupMembers | Where-Object{$_.RecipientType -like '*Group*'}
    $SecurityGroupsFound = $DistributionGroupMembers | Where-Object{$_.RecipientType -like '*Security*'}
    $NestedGroupsFoundCount = $NestedGroupsFound | Measure-Object | Select-Object -ExpandProperty count
    $SecurityGroupsFoundCount = $SecurityGroupsFound | Measure-Object | Select-Object -ExpandProperty count
    WriteTransactionsLogs -Task "Discovery for Nested & Security Groups completed. Nested:($NestedGroupsFoundCount) / Security:($SecurityGroupsFoundCount) - $DLPrimarySmtpAddress" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    }
    Catch {WriteTransactionsLogs -Task "Failed to get distribution group members for $DLPrimarySmtpAddress" -Result ERROR -ErrorMessage "Get-DistributionGroupMember failed: " -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true}

}

# FUNCTION - Check if running commands Online
function CheckOnline () {

    # check if the script is running in cloud else stop
    try {$OnlineCheck = Get-DistributionGroup -resultsize 1 -WarningAction SilentlyContinue -ea stop}
    Catch {WriteTransactionsLogs -Task "Please Connect to Exchange Online first" -Result ERROR -ErrorMessage "No Connected" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        ShowMenu}
    if ($OnlineCheck.DistinguishedName -match "DC=PROD,DC=OUTLOOK,DC=COM"){}
    Else {WriteTransactionsLogs -Task "Script looks to be running against Onprem objects, QUITTING" -Result ERROR -ErrorMessage "Wrong Platform" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        ShowMenu
    }
}

# FUNCTION - Check if running commands OnPrem
function CheckOnPrem () {

    # check if the script is running in cloud else stop
    try {$OnlineCheck = Get-DistributionGroup -resultsize 1 -WarningAction SilentlyContinue -ea stop}
    Catch {WriteTransactionsLogs -Task "Please Connect to Exchange On Premise first" -Result ERROR -ErrorMessage "No Connected" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        ShowMenu}
    if ($OnlineCheck.DistinguishedName -Notmatch  "DC=PROD,DC=OUTLOOK,DC=COM"){}
    Else {WriteTransactionsLogs -Task "Script looks to be running against cloud objects, QUITTING" -Result ERROR -ErrorMessage "Wrong Platform" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        ShowMenu
    }
}

# FUNCTION - Check Distribution Group data exists
function CheckDistributionGroupDataLoaded () {

    if (!($script:DistributionListsDetails)) {WriteTransactionsLogs -Task "PLASE RUN OPTION 1 OR 2 TO LOAD GROUPS!" -Result Error -ErrorMessage "No Distribution group data found to process" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false -ExportData false
        ShowMenu 
    }
    
}

# FUNCTION - Ask for individual Group
function SingleADGroup () {
    # Get a single AD group from AD via search
    $SingleGroup = Read-Host -Prompt "Enter the group SMTP Address"
    Write-Host `n
    if ($SingleGroup -eq "") {WriteTransactionsLogs -Task "No group was entered" -Result Error -ErrorMessage "No SMTP Address entered" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False
        Write-Host `n
        ShowMenu
    }

    # Build String
    $CloudDLPrimarySmtpAddress = "cloud-$SingleGroup"

     # See if the existing cloud DL has been created to skip

    Try {Get-DistributionGroup -identity $CloudDLPrimarySmtpAddress -EA Stop
        WriteTransactionsLogs -Task "Exising Cloud Group found $CloudDLPrimarySmtpAddress - Group wil be skipped" -Result Information -ErrorMessage "ExisingGroupFound" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false}
        Catch {$ExistingCloudDLFound = 'NO'}
    
    IF ($ExistingCloudDLFound -eq 'NO'){
        Try {$script:DistributionListsDetails = Get-DistributionGroup -identity $SingleGroup  -EA Stop
        
        $script:DistributionListsDetailsDisplayName = $script:DistributionListsDetails.DisplayName
        WriteTransactionsLogs -Task "Found Group $script:DistributionListsDetailsDisplayName" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False    
        $script:ProcessedGroups = '0'
        }
        Catch {WriteTransactionsLogs -Task "Group was not found" -Result Error -ErrorMessage "Not Found in AD" -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError True}
    }
}

# FUNCTION - Create PlaceHolder Group
function CreatePlaceHolder () {


    # Add Batch Details to Database 
    Open-LiteDBConnection $DBName -Mode shared | Out-Null

    # Remove existing entry from database if found
    Remove-LiteDBDocument -Collection DLgroups -ID $DLGroup.PrimarySmtpAddress -WarningAction SilentlyContinue
    Remove-LiteDBDocument -Collection DLMembers -ID $DLGroup.PrimarySmtpAddress -WarningAction SilentlyContinue

    # Build data to be added Group
    $DLGroup | Add-Member -MemberType NoteProperty -Name "_id" -Value $DLGroup.PrimarySmtpAddress -force
    $DLGroup | Add-Member -MemberType NoteProperty -Name "BatchID" -Value $Global:nextbatch -force
    $DLGroup | Add-Member -MemberType NoteProperty -Name "CreatePlaceHolder" -Value "True" -force
    $DLGroup | Add-Member -MemberType NoteProperty -Name "CreateContact" -Value "" -force
    $DLGroup | Add-Member -MemberType NoteProperty -Name "FinaliseGroup" -Value "" -force
    
    # Build data to be added members
    $DLMembers = @()
    $DlMembers = [pscustomobject][ordered]@{}
    $DLMembersDNs = (Get-DistributionGroupMember $DLGroup.PrimarySmtpAddress -resultsize unlimited).DistinguishedName
    $DLMembers | Add-Member -MemberType NoteProperty -Name "_id" -Value $DLGroup.PrimarySmtpAddress -force
    $DLMembers | Add-Member -MemberType NoteProperty -Name "Members" -Value $DLMembersDNs -force
    WriteTransactionsLogs -Task "Exporting $DLGroup to Database" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false

    # Collect DL Permissions
    $script:DLGroupPermissions = Get-RecipientPermission -Identity $DLGroup.PrimarySMTPAddress
        
    
    # Create strings for new DL
    $Cloud = 'Cloud-'
    $CloudName                       = $Cloud+$DLGroup.name
    $CloudAlias                      = $Cloud+$DLGroup.Alias
    $CloudDisplayName                = $Cloud+$DLGroup.DisplayName
    $CloudPrimarySMTPAddress         = $Cloud+$DLGroup.PrimarySMTPAddress
    $CloudMembers                    = $DLMembersDNs    
    
    # Find recipient objects by SMTP to be added to new Cloud Group
    [System.Collections.ArrayList]$AcceptMessagesOnlyFromSendersOrMembers = @()
    Foreach ($object in $DLGroup.AcceptMessagesOnlyFromSendersOrMembers){
    $AcceptMessagesOnlyFromSendersOrMembers += Get-Recipient -Filter "Name -eq '$object'" -ea SilentlyContinue | Select-Object -ExpandProperty PrimarySmtpAddress}

    [System.Collections.ArrayList]$RejectMessagesFromSendersOrMembers = @()
    Foreach ($object in $DLGroup.RejectMessagesFromSendersOrMembers){
    $RejectMessagesFromSendersOrMembers += Get-Recipient -Filter "Name -eq '$object'" -ea SilentlyContinue | Select-Object -ExpandProperty PrimarySmtpAddress}

    [System.Collections.ArrayList]$AcceptMessagesOnlyFrom = @()
    Foreach ($object in $DLGroup.AcceptMessagesOnlyFrom){
    $AcceptMessagesOnlyFrom += Get-Recipient -Filter "Name -eq '$object'" -ea SilentlyContinue | Select-Object -ExpandProperty PrimarySmtpAddress}

    [System.Collections.ArrayList]$AcceptMessagesOnlyFromDLMembers = @()
    Foreach ($object in $DLGroup.AcceptMessagesOnlyFromDLMembers){
    $AcceptMessagesOnlyFromDLMembers += Get-Recipient -Filter "Name -eq '$object'" -ea SilentlyContinue | Select-Object -ExpandProperty PrimarySmtpAddress}

    [System.Collections.ArrayList]$BypassModerationFromSendersOrMembers = @()
    Foreach ($object in $DLGroup.BypassModerationFromSendersOrMembers){
    $BypassModerationFromSendersOrMembers += Get-Recipient -Filter "Name -eq '$object'" -ea SilentlyContinue | Select-Object -ExpandProperty PrimarySmtpAddress
    }

    # Check the status of ManagedBy is valid
    $ManagedByCheck = $DLGroup | Select-Object -ExpandProperty Managedby
    $ManagedByCheck = $ManagedByCheck | Where-Object {$_ -NE 'Organization Management'}
    $DLGroup | Add-Member -MemberType NoteProperty -Name "ManagedBy" -Value $ManagedByCheck -Force

    if (([string]::IsNullOrEmpty($ManagedByCheck))){

        WriteTransactionsLogs -Task "ManagedBy is missing or'Organization management', using $DefaultManagedBy" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
        $DLGroup | Add-Member -MemberType NoteProperty -Name "ManagedBy" -Value $DefaultManagedBy -Force
    }

    try {
    New-DistributionGroup `
    -Name $CloudName `
    -Alias $CloudAlias `
    -Members $CloudMembers `
    -DisplayName $CloudDisplayName `
    -ManagedBy $DLGroup.ManagedBy `
    -PrimarySmtpAddress $CloudPrimarySMTPAddress -EA Stop | Out-Null

    Start-Sleep -Seconds 1

    Set-DistributionGroup `
    -Identity $CloudPrimarySMTPAddress `
    -AcceptMessagesOnlyFromSendersOrMembers $AcceptMessagesOnlyFromSendersOrMembers `
    -RejectMessagesFromSendersOrMembers $RejectMessagesFromSendersOrMembers `
    -MailTip $DLGroup.MailTip -EA STOP | Out-Null `


    Set-DistributionGroup `
    -Identity $CloudPrimarySMTPAddress `
    -AcceptMessagesOnlyFrom $AcceptMessagesOnlyFrom `
    -AcceptMessagesOnlyFromDLMembers $AcceptMessagesOnlyFromDLMembers `
    -BypassModerationFromSendersOrMembers $BypassModerationFromSendersOrMembers `
    -BypassNestedModerationEnabled $DLGroup.BypassNestedModerationEnabled `
    -CustomAttribute1 $DLGroup.CustomAttribute1 `
    -CustomAttribute2 $DLGroup.CustomAttribute2 `
    -CustomAttribute3 $DLGroup.CustomAttribute3 `
    -CustomAttribute4 $DLGroup.CustomAttribute4 `
    -CustomAttribute5 $DLGroup.CustomAttribute5 `
    -CustomAttribute6 $DLGroup.CustomAttribute6 `
    -CustomAttribute7 $DLGroup.CustomAttribute7 `
    -CustomAttribute8 $DLGroup.CustomAttribute8 `
    -CustomAttribute9 $DLGroup.CustomAttribute9 `
    -CustomAttribute10 $DLGroup.CustomAttribute10 `
    -CustomAttribute11 $DLGroup.CustomAttribute11 `
    -CustomAttribute12 $DLGroup.CustomAttribute12 `
    -CustomAttribute13 $DLGroup.CustomAttribute13 `
    -CustomAttribute14 $DLGroup.CustomAttribute14 `
    -CustomAttribute15 $DLGroup.CustomAttribute15 `
    -ExtensionCustomAttribute1 $DLGroup.ExtensionCustomAttribute1 `
    -ExtensionCustomAttribute2 $DLGroup.ExtensionCustomAttribute2 `
    -ExtensionCustomAttribute3 $DLGroup.ExtensionCustomAttribute3 `
    -ExtensionCustomAttribute4 $DLGroup.ExtensionCustomAttribute4 `
    -ExtensionCustomAttribute5 $DLGroup.ExtensionCustomAttribute5 `
    -GrantSendOnBehalfTo $DLGroup.GrantSendOnBehalfTo `
    -HiddenFromAddressListsEnabled $True `
    -MailTipTranslations $DLGroup.MailTipTranslations `
    -MemberDepartRestriction $DLGroup.MemberDepartRestriction `
    -MemberJoinRestriction $DLGroup.MemberJoinRestriction `
    -ModeratedBy $DLGroup.ModeratedBy `
    -ModerationEnabled $DLGroup.ModerationEnabled `
    -RejectMessagesFrom $DLGroup.RejectMessagesFrom `
    -RejectMessagesFromDLMembers $DLGroup.RejectMessagesFromDLMembers `
    -ReportToManagerEnabled $DLGroup.ReportToManagerEnabled `
    -ReportToOriginatorEnabled $DLGroup.ReportToOriginatorEnabled `
    -RequireSenderAuthenticationEnabled $DLGroup.RequireSenderAuthenticationEnabled `
    -SendModerationNotifications $DLGroup.SendModerationNotifications `
    -SendOofMessageToOriginatorEnabled $DLGroup.SendOofMessageToOriginatorEnabled `
    -BypassSecurityGroupManagerCheck | Out-Null -EA Stop

    # Setting Recipient Permissions
    WriteTransactionsLogs -Task "Setting SendAs Permissions For: $CloudPrimarySMTPAddress" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    Foreach ($mailbox in $script:DLGroupPermissionss){
        Try {Add-RecipientPermission -identity "$CloudPrimarySMTPAddress" -AccessRights "Sendas" -trustee $mailbox.Trustee -Confirm:$false -EA stop -WarningAction SilentlyContinue | Out-Null }
        Catch {WriteTransactionsLogs -Task "Failed Adding SendAs Permissions For: $CloudPrimarySMTPAddress Trustee:$mailbox.Trustee " -Result Warning -ErrorMessage AddPermissionsError -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError true}
    }   

    
    WriteTransactionsLogs -Task "$CloudDisplayName Group Created" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    # Record all data (if completed without error)
    $DLGroup | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection DLGroups 
    $DLMembers | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection DLMembers

    # Update Status bar
    $script:ProcessGroups ++
    $script:DistributionListsDetailsCount = $script:DistributionListsDetails.count
    $host.ui.RawUI.WindowTitle = "Distribution Group Toolkit | $Connected2 | BatchID $nextbatch | $script:ProcessGroups of $script:DistributionListsDetailsCount" 


    }

    Catch {WriteTransactionsLogs -Task "$CloudDisplayName Group failed to be created" -Result ERROR -ErrorMessage "Failed :" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
    
    $datenow = Get-Date -f g

    $script:DLFailedDetails = @()
    $script:DLFailedDetails = [pscustomobject][ordered]@{}
    $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "DateTime" -Value $datenow -force
    $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "BatchID" -Value $Global:nextbatch -force
    $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "DLGroup" -Value $DLGroup.PrimarySmtpAddress -force   
    $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "Task" -Value 'Failed to create Cloud Group' -force
    $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "TaskError" -Value $script:SysErrorMessage -force
    $script:DLFailedDetails | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection DLFailedTasks
    }

}

# FUNCTION - Get Batch and Collect existing data
function GetBatchIDForExistingData () {

    # Ask for Batch ID
    $AskBatchID = Read-Host -Prompt "Enter the group BatchID of a completed job"
    Write-Host `n
    if ($AskBatchID -eq "") {WriteTransactionsLogs -Task "No Batch ID was entered" -Result Error -ErrorMessage "No BatchID entered" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False
        Write-Host `n
        GetBatchIDForExistingData}

    Else {
        # Open Database Connection
        Open-LiteDBConnection $DBName -Mode shared | Out-Null

        # Find  batch ID
        $script:BatchID = Find-LiteDBDocument -Collection 'BatchInformation' | Where-object {$_._id -eq $AskBatchID}
        $BatchNotes = $script:BatchID | Select-Object -ExpandProperty BatchNotes


        # Check if batch exists and get data from DLGroups Collection
        if ($script:BatchID) {

            # Used to create the contacts list based on the CreatePlaceHolder process been run
            if ($script:CreatePlaceHolderMarker -eq '$true'){ $script:DistributionListsDetails = Find-LiteDBDocument -Collection 'DLGroups' | Where-object {($_.batchid -eq $AskBatchID) -and ($_.CreatePlaceHolder -eq 'True') -and ($_.CreateContact -like $null)}}

            # Used to create the finalisation list to run, if the contacts process has been run
            if ($script:FinalisationMarker -eq '$true'){ $script:DistributionListsDetails = Find-LiteDBDocument -Collection 'DLGroups' | Where-object {($_.batchid -eq $AskBatchID) -and ($_.CreatePlaceHolder -eq 'True') -and ($_.CreateContact -eq 'True') -and ($_.FinaliseGroup -like $null)}}

            $script:DistributionListsDetailsCount = $script:DistributionListsDetails.count
            $script:ProcessedGroups = '0'
            WriteTransactionsLogs -Task "Batch ID contains $script:DistributionListsDetailsCount groups to be finalised | Notes:$BatchNotes" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False
            Pause
            if ($script:DistributionListsDetails.count -eq '0') {ShowMenu}
        }
        Else {WriteTransactionsLogs -Task "No Batch found" -Result Error -ErrorMessage "No Batch found" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False}
    }

}

# FUNCTION - Finalise Cloud DL to be real DL
function FinaliseGroup () {

    # Add Batch Details to Database 
    Open-LiteDBConnection $DBName -Mode shared | Out-Null

    # Build information
    $ExistingProxyAddresses = @()  
    $ExistingProxyAddresses += $DLgroup | Select-Object -ExpandProperty EmailAddresses  
    $ExistingLEDN = $DLgroup | Select-Object -ExpandProperty LegacyExchangeDN
    $ExistingProxyAddresses += $ExistingLEDN = "x500:"+ $ExistingLEDN

    $CloudPrimarySMTPAddress = $DLgroup | Select-Object -ExpandProperty PrimarySMTPAddress
    $CloudPrimarySMTPAddress = "Cloud-$CloudPrimarySMTPAddress"

    $NewProxyAddresses = $ExistingProxyAddresses | ForEach-Object {$_ -Replace("X500","x500")}
    $NewPrimarySmtpAddress = ($NewProxyAddresses | Where-Object {$_ -clike "SMTP:*"}).Replace("SMTP:","")

    $NewDGName = $DLgroup | Select-Object -ExpandProperty Name
    $NewDGDisplayName = $DLgroup | Select-Object -ExpandProperty DisplayName
    $NewDGAlias = $DLgroup | Select-Object -ExpandProperty Alias

    Try {
        WriteTransactionsLogs -Task "Updating $CloudPrimarySMTPAddress to make primary distribution group" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        Set-DistributionGroup `
        -Identity $CloudPrimarySMTPAddress `
        -Name $NewDGName `
        -Alias $NewDGAlias `
        -DisplayName $NewDGDisplayName `
        -PrimarySmtpAddress $NewPrimarySmtpAddress `
        -HiddenFromAddressListsEnabled $False `
        -BypassSecurityGroupManagerCheck -EA Stop

        Set-DistributionGroup `
        -Identity $NewPrimarySmtpAddress `
        -EmailAddresses @{Add=$NewProxyAddresses} `
        -BypassSecurityGroupManagerCheck -EA Stop

        Start-Sleep -Seconds 1

        Set-DistributionGroup `
        -Identity $NewPrimarySmtpAddress `
        -EmailAddresses @{Remove="smtp:$CloudPrimarySMTPAddress"} `
        -BypassSecurityGroupManagerCheck -EA STOP

        $DLGroup | Add-Member -MemberType NoteProperty -Name "FinaliseGroup" -Value "True" -force
        $DLGroup | ConvertTo-LiteDbBSON | Update-LiteDBDocument -Collection DLGroups 

        WriteTransactionsLogs -Task "Completed $NewPrimarySmtpAddress updates and is now live" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        
        # Update Status bar
        $script:ProcessGroups ++
        $script:DistributionListsDetailsCount = $script:DistributionListsDetails.count
        $host.ui.RawUI.WindowTitle = "Distribution Group Toolkit | $Connected2 | BatchID $nextbatch | $script:ProcessGroups of $script:DistributionListsDetailsCount" 

    }
    Catch {WriteTransactionsLogs -Task "Failed Updating $CloudPrimarySMTPAddress with recorded values" -Result ERROR -ErrorMessage UpdateError -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
        
        $datenow = Get-Date -f g

        $script:DLFailedDetails = @()
        $script:DLFailedDetails = [pscustomobject][ordered]@{}
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "DateTime" -Value $datenow -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "BatchID" -Value $Global:nextbatch -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "DLGroup" -Value $CloudPrimarySMTPAddress -force   
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "Task" -Value 'Failed to finalise Cloud Group' -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "TaskError" -Value $script:SysErrorMessage -force
        $script:DLFailedDetails | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection DLFailedTasks

        $DLGroup | Add-Member -MemberType NoteProperty -Name "FinaliseGroup" -Value "False" -force
        $DLGroup | ConvertTo-LiteDbBSON | Update-LiteDBDocument -Collection DLGroups 
        exit
    }

}

# FUNCTION - Create Contact and hide original group
function CreateContact () {

    # Add Batch Details to Database 
    Open-LiteDBConnection $DBName -Mode shared | Out-Null

    # Build data
    $CurrentSMTP        = $DLgroup | Select-Object -ExpandProperty PrimarySmtpAddress
    $CurrentName        = $DLgroup | Select-Object -ExpandProperty Name
    $CurrentDisplayName = $DLgroup | Select-Object -ExpandProperty DisplayName
    $CurrentAlias       = $DLgroup | Select-Object -ExpandProperty Alias
    $GroupEmailAddresses = $DLgroup | Select-Object -ExpandProperty EmailAddresses


    # Create new object data to be stamped
    $NewSMTP        = "old_$CurrentSMTP"
    $NewName        = "old_$CurrentName"
    $NewDisplayName = "old_$CurrentDisplayName"
    $NewAlias       = "old_$CurrentAlias"


    # Find Onmicrosorft address to be used as target 
    $TargetOnMicrosoft = ($GroupEmailAddresses | Where-Object {$_ -clike "smtp:*mail.onmicrosoft.com"}) | Select-Object -First 1 
    If ($TargetOnMicrosoft){WriteTransactionsLogs -Task "Found Onmicrosoft Address for $CurrentSMTP : $TargetOnMicrosoft" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
    If (!($TargetOnMicrosoft)){WriteTransactionsLogs -Task "No Onmicrosoft Address for $CurrentSMTP - Group will be skiped" -Result ERROR -ErrorMessage "No Onmicrosoft Address Found" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false}

    Try {
        Set-DistributionGroup `
        -Identity $CurrentSMTP `
        -Alias $NewAlias `
        -DisplayName $NewDisplayName `
        -PrimarySmtpAddress $NewSMTP `
        -HiddenFromAddressListsEnabled $true `
        -CustomAttribute5 'DO-NOT-SYNC' `
        -EmailAddressPolicyEnabled $False `
        -BypassSecurityGroupManagerCheck `
        -DomainController $DCServer -EA STOP
        WriteTransactionsLogs -Task "Set DistributionGroup details to Old / Not Synced" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    
        Start-Sleep -Seconds 1

        Set-DistributionGroup `
        -Name $NewName `
        -Identity $NewSMTP `
        -EmailAddresses @{Remove=$TargetOnMicrosoft,"smtp:$CurrentSMTP"} `
        -BypassSecurityGroupManagerCheck `
        -DomainController $DCServer -EA STOP
       
    }
    Catch {WriteTransactionsLogs -Task "Failed to Set DistributionGroup details to Old / Not Synced" -Result ERROR -ErrorMessage "Failed to Set DL details to old" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
        
        $script:DLFailedDetails = @()
        $script:DLFailedDetails = [pscustomobject][ordered]@{}
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "DateTime" -Value $datenow -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "BatchID" -Value $Global:nextbatch -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "DLGroup" -Value $CloudPrimarySMTPAddress -force   
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "Task" -Value 'Failed to Set on premise Group values' -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "TaskError" -Value $script:SysErrorMessage -force
        $script:DLFailedDetails | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection DLFailedTasks
    
        $DLGroup | Add-Member -MemberType NoteProperty -Name "CreateContact" -Value "False" -force
        $DLGroup | ConvertTo-LiteDbBSON | Update-LiteDBDocument -Collection DLGroups 

    }

    # Create Contact with forwarder to cloud group 
    Try {
        writeTransactionsLogs -Task "Creating New contact based on Group details for $CurrentSMTP" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $NewContact = New-MailContact -Name  $CurrentName -OrganizationalUnit $ContactGroupOU -DisplayName  $CurrentDisplayname -PrimarySmtpAddress $CurrentSMTP -ExternalEmailAddress $TargetOnMicrosoft -Alias $CurrentAlias -DomainController $DCServer -EA STOP
        }
    Catch {writeTransactionsLogs -Task "Failed to create mail contact object" -Result ERROR -ErrorMessage "Failed creating contact" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        
        $script:DLFailedDetails = @()
        $script:DLFailedDetails = [pscustomobject][ordered]@{}
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "DateTime" -Value $datenow -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "BatchID" -Value $Global:nextbatch -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "DLGroup" -Value $CloudPrimarySMTPAddress -force   
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "Task" -Value 'Failed to create onprem Contact' -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "TaskError" -Value $script:SysErrorMessage -force
        $script:DLFailedDetails | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection DLFailedTasks
        
        $DLGroup | Add-Member -MemberType NoteProperty -Name "CreateContact" -Value "False" -force
        $DLGroup | ConvertTo-LiteDbBSON | Update-LiteDBDocument -Collection DLGroups     
        Exit
    }

    Start-Sleep -Seconds 1

    Try {
        WriteTransactionsLogs -Task "Setting Additional Mail Contact object information for $CurrentSMTP" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        Start-Sleep 1
        Set-MailContact -identity $NewContact.DistinguishedName -HiddenFromAddressListsEnabled:$true -CustomAttribute5 'DO-NOT-SYNC' -DomainController $DCServer -EA STOP
        Set-Contact -identity $NewContact.DistinguishedName -Notes "This contact is used to support the $CurrentName group located in 365. " -DomainController $DCServer -EA STOP
        
        #Set-ADObject -Identity $NewContact.DistinguishedName -add @{Notes="This contact is used to support the $CurrentName group located in 365."; Description="This contact is used to support the $CurrentName group located in 365."; AdminDescription="Group_NoSync"} -Server $DCServer -EA STOP        
        
        # Using DSmod as AD Module requires PS5
        dsmod contact $NewContact.DistinguishedName -desc "This contact is used to support the $CurrentName group located in 365." -server $DCServer -q
        
        $DLGroup | Add-Member -MemberType NoteProperty -Name "CreateContact" -Value "True" -force
        $DLGroup | ConvertTo-LiteDbBSON | Update-LiteDBDocument -Collection DLGroups

        writeTransactionsLogs -Task "Completed Mail Contact object for $CurrentSMTP" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    
        # Update Status bar
        $script:ProcessGroups ++
        $script:DistributionListsDetailsCount = $script:DistributionListsDetails.count
        $host.ui.RawUI.WindowTitle = "Distribution Group Toolkit | $Connected2 | BatchID $nextbatch | $script:ProcessGroups of $script:DistributionListsDetailsCount"     

    }
    Catch {writeTransactionsLogs -Task "Failed Setting additional Mail Contact object Information for $CurrentSMTP" -Result ERROR -ErrorMessage "Failed updating object" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
                       
        $script:DLFailedDetails = @()
        $script:DLFailedDetails = [pscustomobject][ordered]@{}
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "DateTime" -Value $datenow -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "BatchID" -Value $Global:nextbatch -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "DLGroup" -Value $CloudPrimarySMTPAddress -force   
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "Task" -Value 'Failed to update contact with additional settings' -force
        $script:DLFailedDetails | Add-Member -MemberType NoteProperty -Name "TaskError" -Value $script:SysErrorMessage -force
        $script:DLFailedDetails | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection DLFailedTasks
    
        $DLGroup | Add-Member -MemberType NoteProperty -Name "CreateContact" -Value "False" -force
        $DLGroup | ConvertTo-LiteDbBSON | Update-LiteDBDocument -Collection DLGroups
    }
}

# FUNCTION - Connecto to Onprem Exchange Server
function ConnectOnPremExchange () {

    WriteTransactionsLogs -Task "Connecting to Onprem Exchange Server - $OnPremExchange" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false -ExportData false
    Try {$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$OnPremExchange/powershell -Authentication Kerberos
    Import-PSSession $Session -WarningAction SilentlyContinue -DisableNameChecking | Out-Null
    # Update Status bar
    $script:Connected2 = "Connect to Exchange On Premise"
    $host.ui.RawUI.WindowTitle = "Distribution Group Toolkit | $Connected2 | BatchID $nextbatch" }
    Catch {WriteTransactionsLogs -Task "Failed Connecting to Onprem Exchange Server - $OnPremExchange" -Result ERROR -ErrorMessage "Connection " -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true -ExportData false}
}

# FUNCTION - Connecto to Exchange Online
function ConnectOnlineExchange () {

    WriteTransactionsLogs -Task "Connecting Exchange Online" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false -ExportData false
    
    # Import Exchaange Module
    Import-Module ExchangeOnlineManagement -WarningAction SilentlyContinue
    
    Try {
        Try { Get-OrganizationConfig -ea stop | Out-Null;  WriteTransactionsLogs -Task "Existing Exchange Online Connection Found" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false -ExportData false
            # Update Status bar
            $script:Connected2 = "Connect to Exchange Online"
            $host.ui.RawUI.WindowTitle = "Distribution Group Toolkit | $Connected2 | BatchID $nextbatch" }

        Catch {WriteTransactionsLogs -Task "Not Connected to Exchange Online" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false -ExportData false
               Connect-ExchangeOnline  -ErrorAction Stop | Out-Null
               # Update Status bar
               $script:Connected2 = "Connect to Exchange Online"
               $host.ui.RawUI.WindowTitle = "Distribution Group Toolkit | $Connected2 | BatchID $nextbatch"
            }
    }
    Catch {WriteTransactionsLogs -Task "Unable to Connect to Microsoft Exchange Online" -Result Error -ErrorMessage "Connect Error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true -ExportData false}
}

# FUNCTION - Close all Exchange Connections 
function CloseExchangeConnection () {
    Get-PSSession | Remove-PSSession
    Get-Module -Name tmp* | Remove-Module
    WriteTransactionsLogs -Task "Closed any open sessions to Exchange" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false -ExportData false
    $host.ui.RawUI.WindowTitle = "Distribution Group Toolkit | Not Connected | BatchID $nextbatch"
}





# Base Run Functions on Startup
CheckPSVersion
DisplayExtendedInfo
ShowMenu