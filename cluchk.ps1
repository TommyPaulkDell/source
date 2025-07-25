<#
.Synopsis
   S2D Ready Node Deployment Checker
.DESCRIPTION
    Script to verify the Dell S2D Ready Node Deployment guide was implemented fully
.CREATEDBY
    Jim Gandy 
and others
.NOTES
The script can be called without any parameters
examples:
.\CluChk.ps1
.\CluChk.ps1 -runType 3 -uploadToAzure $false -SDDCInputFolder 'C:\customer logs\Customer X - HCI\HealthTest-OL-P-HCI-CL01-20230308-1618\'

.Parameter runType
Specifies what to process, values are the same as menu values
.Parameter SDDCInputFolder
Location of the extracted SDDC collection
.Parameter TSRInputFolder
Locaton of the TSR/Support Assist files
.Parameter STSInputFolder
Location of the Show Tech support files
.Parameter uploadToAzure
Specifies if the collected data should be uploaded in Azure for analysis
.Parameter debug
Specifies to show debug information

.UPDATES
    2025/07/25:v1.69 -  1. New Update: TP - New version 1.69 DEV
                        2. New Update: TP - Added test-environment check for Last Action Plan Failure
                        3. New Feature: TP - Mark mismatched UBR build numbers as errors
                        4. New Update: TP - Pull system information into the html table if using the MS version of the SDDC

    2025/07/10:v1.68 -  1. New Update: TP - New version 1.68 DEV
                        2. New Update: TP - If a vm switch has no virtual adapter will mark RED with Note about health checks failing
                        3. New Update: TP - Reads GetActionPlanInstancesToComplete to find Errors. More to be done.gci 
                        4. Bug Fix: TP - Fixed Update Out of Box drivers now that we use the new method for the support matrix.
    
    2025/06/23:v1.67 -  1. New Update: TP - New version 1.67 DEV
                        2. New Update: JG - Rewrote the Support Matrix scrapter to use API instead of HTML parsing
                        3. New Update: JG - Rewrote Convert-HtmlTableToPsObject to account for possible Null values
                        4. New Update: JG - Physical Disks - Updated for all the changes to SM Scrapter & Convert-HtmlTableToPsObject
                        5. New Update: JG - Physical Disks - Added Kioxia drives to the table
                        6. New Update: JG - $GenerationNodes added AX740XD support
                        7. New Update: TP - $GenerationNodes added R?40?? Storage Spaces Direct support
                        8. New Feature: TP - Change all throw commands to Write-Warning so script will not stop
                        9. New Feature: TP - Only Invoke HTML file if CluChk is run on a system with a browser


    2025/05/16:v1.66 -  1. New Update: TP - New version 1.66 DEV
                        2. New Feature: TP - New tables for updates on Azure local 23H2 and Apex
                        3. New Feature: TP - Added new table Solution and SBE Updates
                        4. New Feature: TP - Check for CAU Role Automatic Updates
    
    2025/05/14:v1.65 -  1. New Update: TP - New version 1.65 DEV
                        2. Bug Fix: SA - Suppress remove old logfile error
                        3. Bug Fix: SA - Only warn if pagefile is larger than expected
                        4. Bug Fix: SA - Include FIPS disks in firmware check
                        5. Bug Fix: SA - Fix Recommended Updates and Hotfixes sorting, first per node (Asc), then per KB (Desc)
                        6. Bug Fix: SA - More minor sort fixes
                        7. Bug Fix: SA - Added Windows Server 2025 version
                        8. New Feature: SA - Add Azure Local version in Cluster node table
                        9. New Feature: SA - Check if Azure version is same on all nodes
                        10. New Feature: TP - Added check and error for mismatched time zones on nodes
                        11. Bug Fix: JG - Repaired the Show Tech Report as is was not working and slow
			
    2025/04/29:v1.64 -  1. New Update: TP - New version 1.64 DEV
                        2. Bug Fix: TP - Support Matrix changed link title from Azure Stack HCI to Azure Local. Updated the code.

    2025/05/xx:v1.63 -  1. New Update: TP - New version 1.63 DEV
                        2. New Feature: TP - If SysInfo cannot be gathered, still show the node in the Cluster Nodes table.
                        3. New Feature: TP - Call out simple virtual disks.
    
    2025/04/02:v1.62 -  1. New Update: TP - New version 1.62 DEV
                        2. Bug Fix: TP - Adding timeoutsec to all Invoke-WebRequests to 30 seconds.
                        3. Bug Fix: TP - Do not show error for E810 adapters using RoCEv2
    
    2025/01/23:v1.61 -  1. New Update: TP - New version 1.61 DEV
                        2. Bug Fix: SA - Corrected APEX ISM IP, should be 169.254.0.1
                        3. New Update: JG - Added new FLTMCXml file processing
                        4. Bug Fix: JL - Updated Jumbo Frames section
                        5. Bug Fix: JL - Added fix and sorting to FLTMCXml processing

    2024/12/11:v1.60 -  1. New Update: TP - New version 1.60 DEV
                        2. Bug Fix: TP - Do not show Management network as Red for LM if it's a regular PowerEdge server.
                        3. New Feature: TP - Cluster only networks show RED if excluded from live migration for Apex/HCI
                        4. Bug Fix: SA - Added Dell Ent NVMe CM7 disk model
                        5. Bug Fix: SA - Trim firmware results from support matrix
                        

    2024/10/31:v1.59 -  1. New Update: JG - New version 1.59 DEV
                        2. Bug Fix: SA - Fixed sorting/label in Host Management VLAN table
                        3. Bug Fix: SA - Fixed some more sorting issues
                        4. Bug Fix: SA - Fixed SMB signing required for 23h2
                        5. Bug Fix: SA - Fixed firmware catalog selection 14g/15g vs 16g
                        6. Bug Fix: SA - Hack for Broadcom driver version
                        7. Bug Fix: SA - Added warning for no SBE package validation on 23h2 and APEX
                        8. Bug Fix: SA - Added AX-45?0 node types to generations
                        9. Bug Fix: SA - Fixed catalog selection on unknown node generation
                        10. Bug Fix: TP - Added "DvFlt","MdvFsFlt","applockerfltr" to show as MS FLTMC drivers.
                        11. New Feature: SA - Added ISM IP checker
                        12. Bug Fix: SA - Some RED fixes
                        13. Bug Fix: SA - Suppress runspace output
    
    2024/08/21:v1.58 -  1. Bug Fix: TP - Fixed FileSystem field for Virtual Disks table
                        2. Bug Fix: TP - Fixed Virtual Disk status not showing on updated systems.
                        3. New Feature: TP - Show warning instead of failure when not able to down support matrix.
                        4. Bug Fix TP - Change Drift run code due to errors on some systems.
                        5. New Feature: TP - Set URL for 2016 HCI solutions to the archive support matrix page.
                        6. Bug Fix: SA - Fixed some text
                        7. New Feature: JG - NetAdapterAdvancedProperty - Exclude Host in Charge check for APEX
                        8. New Feature: JG - Cluster Nodes - Added UBR
    
    2024/06/12:v1.57 -  1. Bug Fix: TP - Ignore NetATC overrides on 23H2
                        2. Bug Fix: TP - Do not gather Windows updates for 23H2
    
    2024/05/09:v1.56 -  1. Bug Fix:JG - Added -UseBasicParsing to all Invoke-WebRequests for Server OS compatability
                        2. Bug Fix:TP - If CSV name and Volume name does not match make file system type blank instead of NTFS
    
    2024/04/09:v1.55 -  1. New Feature:TP - Drift moved to GitHub

    2024/04/05:v1.54 -  1. Bug Fix:TP - Fixed Computer Nodes section missing due to previous fix.

    2024/04/03:v1.53 -  1. Bug Fix:TP - If IOV is enabled ignore the BandwidthReservationMode and BandwidthPercentage settings
                        2. New Feature:TP - On LRs request, check for Mixed Mode clusters and show a problem in the Cluster Name object
                        3. Bug Fix: Using GetComputerInfo.xml instead of SystemInfo.txt because it seems to gather more reliabily
                        4. New Feature: JG - Added Invoke-RunCluChk

    2024/03/13:v1.52 -  1. New Feature: TP - Change HTML Agility pack to only download if not in C:\ProgramData\Dell\htmlagilitypack
                        2. New Feature: JG - Added APEX version to OSVersionNodes so that 23H2 will be red for non-APEX until we support it
                        3. New Feature: JG/TP - Cluster Heartbeat Configuration - Added check for Stretch Cluster heartbeat settings
                        4. Bug Fix: TP - Fixed an issue with System Page file that I caused in 1.51
                        5. Bug Fix: TP - Net Qos Dcbx Property table was not flagging Firmware in Charge as an Error
                        6. New Feature: TP - Added the file system to Cluster Shared Volumes
                        7. Bug Fix: TP - Net QOS Traffic Class was calling out ETS as an error when it is configured correctly.
                        8. New Feature: TP - Wrote PS Function to pull HTML tables. Removed downloading HTMLAgilityPack and changed sections using it.

    2024/02/09:v1.51 -  1. New Feature: TP - Call out a warning for a Cluster network with an IP but the Cluster Role set to None
                        2. New Feature: JG - Jumbo Frames: Changes Slot to Port and added Intel Nics for APEX support
                        3. New Feature: JG - Net Adapter Qos: ONLY check if Quality Of Service for APEX support
                        4. Bug Fix: TP - Qos DBCX settings did not show the Node name
                        5. New Feature: JG - Cluster Name: Added 23H2 for APEX support
                        6. New Feature: JG - Cluster Nodes: Added 25398 to OSBuild for APEX support
                        7. Bug Fix: JG - Net Qos Dcbx Property: Do not show if no dhbxmode found

    2024/01/25:v1.50 -  1. New Feature: JG - Cluster Nodes - Added check for EoL Operation Systems
                        2. Bug Fix: TP - Fixed disk firmware showing older firmware preferred if two are in the support matrix
                        3. Bug Fix: TP - Fixed Storage Network Cards and Node map due to changing to using jobs for all commands
                        4. Bug Fix: TP - Net QOS traffic class not showing correctly
                        5. New Feature: TP - Added unknown Net QOS Traffic Classes to the table
                        6. Bug Fix: TP - Change OEM Support Provider to only call out if it does not contain 'dell'

    See previous version for update notes
  
#>
`
Function Invoke-RunCluChk{
    
param (
[int]$runType = 0,
    [string]$SDDCInputFolder = '',
    [string]$TSRInputFolder = '',
[string]$STSInputFolder = '',
    [boolean]$uploadToAzure = $false,
    [boolean]$debug = $false
)

$CluChkVer="1.69"

#Fix "The response content cannot be parsed because the Internet Explorer engine is not available"
try {Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Value 2} catch {}

#region Variable Cleanup
#The following line causes alot of errors and does not appear to accomplish what is required.
# Remove-Variable * -Scope Local  -ErrorAction SilentlyContinue
# Recommend using a prefix on variables so remove-variable can -include prefix*
$TLogLoc = "C:\programdata\Dell\CluChk"
$DateTime=Get-Date -Format yyyyMMdd_HHmmss;Start-Transcript -NoClobber -Path "$TlogLoc\CluChk_$DateTime.log"
[system.gc]::Collect()
#endregion

# Function to convert HTML table to PowerShell objects
function Convert-HtmlTableToPsObject {
    param (
        [string]$HtmlContent
    )

    $tableRegex = '(?s)<h3 id="([^"]+)">(.*?)</h3>.*?<table[^>]*>(.*?)</table>'
    $tableNoNameRegex = '(?s)<table[^>]*>(.*?)</table>'
    $rowRegex = '<tr[^>]*>(.*?)</tr>'
    $thCellRegex = '<th[^>]*>(.*?)</th>'
    $tdCellRegex = '(.*?)</td>'
    $rowSpanRegex = 'rowspan="(\d+)"'

    $NoName = $false
    $tableMatches = [regex]::Matches($HtmlContent, $tableRegex)
    if ($tableMatches.Count -eq 0) {
        $tableMatches = [regex]::Matches($HtmlContent, $tableNoNameRegex)
        $NoName = $true
    }

    $result = @{}
    $tableNo = 0

    foreach ($tableMatch in $tableMatches) {
        if ($NoName) {
            $tableName = $tableNo
            $tableNo++
            $tableHtml = $tableMatch.Groups[1].Value
        } else {
            $tableName = $tableMatch.Groups[2].Value
            $tableHtml = $tableMatch.Groups[3].Value
        }

        $headers = @()
        [regex]::Matches($tableHtml, $thCellRegex) | ForEach-Object {
            $header = $_.Groups[1].Value.Trim()
            if (-not $header) { $header = "Column$($headers.Count + 1)" }
            $headers += $header
        }

        $rowMatches = [regex]::Matches($tableHtml, $rowRegex)
        $rows = @()
        $rowIndex = 0

        foreach ($rowMatch in ($rowMatches | Select-Object -Skip 1)) {
            $rowHtml = $rowMatch.Groups[1].Value
            $cellMatches = [regex]::Matches($rowHtml, $tdCellRegex)
            $rowData = @{}
            $colIndex = 0

            foreach ($cellMatch in $cellMatches) {
                if ($colIndex -ge $headers.Count) { continue }
                $curHeader = $headers[$colIndex]

                $cell = $cellMatch.Groups[1].Value
                if ($cell) {
                    $cell = $cell -replace "<td[^>]*>", ""
                    $cell = $cell.Trim()
                } else {
                    $cell = $null
                }

                $rowSpan = 1
                $splitCells = $rowHtml -split '</td>'
                if ($colIndex -lt $splitCells.Count) {
                    $rowSpanMatch = [regex]::Match($splitCells[$colIndex], $rowSpanRegex)
                    if ($rowSpanMatch.Success) {
                        $rowSpan = [int]$rowSpanMatch.Groups[1].Value
                    }
                }

                $rowData[$curHeader] = $cell

                if ($rowSpan -gt 1) {
                    for ($j = 1; $j -lt $rowSpan; $j++) {
                        $targetIndex = $rowIndex + $j
                        while ($rows.Count -le $targetIndex) {
                            $rows += @{}
                        }
                        $rows[$targetIndex][$curHeader] = $cell
                    }
                }

                $colIndex++
            }

            while ($rows.Count -le $rowIndex) {
                $rows += @{}
            }

            foreach ($k in $rowData.Keys) {
                $rows[$rowIndex][$k] = $rowData[$k]
            }

            $rowIndex++
        }

        $psRows = $rows | ForEach-Object { [PSCustomObject]$_ }
        $result[$tableName] = $psRows
    }

    return $result
}


#region Cleanup CluChk Transcripts older than 10 days
function TLogCleanup {
    Get-ChildItem $TlogLoc -ErrorAction SilentlyContinue | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-10)} | Foreach {Remove-Item $_.Fullname -force -ErrorAction SilentlyContinue}
  # Cleanup chipset extract
    Remove-Variable IntelChipset* -Scope Local  -ErrorAction SilentlyContinue
    IF($IntelChipsetDownloadLocation){Remove-Item $IntelChipsetDownloadLocation}
    IF($IntelChipsetDupExtractLocation){Remove-Item $IntelChipsetDupExtractLocation -Recurse}
}
#endregion
# Create a unique Guid to use for file names
$CluChkGuid=(New-Guid).GUID
#endregion
<#
#Async download of Agility pack
Start-Job -Name "AgilityPack" -ScriptBlock {
# Download htmlagilitypack
If (!(Get-ChildItem -Path "C:\ProgramData\Dell\htmlagilitypack" -ErrorAction SilentlyContinue -Recurse -Filter 'Net45' | Get-ChildItem | Where-Object{$_.Name -imatch 'HtmlAgilityPack.dll'}).count) {
   try {
   invoke-webrequest -Uri "https://www.nuget.org/api/v2/package/HtmlAgilityPack\1.11.46" -OutFile "$env:TEMP\htmlagilitypack.nupkg.zip" -TimeoutSec 30
   } 
   catch {
   Write-Warning ('Unable to download HTMLAgilityPack')
   }
   finally {
       # Find the htmlagilitypack.dll
       Expand-Archive "$env:TEMP\htmlagilitypack.nupkg.zip" -DestinationPath "C:\ProgramData\Dell\htmlagilitypack" -force -ErrorAction SilentlyContinue
   }
}}
#>
Function EndScript{  
    For ($i=0; $i -le 100; $i++) {
        Start-Sleep -Milliseconds 5
        Write-Progress -Activity "Exit Timer" -Status " " -PercentComplete $i
    }
    break script 
}
$WhatsNew=@"
1. in dev
"@


#region Opening Banner and menu
if (!$runType) {Clear-Host}
$text = @"
v$CluChkVer
   _____ _        _____ _     _         
  / ____| | ___  / ____| |   | |        
 | |    | |_   _| |    | |__ | | __ 
 | |    | | | | | |    | '_ \| |/ /
 | |____| | |_| | |____| | | |   <   
  \_____|_|\__,_|\_____|_| |_|_|\_\   
                      by: Jim Gandy 
"@
$Oops=@"
Oops... Something went wrong. Please try again.
"@
Write-Host $text
Write-Host ""

# Run Menu
Function ShowMenu{
    do
     {
         $Global:selection=""
         Clear-Host
         Write-Host $text
         Write-Host ""
         Write-Host "============ Please make a selection ==================="
         Write-Host ""
         Write-Host "Press '1' to Process Show Tech-Support(s)"
         Write-Host "Press '2' to Process Support Assist Collection(s)"
         Write-Host "Press '3' to Process PrivateCloud.DiagnosticInfo (SDDC)"
         Write-Host "Press '4' to Process S2D\HCI Performance Report"
         Write-Host "Press '5' to Process All the above"
         Write-Host "Press 'H' to Display Help"
         Write-Host "Press 'Q' to Quit"
         Write-Host ""
         $Global:selection = Read-Host "Please make a selection"
     }
    until ($Global:selection -match '[1-5,qQ,hH]')
    $Global:ProcessSTS  = "N"
    $Global:ProcessSDDC = "N"
    $Global:ProcessTSR  = "N"
    $Global:ProcessPerformanceReport = "N"
    IF($selection -imatch 'h'){
        Clear-Host
        Write-Host ""
        Write-Host "What's New in"$CluChkVer":"
        Write-Host $WhatsNew 
        Write-Host ""
        Write-Host "Usage:"
        Write-Host "    Make a selection by entering a comma delimited string of numbers from the menu."
        Write-Host ""
        Write-Host "        Example: 1 will process Show Tech-Support(s) only and create a report."
        Write-Host "                 Show Tech-Support is a log collection from a Dell switch."
        Write-Host ""
        Write-Host "        Example: 1,3 will process Show Tech-Support(s) and "
        Write-Host "                     PrivateCloud.DiagnosticInfo (SDDC) and create a report."
        Write-Host ""
        Pause
        ShowMenu
    }
    IF($Global:selection -match 1){
        Write-Host "Process Show Tech-Support(s)..."
        $Global:ProcessSTS = "Y"
    }

    IF($Global:selection -match 2){
        Write-Host "Process Support Assist Collection(s)..."
        $Global:ProcessTSR  = "Y"
    }
    IF($Global:selection -match 3){
        Write-Host "Process PrivateCloud.DiagnosticInfo (SDDC)..."
        $Global:ProcessSDDC = "Y"
    }
    IF($Global:selection -match 4){
        Write-Host "Process SDDC Performance Report..."
        $Global:ProcessPerformanceReport = "Y"
        #Enabled to get the SDDC then disabled once we have it
        $Global:ProcessSDDC = "Y"
    }
    ElseIF($Global:selection -eq 5){
        Write-Host "Process Show Tech-Support(s) + Support Assist Collection(s) + PrivateCloud.DiagnosticInfo (SDDC) + SDDC Performance Report..."
        $Global:ProcessSTS  = "Y"
        $Global:ProcessSDDC = "Y"
        $Global:ProcessTSR  = "Y"
        $Global:ProcessPerformanceReport = "Y"
    }
    IF($Global:selection -imatch 'q'){
        Write-Host "Bye Bye..."
        EndScript
    }
}#End of ShowMenu



#endregion
if ($runType -eq 0) {
ShowMenu
} else {
$Global:ProcessSTS  = "N"
    $Global:ProcessSDDC = "N"
    $Global:ProcessTSR  = "N"
$Global:ProcessPerformanceReport = "N"

switch ($runType) {
     1 {$Global:ProcessSTS  = "Y"}
     2 {$Global:ProcessTSR  = "Y"}
     3 {$Global:ProcessSDDC  = "Y"}
     4 {
         $Global:ProcessPerformanceReport = "Y"
 $Global:ProcessSDDC = "Y"
   
}
 5 {
         $Global:ProcessSTS  = "Y"
 $Global:ProcessSDDC = "Y"
 $Global:ProcessTSR  = "Y"
   }
}
}

Set-Variable -Name htmlout -Scope Global

#region Telemetry Information
Write-Host "Logging Telemetry Information..."
function add-TableData {
    [CmdletBinding()] 
        param(
            [Parameter(Mandatory = $true)]
            [string] $tableName,

            [Parameter(Mandatory = $true)]
            [string] $PartitionKey,

            [Parameter(Mandatory = $true)]
            [string] $RowKey,

            [Parameter(Mandatory = $true)]
            [array] $data,
            
            [Parameter(Mandatory = $false)]
            [array] $SasToken
        )
if ($uploadToAzure) {
$storageAccount = "gsetools"

# Allow only add and update access via the "Update" Access Policy on the CluChkTelemetryData table
# Ref: az storage table generate-sas --connection-string 'USE YOUR KEY' -n "CluChkTelemetryData" --policy-name "Update" 
If(-not($SasToken)){
$sasWriteToken = "?sv=2017-04-17&si=Update&tn=CluChkTelemetryData&sig=NP2ZQnHuUhOAOyGhzd94GVbKrBqhYlIKjX%2BVrhAjFoE%3D"
}Else{$sasWriteToken=$SasToken}

$resource = "$tableName(PartitionKey='$PartitionKey',RowKey='$Rowkey')"

# should use $resource, not $tableNmae
$tableUri = "https://$storageAccount.table.core.windows.net/$resource$sasWriteToken"
# Write-Host   $tableUri 

# should be headers, because you use headers in Invoke-RestMethod
$headers = @{
Accept = 'application/json;odata=nometadata'
}

$body = $data | ConvertTo-Json
#This will write to the table
#write-host "Invoke-RestMethod -Method PUT -Uri $tableUri -Headers $headers -Body $body -ContentType application/json"
try {
$item = Invoke-RestMethod -Method PUT -Uri $tableUri -Headers $headers -Body $body -ContentType application/json
} catch {
#write-warning ("table $tableUri")
#write-warning ("headers $headers")
}
} # if upload
}# End function add-TableData
    
# Generating a unique report id to link telemetry data to report data
$CReportID=""
$CReportID=(new-guid).guid
# Get the internet connection IP address by querying a public API
#$internetIp = (invoke-webrequest -uri "http://ifconfig.me/ip" -UseBasicParsing).Content 

# Define the API endpoint URL
$geourl = "http://ip-api.com/json" #$geourl = "http://ip-api.com/json/$internetIp"

# Invoke the API to determine Geolocation
$response = Invoke-RestMethod $geourl

$data = @{
    Region=$env:UserDomain
    Version=$CluChkVer
    ReportID=$CReportID
    country=$response.country
    counrtyCode=$response.countryCode
    georegion=$response.region
    regionName=$response.regionName
    city=$response.city
    zip=$response.zip
    lat=$response.lat
    lon=$response.lon
    timezone=$response.timezone
}

$RowKey=(new-guid).guid
$PartitionKey="CluChk"

#Make sure we get telemetry if uploadToAzure = $False
If($uploadToAzure -eq $false){
    $uploadToAzure = $true
    add-TableData -TableName "CluChkTelemetryData" -PartitionKey $PartitionKey -RowKey $RowKey -data $data
    $uploadToAzure=$false
}Else{add-TableData -TableName "CluChkTelemetryData" -PartitionKey $PartitionKey -RowKey $RowKey -data $data}

#endregion End of Telemetry data

# unzip files
function Unzip {
param([string]$zipfile, [string]$outpath)
Write-Host "    Expanding: "
Write-Host "      $SDDCLoc "
Write-Host "    To:"
Write-Host "      $ExtracLoc"
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}


Function Get-FileName([string]$initialDirectory, [string]$infoTxt, [string]$filter) {
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{MultiSelect = $true}
$OpenFileDialog.Title = $infoTxt
$OpenFileDialog.initialDirectory = $initialDirectory
$OpenFileDialog.filter = $filter
$OpenFileDialog.ShowDialog() | Out-Null
$OpenFileDialog.filenames
}

#region Ask for Show Tech-Support files
    #$ProcessSTS = Read-Host "Would you like to process Show Tech-Supports? [y/n]"
    IF($ProcessSTS -ieq "y"){
if ($STSInputFolder -eq '') {
Write-Host "Please Select Show Tech-Support File(s) to use..."
$STSLOC = Get-FileName "$env:USERPROFILE\Documents\SRs" "Please Select Show Tech-Support File(s)." "Logs (*.zip,*.txt,*.log)| *.zip;*.TXT;*.log"
} else {
$STSLOC = $STSInputFolder
}

If ($STSLOC.length -eq 0) {
Write-Host "    ERROR: Empty path given. Existing" -ForegroundColor Red
EndScript
}
if (!(Test-Path -Path $STSLOC)) {
Write-Host "    ERROR: STS Path does not exist. Existing" -ForegroundColor Red
EndScript
}

# unzip files
        $STSFiles=@()
        $CluChkReportLoc=(Split-Path -Path $STSLOC)
        $STSExtracLoc=(Split-Path -Path $STSLOC)+"\"+ (Split-Path -Path $STSLOC -Leaf).Split(".")[0]
        ForEach($STSFile in $STSLOC){
            If($STSFile -imatch '.zip'){                
                unzip $STSFile $STSExtracLoc
                $STSFiles+=(Get-ChildItem $STSExtracLoc).fullName
            }Else{$STSFiles+=$STSFile}
        }
} # if $ProcessSTS
#endregion

#region $ProcessSDDC = Read-Host "Would you like to process SDDC? [y/n]"
IF($ProcessSDDC -ieq "y"){
    Write-Host "Starting SDDC..."
    IF($ProcessPerformanceReport -ieq "y"){
        $SDDCPerf="YES"
    }Else{$SDDCPerf="NO"}

if ($SDDCInputFolder -eq '') {
# no input folder given
    
# Added to Select-Object the extracted SDDC
$SDDCExtracted = Read-Host "    Do you already have an extracted SDDC? [y/n]"
IF($SDDCExtracted -ieq "y") {
$SDDCExtracted="YES"
Write-Host "    Please Provide the path to the extracted SDDC. "
Write-Host "    Do NOT include Quotes even if path includes spaces."
Write-Host "      Ex: c:\SRs\81449725\HealthTest-NL70U00CL02-20200928-1130"
$SDDCPath=Read-Host "    Path"
$IR=1
$CluChkReportLoc=Split-Path -Path $SDDCPath
$ExtracLoc=(Split-Path -Path $SDDCPath) +"\"+ (Split-Path -Path $SDDCPath -Leaf).Split(".")[0]
            $IncompleteSDDC=$False
If(!(Test-Path -Path (Join-Path $SDDCPath 0_CloudHealthSummary.log))){
              If(!(Test-Path -Path (Join-Path $SDDCPath 0_CloudHealthGatherTranscript.log))){  
Do{ Write-Host "    ERROR: Invalid SDDC Path." -ForegroundColor Red 
$SDDCPath = Read-Host "Please try again"
$IR+=$IR++
$CluChkReportLoc=Split-Path -Path $SDDCPath
$ExtracLoc=(Split-Path -Path $SDDCPath) +"\"+ (Split-Path -Path $SDDCPath -Leaf).Split(".")[0]
IF($IR -gt 2){
Write-Host "    ERROR: Failed SDDC Path too many times. Extracing again." -ForegroundColor Red
$SDDCExtracted="NO"
break
   }}
While((!(Test-Path -Path "$SDDCPath\0_CloudHealthSummary.log")))
} else {
                $IncompleteSDDC=$True
                Write-Host "    WARNING: Incomplete SDDC Capture." -ForegroundColor Yellow
                gc (Join-Path $SDDCPath 0_CloudHealthGatherTranscript.log) -tail 10 | %{Write-Host "    $_ " -ForegroundColor Yellow}
            } 
          }
}

If($SDDCExtracted -ne "YES"){
Write-Host "    Please Select-Object SDDC File to use..."
$SDDCLoc=Get-FileName "$env:USERPROFILE\Documents\SRs" "Please Select SDDC File." "ZIP (*.zip)| *.zip"
if(!$SDDCLoc){EndScript}

#Extraction temp location
$CluChkReportLoc=Split-Path -Path $SDDCLoc
$ExtracLoc=(Split-Path -Path $SDDCLoc) +"\"+ (Split-Path -Path $SDDCLoc -Leaf).Split(".")[0]
Try{
If (Test-Path $ExtracLoc -PathType Container){Remove-Item $ExtracLoc -Recurse -Force -ErrorAction Stop | Out-Null}
}Catch{
Write-Host $Oops
Write-Host ""
Write-Host "$Error" -ForegroundColor Red
EndScript
   }
if (!(Test-Path $ExtracLoc -PathType Container)) {New-Item -ItemType Directory -Force -Path $ExtracLoc | Out-Null }
Unzip $SDDCLoc $ExtracLoc
$SDDCPath=$ExtracLoc
            $IncompleteSDDC=$False
If(!(Test-Path -Path (Join-Path $SDDCPath 0_CloudHealthSummary.log))){
              If(!(Test-Path -Path (Join-Path $SDDCPath 0_CloudHealthGatherTranscript.log))){  
 Write-Host "    ERROR: Invalid SDDC Path." -ForegroundColor Red 
break
} else {
                $IncompleteSDDC=$True
                Write-Host "    WARNING: Incomplete SDDC Capture." -ForegroundColor Yellow
                gc (Join-Path $SDDCPath 0_CloudHealthGatherTranscript.log) -tail 10 | %{Write-Host "    $_ " -ForegroundColor Yellow}
            } 
          }
} # if SDDCInputFolder
    } else {
if (!(Test-Path -Path $SDDCInputFolder)) {
Write-Host "    ERROR: SDDC Path does not exist. Existing" -ForegroundColor Red
EndScript
}
$SDDCPath = $SDDCInputFolder
$CluChkReportLoc=Split-Path -Path $SDDCPath
}
IF($selection -eq "4"){$Global:ProcessSDDC = "N"}
 }
#endregion SDDC Locate and Extract

#region  Ask for TSRs
#$ProcessTSR = Read-Host "Would you like to process TSRs? [y/n]"
IF($ProcessTSR -ieq "y"){
# Show Tech-Support Report
if ($TSRInputFolder -eq '') {
Write-Host "Please Select Support Assist Collection(TSR) File(s) to use..."
$TSRLOC = Get-FileName "$env:USERPROFILE\Documents\SRs" "Please Select Support Assist Collection(TSR) File(s) to use." ""
} else {
$TSRLOC = $TSRInputFolder
}
If ($TSRLOC.length -eq 0) {
Write-Host "    ERROR: Empty path given. Existing" -ForegroundColor Red
EndScript
}
if (!(Test-Path -Path $TSRLOC)) {
Write-Host "    ERROR: TSR Path does not exist. Existing" -ForegroundColor Red
EndScript
}
    $CluChkReportLoc=(Split-Path -Path $TSRLOC)
}
#endregion End Ask for TSRs


Function Convert-BytesToSize
    {
    <#
    .SYNOPSIS
    Converts any integer size given to a user friendly size.
    .DESCRIPTION
    Converts any integer size given to a user friendly size.
    .PARAMETER size
    Used to convert into a more readable format.
    Required Parameter
    .EXAMPLE
    ConvertSize -size 134217728
    Converts size to show 128MB

    #>
    #Requires -version 2.0
    [CmdletBinding()]
    Param
    (
    [parameter(Mandatory=$False,Position=0)][int64]$Size
    )

    #Decide what is the type of size
    Switch ($Size)
    {
    {$Size -gt 1PB}
    {
    Write-Verbose "Convert to PB"
    $NewSize = "$([math]::Round(($Size / 1PB),2))PB"
    Break
    }
    {$Size -gt 1TB}
    {
    Write-Verbose "Convert to TB"
    $NewSize = "$([math]::Round(($Size / 1TB),2))TB"
    Break
    }
    {$Size -gt 1GB}
    {
    Write-Verbose "Convert to GB"
    $NewSize = "$([math]::Round(($Size / 1GB),2))GB"
    Break
    }
    {$Size -gt 1MB}
    {
    Write-Verbose "Convert to MB"
    $NewSize = "$([math]::Round(($Size / 1MB),2))MB"
    Break
    }
    {$Size -gt 1KB}
    {
    Write-Verbose "Convert to KB"
    $NewSize = "$([math]::Round(($Size / 1KB),2))KB"
    Break
    }
    Default
    {
    Write-Verbose "Convert to Bytes"
    $NewSize = "$([math]::Round($Size,2))Bytes"
    Break
    }
    }
    Return $NewSize
}

#region Create HTML file and adding CSSfor output
#$DateTime=Get-date
$DTString=Get-Date -Format "yyyyMMdd_HHmmss_"
$htmlout=""
$html=""
$ResultsSummary=@()
$htmlStyle = @"
<style TYPE="text/css">
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
TR:Nth-Child(Even) {Background-Color: #dddddd;}
TR:Hover TD {Background-Color: #C1D5F8;}
h5 {display: block;font-size: .83em;margin-top: 0;margin-bottom: 0;margin-left: 0;margin-right: 0;font-weight: normal;padding: 0;}
h1{border-bottom: 5px solid #f4a460}
mark { 
    background-color: yellow;
    color: black;
  }
</style>
<script>
    function toggle(ele,cElem) {
        var cont = document.getElementById(ele);
        cont.style.display = cont.style.display == 'none' ? 'block' : 'none';
        cElem.innerText = cElem.innerText == '\u21ca\xa0' ? '\u21c9\xa0' : '\u21ca\xa0';
    }
</script>
"@

#endregion

Function Set-ResultsSummary{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Name,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $html
    )
# Add summary info
#$WarningsCount= 0
#$ErrorsCount= 0
            $Table          =@()
$Table   = [PSCustomObject]@{
Name = $name
Warnings    = ($html -split '\<\/tr\>' | Where-Object{$_ -imatch 'ffff00'}| Measure-Object).Count
Errors    = ($html -split '\<\/tr\>' | Where-Object{($_ -imatch 'ffffff') -or($_ -imatch 'font color="red"')}| Measure-Object).Count
}
            return $Table
} #function Set-ResultsSummary

<#
# Download htmlagilitypack
try {
invoke-webrequest -Uri "https://www.nuget.org/api/v2/package/HtmlAgilityPack/1.11.46" -OutFile "$env:TEMP\htmlagilitypack.nupkg.zip"
} 
catch {
Write-Warning ('Unable to download HTMLAgilityPack')
}

# Find the htmlagilitypack.dll
Expand-Archive "$env:TEMP\htmlagilitypack.nupkg.zip" -DestinationPath "$env:TEMP\htmlagilitypack" -force -ErrorAction SilentlyContinue
#>
<#If ((Get-Job -Name "AgilityPack").State -match 'Running') {
write-host ("Waiting on download job to complete")
$loopTimer = 0
$maxTimer = 150
while (((Get-Job -Name "AgilityPack").State -ne 'Completed') -and ($loopTimer -le $maxTimer)){
Sleep -Milliseconds 100
$loopTimer++
}
}
If ((Get-Job -Name "AgilityPack").State -ne 'Completed') {
Write-Host "Could not download Agility Pack" -ForegroundColor Red
Get-Job -Name "AgilityPack" | Receive-Job
Get-Job | Remove-Job -Force
exit
}
Get-Job | Remove-Job -Force
# Import the required libraries windows 10 or above will have .Net 4.5 or greater. Assuming .Net 4.5 lib
Add-Type -Path (Get-ChildItem -Path "C:\ProgramData\Dell\htmlagilitypack" -Recurse -Filter 'Net45' | Get-ChildItem | Where-Object{$_.Name -imatch 'HtmlAgilityPack.dll'}).fullname

#.Net download -
# $webClient.DownloadFile($URL, $OutFile)}
#>
If ($ProcessSDDC -ieq 'y') {

          #$SysInfoFiles=(Get-ChildItem -Path $SDDCPath -Filter "SystemInfo.TXT" -Recurse -Depth 1).FullName
          $GetComputerInfo=gci $SDDCPath -Filter "GetComputerInfo.xml" -Recurse -Depth 1 | Import-Clixml
        #$SystemInfoData=@()
        $SysInfo=@()
        <#ForEach($dFile in $SysInfoFiles){ # Changed 06072023
            $SystemInfoContent=@()
            $SystemInfoContent= Get-Content $dFile
            $osInfo=@{}
            $SystemInfoContent | Where-Object{($_ -like "OS*") -or ($_ -like "Host Name:*") -or ($_ -like "System Model:*")} | %{try {$osInfo.add($_.split(":")[0].trim(),$_.split(":")[1].trim())} catch {}}
            $SysInfo += [PSCustomObject] @{
                                HostName      = $osInfo.'Host Name'
                                OSVersion     = ($osInfo.'OS Version' -split "\s+")[0]
                                OSBuildNumber = ($osInfo.'OS Version' -split "\s+")[-1]
                                OSName        = $osInfo.'OS Name'.Replace('Microsoft','').Trim() # remove Microsoft
                                SysModel      = $osInfo.'System Model'
                            }
        }#>
        ForEach ($osInfo in $GetComputerInfo) {
            $SysInfo += [PSCustomObject] @{
                                HostName      = $osInfo.CsCaption
                                OSVersion     = $osInfo.OsVersion
                                OSBuildNumber = $osInfo.OsBuildNumber
                                OSName        = $osInfo.OSName.Replace('Microsoft','').Trim() # remove Microsoft
				                AzureLocalVersion = ''
                                SysModel      = $osInfo.CsModel
                            }

        }
        if (!($GetComputerInfo)) {
            $GetComputerInfo=gci $SDDCPath -Filter "Win32_OperatingSystem.xml" -Recurse -Depth 1 | Import-Clixml
            
            ForEach ($osInfo in $GetComputerInfo) {
            $SysInfo += [PSCustomObject] @{
                                HostName      = $osInfo.CsName
                                OSVersion     = $osInfo.Version
                                OSBuildNumber = $osInfo.BuildNumber
                                OSName        = $osInfo.Caption.Replace('Microsoft','').Trim() # remove Microsoft
				                AzureLocalVersion = ''
                                SysModel      = (gc -Path "$SDDCPath\Node_$($osInfo.CsName)\SystemInfo.TXT" | Select-String "System Model: ").Line.split(":")[-1].Trim()
                                } 
            }
        }

# determine os version, used later on
$OSVersionNodes = 'Unknown'
# Azure Stack HCI OS versions
If($SysInfo[0].OSName -imatch 'HCI'){
    # update Azure Local version in $SysInfo table, per node
    $StampFiles=gci $SDDCPath -Filter "GetStampInformation.xml" -Recurse -Depth 1
    ForEach ($StampFile in $StampFiles) {
      $NodeName = ($StampFile.DirectoryName | split-path -Leaf).Replace('Node_','')
      $GetStampInformation= $StampFile | Import-Clixml
      ForEach($Sys in $SysInfo){ 
        if ($Sys.HostName -eq $nodeName) {
          $Sys.AzureLocalVersion = $GetStampInformation[0].StampVersion
        }
      }
    }

    #NOT APEX nodes
    IF($SysInfo[0].SysModel -notmatch "^APEX"){
        $OSVersionNodes = Switch ($SysInfo[0].OSBuildNumber) {
        '19042' {'20H2'}
        '20348' {'21H2'}
        '20349' {'22H2'}
        '25398' {'23H2'}
        Default {"RREEDD"+$SysInfo[0].OSBuildNumber}
        }
    }
    #APEX nodes
    IF($SysInfo[0].SysModel -match "^APEX"){
        $OSVersionNodes = Switch ($SysInfo[0].OSBuildNumber) {
        '20349' {'22H2'}
        '25398' {'23H2'}
        Default {"RREEDD"+$SysInfo[0].OSBuildNumber}
        }
    }
} else {
# regular Windows Server versions
$OSVersionNodes = Switch ($SysInfo[0].OSBuildNumber) {
'26100' {'2025'}
'20348' {'2022'}
'17763' {'2019'}
'14393' {'2016'}
Default {"RREEDD"+$SysInfo[0].OSBuildNumber}
}
}
IF($SysInfo[0].SysModel -match "^APEX"){
	$GenerationNodes = "16g"
} elseif (($SysInfo[0].SysModel -like "AX-?60") -or ($SysInfo[0].SysModel -like "AX*45?0")){
	$GenerationNodes = "16g"
} elseif ($SysInfo[0].SysModel -like "AX-?50"){
	$GenerationNodes = "15g"
} elseif (($SysInfo[0].SysModel -like "AX-?40") -or ($SysInfo[0].SysModel -like "AX-?40??") -or ($SysInfo[0].SysModel -like "R?40*Storage Spaces Direct*")){
	$GenerationNodes = "14g"
} else {
	$GenerationNodes ="APEX"
}
}

#region Gather the Support Matrix for Microsoft HCI Solutions
Write-Host "    Gathering Support Matrix for Dell EMC Solutions for Microsoft Azure Stack HCI..."
$repo = "dell/azurestack-docs"
$basePath = "content/en/docs/hci/SupportMatrix"
$headers = @{ "User-Agent" = "PowerShell" }

# Step 1: Get SupportMatrix versions
$topResponse = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/contents/$basePath" -Headers $headers -UseBasicParsing
$latestFolder = $topResponse | Where-Object { $_.type -eq 'dir' -and $_.name -match '^\d{4}$' } |
    Sort-Object name -Descending | Select-Object -First 1
$latestVersion = $latestFolder.name
$latestVersionPath = "$basePath/$latestVersion"

# Step 2: Get platform folders under latest version
$platformFolders = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/contents/$latestVersionPath" -Headers $headers |
    Where-Object { $_.type -eq 'dir' }

# Step 3: Choose platform folder (update this as needed)
$matchedPlatform = $platformFolders | Where-Object { $_.name -imatch $GenerationNodes }
$mdUrl = $null

if ($matchedPlatform) {
    $mdPath = "$latestVersionPath/$($matchedPlatform.name)"
    $mdIndex = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/contents/$mdPath" -Headers $headers |
        Where-Object { $_.name -eq "_index.md" }
    $mdUrl = $mdIndex.download_url
}

if (-not $mdUrl) {
    Write-Warning "No matching _index.md found for platform $GenerationNodes"
}

# Check for archive Support Matrix
if ($OSVersionNodes -eq "2016" -or $OSVersionNodes -match "2012") {$mdUrl='https://raw.githubusercontent.com/dell/azurestack-docs/refs/heads/main/content/en/docs/hci/SupportMatrix/Archive/WS2016/_index.md'}Else{
$mdUrl}
$SMRevHistLatest=($mdURL -replace "https://raw.githubusercontent.com/dell/azurestack-docs/main/content/en","https://dell.github.io/azurestack-docs" -replace "_index.md","").ToLower()


# Step 4: Extract all HTML tables from _index.md
$mdText = (Invoke-WebRequest -Uri $mdUrl -UseBasicParsing).Content
$mdLines = $mdText -split "`n"

# Initialize variables
$NamedTables = @{}
$currentTitle = $null
$tableBuffer = @()
$inTable = $false

foreach ($line in $mdLines) {
    $trimmed = $line.Trim()

    # Match markdown table section heading
    if ($trimmed -match '^###\s+(.*)') {
        # Save previously collected table if any
        if ($currentTitle -and $tableBuffer.Count -gt 0) {
            $html = ($tableBuffer -join "`n").Trim()
            if ($html -like '<table*') {
                Write-Host "Parsed table: $currentTitle"
                $NamedTables[$currentTitle] = Convert-HtmlTableToPsObject -HtmlContent $html
            }
        }
        $currentTitle = $Matches[1].Trim()
        $tableBuffer = @()
        continue
    }

    # Begin collecting table HTML block
    if ($trimmed -eq '{{< rawhtml >}}') {
        $inTable = $true
        $tableBuffer = @()
        continue
    }

    # End collecting table HTML block
    if ($trimmed -eq '{{< /rawhtml >}}') {
        $inTable = $false
        if ($currentTitle -and $tableBuffer.Count -gt 0) {
            $html = ($tableBuffer -join "`n").Trim()
            if ($html -like '<table*') {
                Write-Host "Parsed table: $currentTitle"
                $NamedTables[$currentTitle] = Convert-HtmlTableToPsObject -HtmlContent $html
            }
        }
        $tableBuffer = @()
        continue
    }

    # Accumulate lines inside table block
    if ($inTable) {
        $tableBuffer += $line
    }
}

# Handle any final dangling table at end of file
if ($currentTitle -and $tableBuffer.Count -gt 0) {
    $html = ($tableBuffer -join "`n").Trim()
    if ($html -like '<table*') {
        Write-Host "Parsed final table: $currentTitle"
        $NamedTables[$currentTitle] = Convert-HtmlTableToPsObject -HtmlContent $html
    }
}

# Convert the table data to XML
$SMxml = $NamedTables | ConvertTo-Xml -As String -NoTypeInformation
$SupportMatrixtableData = $NamedTables

#$SupportMatrixtableData.Values.values

#endregion Support Matrix

#Report ID
$htmlout+=""
$htmlout+="ReportID: $CReportID"

#region Prcoess Show Tech-Support Report(s)
If($ProcessSTS -ieq 'y'){
# Write-Host "*******************************************************************************"
# Write-Host "*                                                                             *"
# Write-Host "*                            Show Tech-Support Report                         *"
# Write-Host "*                                                                             *"
# Write-Host "*******************************************************************************"
# Show Tech-Support Report
    # Filter Support Matrix for Microsoft HCI Solutions for Switch firmware information from
$html+='<H1 id="ShowTechSupport ">Show Tech-Support </H1>'
    $Name="Tech-Support"
    Write-Host "    Gathering $Name..."  
    $STS="C:\Users\jim_gandy\OneDrive - Dell Technologies\Documents\SRs\109739084\Switch_Logs\RDUAITSW01 after uplink.log"

    # Output Var
#$ShowTechSupportOut=@()
$parsedshowlldpneighbors=@()
$RunningConfigOut=@()
$ShowVersionOut=@()
$Interfaces=@()
    ForEach($STS in $STSFiles){
        Write-Host "    Gathering Show Tech information from $STS"
        $logText = Get-Content $STS -Raw
        $fileName = Split-Path $STS -Leaf
        
        # Match the header pattern
        #$pattern = "^-{35}\s+show .+?$"
        $pattern = '\s*-+\s+show .+'
        
        # Find all header matches
        $matches = [regex]::Matches($logText, $pattern, [System.Text.RegularExpressions.RegexOptions]::Multiline)
        
        # Build the sections
        $sections = @()
        for ($i = 0; $i -lt $matches.Count; $i++) {
            $headerLine = $matches[$i].Value.Trim()
            $sectionName = ($headerLine -replace '^-{35}\s+', '').Trim()
        
            $startIndex = $matches[$i].Index + $matches[$i].Length
            $endIndex = if ($i -lt $matches.Count - 1) {
                $matches[$i + 1].Index
            } else {
                $logText.Length
            }
        
            # Extract content between this section header and the next
            $sectionContent = $logText.Substring($startIndex, $endIndex - $startIndex).Trim()
        
            $sections += [PSCustomObject]@{
                File = $fileName
                Section = ($SectionName -replace '^-{35}\s+', '') -replace '\s*[-]+$', ''
                Content = $sectionContent
            }
        }
        
        # Output how many were found
        Write-host "Found $($sections.Count) sections."
        
        function Parse-TextVar {
            param (
                [Parameter(Mandatory=$true)]
                [string]$Input1,
        
                [Parameter(Mandatory=$false)]
                [string]$DelimiterPattern = "^!"
            )
        
            # Split the input text into lines
            $lines = $Input1 -split "`r?`n"
            
            # Initialize variables
            $sections = @()
            $currentSection = @()
            $index = 0
        
            # Loop through lines
            foreach ($line in $lines) {
                if ($line -match $DelimiterPattern) {
                    if ($currentSection.Count -gt 0) {
                        # Create a section object and clear the current section
                        $sections += [PSCustomObject]@{
                            Index   = $index++
                            Content = ($currentSection -join "`n").Trim()
                        }
                        $currentSection.Clear()
                    }
                }
                
                # Only add non-empty lines
                if ($line.Trim().Length -gt 0) {
                    $currentSection += $line
                }
            }
        
            # Add the final section if it exists
            if ($currentSection.Count -gt 0) {
                $sections += [PSCustomObject]@{
                    Index   = $index
                    Content = ($currentSection -join "`n").Trim()
                }
            }
        
            # Explicitly return the array of objects
            return ,$sections
        }
        
        #$ShowVersion = $sections | ?{ $_.section -imatch "show Version" } | select -ExpandProperty Content
        
        # Find show running-configuration section
            $showrunningconfiguration = $sections | ?{ $_.section -imatch "show running-configuration" } | select -ExpandProperty Content
            $parsedshowrunningconfiguration = Parse-TextVar $showrunningconfiguration
        
        # Running Configuration: Find the hostname
            $hostname = (($parsedshowrunningconfiguration | Where-Object { $_.Content -match "hostname" } | select -ExpandProperty content) -split "`r?`n" | ?{$_ -imatch "hostname" }) -replace "^hostname\s+", ""
        
        # Running Configuration: Find Switch OS version
            $SwitchOSVersion = [PSCustomObject]@{
                                    Hostname = $hostname
                                    Version = (($parsedshowrunningconfiguration |  Where-Object { $_.Content -match 'Version' }).Content -replace "! " -split 'Version ')[1]
                                  }
            $ShowVersionOut += $SwitchOSVersion

        # show interface
            Write-Host "    Gathering Show Interface...."
            $STSShowInterfaces=@()
            # Find show interface section
            $showinterface = $sections | ?{ $_.section -imatch "show interface"} | select -ExpandProperty Content
            $parsedshowinterfaceEthernet = Parse-TextVar $showinterface[0] -DelimiterPattern "(?i)^.*line protocol is.*$"

            # Running-Configuration: 

        ForEach($item in $parsedshowinterfaceEthernet){
            $STSShowInterfaces += [PSCustomObject]@{
                                  Hostname = $hostname
                                  Content  = $item.Content 
                                  }
        }
        $SwitchHostName=$hostname
        ForEach($STSShowInterface in $STSShowInterfaces){
            $Interface = $STSShowInterface.Content
            IF($Interface.Length -gt 0){
                    $InputStatistics = ((($Interface -split 'Input statistics\:[\r\n]')[1] -split 'Output statistics\:[\r\n]')[0] -replace '^[\r\n]' -replace '[\r\n]\s{5}',',' -split ',').trim()
                    $OutputStatistics = ((($Interface -split 'Output statistics\:[\r\n]')[0] -split 'Input statistics\:[\r\n]')[1] -replace '^[\r\n]' -replace '[\r\n]\s{5}',',' -split ',').trim()
                    $Interfaces+= [PSCustomObject] @{
                        HostName=$SwitchHostName
                        Interface = (($Interface -split '\sis\s')[0] -replace '[\r\n]{1,}').Trim()
                        InterfaceStatus = (($Interface -split '\,\sline\sprotocol\sis\s')[1] -split '[\r\n]')[0].trim()
                        Description = (($Interface -split 'Description\:\s')[1] -split '[\r\n]')[0]
                        MacAddress = (($Interface -split 'Current\saddress\sis\s')[1] -split '[\r\n]')[0].trim()
                        Pluggablemedia = ((($Interface -split 'Pluggable\smedia\s')[1] -split '[\r\n]')[0] -split ',\s')[0]
                        MediaType = ((($Interface -split 'Pluggable\smedia\s')[1] -split '[\r\n]')[0] -split '\sis\s')[1]
                        MTU = ((($Interface -split 'MTU\s')[1] -split '[\r\n]')[0] -split ',\s')[0]
                        LineSpeed = ((($Interface -split 'LineSpeed\s')[1] -split '[\r\n]')[0] -split ',\s')[0]
                        AutoNegotiation = (($Interface -split 'Auto-Negotiation\s')[1] -split '[\r\n]')[0]
                        ConfiguredFEC = ((($Interface -split 'Configured\sFEC\sis\s')[1] -split '[\r\n]')[0] -split ',\s')[0]
                        NegotiatedFEC = (($Interface -split 'Negotiated\sFEC\sis\s')[1] -split '[\r\n]')[0]
                        Flowcontrol = (($Interface -split 'Flowcontrol\s')[1] -split '[\r\n]')[0]
                        LastCleared = (($Interface -split 'Last\sclearing\sof\s\"show\sinterface\"\scounters\:\s')[1] -split '[\r\n]')[0]
                        # Input Statistics
                            InPackets = ($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'packets'}) -replace ' packets'
                            InOctets = ($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'octets'}) -replace ' octets'
                            In64bytepkts = ($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch '64-byte pkts' -and $_ -inotmatch 'over' }) -replace ' 64-byte pkts'
                            InOver64bytepkts = ($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'over 64-byte pkts'}) -replace ' over 64-byte pkts'
                            InOver127bytepkts = ($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'over 127-byte pkts'}) -replace ' over 127-byte pkts'
                            InOver255bytepkts = ($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'over 255-byte pkts'}) -replace ' over 255-byte pkts'
                            InOver511bytepkts = ($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'over 511-byte pkts'}) -replace ' over 511-byte pkts'
                            InOver1023bytepkts = ($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'over 1023-byte pkts'}) -replace ' over 1023-byte pkts'
                            InMulticasts = ($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'Multicasts'}) -replace ' Multicasts'
                            InBroadcasts = ($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'Broadcasts'}) -replace ' Broadcasts'
                            InUnicasts = ($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'Unicasts'}) -replace ' Unicasts'
                            InRunts = "YYEELLLLOOWW"+($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'runts'}) -replace ' runts'
                            InGiants = "YYEELLLLOOWW"+($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'giants'}) -replace ' giants'
                            InThrottles = "YYEELLLLOOWW"+($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'throttles'}) -replace ' throttles'
                            InCRC = "YYEELLLLOOWW"+($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'CRC'}) -replace ' CRC'
                            InOverrun = "YYEELLLLOOWW"+($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'overrun'}) -replace ' overrun'
                            InDiscarded = "YYEELLLLOOWW"+($InputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'discarded'}) -replace ' discarded'
                        # OutputStatistics
                            OutPackets = ($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'packets'}) -replace ' packets'
                            OutOctets = ($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'octets'}) -replace ' octets'
                            Out64bytepkts = ($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch '64-byte pkts' -and $_ -inotmatch 'over' }) -replace ' 64-byte pkts'
                            OutOver64bytepkts = ($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'over 64-byte pkts'}) -replace ' over 64-byte pkts'
                            OutOver127bytepkts = ($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'over 127-byte pkts'}) -replace ' over 127-byte pkts'
                            OutOver255bytepkts = ($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'over 255-byte pkts'}) -replace ' over 255-byte pkts'
                            OutOver511bytepkts = ($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'over 511-byte pkts'}) -replace ' over 511-byte pkts'
                            OutOver1023bytepkts = ($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'over 1023-byte pkts'}) -replace ' over 1023-byte pkts'
                            OutMulticasts = ($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'Multicasts'}) -replace ' Multicasts'
                            OutBroadcasts = ($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'Broadcasts'}) -replace ' Broadcasts'
                            OutUnicasts = ($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'Unicasts'}) -replace ' Unicasts'
                            OutRunts = "YYEELLLLOOWW"+($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'runts'}) -replace ' runts'
                            OutGiants = "YYEELLLLOOWW"+($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'giants'}) -replace ' giants'
                            OutThrottles = "YYEELLLLOOWW"+($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'throttles'}) -replace ' throttles'
                            OutCRC = "YYEELLLLOOWW"+($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'CRC'}) -replace ' CRC'
                            OutOverrun = "YYEELLLLOOWW"+($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'overrun'}) -replace ' overrun'
                            OutDiscarded = "YYEELLLLOOWW"+($OutputStatistics -split '[\r\n]' | Where-Object{$_ -imatch 'discarded'}) -replace ' discarded'                    
                        
                    }
            }
        }
          #$Interfaces | ft

        # Running-Configuration: 
            $AllTheparsedshowrunningconfiguration=@()
            ForEach($item in $parsedshowrunningconfiguration){
            $AllTheparsedshowrunningconfiguration += [PSCustomObject]@{
                                                Hostname     = $hostname
                                                ConfigItem   = ($item.Content -split "`r?`n")[1]
                                                ConfigAttrib = $item.Content 
                                                }
            }
            $RunningConfigOut += $AllTheparsedshowrunningconfiguration
        # show lldp neighbors
            $showlldpneighbors = $sections | ?{ $_.section -imatch "show lldp neighbors" } | select -ExpandProperty Content
            
            # split on carrige return line feed
            $data = $showlldpneighbors[0] -split "`r`n"

            # Remove header and separator lines
            $data = $data | Where-Object { $_ -notmatch "^(-|\s*Loc\s+PortID)" }
            $data = $data | ForEach-Object { $_.TrimEnd() }

            # Define column widths (based on visual spacing)

            $parsedshowlldpneighbors += foreach ($line in $data) {
                $pline = $line -split '(?:\s{2,}|\.{3})' | ForEach-Object { $_.Trim() }
                $LocPortID     = $pline[0]
                $RemHostName   = $pline[1]
                $RemPortID     = $pline[2]
                $RemChassisID  = $pline[3]
            

                [PSCustomObject]@{
                    'SwitchName'        = $hostname
                    'LocalPortID'     = $LocPortID
                    'RemoteHostName'  = $RemHostName
                    'RemotePortID'    = $RemPortID
                    'RemoteChassisID' = $RemChassisID
                }
            }
    }
        #$ShowVersionOut=$ShowVersionOut|Sort-Object UpTime -Unique
        $Name="Show Version"
        $html+='<H2 id="ShowTechsShowVersion">Show Version</H2>'
        $html+=$ShowVersionOut | ConvertTo-html -Fragment
        $html=$html `
            -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
            -replace '<td>YYEELLLLOOWW0</td>','<td>0</td>' `
            -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://www.dell.com/support/kbdoc/en-us/000126014/support-matrix-for-dell-emc-solutions-for-microsoft-azure-stack-hci' target='_blank'>Ref: Support Matrix for Dell EMC Solutions for Microsoft Azure Stack HCI</a></h5>"            
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""

    # Convert show lldp Neighbors for better visual
        $Name="Show lldp Neighbors"
        $html+='<H2 id="ShowlldpNeighbors">Show lldp Neighbors</H2>'
        $SWLLDPNtbl=""
        $SWLLDPNtbl = New-Object System.Data.DataTable "ShowlldpNeighbors"
            $SWLLDPNtbl.Columns.add((New-Object System.Data.DataColumn("RemotePortID")))
            ForEach ($a in ($parsedshowlldpneighbors.SwitchName | Sort-Object -Unique)){
                $SWLLDPNtbl.Columns.Add((New-Object System.Data.DataColumn([string]$a)))}
            $a=$null
            ForEach($b in ($parsedshowlldpneighbors | Sort-Object RemotePortID )){
                IF($b.LocalPortID.length -gt 2 -and $b.LocalPortID -notmatch 'System.__ComObject'){
                    if ($b.RemotePortID -ne $a) {
                        $a=$b.RemotePortID
                        if ($Null -ne $a) {$SWLLDPNtbl.rows.add($row)}
                        $row=$SWLLDPNtbl.NewRow()
                        $row["RemotePortID"]=$b.RemotePortID
                    }
                    $row["$($b.SwitchName)"] = [string]($b.LocalPortID.PSObject.BaseObject -replace '[\r\n]{2,}','&lt;br&gt;')
                }
            }
        $ShowlldpNeighborsOut = $SWLLDPNtbl | Sort-Object RemotePortID | Where-Object{$_ -notmatch 'System.__ComObject'} | Sort-Object ConfigItem | Select-object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors
        $html+=$ShowlldpNeighborsOut | ConvertTo-html -Fragment
        $html=$html `
            -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
            -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $Send2Html=""
        $html=""
        $Name="" 

        $Name="Show Interface"
        $html+='<H2 id="ShowTechsShowInterface">Show Interface</H2>'
        $html+=$Interfaces| Sort-Object Interface| ConvertTo-html -Fragment
        $html=$html `
            -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
            -replace '<td>System.Object\[\]</td>','<td></td>' `
            -replace '<td>YYEELLLLOOWW0</td>','<td>0</td>' `
            -replace '<td>YYEELLLLOOWW</td>','<td></td>' `
            -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""

        $Name="Show Running Configuration"
        $html+='<H2 id="ShowRunningConfiguration">Show Running Configuration</H2>'
        $SRCtbl = New-Object System.Data.DataTable "ShowRuningConfig"
        $SRCtbl.Columns.add((New-Object System.Data.DataColumn("ConfigItem")))
        ForEach ($a in ($RunningConfigOut.HostName | Sort-Object -Unique)){
            $SRCtbl.Columns.Add((New-Object System.Data.DataColumn([string]$a)))}
        $a=$null
        ForEach($b in ($RunningConfigOut | Sort-Object ConfigItem )){
            IF($b.ConfigItem.length -gt 2 -and $b.ConfigItem -notmatch 'System.__ComObject'){
                if ($b.ConfigItem -ne $a) {
                    $a=$b.ConfigItem
                    if ($Null -ne $a) {$SRCtbl.rows.add($row)}
                    $row=$SRCtbl.NewRow()
                    $row["ConfigItem"]=$b.ConfigItem
                }
                $row["$($b.HostName)"] = [string]($b.ConfigAttrib.PSObject.BaseObject -replace '[\r\n]','&lt;br&gt;')
            }
        }
        $Send2Html=$SRCtbl | Where-Object{$_ -notmatch 'System.__ComObject'} | Sort-Object ConfigItem | Select-object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors
        $html+=$Send2Html | ConvertTo-html -Fragment
        $html=$html `
            -replace '&gt;','>' -replace '&lt;','<' -replace '&quot;','"'`
            -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
            -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $Send2Html=""
        $html=""
        $Name=""

}
#endregion End of Process STS
$RunDate=Get-Date
$dstartSDDC=Get-Date
If($ProcessSDDC -ieq "y"){
#[System.Threading.Thread]::CurrentThread.Priority='Highest'
  #(Get-Process -Id $pid).PriorityClass='Normal'
  Write-Host "    Gathering All XML file data..."
  $SDDCFiles=@{}
  gci $SDDCPath -Filter "*.xml" | %{$SDDCFiles.Add($_.basename,(Import-Clixml -Path $_.fullname))}
  $ClusterNodes=$SDDCFiles."GetClusterNode" |`
  Sort-Object Name | Select-Object Name,Model,SerialNumber,State,StatusInformation,Id
  Foreach ($Name in ($ClusterNodes.Name)) {
      gci "$SDDCPath\Node_$Name" -Filter "*.xml" | %{$SDDCFiles.Add("$Name$($_.basename)",(Import-Clixml -Path $_.fullname))}
  }
 #<#
 Write-Host "Creating Recent Cluster Events runspace"
 $Inputs3  = New-Object 'System.Management.Automation.PSDataCollection[PSObject]'
 $Outputs3 = New-Object 'System.Management.Automation.PSDataCollection[PSObject]'
 $initialSessionState3 = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
 $addMessageSessionStateFunction=@()
 $addMessageSessionStateFunction += New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'add-TableData', $(Get-Content Function:\add-TableData -ErrorAction Stop)
 $addMessageSessionStateFunction += New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Convert-BytesToSize', $(Get-Content Function:\Convert-BytesToSize -ErrorAction Stop)
 $addMessageSessionStateFunction += New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Set-ResultsSummary', $(Get-Content Function:\Set-ResultsSummary -ErrorAction Stop)
 $addMessageSessionStateFunction | % {$initialSessionState3.Commands.Add($_)}

 $Runspace3 = [runspacefactory]::CreateRunspace($initialSessionState3)
 $PowerShell3 = [powershell]::Create()
 $PowerShell3.Runspace = $Runspace3
 $Runspace3.Open()
 $Runspace3.SessionStateProxy.SetVariable('ClusterNodes',$ClusterNodes)
 $Runspace3.SessionStateProxy.SetVariable('SDDCPath',$SDDCPath)
 $PowerShell3.AddScript({
# Cluster Events
    #xml events ([xml]$myevents.Objects.Object[4].'#text').event.System.eventid.'#text'
    $dstart2=Get-Date
    $Name="Recent Cluster Events"
    #Write-Host "    Gathering $Name..."
    $ClusLogFiles=Get-ChildItem -Path $SDDCPath -Filter "*cluster.log"
    #Write-Host "        Checking Recent Cluster Events"
    $LogEvents=@()
    #$LogEvents=$ClusLogFiles | %{"Node $((Split-Path $_.Directory -Leaf).Replace('Node_',''))" | Select-Object @{L="Line";E={$_}};" " | Select-Object @{L="Line";E={$_}};Get-Content $_.FullName -tail 100000 | Select-String -SimpleMatch "NetftTwoFifth" | Select-Object Line };$dstop2=Get-Date
    $LogEvents=$ClusLogFiles | %{"BBLLAACCKK";"BBLLAACCKK";"BBLLAACCKK";"Node $(($_.BaseName) -ireplace '_cluster','')" ;(Get-Content $_.FullName -tail 100000 | Select-String -SimpleMatch "NetftTwoFifth" | Select-Object Line).Line };$dstop2=Get-Date
    #Write-Host "Total time parsing cluster log $(($dstop2-$dstart2).totalseconds) secs"

        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$LogEvents
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=YKz%2F0%2BLMBvGU97hc0sH%2FVf0%2F1bi9gyr6vb6r2sjxbHg%3D&tn=CluChkStorageSpacesDriverDiagnosticEvents'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
    $html+='<H2 id="RecentClusterEvents">Recent Cluster Events</H2>'
    $html+="<h5><b>Key:</b></h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;These events indicate a possible node communication issue</h5>"
    $html+=""
    If ($ClusterNodes.Count*5 -ge $LogEvents.count) { 
    #Write-Host "            No recent cluster events found" -ForegroundColor Yellow;
    $html+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;No Recent Cluster Events found</h5>' } else {$LogEvents=@('YYEELLLLOOWWEvents Found!!')+$LogEvents
    $html+=$LogEvents | Select-Object @{L="Text";E={$_}} | ConvertTo-Html -Fragment}
    $html=$html `
            -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
            -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'`
            -replace '<td>BBLLAACCKK','<td style="background-color: #000000">' 
    $ResultsSummary+=Set-ResultsSummary -name $name -html $html
    $htmlout+=$html
    $ResultsSummary | Export-Clixml -Path "$env:temp\ClusterResultSummary2.xml" -Confirm:$false -Force
    $htmlout
    $html=""
    $Name=""
 }) | out-null
 $Job3 = $PowerShell3.BeginInvoke($Inputs3,$Outputs3)
 #>


}



 #Write-Host "*******************************************************************************"
 #Write-Host "*                                                                             *"
 #Write-Host "*                         HCI Performance Report                              *"
 #Write-Host "*                                                                             *"
 #Write-Host "*******************************************************************************"


IF($SDDCPerf -imatch "YES"){
 Write-Host "Creating HCI Performance runspace"
 $Inputs2  = New-Object 'System.Management.Automation.PSDataCollection[PSObject]'
 $Outputs2 = New-Object 'System.Management.Automation.PSDataCollection[PSObject]'
 $initialSessionState2 = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
 $addMessageSessionStateFunction=@()
 $addMessageSessionStateFunction += New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'add-TableData', $(Get-Content Function:\add-TableData -ErrorAction Stop)
 $addMessageSessionStateFunction += New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Convert-BytesToSize', $(Get-Content Function:\Convert-BytesToSize -ErrorAction Stop)
 $addMessageSessionStateFunction += New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Set-ResultsSummary', $(Get-Content Function:\Set-ResultsSummary -ErrorAction Stop)
 $addMessageSessionStateFunction | % {$initialSessionState2.Commands.Add($_)}

 $Runspace2 = [runspacefactory]::CreateRunspace($initialSessionState2)
 $PowerShell2 = [powershell]::Create()
 $PowerShell2.Runspace = $Runspace2
 $Runspace2.Open()
 $Runspace2.SessionStateProxy.SetVariable('ClusterNodes',$ClusterNodes)
 $Runspace2.SessionStateProxy.SetVariable('SysInfo',$SysInfo)
 $Runspace2.SessionStateProxy.SetVariable('OSVersionNodes',$OSVersionNodes)
 #$Runspace2.SessionStateProxy.SetVariable('SupportMatrixtableData',$SupportMatrixtableData)
 $Runspace2.SessionStateProxy.SetVariable('CluChkReportLoc',$CluChkReportLoc)
 $Runspace2.SessionStateProxy.SetVariable('SDDCFiles',$SDDCFiles)
 $Runspace2.SessionStateProxy.SetVariable('SDDCPath',$SDDCPath)
 $Runspace2.SessionStateProxy.SetVariable('SMRevHistLatest',$SMRevHistLatest)
 $Runspace2.SessionStateProxy.SetVariable('CluChkVer',$CluChkVer)
 $Runspace2.SessionStateProxy.SetVariable('htmlStyle',$htmlStyle)
 $Runspace2.SessionStateProxy.SetVariable('RunDate',$RunDate)
 $Runspace2.SessionStateProxy.SetVariable('DTString',$DTString)
 $PowerShell2.AddScript({
    $htmloutPerf=@()
    $htmloutPerfReport=@() 
    $PerfResultsSummary=@()
    $htmloutPerf+='<H1 id="PerformanceReport">S2D/HCI Performance</H1>'
#S2D Disk Performance Data
    $Name="S2D Disk Performance Data"
    Write-Host "    Gathering $Name..."
    # SDDC KPI's 
    # https://docs.microsoft.com/en-us/windows-server/storage/storage-spaces/performance-history-scripting
    # Get log subset
        $S2DDPD=Get-ChildItem -Path $SDDCPath -Filter "GetCounters.blg" -Recurse -Depth 1
        If($S2DDPD.Length -gt 0){
        $d2start=Get-Date
        #relog $S2DDPD.fullname -c "\Cluster Disk Counters(*)\ExceededLatencyLimit/sec" "\Cluster Disk Counters(*)\ExceededLatencyLimit" "\Cluster Disk Counters(*)\IO (> 10,000ms)/sec" "\Cluster Disk Counters(*)\IO (<= 10,000ms)/sec" "\Cluster Disk Counters(*)\IO (<= 1000ms)/sec" "\Cluster Disk Counters(*)\IO (<= 100ms)/sec" "\Cluster Disk Counters(*)\IO (<= 100ms)/sec" "\Cluster Disk Counters(*)\IO (<= 10ms)/sec" "\Cluster Disk Counters(*)\IO (<= 5ms)/sec" "\Cluster Disk Counters(*)\IO (<= 1ms)/sec" "\Cluster Storage Cache Stores(*)\Cache Usage %" "\Cluster Storage Hybrid Disks(*)\Cache Miss Reads/sec" "\Cluster Storage Hybrid Disks(*)\Disk Transfers/sec" -o "$SDDCPath\Counters0.blg" | Out-Null
        #relog $S2DDPD.fullname -y -f csv -c "\Cluster Disk Counters(*)\ExceededLatencyLimit" "\Cluster Disk Counters(*)\IO (> 10,000ms)/sec" "\Cluster Disk Counters(*)\IO (<= 10,000ms)/sec" "\Cluster Disk Counters(*)\IO (<= 1000ms)/sec" "\Cluster Disk Counters(*)\IO (<= 100ms)/sec" "\Cluster Disk Counters(*)\IO (<= 10ms)/sec" "\Cluster Disk Counters(*)\IO (<= 5ms)/sec" "\Cluster Disk Counters(*)\IO (<= 1ms)/sec" "\Cluster Storage Hybrid Disks(*)\Cache Miss Reads/sec" -o "$SDDCPath\Counters0.csv" | Out-Null
        relog $S2DDPD.fullname -y -f csv -c "\Cluster Disk Counters(*)\ExceededLatencyLimit" "\Cluster Disk Counters(*)\IO (> 10,000ms)/sec" "\Cluster Disk Counters(*)\IO (<= 10,000ms)/sec" "\Cluster Disk Counters(*)\IO (<= 1000ms)/sec" "\Cluster Disk Counters(*)\IO (<= 100ms)/sec" "\Cluster Storage Hybrid Disks(*)\Cache Miss Reads/sec" -o "$SDDCPath\Counters0.csv" | Out-Null
        $d2stop=Get-Date
        #Write-Host "Relog total time $(($d2stop-$d2start).TotalMilliseconds)"

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

public class Counters
{
    public class ResultClass
    {
         public string Node {get; set;}
         public string CNId {get; set;}
         public string Object {get; set;}
         public string Counter {get; set;}
         public string Avg {get; set;}
         public string Min {get; set;}
         public string Max {get; set;}
    }
    public static dynamic[] Import(string filePath, string[] NodeIds)
    {
        Dictionary<string, string> clusterIds = new Dictionary<string, string>();
        foreach (string NodeId in NodeIds)
        {
            clusterIds[(NodeId.Split(':')[0].ToLower())] = NodeId.Split(':')[1];
        }

        // Create a list to hold the data rows
        List<Dictionary<string, decimal>> dataRows = new List<Dictionary<string, decimal>>();
        string[] headers;

        // Read the CSV file
        using (StreamReader reader = new StreamReader(@filePath))
        {
            // Get the column headers
            string headerLine = reader.ReadLine();
            string pattern = "(?:^|,)(\"(?:[^\"])*\"|[^,]*)";
            headers = Regex.Split(headerLine, pattern).Skip(2).ToArray();
            headers = headers.Where(x => x.Length > 3).ToArray();
            for (int i = 0; i < headers.Length; i++)
            {
                headers[i] = headers[i].Replace("\"","");
            }
            
            // Process the data rows
            string dataLine;
            while ((dataLine = reader.ReadLine()) != null)
            {
                dataLine=dataLine.Replace("\"","");
                string[] fields = dataLine.Split(',').Skip(1).ToArray();
                // Create a dictionary to hold the row data
                Dictionary<string, decimal> rowData = new Dictionary<string, decimal>();
                for (int i = 0; i < fields.Length; i++)
                {
                    decimal ddata;
                    if (Decimal.TryParse(fields[i], out ddata))
                    {
                        rowData[headers[i]] = ddata;
                    }
                    else
                    {
                        rowData[headers[i]] = 0;
                    }
                }

                // Add the row data to the list
                dataRows.Add(rowData);
            }
        }
        //Console.WriteLine("Got the data...");
        // Calculate Min, Avg, and Max for each number field
        Dictionary<string, string> minValues = new Dictionary<string, string>();
        Dictionary<string, string> avgValues = new Dictionary<string, string>();
        Dictionary<string, string> maxValues = new Dictionary<string, string>();
        Dictionary<string, string> Node = new Dictionary<string, string>();
        Dictionary<string, string> Object = new Dictionary<string, string>();
        Dictionary<string, string> Counter = new Dictionary<string, string>();
        List<decimal> row = new List<decimal>();
        foreach (string header in headers)
        {
            row.Clear();
            try {row = dataRows.Select(drow => drow[header]).ToList();} 
            catch {}
            if (row != null)
            {
                try
                {
                    minValues[header] = Regex.Replace(Math.Round(row.Min(),2).ToString(),@"\b0\b","");
                    avgValues[header] = Regex.Replace(Math.Round(row.Average(),2).ToString(),@"\b0\b","");
                    maxValues[header] = Regex.Replace(Math.Round(row.Max(),2).ToString(),@"\b0\b","");

                }
                catch {}
                try 
                {
                    Node[header] = header.Split('\\')[2];
                    Object[header] = header.Split('\\')[3];
                    Counter[header] = header.Split('\\')[4];
                }
                catch
                {
                    try {
                        minValues.Remove(header);
                        avgValues.Remove(header);
                        maxValues.Remove(header);
                        Node.Remove(header);
                        Object.Remove(header);
                        Counter.Remove(header);
                        }
                    catch {}
                }
                if ((minValues[header] == "" && avgValues[header] == "" && maxValues[header] == "") || header.Contains(@"_total") && !header.Contains("Cache Miss Reads"))
                {
                    try {
                        minValues.Remove(header);
                        avgValues.Remove(header);
                        maxValues.Remove(header);
                        Node.Remove(header);
                        Object.Remove(header);
                        Counter.Remove(header);
                        }
                    catch {}
                }
                if (header.Contains("ExceededLatencyLimit") && avgValues.ContainsKey(header)) 
                {
                    if (minValues[header] != "")
                    {
                        minValues[header] = "RREEDD" + minValues[header];
                    }
                    if (maxValues[header] != "")
                    {
                        maxValues[header] = "RREEDD" + maxValues[header];
                    }
                    if (avgValues[header] != "")
                    {
                        avgValues[header] = "RREEDD" + avgValues[header];
                    }
                }
                if ((header.Contains("Cache Miss Reads") || header.Contains("10,000ms")) && avgValues.ContainsKey(header)) 
                {
                    if (minValues[header] != "")
                    {
                        minValues[header] = "YYEELLLLOOWW" + minValues[header];
                    }
                    if (maxValues[header] != "")
                    {
                        maxValues[header] = "YYEELLLLOOWW" + maxValues[header];
                    }
                    if (avgValues[header] != "")
                    {
                        avgValues[header] = "YYEELLLLOOWW" + avgValues[header];
                    }
                }

            }
        }

        // Create a dynamic object with the desired properties
        ResultClass[] result = new ResultClass[]{};
        foreach (string header in headers)
        {
            try
            {
                result = result.Append(new ResultClass {Node = Node[header], CNId = clusterIds[Node[header]], Object = Object[header], Counter = Counter[header], Avg = avgValues[header], Min = minValues[header], Max = maxValues[header]}).ToArray();
            }
            catch {}
        }
        return result;
    }
}
'@ -ReferencedAssemblies "Microsoft.Csharp"
        $S2DStoragePerfData=@()
        $ClusterNodes=Get-ChildItem -Path $SDDCPath -Filter "GetClusterNode.XML" -Recurse -Depth 1 | import-clixml |`
            Sort-Object Name | Select-Object Name,Model,SerialNumber,State,StatusInformation,Id
        $S2DStoragePerfData=[Counters]::Import("$SDDCPath\Counters0.csv", @($ClusterNodes | %{"$($_.Name):$($_.Id)"}))
        Remove-Item -Path "$SDDCPath\Counters0.csv"
        $keepthese= $S2DStoragePerfData
        #Write-Host "Import and group objects total time $(((Get-Date)-$d2stop).TotalMilliseconds)"
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$keepthese|Select-Object -Property `
                @{L='Node';E={[string]$_.Node}},
                @{L='CNId';E={[string]$_.CNId}},
                @{L='Object';E={[string]$_.Object}},
                @{L='Counter';E={[string]$_.Counter}},
                @{L='Avg';E={[string]$_.Avg}},
                @{L='Min';E={[string]$_.Min}},
                @{L='Max';E={[string]$_.Max}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=TzWN1pRbVflAvGc5g9PpgtDzp1sNFWEwRQ9hXuhT3f8%3D&tn=CluChkS2DDiskPerformanceData'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
    #$div=Get-Random 10000
    #$html+='<H2 id="S2DDiskPerformanceData"><a onclick="toggle(''d{0}'',this)">&#8649;&nbsp;</a>S2D Disk Performance Data</H2>' -f $div
    #$html+="<div id='d{0}' style='display:block;'><h5>&nbsp;&nbsp;&nbsp;&nbsp;-Zero values and anything below 100ms removed to make it easily spot problems.</h5>" -f $div
    $html+='<H2 id="S2DDiskPerformanceData">S2D Disk Performance Data</H2>'
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Zero values and anything below 100ms removed to make it easily spot problems.</h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-io >= 1000/sec exceeds SpacePort HwTimeout value of 1000 and indicates a failing hard drive.</h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Cache Miss Reads/sec: If too many reads are missing the cache, it may be undersized and you should consider adding cache drives to expand your cache.</h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://docs.microsoft.com/en-us/windows-server/storage/storage-spaces/understand-the-cache#sizing-the-cache' target='_blank'>Ref: Sizing the cache</a></h5>"
    $html+=$keepthese |Sort-Object Node,Object,Counter| ConvertTo-html -Fragment
    $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
    #$html+="</div>"
    $PerfResultsSummary+=Set-ResultsSummary -name $name -html $html
    $htmloutPerf+=$html
    $html=""
    $Name=""

    $dstop=Get-Date
#Write-Host "Total time $(($dstop-$dstart).TotalMilliseconds)"


#region Cluster Disk Counters
    $Name="Cluster Disk Counters"
    Write-Host "    Gathering $Name..."
    # Clean up BLG
    Remove-Item -Path "$SDDCPath\Counters1.csv" -ErrorAction SilentlyContinue
    # Get log subset
    $S2DDPD=Get-ChildItem -Path $SDDCPath -Filter "GetCounters.blg" -Recurse -Depth 1
    If($S2DDPD.Length -gt 0){
        relog $S2DDPD.fullname -y -f csv -c "\cluster disk counters(_total)\local: read/sec" "\cluster disk counters(_total)\local: writes/sec" "\cluster disk counters(_total)\remote: read/sec" "\cluster disk counters(_total)\remote: writes/sec" -o "$SDDCPath\Counters1.csv" | Out-Null
    # Imports BLG data
        $Data=@()
        $Data=Import-Csv -Path "$SDDCPath\Counters1.csv" 
    # Group-Objects and Averages the data points
        #$CookedData=@()
        #$CookedData_total=$Data.CounterSamples | select @{L="Node";E={($_.Path -split '\\')[2]}},@{L="Object";E={($_.Path -split '\\')[3]}},@{L="Counter";E={($_.Path -split '\\')[4]}},InstanceName,@{L="Value";E={[math]::round($_.CookedValue,2)}}
        #$CookedData_total_bynode=$CookedData_total | Group-Object -Property Node
        #$nodes = $CookedData_total | Select-Object Node -Unique
        $results=@()
        $n=$false;
        $results = foreach ($entry in ($data | gm | ? MemberType -eq 'NoteProperty').Name) {
            if ($entry.Contains("local: read/sec")) {$local_read_sum=$null;$data."$entry" | % {[float]$local_read_sum+=$_}}
            if ($entry.Contains("local: writes/sec")) {$local_writes_sum=$null;$data."$entry" | % {[float]$local_writes_sum+=$_}}
            if ($entry.Contains("remote: read/sec")) {$remote_read_sum=$null;$data."$entry" | % {[float]$remote_read_sum+=$_}}
            if ($entry.Contains("remote: writes/sec")) {$n=$true;$remote_writes_sum=$null;$data."$entry" | % {[float]$remote_writes_sum+=$_}}
            if ($n) {
                $node=$entry.split('\\')[2]
                IF ($local_read_sum -gt 0 -and $remote_read_sum -gt 0){
                    if ($local_read_sum -gt $remote_read_sum) {
                        $LocalGreaterValue = $local_read_sum
                        $ReadGreater= "Local"
                    } Else{
                        $LocalGreaterValue = $remote_read_sum
                        $ReadGreater= "Remote"
                    }
                } Else{$ReadGreater= "Equal"}

                IF ($local_writes_sum -gt 0 -and $remote_writes_sum -gt 0){
                    if ($local_writes_sum -gt $remote_writes_sum) {
                        $WriteGreaterValue = $local_writes_sum
                        $writesGreater= "Local"
                    } Else{ 
                        $WriteGreaterValue = $remote_writes_sum
                        $writesGreater= "Remote"
                    }
                } Else{$writesGreater= "Equal"}
                    
                $ReadPercentDiff = [math]::Round((($LocalGreaterValue / ($local_read_sum + $remote_read_sum)))  * 100, 2)
                $WritesPercentDiff = [math]::Round((($WriteGreaterValue / ($local_writes_sum + $remote_writes_sum)))  * 100, 2)
            
                $n=$false
                [PSCustomObject]@{
                    Node = $node
                    LocalReadPerSecSum =  [math]::Round($local_read_sum,2)
                    RemoteReadPerSecSum =  [math]::Round($remote_read_sum,2)
                    ReadGreater = $ReadGreater
                    'Read%Diff' = $ReadPercentDiff
                    LocalwritesPerSecSum =  [math]::Round($local_writes_sum,2)
                    RemotewritesPerSecSum =  [math]::Round($remote_writes_sum,2)
                    writesGreater = $writesGreater
                    'Writes%Diff' = $writesPercentDiff
                }
            }
        }
               
        $dstop2=Get-Date
        #Write-host "Total Time is $(($dstop2-$dstart).totalmilliseconds)" 
        #$results | Format-Table -AutoSize

      # Clean up BLG
        Remove-Item -Path "$SDDCPath\Counters1.csv"

    }
   
  # HTML Report
    $html+='<H2 id="ClusterDiskCounters">Cluster Disk Counters</H2>'
    $html+=$results|Sort-Object Node, -Descending| ConvertTo-html -Fragment
    $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
    $PerfResultsSummary+=Set-ResultsSummary -name $name -html $html
    $htmloutPerf+=$html
    $html=""
    $Name=""
#endregion Cluster Disk Counters

    #region CPU, I see you
        $Name="Sample 1: CPU, I see you!"
        Write-Host "    Gathering $Name..."
        # Get input
            $CPUIseeyou=""
            $CPUIseeyou=Get-ChildItem -Path $SDDCPath -Filter "CPUIseeyou.xml" -Recurse -Depth 1 | Import-Clixml
        Function Format-Hours {
            Param (
                $RawValue
            )
            # Weekly timeframe has frequency 15 minutes = 4 points per hour
            [Math]::Round($RawValue/4)
        }

        Function Format-Percent {
            Param (
                $RawValue
            )
            [String][Math]::Round($RawValue) + " " + "%"
        }
        $CPUIseeyouOutput=@()
        $NodesFound=($CPUIseeyou | Group-Object -Property ObjectDescription -NoElement).name -replace "ClusterNode "| Sort-Object -Unique
        ForEach($Node in $NodesFound){
            $Measure = $CPUIseeyou | Where-Object {$_.ObjectDescription -imatch $Node}| Measure-Object -Property Value -Minimum -Maximum -Average
            $Min = $Measure.Minimum
            $Max = $Measure.Maximum
            $Avg = $Measure.Average

            $CPUIseeyouOutput += [PsCustomObject]@{
                "ClusterNode"    = $Node
                "MinCpuObserved" = Format-Percent $Min
                "MaxCpuObserved" = Format-Percent $Max
                "AvgCpuObserved" = Format-Percent $Avg
                "HrsOver25%"     = Format-Hours ($CPUIseeyou | Where-Object {$_.ObjectDescription -imatch $Node -and $_.Value -Gt "25"}).Length
                "HrsOver50%"     = Format-Hours ($CPUIseeyou | Where-Object {$_.ObjectDescription -imatch $Node -and $_.Value -Gt "50"}).Length
                "HrsOver75%"     = Format-Hours ($CPUIseeyou | Where-Object {$_.ObjectDescription -imatch $Node -and $_.Value -Gt "75"}).Length
            }
        }

        #$Output | Sort-Object ClusterNode | Format-Table        
            # HTML Report
        $html+='<H2 id="Sample1CPUIseeyou">Sample 1: CPU, I see you</H2>'
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;This sample uses the ClusterNode.Cpu.Usage series from the LastWeek timeframe to show the maximum (high water mark), </h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;  minimum, and average CPU usage for every server in the cluster. It also does simple quartile analysis to show how </h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;  many hours CPU usage was over 25%, 50%, and 75% in the last 8 days.</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://learn.microsoft.com/en-us/windows-server/storage/storage-spaces/performance-history-scripting#sample-1-cpu-i-see-you' target='_blank'>Ref: Sample 1: CPU, I see you!</a></h5>"
        $html+=$CPUIseeyouOutput |Sort-Object ClusterNode | ConvertTo-html -Fragment
        $html=$html `
             -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
             -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        $PerfResultsSummary+=Set-ResultsSummary -name $Name -html $html
        $htmloutPerf+=$html
        $html=""
        $Name=""
    #endregion CPU, I see you

    #region Fire, fire, latency outlier
        $Name="Sample 2: Fire, fire, latency outlier"
        Write-Host "    Gathering $Name..."
        # Get input
            $Firefirelatencyoutlier=""
            $FirefirelatencyoutlierOutput=""
            $Firefirelatencyoutlier=Get-ChildItem -Path $SDDCPath -Filter "latencyoutlier.xml" -Recurse -Depth 1 | Import-Clixml
            $FirefirelatencyoutlierOutput=$Firefirelatencyoutlier | Select PSComputerName,FriendlyName,SerialNumber,MediaType,@{L="AvgLatencyPopulation";E={$_.AvgLatencyPopulation.replace("$([char]195)$([char]381)$([char]194)$([char]188)s", "$([char]206)$([char]188)s")}},@{L="AvgLatencyThisHDD";E={$_.AvgLatencyThisHDD.replace("$([char]195)$([char]381)$([char]194)$([char]188)s", "$([char]206)$([char]188)s")}},@{L="Deviation";E={
                $pattern = '[-+]?\d+(\.\d+)?' # Matches any decimal number
                $match = [regex]::Matches($_.Deviation, $pattern)
                $number = [double]$match.Value
                if ($number -gt 3 -or $number -lt -3) {
                    "RREEDD"+$_.Deviation
                }Else{$_.Deviation}}}
       
            #$FirefirelatencyoutlierOutput | Sort-Object PSComputerName | FT
        # HTML Report
            $html+='<H2 id="Sample2Firefirelatencyoutlier">Sample 2: Fire, fire latency outlier</H2>'
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;This sample uses the PhysicalDisk.Latency.Average series from the LastWeek timeframe to look for statistical outliers,  </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;  defined as drives with an hourly average latency exceeding +3&#x3C3 (three standard deviations) above the population average. </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://learn.microsoft.com/en-us/windows-server/storage/storage-spaces/performance-history-scripting#sample-2-fire-fire-latency-outlier' target='_blank'>Ref: Sample 2: Fire, fire, latency outlier</a></h5>"
            $html+=$FirefirelatencyoutlierOutput |Sort-Object PSComputerName | ConvertTo-html -Fragment
            $html=$html `
                 -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                 -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'`
                 -replace "$([char]956)s", '&mu;s'`
                 -replace '&#206;&#188;s', '&mu;s'
            $PerfResultsSummary+=Set-ResultsSummary -name $Name -html $html
            $htmloutPerf+=$html
            $html=""
            $Name=""
    #endregion Fire, fire, latency outlier

    #region Noisy neighbor? That's write!
        $Name="Sample 3: Noisy neighbor? That's write!"
        Write-Host "    Gathering $Name..."
        # Get input
            $Noisyneighbor=""
            $NoisyneighborOutput=""
            $Noisyneighbor=Get-ChildItem -Path $SDDCPath -Filter "Noisyneighbor.xml" -Recurse -Depth 1 | Import-Clixml
            $NoisyneighborOutput = $Noisyneighbor | Sort-Object PSComputerName,IopsTotal | Select-Object PSComputerName, VM, IopsTotal, IopsRead, IopsWrite
            #$NoisyneighborOutput | Sort-Object PSComputerName | FT
        # HTML Report
            $html+='<H2 id="Sample3NoisyneighborThatswrite">Sample 3: Noisy neighbor</H2>'
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;This sample uses the VHD.Iops.Total series from the MostRecent timeframe to identify the busiest (some might say noisiest) </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;  virtual machines consuming the most storage IOPS, across every host in the cluster, and show the read/write breakdown of their activity. </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://learn.microsoft.com/en-us/windows-server/storage/storage-spaces/performance-history-scripting#sample-3-noisy-neighbor-thats-write' target='_blank'>Ref: Sample 3: Noisy neighbor? That's write!</a></h5>"
            $html+=$NoisyneighborOutput |Sort-Object PSComputerName | ConvertTo-html -Fragment
            $html=$html `
                 -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                 -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'`
                 -replace "&#206;&#188;", '&mu;s'
            $PerfResultsSummary+=Set-ResultsSummary -name $Name -html $html
            $htmloutPerf+=$html
            $html=""
            $Name=""
    #endregion Noisy neighbor? That's write!

    #region As they say, "25-gig is the new 10-gig
        $Name="Sample 4: As they say, 25-gig is the new 10-gig"
        Write-Host "    Gathering $Name..."
        # Get input
            $25gigisthenew10gig=""
            $25gigisthenew10gigOutput=""
            $25gigisthenew10gig=Get-ChildItem -Path $SDDCPath -Filter "25gigisthenew10gig.xml" -Recurse -Depth 1 | Import-Clixml
            $25gigisthenew10gigOutput = $25gigisthenew10gig | Sort-Object PSComputerName,NetAdapter | select PSComputerName,NetAdapter,LinkSpeed,MaxInbound,MaxOutbound,@{L="Saturated";E={If($_.Saturated -imatch "true" -and $_.NetAdapter -inotmatch "NDIS"){"RREEDD"+$_.Saturated}Else{$_.Saturated}}}
            #$25gigisthenew10gigOutput | Sort-Object PSComputerName | FT
        # HTML Report
            $html+='<H2 id="Sample4Astheysay25gigisthenew10gig">Sample 4: As they say, 25-gig is the new 10-gig</H2>'
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;This sample uses the NetAdapter.Bandwidth.Total series from the LastDay timeframe to look for signs of network saturation, defined as >90% </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;  of theoretical maximum bandwidth. For every network adapter in the cluster, it compares the highest observed bandwidth usage in the last </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;  of theoretical day to its stated link speed. </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://learn.microsoft.com/en-us/windows-server/storage/storage-spaces/performance-history-scripting#sample-4-as-they-say-25-gig-is-the-new-10-gig' target='_blank'>Ref: Sample 4: As they say, 25-gig is the new 10-gig</a></h5>"
            $html+=$25gigisthenew10gigOutput |Sort-Object PSComputerName | ConvertTo-html -Fragment
            $html=$html `
                 -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                 -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'`
                 -replace "$([char]206)$([char]188)s", '&mu;s'
            $PerfResultsSummary+=Set-ResultsSummary -name $Name -html $html
            $htmloutPerf+=$html
            $html=""
            $Name=""
    #endregion As they say, "25-gig is the new 10-gig

    #region Make storage trendy again!
        $Name="Sample 5: Make storage trendy again!"
        Write-Host "    Gathering $Name..."
        # Get input
            $trendyagain=""
            $trendyagainOutput=@()
            $trendyagain=Get-ChildItem -Path $SDDCPath -Filter "trendyagain.xml" -Recurse -Depth 1 | Import-Clixml
            Function Format-Bytes {
                Param (
                    $RawValue
                )
                $i = 0 ; $Labels = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
                Do { $RawValue /= 1024 ; $i++ } While ( $RawValue -Gt 1024 )
                # Return
                [String][Math]::Round($RawValue) + " " + $Labels[$i]
            }

            Function Format-Trend {
                Param (
                    $RawValue
                )
                If ($RawValue -Eq 0) {
                    "0"
                }
                Else {
                    If ($RawValue -Gt 0) {
                        $Sign = "+"
                    }
                    Else {
                        $Sign = "-"
                    }
                    # Return
                    #Write-Host $RawValue
                    #$Sign + $(Format-Bytes [Math]::Abs($RawValue)) + "/day"
                    try {
                        $absValue = [Math]::Abs($RawValue)
                        $Sign + $(Format-Bytes $absValue) + "/day"
                    }
                    catch {
                        # Handle the error
                        Write-Host "Error occurred while rounding the value: $RawValue"
                    }
                }
            }
            Function Format-Days {
                Param (
                    $RawValue
                )
                [Math]::Round($RawValue)
            }

        $CSVsFound=($trendyagain | Group-Object -Property ObjectDescription -NoElement).name | Sort-Object -Unique
        $CSVs=Get-ChildItem -Path $SDDCPath -Filter "getvolume.xml" -Recurse -Depth 1 | Import-Clixml
        $trendyagainByCSVPerfData=$trendyagain | Group-Object -Property ObjectDescription
        $N = 14 # Require 14 days of history
        ForEach($CSV in $CSVs){
            ForEach($CSVPD in $trendyagainByCSVPerfData){
                if($($CSV.FileSystemLabel) -imatch $(($CSVPD.GROUP.ObjectDescription | Sort-Object -Unique) -replace "Volume ")){
                    #Write-Host "MATCH:"$($CSV.FileSystemLabel)" : " $(($CSVPD.GROUP.ObjectDescription | Sort-Object -Unique) -replace "Volume ")
                    # Last N days as (x, y) points
                    $PointsXY = @()
                    1..$N | ForEach-Object {
                        $PointsXY += [PsCustomObject]@{ "X" = $_ ; "Y" = $CSVPD.group[$_-1].value }
                    }

                    # Linear (y = ax + b) least squares algorithm
                    $MeanX = ($PointsXY | Measure-Object -Property X -Average).Average
                    $MeanY = ($PointsXY | Measure-Object -Property Y -Average).Average
                    $XX = $PointsXY | ForEach-Object { $_.X * $_.X }
                    $XY = $PointsXY | ForEach-Object { $_.X * $_.Y }
                    $SSXX = ($XX | Measure-Object -Sum).Sum - $N * $MeanX * $MeanX
                    $SSXY = ($XY | Measure-Object -Sum).Sum - $N * $MeanX * $MeanY
                    $A = ($SSXY / $SSXX)
                    $B = ($MeanY - $A * $MeanX)
                    $RawTrend = -$A # Flip to get daily increase in Used (vs decrease in Remaining)
                    $Trend = Format-Trend $RawTrend

                    If ($RawTrend -Gt 0) {
                        $DaysToFull = Format-Days ($CSV.SizeRemaining / $RawTrend)
                    }
                    Else {
                        $DaysToFull = "-"
                    }

                   $trendyagainOutput += [PsCustomObject]@{
                        "Volume"     = $CSV.FileSystemLabel
                        "Size"       = Format-Bytes ($CSV.Size)
                        "Used"       = Format-Bytes ($CSV.Size - $CSV.SizeRemaining)
                        "Trend"      = $Trend
                        "DaysToFull" = $DaysToFull
                    }
                }
            }
        }
        #$trendyagainOutput | Sort-Object PSComputerName | FT

        # HTML Report
            $html+='<H2 id="Sample5Makestoragetrendyagain">Sample 5: Make storage trendy again!</H2>'
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;To look at macro trends, performance history is retained for up to 1 year. This sample uses the Volume.Size.Available series from the LastYear </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;  timeframe to determine the rate that storage is filling up and estimate when it will be full. </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://learn.microsoft.com/en-us/windows-server/storage/storage-spaces/performance-history-scripting#sample-5-make-storage-trendy-again' target='_blank'>Ref: Sample 4: As they say, 25-gig is the new 10-gig</a></h5>"
            $html+=$trendyagainOutput |Sort-Object PSComputerName | ConvertTo-html -Fragment
            $html=$html `
                 -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                 -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'`
                 -replace  "$([char]206)$([char]188)s", '&mu;s'
            $PerfResultsSummary+=Set-ResultsSummary -name $Name -html $html
            $htmloutPerf+=$html
            $html=""
            $Name=""
    #endregion Make storage trendy again!
   

    #region Sample 6: Memory hog, you can run but you can't hide
        $Name="Sample 6: Memory hog, you can run but you can't hide"
        Write-Host "    Gathering $Name..."
        # Get input
            $memoryhog=""
            $memoryhogOutput=""
            $memoryhog=Get-ChildItem -Path $SDDCPath -Filter "memoryhog.xml" -Recurse -Depth 1 | Import-Clixml
            $memoryhogOutput = $memoryhog | Sort-Object PSComputerName | Select-Object PSComputerName,VM,AvgMemoryUsage
            #$memoryhogOutput | Sort-Object PSComputerName | FT
        # HTML Report
            $html+='<H2 id="Sample6Memoryhogyoucanrunbutyoucanthide">Sample 6: Memory hog</H2>'
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;Because performance history is collected and stored centrally for the whole cluster, you never need to stitch together data from different </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;  machines, no matter how many times VMs move between hosts. This sample uses the VM.Memory.Assigned series from the LastMonth timeframe </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;  to identify the virtual machines consuming the most memory over the last 35 days.</h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://learn.microsoft.com/en-us/windows-server/storage/storage-spaces/performance-history-scripting#sample-6-memory-hog-you-can-run-but-you-cant-hide' target='_blank'>Ref: Sample 6: Memory hog, you can run but you can't hide</a></h5>"
            $html+=$memoryhogOutput |Sort-Object PSComputerName | ConvertTo-html -Fragment
            $html=$html `
                 -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                 -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'`
                 -replace "$([char]956)s", '&mu;s'
            $PerfResultsSummary+=Set-ResultsSummary -name $Name -html $html
            $htmloutPerf+=$html
            $html=""
            $Name=""
    #endregion Sample 6: Memory hog, you can run but you can't hide

    }

    #region Create CluChk Performance Html Report

    $htmloutPerfReport = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml" lang="en">
    <head>'
    $htmloutPerfReport+=$htmlStyle
    $htmloutPerfReport+='<title>CluChk Performance Report</title>'
    $htmloutPerfReport+='<meta charset="UTF-8">'
    $htmloutPerfReport+="</head>"
    $htmloutPerfReport+="<body>"

    $html='<h1>CluChk Performance Report</h1>'
    $html+='<h3>&nbsp;Version: ' + $CluChkVer +' </h3>'
    #$RunDate=Get-Date
    $html+='<h3>&nbsp;Run Date: ' + $RunDate +' </h3>'
    $html+=''

    $html+='<h1>Results Summary</h1>'
    $html+= $PerfResultsSummary | Sort-Object Name | Select-Object `
    @{Label='Name';Expression={
    $Part1='<A href="#'
    $Part2=$_.Name -replace '\s',"" -replace '[^a-zA-Z0-9\s]', ''
    $Part3='">'
    $Part4=$_.Name
    $Part5='</A>'
    $Part1+$Part2+$Part3+$Part4+$Part5
    }},`
    @{Label='Warnings';Expression={IF($_.Warnings -gt 0){"YYEELLLLOOWW"+$_.Warnings}Else{$_.Warnings}}},`
    @{Label='Errors';Expression={IF($_.Errors -gt 0){"RREEDD"+$_.Errors}Else{$_.Errors}}}| ConvertTo-html -Fragment
    $html=$html `
     -replace '&gt;','>' -replace '&lt;','<' -replace '&quot;','"'`
     -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
     -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
    $htmloutPerfReport+=$html

    #$htmlout=$htmlout -replace '<table>','<table class="Sort-Objectable">'
    $htmloutPerf=$htmloutPerf `
    -replace [Regex]::Escape("***"),''`
    -replace '&amp;gt;','>'`
    -replace '&amp;lt;','<'`
    -replace '&quot;','"'`
    -replace '&lt;br&gt;','<br>'
    $htmloutPerfReport+=$htmloutPerf

    # Close body
    $htmloutPerfReport+='</body></html>'

    # Generate HTML Report
    $SDDCFileName=($SDDCPath -split '\\')[-1]
    If($CluChkReportLoc.Count -gt 1) {$CluChkReportLoc = $CluchkReportLoc[0]}
    $HtmlReport= Join-Path -Path $CluChkReportLoc -ChildPath CluChkPerfReport_v$CluChkVer-$DTString$SDDCFileName.html
    Write-Host ("Report Output location: " + $HtmlReport)
    if (Test-Path "$HtmlReport") {Remove-Item $HtmlReport}
    Out-File -FilePath $HtmlReport -InputObject $htmloutPerfReport -Encoding ASCII

})
$Job2 = $PowerShell2.BeginInvoke($Inputs2,$Outputs2)
}


#region Process SDDC
If($ProcessSDDC -ieq 'y'){
#Write-Host "Cluster Summary"
# Write-Host "*******************************************************************************"
# Write-Host "*                                                                             *"
# Write-Host "*                            Cluster Summary                                  *"
# Write-Host "*                                                                             *"
# Write-Host "*******************************************************************************"
$RunClusSum=""

If($RunClusSum -ne "No"){

 Write-Host "Creating Cluster Summary runspace"
 $Inputs  = New-Object 'System.Management.Automation.PSDataCollection[PSObject]'
 $Outputs = New-Object 'System.Management.Automation.PSDataCollection[PSObject]'
 $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
 $addMessageSessionStateFunction=@()
 $addMessageSessionStateFunction += New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'add-TableData', $(Get-Content Function:\add-TableData -ErrorAction Stop)
 $addMessageSessionStateFunction += New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Convert-BytesToSize', $(Get-Content Function:\Convert-BytesToSize -ErrorAction Stop)
 $addMessageSessionStateFunction += New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Set-ResultsSummary', $(Get-Content Function:\Set-ResultsSummary -ErrorAction Stop)
 $addMessageSessionStateFunction | % {$initialSessionState.Commands.Add($_)}

 $Runspace = [runspacefactory]::CreateRunspace($initialSessionState)
 $PowerShell = [powershell]::Create()
 $PowerShell.Runspace = $Runspace
 $Runspace.Open()
 $Runspace.SessionStateProxy.SetVariable('ClusterNodes',$ClusterNodes)
 $Runspace.SessionStateProxy.SetVariable('SysInfo',$SysInfo)
 $Runspace.SessionStateProxy.SetVariable('OSVersionNodes',$OSVersionNodes)
 $Runspace.SessionStateProxy.SetVariable('AzureLocalVersion',$AzureLocalVersion)
 $Runspace.SessionStateProxy.SetVariable('SupportMatrixtableData',$SupportMatrixtableData)
 $Runspace.SessionStateProxy.SetVariable('SDDCFiles',$SDDCFiles)
 $Runspace.SessionStateProxy.SetVariable('SDDCPath',$SDDCPath)
 $Runspace.SessionStateProxy.SetVariable('SMRevHistLatest',$SMRevHistLatest)
 $Runspace.SessionStateProxy.SetVariable('DTString',$DTString)
 $Runspace.SessionStateProxy.SetVariable('IncompleteSDDC',$IncompleteSDDC)
 $PowerShell.AddScript({
    #(Get-Process -Id $pid).PriorityClass='AboveNormal'
    #$ClusterNodes=$using:ClusterNodes;
    #$function:Convert-BytesToSize=$using:function:Convert-BytesToSize;$function:add-TableData=$using:function:add-TableData
    #$SDDCPath=$using:SDDCPath
    #$SysInfo=$using:SysInfo
    #$OSVersionNodes=$using:OSVersionNodes
    #$SupportMatrixtableData=$using:SupportMatrixtableData
    #$SDDCFiles=$using:SDDCFiles
    $htmlout=""
    $html=""
    $ResultsSummary=@()

#$htmlout+='<h1>Cluster Summary</h1>'
$htmlout+='<H1 id="ClusterSummary">Cluster Summary</H1>'

#Dell SDDC Check
    $name="Dell SDDC Version"
    $CloudHealthGatherTranscript=Get-ChildItem -Path $SDDCPath -Filter "0_CloudHealthGatherTranscript.log" -Recurse -Depth 1
    $DellSDDCVersionVersion = Get-Content $CloudHealthGatherTranscript.FullName | Where-Object{$_ -imatch "Dell SDDC Version"} 
    If($DellSDDCVersionVersion){$SDDCVersion="Dell"}Else{$SDDCVersion="Microsoft"}
    #HTML Report
$html+='<H2 id="DellSDDCVersion">Dell SDDC Version</H2>'
If($SDDCVersion -eq "Microsoft"){$html+='<h5><span style="background-color: #ffff00">&nbsp;&nbsp;&nbsp;&nbsp;Dell SDDC version used: False</span></h5>'}
If($SDDCVersion -eq "Dell"){$html+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;Dell SDDC version used: True</h5>'}          
        If($IncompleteSDDC -eq $True){$html+='<h5><span style="background-color: #ffff00">&nbsp;&nbsp;&nbsp;&nbsp;INCOMPLETE SDDC CAPTURE</span></h5>'}
$ResultsSummary+=Set-ResultsSummary -name $name -html $html
$htmlout+=$html
$html=""
$Name=""

#Cluster Name
        $Name="Cluster Name"
        #Write-Host "    Gathering $Name..."
        $ClusterName=$SDDCFiles."GetCluster" | `
        Select-Object @{Label="Name";Expression={If ($_.Name.length -gt 15) {"YYEELLLLOOWW"+"$($_.Name)"} else {$_.name}}},S2DEnabled,BlockCacheSize,`
        @{Label="ClusterFunctionalLevel";Expression={Switch($_.ClusterFunctionalLevel){`
            '9'{'2016'}`
            '10'{'2019'}`
            '11'{'2022'}`
            '12'{'23H2'}`
            default{('YYEELLLLOOWW' + $_.ClusterFunctionalLevel) }
        }}},`
        @{Label="CsvBalancer";Expression={
            $S2DEnabled=$_.S2DEnabled
            $CsvBalancer=$_.CsvBalancer
            IF($S2DEnabled -eq 1){
                Switch($CsvBalancer){
                '0'{'Disabled'}
                '1'{'YYEELLLLOOWWEnabled'}}}
            ElseIF($S2DEnabled -eq 0){
                Switch($CsvBalancer){
                '0'{'Disabled'}
                '1'{'Enabled'}}}
        }},@{Label="Quorum";Expression={@("Static","Dynamic","Unknown")[$_.DynamicQuorum]}},
        @{Label="Mixed Mode";Expression={if (($SDDCFiles."GetClusterNodeSupportedVersion").ClusterFunctionalLevel.count -gt 1) {'RREEDDTrue'} else {'False'}}},
        @{L="CAU Auto Update";E={$CauRole=$SDDCFiles."GetCauClusterRole";if ($CauRole) {if ($CauRole[1].value.value -eq "Online" -and $CauRole[-1].value -gt -1 -and $CauRole[-2].value.value -gt '' -and $CauRole[2].value -lt (Get-Date).AddDays(30)) {"RREEDDEnabled"}else{"Disabled"}}}}
        #$ClusterName | FT -AutoSize -Wrap

        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$ClusterName| Select-Object *,@{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=1ffFHyblkyR9Eo9NmSSs4npQJ0Q2TpfKP9cyKpXouQ0%3D&tn=CluChkClusterName'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
            $html+='<H2 id="ClusterName">Cluster Name</H2>'
            $html+="<h5><b>Should be:</b></h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;If S2DEnabled then CSVBalance should be disabled <a href='https://docs.microsoft.com/en-us/answers/questions/228498/csvbalancer-windows-server-2019-with-s2d-enabled.html' target='_blank'>Ref: CSVBalancer Windows Server 2019 with S2D enabled</a> </h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;If Mixed Mode is True, the Cluster and Storage Pool should be updated with Update-ClusterFunctionalLevel and Update-StoragePool -InputObject (Get-StoragePool -IsPrimordial &#36;False) </h5>"
If($ClusterName.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No Cluster found</span></h5>'}
            $html+=$ClusterName | ConvertTo-html -Fragment
            $html=$html `
                -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'          
$ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name=""
        
#Cluster IP
        $Name="Cluster IP"
        #Write-Host "    Gathering $Name..."
        
        $ClusterIPPrefix=$SDDCFiles."GetClusterNetwork" | ? Role -like "ClusterAndClient"
        $ClusterIP=$SDDCFiles."GetClusterResourceParameters" |`
        Where-Object{($_.Name -eq 'Address') -and ($_.InterfaceAlias -notmatch 'isatap')}| Select-Object ClusterObject,Name,Value,
        @{L="Prefix";E={foreach($aa in $ClusterIPPrefix) {if (([ipaddress](([ipaddress]$_.Value).address -band ([ipaddress]$aa.AddressMask).address)).ipaddresstostring -eq $aa.Address) {$aa.Ipv4PrefixLengths[0]}}}}

        #$ClusterIP | FT -AutoSize -Wrap

        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$ClusterIP|select -Property @{L='ClusterObject';E={[string]$_.ClusterObject}},Name,Value,Prefix,@{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=xK3wczXSgrGOibN%2FVReaIsRLsUeJugYw3yyFmtf%2FH84%3D&tn=CluChkClusterIp'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
            $html+='<H2 id="ClusterIP">Cluster IP</H2>'
If($ClusterIP.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No Cluster IP found</span></h5>'}
            $html+=$ClusterIP | ConvertTo-html -Fragment
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name=""

#Cluster Owner
        $Name="Cluster Owner"
        #Write-Host "    Gathering $Name..."
        $ClusterOwner=$SDDCFiles."GetClusterResource" |`
        Where-Object{$_.Name -eq 'Cluster IP Address'}|Select-Object Cluster,OwnerNode,State
        #$ClusterOwner | FT -AutoSize -Wrap

        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$ClusterOwner|select -Property @{L='Cluster';E={[string]$_.Cluster}},@{L='OwnerNode';E={[string]$_.OwnerNode}},@{L='State';E={[string]$_.State}},@{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=BBtyj%2BmV5djAbEOiTBGtM0oOT%2BoWujvg3jrYJbkIaL0%3D&tn=CluChkClusterOwner'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
            $html+='<H2 id="ClusterOwner">Cluster Owner</H2>'
If($ClusterOwner.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No Cluster IP found</span></h5>'}
            $html+=$ClusterOwner | ConvertTo-html -Fragment
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html 
            $html=""
            $Name=""

#Cluster Witness
        $Name="Cluster Witness"
        #Write-Host "    Gathering $Name..."
        $ClusterNodeCount=($SDDCFiles."GetClusterNode" |Measure-Object).count
        $ClusterWitness=$SDDCFiles."GetClusterResource" |`
        Where-Object{$_.ResourceType -imatch "witness"} |Select-Object Cluster,Name,ResourceType,OwnerNode,
        @{L="State";E={
            IF([string]$_.State -inotmatch "Online"){"RREEDD"+[string]$_.State}Else{[string]$_.State}
                    }}
        #$ClusterWitness | FT -AutoSize -Wrap

        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$ClusterWitness|Select-Object -Property `
                @{L='Cluster';E={[string]$_.Cluster}},
                @{L='Name';E={[string]$_.name}},
                @{L='OwnerNode';E={[string]$_.OwnerNode}},
                @{L='State';E={[string]$_.State}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=6bjScxiuJA9Jyngg%2BkfBIVJ0522SsDR2zEpzjeSHklk%3D&tn=CluChkClusterWitness'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
            $html+='<H2 id="ClusterWitness">ClusterWitness</H2>'
            $html+="<h5><b>Should be:</b></h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://docs.microsoft.com/en-us/windows-server/storage/storage-spaces/understand-quorum#cluster-quorum-overview' target='_blank'>Ref: Microsoft Docs - Cluster quorum overview</a> </h5>"
            $html+=$ClusterWitness | ConvertTo-html -Fragment
            IF($ClusterNodeCount -lt 5){
               IF($ClusterNodeCount -eq 2){
                   IF(!($ClusterWitness) -and $SysInfo[0].SysModel -notmatch "^APEX"){
                       $html+='<p style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;ERROR: No Cluster Witness found. Cluster Witness is Required for Two(2) node Clusters.</p>' 
                   }}
               IF($ClusterNodeCount -gt 2){
                  IF(!($ClusterWitness) -and $SysInfo[0].SysModel -notmatch "^APEX"){
                       $html+='<p style="color: #000000; background-color: #ffff00">&nbsp;&nbsp;&nbsp;&nbsp;WARNING: Missing Cluster Witness. It is Recommened to have a Cluster Witness for clusters with Less Than Five(5) Nodes to improve resilancy.</p>' 
                  }}}
            IF($ClusterNodeCount -gt 5){
                IF(!($ClusterWitness) -and $SysInfo[0].SysModel -notmatch "^APEX"){
                    $html+='<p>&nbsp;&nbsp;&nbsp;&nbsp;Cluster witness missing, but more than five(5) cluster nodes detected. No cluster witness required.</p>.' 
                }}

            $html=$html `
                -replace '<td>Offline</td>','<td style="background-color: #ffff00">Offline</td>'`
                -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'  
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name=""                

#Cluster Groups
        $Name="Cluster Groups"
        #Write-Host "    Gathering $Name..."
        $ClusterGroup=$SDDCFiles."GetClusterGroup" |`
        Sort-Object GroupType,Name | Select-Object GroupType,Name,OwnerNode,`
        @{Name="State";Expression={
            If($_.Name -inotmatch "Available Storage"){
                Switch($_.State.tostring()){
                        'Online'{'Online'}
                        'Offline'{'YYEELLLLOOWWOffline'}
                        'PartialOnline'{'YYEELLLLOOWWPartialOnline'}
                        'Failed'{'RREEDDFailed'}
                        }}
            ElseIf($_.Name -imatch "Available Storage"){
                Switch($_.State.tostring()){
                    'Online'{'Online'}`
                    'Offline'{'Offline'}`
                    'PartialOnline'{'YYEELLLLOOWWPartialOnline'}`
                    'Failed'{'RREEDDFailed'}}
                }}},`
        @{Label='StatusInformation';Expression={
            Switch($_.StatusInformation.ToString()){`
                '0' {'Healthy'}`
                '1' {'YYEELLLLOOWWWarning'}`
                '2' {'RREEDDUnhealthy'}`
                '5' {'Unknown'}`
                '512' {'AppUnknown'}`
                '1024' {'YYEELLLLOOWWHB Warning'}`
                '1536' {'Healthy'}`
                Default {$_.StatusInformation}

        }}},Id
        #$ClusterGroup-Objects | FT -AutoSize -Wrap

        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$ClusterGroup|Select-Object -Property `
                @{L='GroupType';E={[string]$_.GroupType}},
                @{L='Name';E={[string]$_.name}},
                @{L='OwnerNode';E={[string]$_.OwnerNode}},
                @{L='State';E={[string]$_.State}},
                @{L='StatusInformation';E={[string]$_.StatusInformation}},
                @{L='Id';E={[string]$_.Id}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=S2UBJb%2FzM0M4miaFAnPVda3ub6QOawh%2F0oJ5YV8vwWc%3D&tn=CluChkClusterGroups'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
            $html+='<H2 id="ClusterGroups">Cluster Groups</H2>'
            $html+=$ClusterGroup | ConvertTo-html -Fragment 
            $html=$html `
                -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'          
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name=""
        

#VM Info
        $Name="VM Info"
        #Write-Host "    Gathering $Name..."
        # Gather Host VM Configuration version based on cluster version
        $HCV=Switch(($SDDCFiles."GetCluster").ClusterFunctionalLevel){
                '11'{'10'}#2022
                '10'{'9'}#2019
                '9'{'8'}#2016
                '8'{'5'}#2012
            }
        $VMInfo=foreach ($key in ($SDDCFiles.keys -like "*GetVM")) { $SDDCFiles."$key" |`
        Sort-Object ComputerName,Name | Select-Object @{L="Host";E={$_.ComputerName}},VMName,State,CPUUsage,ProcessorCount,`
        @{Name="MemoryAssigned";Expression={Convert-BytesToSize $_.MemoryAssigned}},`
        Uptime,Status,@{Name="ConfigurationVersion";Expression={
            # Check for VM Configuration Version Less Than Host
            $VMCV=$_.Version
            IF($VMCV -gt 0){
                IF($HCV -gt $VMCV){'YYEELLLLOOWW'+$VMCV}`
                Else{$VMCV}
            }}},Generation,`
        @{Name="IsClustered";Expression={Switch($_.IsClustered){
            'True'{'True'}`
            'False'{'YYEELLLLOOWWFalse'}`
        }}},@{Name="NetworkAdapter";Expression={($_.NetworkAdapters |Select-Object macaddress).macaddress}},`
        @{Name="MacAddressSpoofing"; Expression={($_.NetworkAdapters | Select-Object MacAddressSpoofing).MacAddressSpoofing }},`
        @{Name="VMSwitchName"; Expression={($_.NetworkAdapters | Select-Object SwitchName).SwitchName}},`
        @{Name="HardDrives";Expression={
            $i=0
            ForEach($VHDX in $_.harddrives.path){
                $i++
                $VHDX="$($i). $VHDX`r`n"
                IF(-Not $VHDX.tolower().contains("c:\clusterstorage")){
                    $VHDX='&lt;span style="background-color: #FFFF00"&gt;'+$VHDX+'&lt;/span&gt;'
                }
                $VHDX
         }}},`
         @{Name="LocalISOAttached";Expression={('RREEDD'+$_.DVDDrives.path)*($_.DVDDrives.Path -ne $null -and -not ($_.DVDDrives.path -match 'C:\\ClusterStorage'))}},`
         @{Name="Repl Health";E={"RREEDD"*($_.ReplicationHealth.Value -notmatch "NotApplicable|Normal")+$_.ReplicationHealth.Value}} 
         }
        
        #$VMInfo| FT -AutoSize -Wrap

        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$VMInfo|Select-Object -Property `
                @{L='Host';E={[string]$_.Host}},
                @{L='VMName';E={[string]$_.VMName}},
                @{L='State';E={[string]$_.State}},
                @{L='CPUUsage';E={[string]$_.CPUUsage}},
                @{L='ProcessorCount';E={[string]$_.ProcessorCount}},
                @{L='MemoryAssigned';E={[string]$_.MemoryAssigned}},
                @{L='Uptime';E={[string]$_.Uptime}},
                @{L='Status';E={[string]$_.Status}},
                @{L='ConfigurationVersion';E={[string]$_.ConfigurationVersion}},
                @{L='Generation';E={[string]$_.Generation}},
                @{L='IsClustered';E={[string]$_.IsClustered}},
                @{L='NetworkAdapter';E={[string]$_.NetworkAdapter}},
                @{L='MacAddressSpoofing';E={[string]$_.MacAddressSpoofing}},
                @{L='SwitchName';E={[string]$_.SwitchName}},
                @{L='HardDrives';E={[string]$_.HardDrives}},
                @{L='LocalISOAttached';E={[string]$_.LocalISOAttached}},
                @{L='ReplHealth';E={[string]$_.'Repl Health'}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=nz3PDhV7QLcuhKjwHbn4ZofKmV71kJsW6RvbaIJK1og%3D&tn=CluChkVMInfo'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
            $html+='<H2 id="VMInfo">Virtual Machine Information</H2>'
If($VMInfo.count -eq 0){$html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;No Virtual Machines found</h5>"}
            $html+=$VMInfo | ConvertTo-html -Fragment 
            $html=$html`
                -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'          
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html        
    $htmlout+=$html
            $html=""
            $Name=""
#Cluster Nodes
            $dstop=Get-Date

        $Name="Cluster Nodes"
        #$SDDCPath
        $H2id="Cluster Nodes" -replace '\s'
        #Write-Host "    Gathering $Name..."

        #$ClusterNodes | FT -AutoSize -Wrap
        #Long Term Servicesing OS Version check

        $ClusterNodesOut=@()
# Reference: https://docs.microsoft.com/en-us/windows-server/get-started/windows-server-release-info
$filePath = gci $SDDCPath -Filter "systeminfo.txt" -Recurse

$allproperties=@()
# Initialize an empty hashtable
Foreach ($file in $filepath.fullname) {
    $properties = @{}

    # Read each line and extract key-value pairs
    Get-Content $file | ForEach-Object {
        if ($_ -match "^(.*?):\s*(.*)$") {
           $key = $matches[1].Trim()
           $value = $matches[2].Trim()
           $properties[$key] = $value
        }
    }
    $allproperties+=$properties | sort name
}
$primarytz=($allproperties.'Time Zone' | Group-Object | sort count)[-1].name
$HighestVersion = ($SysInfo | sort AzureLocalVersion -desc | select -first 1).AzureLocalVersion
$NewestUBR=Foreach ($key in ($SDDCFiles.Keys | ? {$_ -match "GetCurrentVersion"})) {$SDDCFiles.$key.ubr}
$NewestUBR=($NewestUBR | sort)[-1]
ForEach($Sys in $SysInfo){
    $ClusterNodesOut+=$ClusterNodes | Where-Object{$_.Name -imatch $Sys.HostName}|Sort-Object Name | Select-Object Name,Model,SerialNumber,State,StatusInformation,Id,`
        @{L="OSName";E={$Sys.OSName}},`
        @{L="OSVersion";E={$Sys.OSVersion}},`
        @{L="UBR";E={
            "RREEDD"*(($SDDCFiles."$($_.Name)GetCurrentVersion").UBR -ne $NewestUBR)+($SDDCFiles."$($_.Name)GetCurrentVersion").UBR
        }},
        @{L="OSBuild";E={
            IF($sys.OSName -imatch "Server"){
                "RREEDD"*!(@("17784","17763","14393","19042","20348","20349","25398","26100").contains($Sys.OSBuildNumber))+$Sys.OSBuildNumber
            }
            IF($sys.OSName -imatch "Stack"){
                "RREEDD"*!(@("17784","17763","14393","20349","25398","26100").contains($Sys.OSBuildNumber))+$Sys.OSBuildNumber
            }
        }},
        @{L="AzureLocal";E={
          IF($Sys.OSName -imatch "Server"){
            "N/A"
          }
          IF($Sys.OSName -imatch "Stack"){
            If ($Sys.AzureLocalVersion -ne $HighestVersion) {
	      "RREEDD"+$Sys.AzureLocalVersion
            } else {
              $Sys.AzureLocalVersion
            }
          }
        }},	
        @{L="Time Zone";E={
            $allproperties | ? {$_.'Host Name' -ieq $Sys.Hostname} | %{if ($_.'Time Zone' -ne $primarytz) {"RREEDD$($_['Time Zone'])"} else {"$($_['Time Zone'])"} }
        }}
        $ClusterNodesOut+=$ClusterNodes | Where-Object {!($Sysinfo.HostName -match $_.name)}
}
$NewestUBR=($ClusterNodesOut | sort UBR)[-1].UBR
#$ClusterNodesOut = $ClusterNodesOut | Select-Object Name,Model,SerialNumber,State,StatusInformation,Id,OSName,OSVersion,OSBuild,AzureLocal,@{L="UBR";E={"RREEDD"*($_.UBR -ne $NewestUBR)+$_.UBR}},'Time Zone'

#Azure Table CluChkClusterNodes
    $AzureTableData=@()
    $AzureTableData=$ClusterNodesOut| Select-Object Name,Model,SerialNumber,Id,OSVersion,OSBuild,@{L='State';E={$_.State.value}},@{L='StatusInformation';E={$_.StatusInformation.value}},@{L='ReportID';E={$CReportID}}
    $PartitionKey=$H2id
    $TableName="CluChk$($H2id)"
    $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=jsIz8RMBpn6Oo7z%2BNZ136xXHUlZN2LzdGM7ckGKohEU%3D&tn=CluChkClusterNodes'
    $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

#HTML Report
    $html+='<H2 id="ClusterNodes">Cluster Nodes</H2>'
    $html+="<h5><b>Should be:</b></h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;OSBuildVersion = 25398(HCI OS 23H2), 20349(HCI OS 22H2), 20348(Server 2022 LTSB or HCI OS 21H2), 19042(HCI OS 20H2), 2610(Server 2025), 20348(Server 2022), 17763(Server 2019), 14393(Server 2016) </h5>"
    If($ClusterNodesOut.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No Cluster Nodes found</span></h5>'}
    $html+=$ClusterNodesOut | ConvertTo-html -Fragment 
    IF(($sys.OSName -imatch "Stack") -and ($ClusterNodesOut.OSBuild -imatch "RREEDD")){$html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;NOTE: 20H2,21H2 are EoL MUST upgrade to 22H2 for Support. Return to Rework</h5>"}
    $html=$html `
     -replace '<td>Paused</td>','<td style="background-color: #ffff00">Paused</td>'`
     -replace '<td>Quarantined</td>','<td style="background-color: #ffff00">Quarantined</td>'`
     -replace '<td>Isolated</td>','<td style="background-color: #ffff00">Isolated</td>'`
     -replace '<td>Down</td>','<td style="color: #ffffff; background-color: #ff0000">Down</td>'`
     -replace '<td>Unknown</td>','<td style="color: #ffffff; background-color: #ff0000">Unknown</td>'`
     -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'
    $ResultsSummary+=Set-ResultsSummary -name $name -html $html        
    $htmlout+=$html
    $html=""
    $Name=""


#Storage Pool
        $Name="Storage Pool"
        #Write-Host "    Gathering $Name..."
        # Getting the total foot print on pool to use to calc free space in the pool
$TotalFootprintOnPool=(($SDDCFiles."GetStorageTier" |`
Where-Object{$_.AllocatedSize -gt 0}|Select-Object FootprintOnPool).FootprintOnPool| Measure-Object -sum).sum

if ($TotalFootprintOnPool -eq 0) {
# try the allocated size from the pool itself
$TotalFootprintOnPool=(($SDDCFiles."GetStoragePool" |`
Where-Object{$_.AllocatedSize -gt 0}|Select-Object AllocatedSize).AllocatedSize| Measure-Object -sum).sum
}

        # Gathering the largest auto-Select-Object physical disk to be used to detirm if the storage pool has enough free space for repairs
$SizeOfLargestDisk=($SDDCFiles."GetPhysicalDisk" | Where-Object {$_.usage -match 'Auto-Select-Object' -or $_.usage -match 1} | Sort-Object Size -Descending | Select-Object Size -first 1 ).size
$InPlaceRepairFreeSpaceNeededInStoragePool=IF(($ClusterNodes|Measure-Object).count -ge 4){$SizeOfLargestDisk*4}Else{($SizeOfLargestDisk*($ClusterNodes|Measure-Object).count)}

        $ClusterPool=$SDDCFiles."GetStoragePool" | `
        Where-Object {$_.FriendlyName -inotmatch "Primordial"}|Sort-Object HealthStatus -Descending | Select-Object PSComputerName,FriendlyName,@{Label="FaultDomainAwarenessDefault";Expression={Switch($_.FaultDomainAwarenessDefault){`
            '1'{'PhysicalDisk'}`
            '2'{'StorageEnclosure'}`
            '3'{'StorageScaleUnit'}`
            '4'{'StorageChassis'}`
            '5'{'StorageRack'}`
        }}},`
        @{Label='OperationalStatus';Expression={Switch($_.OperationalStatus -replace [regex]::match($_.OperationalStatus,"\\d+")){`
          '0'{'Unknown'}`
          '1'{'Other'}`
          '2'{'OK'}`
          '3'{'Degraded'}`
          '4'{'Stressed'}`
          '5'{'Predictive Failure'}`
          '6'{'Error'}`
          '7'{'Non-Recoverable Error'}`
          '8'{'Stopping'}`
          '9'{'Stopping'}`
          '10'{'Stopped'}`
          '11'{'In Service'}`
          '12'{'No Contact'}`
          '13'{'Lost Communication'}`
          '14'{'Aborted'}`
          '15'{'Dormant'}`
          '16'{'Supporting Entity in Error'}`
          '17'{'Completed'}`
          '18'{'Power Mode'}`
        }}},`
        @{Label='HealthStatus';Expression={Switch($_.HealthStatus){`
          '0'{'Healthy'}`
          '1'{'Warning'}`
          '2'{'Unhealthy'}`
          '5'{'Unknown'}`
        }}},IsPrimordial,IsReadOnly,@{Name="FreeSpace";Expression={
                $PoolCap=$_.size
                $FreePS=($PoolCap - $_.AllocatedSize)
                IF($FreePS -lt $InPlaceRepairFreeSpaceNeededInStoragePool){
                    $CFreePS=Convert-BytesToSize ($FreePS)
                    "YYEELLLLOOWW"+$CFreePS
                    Set-Variable -Name "SPNote" -Value "Free Space: Must have the equivalent of one capacity drive per server, up to 4 drives, for in-place recovery. This guarantees an immediate, in-place, parallel repair can succeed after the failure of any drive, even before it is replaced. Ref: https://docs.microsoft.com/en-us/windows-server/storage/storage-spaces/plan-volumes?redirectedfrom=MSDN#reserve-capacity" -Scope global -Force
                }Else{Convert-BytesToSize ($FreePS)}
            }},`
        @{Name="AllocatedSpace";Expression={IF($TotalFootprintOnPool -ne 0){Convert-BytesToSize ($TotalFootprintOnPool)}Else{"Not Available"}}},`
        @{Name="Capacity";Expression={Convert-BytesToSize $_.Size}},
        @{Name="Note";Expression={$SPNote}}
        #$ClusterPool |FL #FT -AutoSize -Wrap

        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$ClusterPool|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='FriendlyName';E={[string]$_.FriendlyName}},
                @{L='FaultDomainAwarenessDefault';E={[string]$_.FaultDomainAwarenessDefault}},
                @{L='OperationalStatus';E={[string]$_.OperationalStatus}},
                @{L='HealthStatus';E={[string]$_.HealthStatus}},
                @{L='IsPrimordial';E={$_.IsPrimordial}},
                @{L='IsReadOnly';E={$_.IsReadOnly}},
                @{L='FreeSpace';E={[string]$_.FreeSpace}},
                @{L='UsedSpace';E={[string]$_.UsedSpace}},
                @{L='Capacity';E={[string]$_.Capacity}},
                @{L='Note';E={[string]$_.Note}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=8nvyZ4j5X5HfLQQrngn10RCnG7UbZ21Kbna8MxAiQ6g%3D&tn=CluChkStoragePool'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
            $html+='<H2 id="StoragePool">Storage Pool</H2>'
If($ClusterPool.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No Storage Pools found</span></h5>'}
            $html+=$ClusterPool | ConvertTo-html -Fragment 
            If($ClusterName.S2DEnabled -eq 1){
                #Sets the FaultDomainAwarenessDefault red if S2D is enabled and the value is not StorageScaleUnit
                $html=$html -replace '<td>PhysicalDisk</td>','<td style="color: #ffffff; background-color: #ff0000">PhysicalDisk</td>'`
                            -replace '<td>PhysicalDisk</td>','<td style="color: #ffffff; background-color: #ff0000">PhysicalDisk</td>'`
                            -replace '<td>StorageEnclosure</td>','<td style="color: #ffffff; background-color: #ff0000">StorageEnclosure</td>'`
                            -replace '<td>StorageChassis</td>','<td style="color: #ffffff; background-color: #ff0000">StorageChassis</td>'`
                            -replace '<td>StorageRack</td>','<td style="color: #ffffff; background-color: #ff0000">StorageRack</td>'
            }
            $html=$html `
             -replace '<td>Warning</td>','<td style="background-color: #ffff00">Warning</td>'`
             -replace '<td>Degraded</td>','<td style="background-color: #ffff00">Degraded</td>'`
             -replace '<td>Stressed</td>','<td style="background-color: #ffff00">Stressed</td>'`
             -replace '<td>Aborted</td>','<td style="background-color: #ffff00">Aborted</td>'`
             -replace '<td>Predictive Failure</td>','<td style="background-color: #ffff00">Predictive Failure</td>'`
             -replace '<td>Error</td>','<td style="color: #ffffff; background-color: #ff0000">Error</td>'`
             -replace '<td>Non-Recoverable Error</td>','<td style="color: #ffffff; background-color: #ff0000">Non-Recoverable Error</td>'`
             -replace '<td>No Contact</td>','<td style="color: #ffffff; background-color: #ff0000">No Contact</td>'`
             -replace '<td>Lost Communication</td>','<td style="color: #ffffff; background-color: #ff0000">Lost Communication</td>'`
             -replace '<td>Unhealthy</td>','<td style="color: #ffffff; background-color: #ff0000">Unhealthy</td>'`
             -replace '<td>Unknown</td>','<td style="color: #ffffff; background-color: #ff0000">Unknown</td>'`
             -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
             -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'  
    $ResultsSummary+=Set-ResultsSummary -name $name -html $html        
    $htmlout+=$html
            $html=""
            $Name=""

#Storage Tiers
        $Name="Storage Tiers"
        #Write-Host "    Gathering $Name..."
        $StorageTiers=$SDDCFiles."GetStorageTier" |`
        Where-Object{$_.AllocatedSize -gt 0}|`
        Select-Object FriendlyName,ResiliencySettingName,
        @{Label="FaultDomainAwareness";Expression={Switch($_.FaultDomainAwareness){`
            '1'{'PhysicalDisk'}`
            '2'{'StorageEnclosure'}`
            '3'{'StorageScaleUnit'}`
            '4'{'StorageChassis'}`
            '5'{'StorageRack'}`
        }}},@{Name="AllocatedSize";Expression={Convert-BytesToSize $_.AllocatedSize}},`
        @{Name="FootprintOnPool(TB)";Expression={Convert-BytesToSize $_.FootprintOnPool}},NumberOfColumns,NumberOfDataCopies,PhysicalDiskRedundancy 
         
        #$StorageTiers | FT -AutoSize -Wrap

         #Azure Table
            $AzureTableData=@()
            $AzureTableData=$StorageTiers|Select-Object -Property `
                @{L='FriendlyName';E={[string]$_.FriendlyName}},
                @{L='ResiliencySettingName';E={[string]$_.ResiliencySettingName}},
                @{L='FaultDomainAwareness';E={[string]$_.FaultDomainAwareness}},
                @{L='AllocatedSize';E={[string]$_.AllocatedSize}},
                @{L='FootprintOnPool';E={$_.FootprintOnPool}},
                @{L='NumberOfColumns';E={$_.NumberOfColumns}},
                @{L='NumberOfDataCopies';E={[string]$_.NumberOfDataCopies}},
                @{L='PhysicalDiskRedundancy';E={[string]$_.PhysicalDiskRedundancy}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=%2FXHuueaYo%2F4bgQCdQJ1zgGfJAc3GnJXci9WE42rS4%2BU%3D&tn=CluChkStorageTiers'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
            $html+='<H2 id="StorageTiers">Storage Tiers</H2>'
            $html+="<h5><b>Should be:</b></h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;ResiliencySettingName should be Blank which means the Extent size is 256MB</h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;ResiliencySettingName = Mirror means Extent size of 1024MB. (Recreate using WAC to reduce to 256MB)</h5>"
            If($StorageTiers.count -eq 0){$html+="<h5><br>&nbsp;&nbsp;&nbsp;&nbsp;No Storage Tiers found</h5>"}
            $html+=$StorageTiers | ConvertTo-html -Fragment

            If($ClusterName.S2DEnabled -eq 1){
                #Sets the FaultDomainAwarenessDefault red if S2D is enabled and the value is PhysicalDisk
                $html=$html -replace '<td>PhysicalDisk</td>','<td style="color: #ffffff; background-color: #ff0000">PhysicalDisk</td>'`
                            -replace '<td>PhysicalDisk</td>','<td style="color: #ffffff; background-color: #ff0000">PhysicalDisk</td>'`
                            -replace '<td>StorageEnclosure</td>','<td style="color: #ffffff; background-color: #ff0000">StorageEnclosure</td>'`
                            -replace '<td>StorageChassis</td>','<td style="color: #ffffff; background-color: #ff0000">StorageChassis</td>'`
                            -replace '<td>StorageRack</td>','<td style="color: #ffffff; background-color: #ff0000">StorageRack</td>'
            } 
            $html=$html `
             -replace '<td>StorageEnclosure</td>','<td style="color: #ffffff; background-color: #ff0000">StorageEnclosure</td>'`
             -replace '<td>StorageChassis</td>','<td style="color: #ffffff; background-color: #ff0000">StorageChassis</td>'`
             -replace '<td>StorageRack</td>','<td style="color: #ffffff; background-color: #ff0000">StorageRack</td>'`
             -replace '<td>Warning</td>','<td style="background-color: #ffff00">Warning</td>'`
             -replace '<td>Unhealthy</td>','<td style="color: #ffffff; background-color: #ff0000">Unhealthy</td>'`
             -replace '<td>Unknown</td>','<td style="color: #ffffff; background-color: #ff0000">Unknown</td>'`
             -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
             -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
    $ResultsSummary+=Set-ResultsSummary -name $name -html $html         
    $htmlout+=$html
            $html=""
            $Name=""

#Virtual Disks
        $Name="Virtual Disks"
        #Write-Host "    Gathering $Name..." 
        #$DeDupTasks=foreach ($key in ($SDDCFiles.keys -like "*GetScheduledTask")) { $SDDCFiles."$key" | Where {$_.TaskName -match "^DeDup_.*Optimization" -and $_.NextRunTime}}
        #$DeDupTask=$DeDupTasks | Sort-Object LastRunTime | Select -First 1
        $SSDDE=Get-ChildItem -Path $SDDCPath -Filter "Microsoft-Windows-Deduplication-Operational.EVTX" -Recurse -Depth 1
        $LogPath=""
        $LogPath=$SSDDE.FullName
        $LogName="Microsoft-Windows-Deduplication/Operational"
        $SSDDEvents=@()
        $SSDDEventsOut=@()
        $DedupTask = Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable @{Path=$LogPath;Id="6153"} -MaxEvents 1

        $GetVolumes=$SDDCFiles."GetVolume"
        $VirtualDisks=$SDDCFiles."GetVirtualDisk"
        If (($VirtualDisks | Where {$_.IsDeduplicationEnabled -eq $false}).count -eq $VirtualDisks.count) {$DedupDisabled=$True} else {$DedupDisabled=$False}
        $VirtualDisks=$SDDCFiles."GetVirtualDisk" |`
        Sort-Object HealthStatus -Descending | Select-Object FriendlyName,`
        @{Label='OperationalStatus';Expression={If ($_.OperationalStatus -match "\d") {Switch($_.OperationalStatus -replace [regex]::match($_.OperationalStatus,"\\d+")){`
          '0'{'Unknown'}`
          '1'{'Other'}`
          '2'{'OK'}`
          '3'{'Degraded'}`
          '4'{'Stressed'}`
          '5'{'Predictive Failure'}`
          '6'{'Error'}`
          '7'{'Non-Recoverable Error'}`
          '8'{'Stopping'}`
          '9'{'Stopping'}`
          '10'{'Stopped'}`
          '11'{'In Service'}`
          '12'{'No Contact'}`
          '13'{'Lost Communication'}`
          '14'{'Aborted'}`
          '15'{'Dormant'}`
          '16'{'Supporting Entity in Error'}`
          '17'{'Completed'}`
          '18'{'Power Mode'}`
          '53250'{'Detached'}
          }} else {$_.OperationalStatus}
        }},`
        @{Label='HealthStatus';Expression={If ($_.HealthStatus -match "\d") {Switch($_.HealthStatus){`
          '0'{'Healthy'}`
          '1'{'Warning'}`
          '2'{'Unhealthy'}`
          '5'{'Unknown'}`
        }} else {$_.HealthStatus}}},`
        @{Label='DetachedReason';Expression={If ($_.DetachedReason -match "\d") {Switch($_.DetachedReason){`
          '0'{'Not Detached'}`
          '1'{'Operational'}`
          '2'{'By Policy'}`
          '5'{'Unknown'}`
        }} else {$_.DetachedReason}}},
        FileSystem,
        IsSnapshot,@{Label='Access';Expression={If ($_.Access -match "\d") {@('Unknown', 'RREEDDRead Only', 'RREEDDWrite Only', 'Read/Write', 'RREEDDWrite Once')[$_.Access]} else {"RREEDD"*($_.Access -ne "Read/Write")+$_.Access}}},@{Label='Dedup Enabled';Expression={if ($DedupDisabled -and $DeDupTask -and $SysInfo[0].SysModel -notmatch "^APEX") {"RREEDD"+$_.IsDeduplicationEnabled} else {$_.IsDeduplicationEnabled}}},@{Label='DeDup Last Run';Expression={If ($DeDupTask) {$DeDupTask.TimeCreated.GetDateTimeFormats('s')}}}
        #$VirtualDisks | FT -AutoSize -Wrap

         #Azure Table
            $AzureTableData=@()
            $AzureTableData=$VirtualDisks|Select-Object -Property `
                @{L='FriendlyName';E={[string]$_.FriendlyName}},
                @{L='OperationalStatus';E={[string]$_.OperationalStatus}},
                @{L='HealthStatus';E={[string]$_.HealthStatus}},
                @{L='DetachedReason';E={[string]$_.DetachedReason}},
                @{L='IsSnapshot';E={$_.IsSnapshot}},
                @{L='IsReadOnly';E={$_.IsReadOnly}},
                @{L='IsDeduplicationEnabled';E={[string]$_.IsDeduplicationEnabled}},
                @{L='DeDup Last Rnn';E={[string]$_.'DeDup Last Run'}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=k%2Bg4CssfAkHj8%2Fj4BMS8XEv2tVlxwRK0tzxM4Qkkpcg%3D&tn=CluChkVirtualDisks'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
        $html+='<H2 id="VirtualDisks">Virtual Disks</H2>'
If($VirtualDisks.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No Virtual Disks found</span></h5>'}
        $html+=$VirtualDisks | ConvertTo-html -Fragment 
        $html=$html `
         -replace '<td>Warning</td>','<td style="background-color: #ffff00">Warning</td>'`
         -replace '<td>Degraded</td>','<td style="background-color: #ffff00">Degraded</td>'`
         -replace '<td>Stressed</td>','<td style="background-color: #ffff00">Stressed</td>'`
         -replace '<td>Aborted</td>','<td style="background-color: #ffff00">Aborted</td>'`
         -replace '<td>Predictive Failure</td>','<td style="background-color: #ffff00">Predictive Failure</td>'`
         -replace '<td>Error</td>','<td style="color: #ffffff; background-color: #ff0000">Error</td>'`
         -replace '<td>Non-Recoverable Error</td>','<td style="color: #ffffff; background-color: #ff0000">Non-Recoverable Error</td>'`
         -replace '<td>No Contact</td>','<td style="color: #ffffff; background-color: #ff0000">No Contact</td>'`
         -replace '<td>Lost Communication</td>','<td style="color: #ffffff; background-color: #ff0000">Lost Communication</td>'`
         -replace '<td>Unhealthy</td>','<td style="color: #ffffff; background-color: #ff0000">Unhealthy</td>'`
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'`
         -replace '<td>Unknown</td>','<td style="color: #ffffff; background-color: #ff0000">Unknown</td>'
$ResultsSummary+=Set-ResultsSummary -name $name -html $html         
$htmlout+=$html
        $html=""
        $Name=""

#Virtual Disk Resiliency
        $Name="Virtual Disk Resiliency"
        #Write-Host "    Gathering $Name..."
        $html+='<H2 id="VirtualDiskResiliency">Virtual Disk Resiliency</H2>'
        $GetStorageTiers=@()
        $GetStorageTiers = $SDDCFiles."GetStorageTier"
        #$GetStorageTiers|Sort-Object FriendlyName|Select-Object FriendlyName,ResiliencySettingName,PhysicalDiskRedundancy
        $GetVirtualDisks=@()
        $GetVirtualDisks = $SDDCFiles."GetVirtualDisk"
        #$GetVirtualDisks|FL *
        #$GetVirtualDisks|Sort-Object FriendlyName|Select-Object FriendlyName,ResiliencySettingName,PhysicalDiskRedundancy
        
        #$GetVolumes|FL *

        $VDRData=@()
        $VDROutput=@()
        $VDResiliency=""
        $VolData=@()
        ForEach($VD in $GetVirtualDisks){
            
                ForEach($V in $GetVolumes){
                    IF($V.FileSystemLabel -eq $VD.FriendlyName){
                    $VolUsed=[Math]::Round(($V.SizeRemaining/$V.Size)*100,2)
                    $VolPercentUsed="$VolUsed%"
                    $VolData += [PSCustomObject]@{
                        "VolName" = $V.FileSystemLabel
                        "%Available"    = $VolPercentUsed
                        }
                }
            }
            #Size Remaining on Volume
            #ResiliencySettingName PhysicalDiskRedundancy Resiliency
            #Mirror                1                      2-way mirror
            #Mirror                2                      3-way mirror
            #Parity                1                      single parity
            #Parity                2                      dual parity
            #Mirror,Parity         2,2                    Mirror-accelerated parity(MAP)
            $VDMirrorSize=""
            $VDParitySize=""
            $VDFriendlyName = $VD.FriendlyName
            $VDFootprint  = $VD.FootprintOnPool
            $VDEfficiency = [math]::Round($VD.Size/$VD.FootprintOnPool*100,2)

            #Non Teired VirtualDisk will have ResiliencySettingName and PhysicalDiskRedundancy properties
            IF($VD.ResiliencySettingName -and $VD.PhysicalDiskRedundancy){
                #PhysicalDiskRedundancy
                    $VDPhysicalDiskRedundancy = $VD.PhysicalDiskRedundancy
                #FaultDomainAwareness
                    IF(($VD.FaultDomainAwareness).length -eq 1){ 
                        IF($VD.FaultDomainAwareness -eq 1){$VDFaultDomainAwareness='PhysicalDisk'}
                        IF($VD.FaultDomainAwareness -eq 2){$VDFaultDomainAwareness='StorageEnclosure'}
                        IF($VD.FaultDomainAwareness -eq 3){$VDFaultDomainAwareness='StorageScaleUnit'}
                        IF($VD.FaultDomainAwareness -eq 4){$VDFaultDomainAwareness='StorageChassis'}
                        IF($VD.FaultDomainAwareness -eq 5){$VDFaultDomainAwareness='StorageRack'}
                    }
                    IF(($VD.FaultDomainAwareness).length -gt 1){
                            $VDFaultDomainAwareness=$VD.FaultDomainAwareness
                    }
                #Resiliency 
                    IF($VD.ResiliencySettingName -eq 'Mirror' -and $VD.PhysicalDiskRedundancy -eq 1){
                        #2-way and 3+ nodes mark yellow as we can only have 1 fault
                        IF($ClusterNodeCount -ge 3){$VDResiliency="YYEELLLLOOWW2-way Mirror"}
                        Else{$VDResiliency="2-way Mirror"}}
                    IF($VD.ResiliencySettingName -eq 'Mirror' -and $VD.PhysicalDiskRedundancy -eq 2){$VDResiliency="3-Way Mirror"}
                    IF($VD.ResiliencySettingName -eq 'Parity' -and $VD.PhysicalDiskRedundancy -eq 1){$VDResiliency="Single Parity"}
                    IF($VD.ResiliencySettingName -eq 'Parity' -and $VD.PhysicalDiskRedundancy -eq 2){$VDResiliency="Dual Parity"}
                #MirrorSize
                    IF($VD.ResiliencySettingName -eq 'Mirror'){$VDMirrorSize = $VD.size}
                #ParitySize
                    IF($VD.ResiliencySettingName -eq 'Parity'){$VDParitySize = $VD.size}
 
            }

            #Teired VirtualDisk will have ResiliencySettingName and PhysicalDiskRedundancy properties completely blank
            #ResiliencySettingName and PhysicalDiskRedundancy properties can be found in GetStorageTier
            IF(!($VD.ResiliencySettingName -and $VD.PhysicalDiskRedundancy)){
                #StorageTierName
                $VD2STMatch=@()

                $VD2STMatch=$GetStorageTiers | Where-Object{$_.FriendlyName -imatch $VD.FriendlyName}
                $VD2STMatch=$GetStorageTiers | Where-Object{$_.FriendlyName -imatch $VD.FriendlyName}
                #$VDResiliencyMap=""
                $VDResiliency=""
                ForEach($VDMatch in $VD2STMatch){
                    $MultiVD2STMatches="No"
                    #PhysicalDiskRedundancy
                        $VDPhysicalDiskRedundancy = $VDMatch.PhysicalDiskRedundancy
                    #FaultDomainAwareness
                        IF(($VDMatch.FaultDomainAwareness).length -eq 1){
                            IF($VDMatch.FaultDomainAwareness -eq 1){$VDFaultDomainAwareness='PhysicalDisk'}
                            IF($VDMatch.FaultDomainAwareness -eq 2){$VDFaultDomainAwareness='StorageEnclosure'}
                            IF($VDMatch.FaultDomainAwareness -eq 3){$VDFaultDomainAwareness='StorageScaleUnit'}
                            IF($VDMatch.FaultDomainAwareness -eq 4){$VDFaultDomainAwareness='StorageChassis'}
                            IF($VDMatch.FaultDomainAwareness -eq 5){$VDFaultDomainAwareness='StorageRack'}
                        }
                        IF(($VDMatch.FaultDomainAwareness).length -gt 1){
                            $VDFaultDomainAwareness=$VDMatch.FaultDomainAwareness
                        }
                    IF($VD2STMatch -is [array]){
                        $MultiVD2STMatches="YES"
                        #Mirror
                            If ($VDMatch.ResiliencySettingName -Eq "Mirror") {
                                $VDMirrorSize=$VDMatch.Size
                                If ($VDMatch.PhysicalDiskRedundancy -Eq 1) { $VDResiliencyMapMirror2 = "2-Way Mirror" }
                                ElseIf ($VDMatch.PhysicalDiskRedundancy -Eq 2) { $VDResiliencyMapMirror2 = "3-Way Mirror" }
                            }
                        #Parity
                            ElseIf ($VDMatch.ResiliencySettingName -Eq "Parity") {
                                $VDParitySize=$VDMatch.Size
                                If ($VDMatch.PhysicalDiskRedundancy -Eq 1) { $VDResiliencyMapParity2 = "+ Single Parity" }
                                ElseIf ($VDMatch.PhysicalDiskRedundancy -Eq 2) { $VDResiliencyMapParity2 = "+ Dual Parity" }
                            }}
                    ElseIF($VD2STMatch -isnot [array]){
                        #Resiliency 
                            IF($VDMatch.ResiliencySettingName -eq 'Mirror' -and $VDMatch.PhysicalDiskRedundancy -eq 1){
                                #2-way and 3+ nodes mark yellow as we can only have 1 fault
                                IF($ClusterNodeCount -ge 3){$VDResiliency1="YYEELLLLOOWW2-way Mirror"}
                                Else{$VDResiliency1="2-way Mirror"}}
                            IF($VDMatch.ResiliencySettingName -eq 'Mirror' -and $VDMatch.PhysicalDiskRedundancy -eq 2){$VDResiliency1="3-way Mirror"}
                            IF($VDMatch.ResiliencySettingName -eq 'Parity' -and $VDMatch.PhysicalDiskRedundancy -eq 1){$VDResiliency1="Single Parity"}
                            IF($VDMatch.ResiliencySettingName -eq 'Parity' -and $VDMatch.PhysicalDiskRedundancy -eq 2){$VDResiliency1="Dual Parity"}
                        #MirrorSize
                            IF($VDMatch.ResiliencySettingName -eq 'Mirror'){$VDMirrorSize = $VDMatch.size}
                        #ParitySize
                            IF($VDMatch.ResiliencySettingName -eq 'Parity'){$VDParitySize = $VDMatch.size}
                    }}
                
                #Resiliency Output
                Switch($MultiVD2STMatches){
                    'YES'{
                            IF($VDResiliencyMapMirror2 -and $VDResiliencyMapParity2){
                                $VDResiliency=""
                                $VDResiliencyMapMirror2=$VDResiliencyMapMirror2 -replace "Mirror",""
                                $VDResiliencyMapParity2=$VDResiliencyMapParity2 -replace "Parity ",""
                                $VDResiliency+="MAP("+$VDResiliencyMapMirror2+$VDResiliencyMapParity2+")"
                            }
                            Else{
                                $VDResiliency=""
                                $VDResiliency+=$VDResiliencyMapMirror2+$VDResiliencyMapParity2
                                }
                            $VDResiliencyMapMirror2=""
                            $VDResiliencyMapParity2=""
                         }
                    'NO' {$VDResiliency=$VDResiliency1;$VDResiliency1=""}
                }}
            $Size=$VDMirrorSize+$VDParitySize
            $Size=Convert-BytesToSize $Size
            IF($VDMirrorSize -gt 0){$VDMirrorSize=Convert-BytesToSize $VDMirrorSize}Else{$VDMirrorSize=""}
            IF($VDParitySize -gt 0){$VDParitySize=Convert-BytesToSize $VDParitySize}Else{$VDParitySize=""}
            IF($VDResiliency -imatch "map"){$VDResiliency="Mirror-accelerated parity"}
            
            #Simple
            IF($VD.ResiliencySettingName -eq "Simple"){$VDResiliency="RREEDDSimple"}
            
            $VDRData = [PSCustomObject]@{
                "VirtualDiskName"        = $VDFriendlyName
                "TotalSize"              = $Size
                "%Free"                  = $VolPercentUsed
                "PhysicalDiskRedundancy" = $VDPhysicalDiskRedundancy
                "FaultDomainAwareness"   = $VDFaultDomainAwareness
                "Resiliency"             = $VDResiliency
                "MirrorSize"             = $VDMirrorSize
                "ParitySize"             = $VDParitySize
                "StorageFootprint"       = Convert-BytesToSize $VDFootprint
                "Efficiency"             = "RREEDD"*($VDEfficiency -eq "100")+"$VDEfficiency%"
                "Note"                   = IF($VDResiliency -imatch "yyeellllooww"){"Resiliency is 2-way with more than two nodes. Default is 3-way"}elseif($VDResiliency -imatch "RREEDD"){"Simple volumes are not supported in HCI/S2D"}}
            $VDROutput += $VDRData
        }

        #$VDROutput |Sort-Object VirtualDiskName |Format-Table

         #Azure Table
            $AzureTableData=@()
            $AzureTableData=$VDROutput|Select-Object -Property `
                @{L='VirtualDiskName';E={[string]$_.VirtualDiskName}},
                @{L='TotalSize';E={[string]$_.TotalSize}},
                @{L='PercentFree';E={[string]$_.'%Free'}},
                @{L='PhysicalDiskRedundancy';E={[string]$_.PhysicalDiskRedundancy}},
                @{L='FaultDomainAwareness';E={$_.FaultDomainAwareness}},
                @{L='Resiliency';E={$_.Resiliency}},
                @{L='MirrorSize';E={[string]$_.MirrorSize}},
                @{L='ParitySize';E={[string]$_.ParitySize}},
                @{L='StorageFootprint';E={[string]$_.StorageFootprint}},
                @{L='Efficiency';E={[string]$_.Efficiency}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=2YjSfvYP6vWU0%2FCy5wyWsnvVfKDYH40xf5cDBic9c4U%3D&tn=CluChkVirtualDiskResiliency'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
            $html+=$VDROutput | ConvertTo-html -Fragment 
            If($ClusterName.S2DEnabled -eq 1){
                #Sets the FaultDomainAwarenessDefault red if S2D is enabled and the value is PhysicalDisk
                $html=$html -replace '<td>PhysicalDisk</td>','<td style="color: #ffffff; background-color: #ff0000">PhysicalDisk</td>'`
                            -replace '<td>PhysicalDisk</td>','<td style="color: #ffffff; background-color: #ff0000">PhysicalDisk</td>'`
                            -replace '<td>StorageEnclosure</td>','<td style="color: #ffffff; background-color: #ff0000">StorageEnclosure</td>'`
                            -replace '<td>StorageChassis</td>','<td style="color: #ffffff; background-color: #ff0000">StorageChassis</td>'`
                            -replace '<td>StorageRack</td>','<td style="color: #ffffff; background-color: #ff0000">StorageRack</td>'`
                            -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                            -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
            }               
            $html=$html `
                -replace '<td>Warning</td>','<td style="background-color: #ffff00">Warning</td>'
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html 
            $htmlout+=$html 
            $html=""
            $Name=""

#ClusterSharedVolume - JF
        $Name="Cluster Shared Volumes"
        Write-Host "    Gathering $Name..."
        $ClusterSharedVolume=$SDDCFiles.GetClusterSharedVolume   

        $CSVOutPut = $ClusterSharedVolume |`
        Select-Object Name, State,
            @{L='MountPath';E={$_.SharedVolumeInfo.FriendlyVolumeName}},
            @{L='FaultState';E={If($_.SharedVolumeInfo.FaultState -ine "NoFaults"){"RREEDD"+$_.SharedVolumeInfo.FaultState}Else{$_.SharedVolumeInfo.FaultState}}},
            @{L='MaintenanceMode';E={$_.SharedVolumeInfo.MaintenanceMode}},
            #@{L='FileSystem';E={$sv=$_.Name;$ftype=($GetVolumes | ? {$sv -match $_.FileSystemLabel}).FileSystemType;if ($ftype -eq $null) {$ftype=2};@("CSVFS_NTFS","CSVFS_REFS",,,"FAT","FAT16","FAT32","NTFS4","NTFS5",,,"EXT2","EXT3","ReiserFS","NTFS","REFS")[$ftype -band 15]}},
            @{L='RedirectedAccess';E={If($_.SharedVolumeInfo.RedirectedAccess -inotmatch "False"){"YYEELLLLOOWW"+$_.SharedVolumeInfo.RedirectedAccess}Else{$_.SharedVolumeInfo.RedirectedAccess}}},
            @{Label='Note';Expression={IF($_.SharedVolumeInfo.RedirectedAccess -inotmatch "False"){"CSV in Redirected Mode might be caused by a Filter Driver."}}}

        #HTML Report
        $html+='<H2 id="ClusterSharedVolumes">Cluster Shared Volumes (CSV)</H2>'
        If($CSVOutPut.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No CSVs found</span></h5>'}
        $html+=$CSVOutPut | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        #$html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;<a href='$($SMRevHistLatest.link)' target='_blank'>Ref: Support Matrix for Dell EMC Solutions for Microsoft Azure Stack HCI</a></h5>"
        $htmlout+=$html 
        $html=""
        $Name="" 
         
#Physical Disks
        $dstop=Get-Date
        $Name="Physical Disks"
        #Write-Host "    Gathering $Name..."
        $PhysicalDisks=$SDDCFiles."GetPhysicalDisk" |`
        Select-Object @{Label='Node';Expression={""}},UniqueID,@{L='ID';E={$_.DeviceId}},FriendlyName,@{L='Model';E={
        Switch -Regex ($_.Model) {
            "Dell Ent NVMe v2 AGN RI U.2"       {"MZWLR6T4HALA-00AD3"}
            "Dell Ent NVMe PM1733a RI"          {"MZWLR15THBLAAD3"}
            "Dell Ent NVMe CM6"                 {"KCM6XVUL1T60"}
	        "Dell Ent NVMe CM7"                 {"KCM7XVUG1T60"}
            "Dell Ent NVMe PM1735a"             {"MZWLR6T4HBLAAD3"}
            "Dell Ent NVMe P5600 MU U.2"        {"D7 P5600 Series 1.6TB"}
            "Dell Express Flash CD5"            {"KCD5XLUG3T84"}
            "Dell DC NVMe CD8 U.2 960GB"        {"KCD8XRUG960G"}
            "Dell DC NVMe CD8 U.2 1.92TB"       {"KCD8XRUG1T92"}
            "Dell DC NVMe CD8 U.2 3.84TB"       {"KCD8XRUG3T84"}
            "Dell DC NVMe CD8 U.2 7.68TB"       {"KCD8XRUG7T68"}
            default {$_}
        }}},`
        @{Label='SerialNumber';Expression={
            If($_.SerialNumber -imatch '^[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*.'){$_.AdapterSerialNumber -replace " ",""}
            Else{$_.SerialNumber -replace " ",""}
        }},@{Label='Slot';Expression={([REGEX]::Match($_.PhysicalLocation ,'^*Slot\s\d*')).value -replace 'Slot ',""}},`
        @{Label='MediaType';Expression={`
            @('','','','HDD','SSD','SCM')[$_.MediaType]
        }},@{Label='DriveCount';Expression={""}},CanPool,`
        @{Label='OperationalStatus';Expression={`
        IF ($_.OperationalStatus -eq '53270') {'In Maintenance Mode'} elseif ($_.OperationalStatus -eq '53285') {'Threshold Exceeded'}
        else {
             @('Unknown','Other','OK','Degraded','Stressed','Predictive Failure','Error','Non-Recoverable Error','Stopping',`
             'Stopping','Stopped','In Service','No Contact','Lost Communication','Aborted','Dormant',`
             'Supporting Entity in Error','Completed','Power Mode')[$_.OperationalStatus]
        }
        }},OperationalDetails,`
        @{Label='HealthStatus';Expression={`
            @('Healthy','Warning','Unhealthy','','','Unknown')[$_.HealthStatus]
        }},`
        @{Label='Usage';Expression={`
            @('Unknown','AutoSelect-Object','ManualSelect-Object','HotSpare','Retired','Journal')[$_.Usage]
        }},@{Label='Size';Expression={Convert-BytesToSize $_.Size}},@{Label='AllocatedSize';Expression={Convert-BytesToSize $_.AllocatedSize}},@{Label='Utilization';Expression={"{0:N2}" -f ($_.AllocatedSize/$_.Size)}},FirmwareVersion|Sort-Object SlotNumber,HealthStatus -Descending
    
    #Check if each disk is no more than 10% different that the max Utilization
    $UtlMax=($PhysicalDisks |?{$_.Usage -imatch "autoselect"} | Measure-Object -Property Utilization -Maximum).Maximum
    $PhysicalDisks = $PhysicalDisks | Foreach{ 
    IF($_.Usage -imatch "autoselect" -and ($UtlMax - $_.Utilization) -ge .1){$_.Utilization = "YYEELLLLOOWW"+$_.Utilization}$_}

    $diskmdls=($PhysicalDisks | Select-Object Model -Unique).Model
#Find all drives from Dell Support matrix
        $resultObject=@()
        ForEach ($diskmdl in $diskmdls) {
        try {
            $SMDrive=$null
            $SMDrive= $SupportMatrixtableData.Values.Values | ForEach-Object {$_ | Where-Object { $_.Model -eq $diskmdl }}

            $SMDrive=$SMDrive | Where Model -eq $diskmdl | Sort 'Firmware Min*' -Descending | Select -First 1
                    #Create a customer object
                        $resultObject += [PSCustomObject] @{
                                        Type             = $SMDrive.Type
                                        DriveType        = $SMDrive.'Drive Type'
                                        FormFactor       = $SMDrive.'Form Factor'
                                        Endurance        = $SMDrive.Endurance
                                        Vendor           = $SMDrive.Vendor
                                        Series           = $SMDrive.Series
                                        Model            = $SMDrive.Model
                                        DevicePartNumber = $SMDrive.'Device Part Number (P/N)'
                                        SoftwareBundle   = $SMDrive.'Firmware Software Bundle'
                                        Firmware         = $SMDrive.'Firmware Minimum Supported Version'
                                        Capacity         = $SMDrive.Capacity
                                        Use              = $SMDrive.Use
    }
              
         } catch {}

        }
        $SMFWDiskData = $resultObject

    # Find host physcially connected
      $GetStorageFaultDomain=@()
        $ConnectHost2PhysicalDisk=@()
        $GetStorageFaultDomain=foreach ($a in ($SDDCFiles.keys -like "*GetStorageFaultDomain")){ $SDDCFiles."$a"|`
            Where-Object {$_.SerialNumber -ne $null}| Select-Object @{Label='Node';Expression={$a.replace("GetStorageFaultDomain","")}},`
                @{Label='SerialNumber';Expression={
If($_.SerialNumber -imatch '^[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*_[A-Z,0-9]*.'){$_.AdapterSerialNumber -replace " ",""}
                    Else{$_.SerialNumber -replace " ",""}}},`
PhysicalLocation,FirmwareVersion,bustype}
        # Check for all NVMe
                   #Write-Host "Total time taken in Getfaultdomain XMLs $(((Get-Date)-$dstop).totalmilliseconds)"
           # $dstop=Get-Date
$AllNVMe=$False
$GetStorageFaultDomainbustype=$GetStorageFaultDomain.bustype | Sort-Object -Unique
IF($GetStorageFaultDomainbustype.count -eq 1){
IF($GetStorageFaultDomainbustype -eq '17'){
# type 17 = NVMe
# ref: https://wutils.com/wmi/root/microsoft/windows/storage/providers_v2/spaces_physicaldisk/#bustype_properties
$AllNVMe=$True
}}
        $diskmdlsfirm=@{}
        Foreach($diskmdl in $diskmdls) {
            $SMFWDiskfirm=$null
	    # first try without FIPS
            try {$SMFWDiskfirm=(($SMFWDiskData | Where-Object {$_.Model -like $diskmdl -and $_.Series -notmatch "FIPS"} | select Firmware | sort -Descending | Select-Object -First 1).Firmware)} catch {}
            IF ($SMFWDiskfirm.count -gt 0) {
              $diskmdlsfirm.add($diskmdl,$SMFWDiskfirm.trim())
            } Else {
              # next try with FIPS included in search
	      try {$SMFWDiskfirm=(($SMFWDiskData | Where-Object {$_.Model -like $diskmdl} | select Firmware | sort -Descending | Select-Object -First 1).Firmware)} catch {}
              IF ($SMFWDiskfirm.count -gt 0) {
                $diskmdlsfirm.add($diskmdl,$SMFWDiskfirm.trim())
              } Else {       
                $diskmdlsfirm.add($diskmdl,("YYEELLLLOOWWNot found in matrix"*($SysInfo[0].SysModel -notmatch "^APEX")))
              }
            }
        }
        $ConnectHost2PhysicalDisk=@()
            ForEach($Disk in $PhysicalDisks){
                $ConnectHost2PhysicalDisk+=$Disk | Select-Object `
                @{Label='Node';Expression={$NodeD=($GetStorageFaultDomain|Where-Object{$_.SerialNumber -eq $disk.SerialNumber}|Select-Object -expandproperty Node -first 1);IF($NodeD){$NodeD}Else{"Missing"}}`
                } ,ID,FriendlyName,UniqueID,SerialNumber,Slot,MediaType,CanPool,OperationalStatus,HealthStatus,Usage,Size,AllocatedSize,Utilization,Outlier,`
                @{L='MatrixVersion';E={$diskmdlsfirm[$_.model]}},
@{L='InstalledVersion';E={
IF (($_.FirmwareVersion -gt $diskmdlsfirm[$_.Model]) -and $diskmdlsfirm[$_.Model] -notmatch "Not found in matrix" -and $SysInfo[0].SysModel -notmatch "^APEX"){
# newer firmware version found on disk
'YYEELLLLOOWW'+$_.FirmwareVersion 
} Else {
IF (($_.FirmwareVersion -lt $diskmdlsfirm[$_.Model]) -and $diskmdlsfirm[$_.Model] -notmatch "Not found in matrix" -and $SysInfo[0].SysModel -notmatch "^APEX"){
# disk has an older firmware version
'RREEDD'+$_.FirmwareVersion
} Else {
$_.FirmwareVersion
}
}
}}
        }
        #$ConnectHost2PhysicalDisk| Sort-Object Node,FriendlyName | FT -AutoSize
        $PhysicalDisks = $ConnectHost2PhysicalDisk | Sort-Object Node,Slot
        
<#$StorPortLog=gci -Path $SDDCPath -Filter "Microsoft-Windows-Storage-Storport-Operational.EVTX" -Depth 1
$DisksLatency=Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable @{Path=$StorPortLog.Fullname;Id="505"}

$LatencyCount=@()
ForEach ($SN in ($PhysicalDisks.SerialNumber))
{
  $DiskCounter=$DisksLatency | %{if ($_.Properties.Value -match $SN) {$_}}
  $LatencyCount +=[PSCustomObject] @{
                        SerialNumber =$SN
                        Lat6 = ([ScriptBlock]{($DiskCounter | %{$_.Properties[60].Value} | Measure-Object -Sum).Sum}).InvokeReturnAsIs()
                        Lat7 = ([ScriptBlock]{($DiskCounter | %{$_.Properties[61].Value} | Measure-Object -Sum).Sum}).InvokeReturnAsIs()
                        Lat8 = ([ScriptBlock]{($DiskCounter | %{$_.Properties[62].Value} | Measure-Object -Sum).Sum}).InvokeReturnAsIs()
                        Lat9 = ([ScriptBlock]{($DiskCounter | %{$_.Properties[63].Value} | Measure-Object -Sum).Sum}).InvokeReturnAsIs()
                        Lat10 = ([ScriptBlock]{($DiskCounter | %{$_.Properties[64].Value} | Measure-Object -Sum).Sum}).InvokeReturnAsIs()
                        Lat11 = ([ScriptBlock]{($DiskCounter | %{$_.Properties[65].Value} | Measure-Object -Sum).Sum}).InvokeReturnAsIs()
                        Lat12 = ([ScriptBlock]{($DiskCounter | %{$_.Properties[66].Value} | Measure-Object -Sum).Sum}).InvokeReturnAsIs()
                        }

}
#>

        # $PhysicalDisks counts
$DiskCountPerNode=@()
ForEach ($PDisk in ($PhysicalDisks | Group-Object Node)){
$DiskCountPerNode+=$PDisk|Select-Object @{Label='Node';Expression={$_.Name}},@{Label='FriendlyName';Expression={'Total'}},count
ForEach($FName in ($PDisk.Group | Group-Object Node,FriendlyName)){
$DiskCountPerNode+=$FName|Select-Object @{Label='Node';Expression={$PDisk.Name}},@{Label='FriendlyName';Expression={($_.Name -split ', ')[1]}},@{L='count';E={If($PDisk.Name -imatch 'missing'){'RREEDD'+$_.count}Else{$_.count}}}
}}

    # Add new output format
$DiskCountPerNodetbl = New-Object System.Data.DataTable "DiskCountPerNode"
$DiskCountPerNodetbl.Columns.add((New-Object System.Data.DataColumn("FriendlyName")))
ForEach ($a in ($DiskCountPerNode.Node | Sort-Object -Unique)){
$DiskCountPerNodetbl.Columns.Add((New-Object System.Data.DataColumn([string]$a)))}
ForEach ($a in ($DiskCountPerNode.FriendlyName | Sort-Object -Unique)){
IF (($a.length -gt 2) -and ($a -inotmatch 'System.__ComObject') -and ($a -ne "Total")) {
$row=$DiskCountPerNodetbl.NewRow()
$row["FriendlyName"]=($a | Out-String).Trim()
ForEach($b in ($DiskCountPerNode | where-object {$_.FriendlyName -eq $a} | Sort-Object Node)){
 $row["$($b.Node)"] = $b.Count
}
$DiskCountPerNodetbl.rows.add($row)
}
}

    #$DiskCountPerNodetbl |Format-Table 
    $DiskCountPerNodeOut = $DiskCountPerNodetbl|Where-Object{$_.FriendlyName -inotmatch 'System.__ComObject'}|Sort-Object FriendlyName | Select-object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors
    $DiskCountPerNodetbl.Columns.Clear()
    $DiskCountPerNodetbl.Columns.Remove.Name | Out-Null
    $DiskCountPerNodetbl=""

    #Write-Host "Physical Disk time is $(((Get-Date)-$dstop).totalmilliseconds)"

         #Azure Table
            $AzureTableData=@()
            $AzureTableData=$PhysicalDisks|Select-Object -Property `
                @{L='Node';E={[string]$_.Node}},
                @{L='ID';E={[string]$_.ID}},
                @{L='FriendlyName';E={[string]$_.FriendlyName}},
                @{L='UniqueID';E={[string]$_.UniqueID}},
                @{L='SerialNumber';E={[string]$_.SerialNumber}},
                @{L='Slot';E={$_.Slot}},
                @{L='MediaType';E={$_.MediaType}},
                @{L='CanPool';E={$_.CanPool}},
                @{L='OperationalStatus';E={$_.OperationalStatus}},
                @{L='HealthStatus';E={$_.HealthStatus}},
                @{L='Usage';E={$_.Usage}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=EjhEHQqt6cCy1oTcn5gCiwp9Ahb8oAetfE65jNhnMIQ%3D&tn=CluChkPhysicalDisks'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
        $html+='<H2 id="PhysicalDisks">Physical Disks</H2>'
If($PhysicalDisks.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No Physical Disks found</span></h5>'}
        $html+=$PhysicalDisks | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>Warning</td>','<td style="background-color: #ffff00">Warning</td>'`
         -replace '<td>Degraded</td>','<td style="background-color: #ffff00">Degraded</td>'`
         -replace '<td>Stressed</td>','<td style="background-color: #ffff00">Stressed</td>'`
         -replace '<td>Aborted</td>','<td style="background-color: #ffff00">Aborted</td>'`
         -replace '<td>Predictive Failure</td>','<td style="background-color: #ffff00">Predictive Failure</td>'`
         -replace '<td>Error</td>','<td style="color: #ffffff; background-color: #ff0000">Error</td>'`
         -replace '<td>Non-Recoverable Error</td>','<td style="color: #ffffff; background-color: #ff0000">Non-Recoverable Error</td>'`
         -replace '<td>No Contact</td>','<td style="color: #ffffff; background-color: #ff0000">No Contact</td>'`
         -replace '<td>Lost Communication</td>','<td style="color: #ffffff; background-color: #ff0000">Lost Communication</td>'`
         -replace '<td>Threshold Exceeded</td>','<td style="color: #ffffff; background-color: #ff0000">Threshold Exceeded</td>'`
         -replace '<td>Unhealthy</td>','<td style="color: #ffffff; background-color: #ff0000">Unhealthy</td>'`
         -replace '<td>Unknown</td>','<td style="color: #ffffff; background-color: #ff0000">Unknown</td>'`
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;<a href='$($SMRevHistLatest.link)' target='_blank'>Ref: Support Matrix for Dell EMC Solutions for Microsoft Azure Stack HCI</a></h5>"
        $htmlout+=$html 
        $html=""
        $Name=""

        $html+='<H2 id="PhysicalDiskCounts">Physical Disk Counts</H2>'
        $html+=$DiskCountPerNodeOut | ConvertTo-html -Fragment 
        $html=$html `
            -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
            -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
        $Name="Physical Disk Counts"
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html        
$htmlout+=$html 
        $html=""
        $Name=""

#Storage Enclosure
        $Name="Storage Enclosure"
        #Write-Host "    Gathering $Name..."
        $StorageEnclosure=$SDDCFiles."GetStorageEnclosure" |`
        Sort-Object HealthStatus -Descending | Select-Object FriendlyName,SerialNumber,`
        @{Label='OperationalStatus';Expression={Switch($_.OperationalStatus -replace [regex]::match($_.OperationalStatus,"\\d+")){`
          '0'{'Unknown'}`
          '1'{'Other'}`
          '2'{'OK'}`
          '3'{'Degraded'}`
          '4'{'Stressed'}`
          '5'{'Predictive Failure'}`
          '6'{'Error'}`
          '7'{'Non-Recoverable Error'}`
          '8'{'Stopping'}`
          '9'{'Stopping'}`
          '10'{'Stopped'}`
          '11'{'In Service'}`
          '12'{'No Contact'}`
          '13'{'Lost Communication'}`
          '14'{'Aborted'}`
          '15'{'Dormant'}`
          '16'{'Supporting Entity in Error'}`
          '17'{'Completed'}`
          '18'{'Power Mode'}`

        }}},`
        @{Label='HealthStatus';Expression={Switch($_.HealthStatus){`
          '0'{'Healthy'}`
          '1'{'Warning'}`
          '2'{'Unhealthy'}`
          '5'{'Unknown'}`
        }}},NumberOfSlots,ElementTypesInError
        #$StorageEnclosure | FT -AutoSize -Wrap
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$StorageEnclosure|Select-Object -Property `
                @{L='FriendlyName';E={[string]$_.FriendlyName}},
                @{L='SerialNumber';E={[string]$_.SerialNumber}},
                @{L='OperationalStatus';E={[string]$_.OperationalStatus}},
                @{L='HealthStatus';E={[string]$_.HealthStatus}},
                @{L='NumberOfSlots';E={$_.NumberOfSlots}},
                @{L='ElementTypesInError';E={$_.ElementTypesInError}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=%2FDPFOOMr200C%2FZyCwQ%2BcvAEJFSQ4GOeJ6uv85xxNMM4%3D&tn=CluChkStorageEnclosure'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
            $html+='<H2 id="StorageEnclosure">Storage Enclosure</H2>'
If($StorageEnclosure.count -eq 0){
if ($AllNVMe) {
$html+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;No storage enclosures found, this is expected for a full NVMe setup</span></h5>'
} else {
$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No storage enclosures found</span></h5>'
}
}
            $html+=$StorageEnclosure | ConvertTo-html -Fragment
            $html=$html `
             -replace '<td>Warning</td>','<td style="background-color: #ffff00">Warning</td>'`
             -replace '<td>Degraded</td>','<td style="background-color: #ffff00">Degraded</td>'`
             -replace '<td>Stressed</td>','<td style="background-color: #ffff00">Stressed</td>'`
             -replace '<td>Aborted</td>','<td style="background-color: #ffff00">Aborted</td>'`
             -replace '<td>Predictive Failure</td>','<td style="background-color: #ffff00">Predictive Failure</td>'`
             -replace '<td>Error</td>','<td style="color: #ffffff; background-color: #ff0000">Error</td>'`
             -replace '<td>Non-Recoverable Error</td>','<td style="color: #ffffff; background-color: #ff0000">Non-Recoverable Error</td>'`
             -replace '<td>No Contact</td>','<td style="color: #ffffff; background-color: #ff0000">No Contact</td>'`
             -replace '<td>Lost Communication</td>','<td style="color: #ffffff; background-color: #ff0000">Lost Communication</td>'`
             -replace '<td>Unhealthy</td>','<td style="color: #ffffff; background-color: #ff0000">Unhealthy</td>'`
             -replace '<td>Unknown</td>','<td style="color: #ffffff; background-color: #ff0000">Unknown</td>'
    $ResultsSummary+=Set-ResultsSummary -name $name -html $html        
    $htmlout+=$html 
            $html=""
            $Name=""        

# Check for SCSI Sense Keys
    $Name="SCSI Sense Key Information"
    #Write-Host "    Gathering $Name..."
    $Baddisks=$PhysicalDisks|Where-Object{($_.HealthStatus -ne 'Healthy'-and $_.HealthStatus -ne '0' -and $_.OperationalStatus -notmatch 'Maintenance'-and $_.OperationalStatus -ne '53270') -and $_.SerialNumber.length -gt "0"} | Sort-Object SerialNumber -Unique
    IF($Baddisks){
        $BDSKs=@()
        $StorDiags=Get-ChildItem -Path $SDDCPath -Filter "Microsoft-Windows-Storage-ClassPnP-Operational.EVTX" -Recurse -Depth 1
        $LogPath=$StorDiags.FullName
        $LogName='Microsoft-Windows-Storage-ClassPnP/Operational'
        $LogID='505'
        Write-Host "        Checking Event Log $LogName for ID $LogID..."
            $StorDiagEvents = Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable @{Path=$LogPath;LogName=$LogName;Id=$LogID}
            If ($Null -eq $StorDiagEvents) { Write-Host "            No such EventId $LogId exists" -ForegroundColor Yellow } Else { Write-Host "            Found "($StorDiagEvents).Count" Events for ID $LogId"}
        $FoundEvents=@()
        ForEach($BadDisk in $Baddisks){
        Write-Host "        Checking for 505's with Disk Serial Number" $Baddisk.SerialNumber"..."
            $FoundEvents+=$StorDiagEvents | ForEach-Object { `
            $StorDiagEvent = $_; `
            $StorDiagXMLData = [xml]$StorDiagEvent.ToXml(); `
            $StorDiagXMLDataCount = $StorDiagXMLData.Event.EventData.Data.Count; `
            0..($StorDiagXMLDataCount - 1) | ForEach-Object { `
            Add-Member -InputObject $StorDiagEvent -MemberType NoteProperty -Force -Name $StorDiagXMLData.Event.EventData.Data[$_].Name -Value $StorDiagXMLData.Event.EventData.Data[$_].'#text' }; `
            $StorDiagEvent|Where-Object{$_.SerialNumber -match $Baddisk.serialnumber}|`
            Select-Object TimeCreated,DeviceNumber,SerialNumber,`
            @{Label='Key';Expression={$_.SenseKey}},`
            @{Label='ASC';Expression={$_.AdditionalSenseCode}},`
            @{Label='ASCQ';Expression={$_.AdditionalSenseCodeQualifier}}}
            If(!($FoundEvents)){
                Write-Host "            None found."
                $BDSKs+="No 505 events found for $($Baddisk.SerialNumber)<br>"
                }}
        $html+='<H2 id="SCSISenseKeyInformation">SCSI Sense Key Information</H2>'
        $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;Checked Event Log $LogName for ID $LogID</b></h5>"
        $html+=$FoundEvents | ConvertTo-html -Fragment 
        $html+=$BDSKs
    }Else{
        $html+="<h2>Sense Key Problems for Physical Disks</h2>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;No unhealthy Disks to lookup.</h5>"
    }
$ResultsSummary+=Set-ResultsSummary -name $name -html $html    
$htmlout+=$html 
    $html=""
    $Name=""

#Storage Jobs
        $Name="Storage Jobs"
        #Write-Host "    Gathering $Name..."
        $StorageJobs=$SDDCFiles."GetStorageJob" |`
        Sort-Object PSComputerName | Select-Object Name,ElapsedTime,`
        @{Label='JobState';Expression={Switch($_.JobState){`
          '2'{'New'}`
          '3'{'Starting'}`
          '4'{'Running'}`
          '5'{'Suspended'}`
          '6'{'ShuttingDown'}`
          '7'{'Completed'}`
          '8'{'Terminated'}`
          '9'{'Killed'}`
          '10'{'Exception'}`
          '32768'{'CompletedWithWarnings'}`
        }}},PercentComplete,BytesProcessed,BytesTotal
        #$StorageJobs | FT -AutoSize -Wrap
         #Azure Table
            $AzureTableData=@()
            $AzureTableData=$StorageJobs|Select-Object -Property `
                @{L='Name';E={[string]$_.Name}},
                @{L='ElapsedTime';E={[string]$_.ElapsedTime}},
                @{L='JobState';E={[string]$_.JobState}},
                @{L='PercentComplete';E={[string]$_.PercentComplete}},
                @{L='BytesProcessed';E={$_.BytesProcessed}},
                @{L='BytesTotal';E={$_.BytesTotal}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2018-03-28&si=CluChkStorageJobs-17FF391F830&tn=cluchkstoragejobs&sig=H1h0Fcame8NRLD%2FkZ0DryrTtR6tQyFh6CYuoQ%2F7VdIs%3D'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
        $html+='<H2 id="StorageJobs">Storage Jobs</H2>'
If($StorageJobs.count -eq 0){$html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;No Storage Jobs found</h5>"}
        $html+=$StorageJobs | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>Terminated</td>','<td style="background-color: #ffff00">Terminated</td>'`
         -replace '<td>CompletedWithWarnings</td>','<td style="background-color: #ffff00">CompletedWithWarnings</td>'`
         -replace '<td>Killed</td>','<td style="color: #ffffff; background-color: #ff0000">Killed</td>'`
         -replace '<td>Exception</td>','<td style="color: #ffffff; background-color: #ff0000">Exception</td>'
$ResultsSummary+=Set-ResultsSummary -name $name -html $html        
$htmlout+=$html
        $html=""
        $Name=""

#DebugStorageSubsystem
        $Name="Debug Storage Subsystem"
        #Write-Host "    Gathering $Name..."
        $DebugStorageSubsystem=$SDDCFiles."DebugStorageSubsystem" |`
        Sort-Object PerceivedSeverity,PSComputerName | Select-Object PSComputerName,FaultType,`
        @{Label='PerceivedSeverity';Expression={Switch($_.PerceivedSeverity -replace [regex]::match($_.PerceivedSeverity,"\\d+")){`
          '0'{'Unknown'}`
          '2'{'Information'}`
          '3'{'Degraded'}`
          '4'{'Minor'}`
          '5'{'Major'}`
          '6'{'Critical'}`
          '7'{'Fatal'}`
        }}}`
        ,Reason,`
        @{Label='RecommendedActions';Expression={$_.RecommendedActions -replace [regex]::match($_.OperationalStatus,"\\d+")}}
        #$DebugStorageSubsystem | FT -AutoSize -Wrap
         #Azure Table
            $AzureTableData=@()
            $AzureTableData=$DebugStorageSubsystem|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='FaultType';E={[string]$_.FaultType}},
                @{L='PerceivedSeverity';E={[string]$_.PerceivedSeverity}},
                @{L='Reason';E={[string]$_.Reason}},
                @{L='RecommendedActions';E={$_.RecommendedActions}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=s8pIL%2F06eOqeGy6NJIU7T38naoE%2FZt14xWAI8O2upew%3D&tn=CluChkDebugStorageSubsystem'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
        $html+='<H2 id="DebugStorageSubsystem">Debug Storage Subsystem</H2>'
If($DebugStorageSubsystem.count -eq 0){$html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;No Storage Debug entries found</h5>"}
        $html+=$DebugStorageSubsystem | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>Degraded</td>','<td style="background-color: #ffff00">Degraded</td>'`
         -replace '<td>Minor</td>','<td style="background-color: #ffff00">Minor</td>'`
         -replace '<td>Fatal</td>','<td style="color: #ffffff; background-color: #ff0000">Fatal</td>'`
         -replace '<td>Critical</td>','<td style="color: #ffffff; background-color: #ff0000">Critical</td>'`
         -replace '<td>Major</td>','<td style="color: #ffffff; background-color: #ff0000">Major</td>'`
         -replace '<td>Unknown</td>','<td style="color: #ffffff; background-color: #ff0000">Unknown</td>'
$ResultsSummary+=Set-ResultsSummary -name $name -html $html        
$htmlout+=$html 
        $html=""
        $Name=""

#Networks
        $Name="Cluster Networks"
        #Write-Host "    Gathering $Name..."  
        $ClusterNetworks=$SDDCFiles."GetClusterNetwork" |`
        Sort-Object Name | Select-Object Name,State,Metric,`
        @{Label='Role';Expression={
            IF($_.Address.length -eq 0 -and $_.Role -notmatch 'None'){"RREEDD"+$_.Role.ToString()}Else{
            IF($ClusterName.S2DEnabled -eq 1 -and $_.Address.length -ne 0 -and $_.Role -match 'None'){"YYEELLLLOOWW"+$_.Role.ToString()}Else{$_.Role.ToString()}
        }}},`
        Address,@{Label='Note';Expression={
                IF($_.Address.length -eq 0 -and $_.Role -notmatch 'None'){"No IP address found: Role should be None to prevent cluster from using this network."}
                IF($ClusterName.S2DEnabled -eq 1 -and $_.Address.length -ne 0 -and $_.Role -match 'None'){"If S2D/HCI Network is for Storage, Role should NOT be set to None."}
            }}
        #$ClusterNetworks | FT -AutoSize -Wrap
         #Azure Table
            $AzureTableData=@()
            $AzureTableData=$ClusterNetworks|Select-Object -Property `
                @{L='Name';E={[string]$_.Name}},
                @{L='State';E={[string]$_.State}},
                @{L='Metric';E={[string]$_.Metric}},
                @{L='Role';E={[string]$_.Role}},
                @{L='Address';E={$_.Address}},
                @{L='Note';E={IF($_.Note.length -le 3){$Null}Else{$_.Note}}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=qLzqkR5GaQTHpC2Bb4mSqohcuCfpUgpgXCT2Di309ss%3D&tn=CluChkClusterNetworks'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
        $html+='<H2 id="ClusterNetworks">Cluster Networks</H2>'
If($ClusterNetworks.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No Cluster Networks found</span></h5>'}
        $html+=$ClusterNetworks | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>Down</td>','<td style="color: #ffffff; background-color: #ff0000">Down</td>'`
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
$ResultsSummary+=Set-ResultsSummary -name $name -html $html        
$htmlout+=$html
        $html="" 
        $Name=""

#IP Addr Info
        $Name="Cluster Node IP Addresses"
        #Write-Host "    Gathering $Name..."
        $ClusterNodeIPAddresses=foreach ($key in ($SDDCFiles.keys -like "*GetNetIpAddress")) { $SDDCFiles."$key" |`
        Where-Object{($_.InterfaceAlias -notmatch 'isatap') -and ($_.InterfaceAlias -notmatch 'Pseudo')}|Sort-Object PSComputerName | Select-Object PSComputerName,InterfaceAlias,ifIndex,`
        @{Label='AddressState';Expression={Switch($_.AddressState){`
          '0'{'Invalid'}`
          '1'{'Tentative'}`
          '2'{'Duplicate'}`
          '3'{'Deprecated'}`
          '4'{'Preferred'}`
        }}},IPv4Address,IPv6Address
        }
        #$ClusterNodeIPAddresses | FT -AutoSize -Wrap
         #Azure Table
            $AzureTableData=@()
            $AzureTableData=$ClusterNodeIPAddresses|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='InterfaceAlias';E={[string]$_.InterfaceAlias}},
                @{L='ifIndex';E={[string]$_.ifIndex}},
                @{L='AddressState';E={[string]$_.AddressState}},
                @{L='IPv4Address';E={$_.IPv4Address}},
                @{L='IPv6Address';E={$_.IPv6Address}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=0ZQPLKPqVvbM8RfjvPenmTFbguOGfBROJBhVnpX%2Bo6Y%3D&tn=CluChkClusterNodeIPAddresses'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
        $html+='<H2 id="ClusterNodeIPAddresses">Cluster Node IP Addresses</H2>'
        $html+=$ClusterNodeIPAddresses | ConvertTo-html -Fragment 
        $html=$html `
         -replace '<td>Invalid</td>','<td style="color: #ffffff; background-color: #ff0000">Invalid</td>'`
         -replace '<td>Duplicate</td>','<td style="color: #ffffff; background-color: #ff0000">Duplicate</td>'`
         -replace '<td>4</td>','<td>Preferred</td>'
$ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $ClusterNodeIPAddressesKey=""
        $ClusterNodeIPAddressesKey+="<h5><b>AddressState Key:</b></h5>"
        $ClusterNodeIPAddressesKey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-- Invalid. IP address configuration information for addresses that are not valid and will not be used.</h5>"
        $ClusterNodeIPAddressesKey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-- Tentative. IP address configuration information for addresses that are not used for communication, as the uniqueness of those IP addresses is being verified.</h5>"
        $ClusterNodeIPAddressesKey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-- Duplicate. IP address configuration information for addresses for which a duplicate IP address has been detected and the current IP address will not be used.</h5>"
        $ClusterNodeIPAddressesKey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-- Deprecated. IP address configuration information for addresses that will no longer be used to establish new connections, but will continue to be used with existing connections.</h5>"
        $ClusterNodeIPAddressesKey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-- Preferred. IP address configuration information for addresses that are valid and available for use.</h5>"
        $ClusterNodeIPAddressesKey+="&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://docs.microsoft.com/en-us/powershell/module/nettcpip/get-netipaddress?view=win10-ps' target='_blank'>Ref: Microsoft Docs - Get-NetIPAddress -AddressState</a>"
        $htmlout+=$ClusterNodeIPAddressesKey
        $html=""
        $Name=""

# FLTMC
    $Name="FLTMC Logs"
    #Write-Host "    Gathering $Name..."  
    $html=""
$html+='<H2 id="FLTMCLogs">FLTMC Logs</H2>'
    $html+="<h5><b>Should be:</b></h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;Suspect anything NOT Microsoft</h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://raw.githubusercontent.com/MicrosoftDocs/windows-driver-docs/staging/windows-driver-docs-pr/ifs/allocated-altitudes.md' target='_blank'>Ref: https://raw.githubusercontent.com/MicrosoftDocs/windows-driver-docs/staging/windows-driver-docs-pr/ifs/allocated-altitudes.md</a></h5>"
    $URL="https://raw.githubusercontent.com/MicrosoftDocs/windows-driver-docs/staging/windows-driver-docs-pr/ifs/allocated-altitudes.md"
    
    # Added to process FLTMCXml files
    $FLTMCXmlLogs=@()
    $FLTMCXmlLogsFiles=Get-ChildItem -Path $SDDCPath -Filter "FLTMC.XML" -Recurse -Depth 1
    IF($FLTMCXmlLogsFiles){
        ForEach($FLTMCXmlLogsFile in $FLTMCXmlLogsFiles){
            $NodeName=((Split-Path -Path $FLTMCXmlLogsFile.FullName).Split("\")[-1]).split("_")[-1]
            $FLTMCXmlLogs+= Import-Clixml -Path $FLTMCXmlLogsFile.FullName | Select-Object *,@{L="PSComputerName";E={$NodeName}}
        }
    }
    IF(!($FLTMCXmlLogs)){    
        $LocalFltmc=@()
        Function Get-FLTMC($InputFile){
            $parts=@()
            $output=@()
            If($InputFile -match "Run it locally"){$output = FLTMC}
            Else{$output = Get-Content $InputFile}
            If($output -match "Filter Name"){
            $Found=($output|Select-String "Filter Name" -Context 0,1).LineNumber
            $output=$output[($Found+1)..($output.length)]
            }
            $output | ForEach-Object {$parts = $_ -split "\s+", 6
                New-Object -Type PSObject -Property @{
                            PSComputerName =($NodeName)
                            FilterName = ($parts[0])
                            NumInstances = $parts[1]
                            Altitude = $parts[2]
                            Frame = $parts[3]}
            }
        }
        # use the credentials of the current user to authenticate on the proxy server
            $Wcl = new-object System.Net.WebClient
            $Wcl.Headers.Add("user-agent", "PowerShell Script")
            $Wcl.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials         
        
        # Gets the the list of known file system minifilter drivers
                $KnownFltrs=invoke-webrequest -Uri $URL -UseBasicParsing -TimeoutSec 30 | Select-Object RawContent
                $RawKnownFltrs=$KnownFltrs.RawContent.split("`n")
                $RawKnownFltrs+=

        # Get FLTMC files
                $FLTMCLogs=Get-ChildItem -Path $SDDCPath -Filter "FLTMC.TXT" -Recurse -Depth 1
        # Grab the contents of each FLTMC
            $FilterData=@()
            $FilterData="PSComputerName,FilterName,NumInstances,Altitude,Frame,WindowsDriver,CompanyName`r`n"
            ForEach($FLTMCLog in $FLTMCLogs){
                $NodeName=((Split-Path -Path $FLTMCLog.FullName).Split("\")[-1]).split("_")[-1]
                $LocalFltmc=Get-FLTMC($FLTMCLog.fullname) | Select-Object @{Label='PSComputerName';Expression={$NodeName}},FilterName,NumInstances,Altitude,Frame
                # Parse the output
                ForEach($Driver in $LocalFltmc){
                    $FilterName=@()
                    $FilterName=$Driver.FilterName.Trim()
                    $FilterData0=""
                    $FilterData0+=$Driver.PSComputerName+","+$FilterName+","+$Driver.NumInstances+","+$Driver.Altitude+","+$Driver.Frame+","
                    ForEach($Line in $RawKnownFltrs){
                        If($Line -imatch '^\|\s'+[regex]::escape($FilterName)+'.sys' -or @("UnionFS","DvFlt","MdvFsFlt","applockerfltr") -contains $FilterName){
                            $S= $Line -split "\s+" , 6
                            $CompanyName=$S[5].Replace("|","").Trim()
                            If (@("UnionFS","DvFlt","MdvFsFlt","applockerfltr") -contains $FilterName) {$CompanyName="Microsoft"}
                            IF($CompanyName -inotmatch "Microsoft"){$CompanyName="RREEDD"+$CompanyName}
                            $FilterData1=""
                            $FilterData1=$S[1]+","+$CompanyName+"`r`n"
                            $FilterData+=$FilterData0+$FilterData1
                        }}
                    If($FilterData -inotmatch $FilterData0){
                        $FilterData+=$FilterData0+"YYEELLLLOOWWN/A,YYEELLLLOOWWN/A`r`n"
                    }}}
         $FilterDataOut=@()
         $FilterDataOut=$FilterData|ConvertFrom-Csv|Sort-Object FilterName,PSComputerName -Unique | Select-Object PSComputerName,FilterName,NumInstances,Altitude,Frame,CompanyName
        }


        $FilterDataOut = foreach ($item in $FLTMCXmlLogs) {
            
            if ($item.company -eq "Unknown") {
                $item.company = "YYEELLLLOOWWN" + $item.company
            } elseif ($item.company -inotmatch "Microsoft") {
                $item.company = "RREEDD" + $item.company
            }
            $item
        }

        $FilterDataOut = $FilterDataOut | Sort-Object FilterName,PSComputerName -Unique

        #HTML Report
        If($FilterDataOut.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No FLTMC found</span></h5>'}
        $html+=$FilterDataOut|ConvertTo-html -Fragment

        $html=$html`
                    -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                    -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
$ResultsSummary+=Set-ResultsSummary -name $name -html $html
$htmlout+=$html
        $html=""
        $Name=""

#Cluster heartbeat configuration
    $Name="Cluster Heartbeat Configuration"
        #Write-Host "    Gathering $Name..."
        $ClusterHeartbeatConfigurationXml=$SDDCFiles."GetCluster" |`
        Sort-Object PSComputerName | Select-Object *subnet*,RouteHistoryLength,*sit*
        $ClusterFaultDomain=$sddcfiles.GetClusterFaultDomain
        $IsStretchedCluster = ((($ClusterfaultDomain).Type.Value -eq "Site").count -gt 1)
        #Stretch Cluster
        IF($IsStretchedCluster){
            <#Check for settings
                SameSubnetThreshold = 20 
                CrossSiteDelay = 4000
                CrossSiteThreshold = 120
                CrossSubnetDelay = 4000
                SameSubnetDelay = 2000
            #>
            $ClusterHeartbeatConfigurationXml = $ClusterHeartbeatConfigurationXml | Select `
                @{L="SameSubnetThreshold";e={IF($_.SameSubnetThreshold -lt 20){'RREEDD'+$_.SameSubnetThreshold}Else{$_.SameSubnetThreshold}}},
                @{L="CrossSiteDelay";e={IF($_.CrossSiteDelay -lt 4000){'RREEDD'+$_.CrossSiteDelay}Else{$_.CrossSiteDelay}}},
                @{L="CrossSiteThreshold";e={IF($_.CrossSiteThreshold -lt 120){'RREEDD'+$_.CrossSiteThreshold}Else{$_.CrossSiteThreshold}}},CrossSubnetThreshold,
                @{L="CrossSubnetDelay";e={IF($_.CrossSubnetDelay -lt 4000){'RREEDD'+$_.CrossSubnetDelay}Else{$_.CrossSubnetDelay}}},
                @{L="SameSubnetDelay";e={IF($_.SameSubnetDelay -lt 2000){'RREEDD'+$_.SameSubnetDelay}Else{$_.SameSubnetDelay}}},
                PlumbAllCrossSubnetRoutes,RouteHistoryLength,AutoAssignNodeSite,PreferredSite
        }
        #$ClusterHeartbeatConfigurationXml | FT -AutoSize -Wrap
         #Azure Table
            $AzureTableData=@()
            $AzureTableData=$ClusterHeartbeatConfigurationXml | Select *,
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=f69DyaUGlwcc%2BUujAbpnJ%2B4VPk3PigwCgjIIa0DQCQY%3D&tn=CluChkClusterHeartbeatConfiguration'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
        $html+='<H2 id="ClusterHeartbeatConfiguration">Cluster Heartbeat Configuration</H2>'
        If($IsStretchedCluster){
            #Added for Stretch clusters
            $html+="<h5><b>Should be:</b></h5>"
            $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;SameSubnetThreshold = 20</b></h5>"
            $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;CrossSiteDelay = 4000</b></h5>"
            $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;CrossSiteThreshold = 120</b></h5>"
            $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;CrossSubnetDelay = 4000</b></h5>"
            $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;SameSubnetDelay = 2000</b></h5>"
        }
        $html+=$ClusterHeartbeatConfigurationXml | ConvertTo-html -Fragment -As List
        $html=$html `
        -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
        -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
        If($IsStretchedCluster){
            #Added for Stretch clusters
            $html+="<h5>&nbsp;&nbsp;<a href='https://infohub.delltechnologies.com/en-US/l/e2e-deployment-and-operations-guide-with-stretched-cluster-integrated-system-for-microsoft-azure-stack-hci-1/deployment-prerequisites-for-stretched-clusters-10/' target='_blank'>Ref: https://infohub.delltechnologies.com/en-US/l/e2e-deployment-and-operations-guide-with-stretched-cluster-integrated-system-for-microsoft-azure-stack-hci-1/deployment-prerequisites-for-stretched-clusters-10/</a></h5>"
        }Else{$html+="<h5>&nbsp;&nbsp;<a href='https://techcommunity.microsoft.com/t5/failover-clustering/tuning-failover-cluster-network-thresholds/ba-p/371834' target='_blank'>Ref: https://techcommunity.microsoft.com/t5/failover-clustering/tuning-failover-cluster-network-thresholds/ba-p/371834</a></h5>"}
$ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""


# SBL Disks
    $Name="Cluster Log: SBL Disks"
    $SBLDisksInfoOutAll=@()
    #Write-Host "    Gathering $Name..." 
    $html='<H2 id="ClusterLog:SBLDisks">Cluster Log: SBL Disks</H2>'
    $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;HealthCounters Key</b></h5>"
    $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;-----------------------------</b></h5>"
    $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;R/M=Read Media Errors</b></h5>"
    $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;R/U=Read Unrecoverable Errors</b></h5>"
    $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;R/T=Read Total Errors</b></h5>"
    $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;W/M=Write Media Errors</b></h5>"
    $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;W/U=Write Unrecoverable Errors</b></h5>"
    $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;W/T=Write Total Errors</b></h5>"
$Cnt=0
    ForEach($Log in $ClusterLogFiles){  
# fancy progress bar in case there large cluster files imported
$i = [math]::floor(($Cnt / $ClusterLogFiles.count) * 100)
Write-Progress -Activity "Reading Cluster Logs" -Status ("$i% Complete: "+$Cnt+"/"+$ClusterLogFiles.count) -PercentComplete $i;
$Cnt++

$LogName=$Log.Name
$SBLDisksLineNumber=(Select-String -Path $Log.FullName -Pattern '^\[===\sSBL\sDisks\s\===]'|Select-Object linenumber).LineNumber + 1
If($ClusterName.ClusterFunctionalLevel -eq 2016){
$SystemLineNumber=(Select-String -Path $Log.FullName -Pattern '^\[===\sSYSTEM\s\===]'|Select-Object linenumber).LineNumber -2
}
If($ClusterName.ClusterFunctionalLevel -ge 2019){
$SystemLineNumber=(Select-String -Path $Log.FullName -Pattern '^\[===\sCertificates\s\===]'|Select-Object linenumber).LineNumber -2
}
$SBLDisksInfo=@()
$SBLDisksInfo=(Get-Content -Path $Log.FullName | Select-Object -Index ($SBLDisksLineNumber..$SystemLineNumber)).replace(",",'","')
$SBLDisksInfoOut=@()
If($ClusterName.ClusterFunctionalLevel -eq 2016){
$SBLDisksInfoOut=$SBLDisksInfo | ConvertFrom-String -Delimiter '","' -PropertyNames DiskId, DeviceNumber, IsSblCacheDevice, HasSeekPenalty, NumPaths, PathId, CacheDeviceId, DiskState, BindingAttributes, DirtyPages, DirtySlots, IsMaintenanceMode, IsOrphan, Manufacturer, ProductId, Serial, Revision, PoolId, HealthCounters
}
ElseIf($ClusterName.ClusterFunctionalLevel -ge 2019){
$SBLDisksInfoOut=$SBLDisksInfo | ConvertFrom-String -Delimiter '","' -PropertyNames DiskId, DeviceNumber, IsSblCacheDevice, HasSeekPenalty, NumPaths, PathId, CacheDeviceId, DiskState, BindingAttributes, DirtyPages, DirtySlots, IsMaintenanceMode, IsOrphan, SblAttributes, Manufacturer, ProductId, Serial, Revision, PoolId, HealthCounters
}
$SBLDisksInfoOutAll+=$SBLDisksInfoOut|Select-Object @{L='Node';E={$LogName -replace '_.+'}},*
    }
Write-Progress -Activity "Parsing Cluster Logs" -Completed

    # Check for unbound cache disks
    #$SBLDisksInfoOutAll[0]|gm
    $SBLDisks=""
    $SBLDisks=$SBLDisksInfoOutAll|sort-Object Node,IsSblCacheDevice,DiskState | `
        Select-Object Node,ProductId,`
            @{L='Serial';E={$_.Serial -replace '\s+' }},`
            Revision,IsSblCacheDevice,`
            @{L='DiskState';E={
                    IF(($_.DiskState -replace '\s+' -inotmatch 'CacheDiskStateInitializedAndBound') -and (($PhysicalDisks.mediatype | sort -Unique).count -gt 1)){'RREEDD'+$_.DiskState}
                    Else{$_.DiskState}}},`
            @{L='IsMaintenanceMode';E={Switch($_.IsMaintenanceMode){'True'{'YYEELLLLOOWW True'}'False'{'False'}}}},IsOrphan,`
            @{L='HealthCounters';E={
                IF((([regex]'\D\/\D\s[1-9]').Matches($_.HealthCounters)).Count -gt 0){"RREEDD"+$_.HealthCounters}
                Else{$_.HealthCounters}}}
    #$SBLDisks|ft
         #Azure Table
            $AzureTableData=@()
            $AzureTableData=$SBLDisks | Select *,
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s' -replace '\:'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=vFfW%2F9D%2FlKfDY9GMCiP7gywK8e01nbDbCIKOsP9k7zo%3D&tn=CluChkClusterLogSBLDisks'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        #HTML Report
            $html+=$SBLDisks | ConvertTo-html -Fragment
            $html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://jtpedersen.com/2017/11/how-to-rebind-mirror-or-performance-drives-back-to-s2d-cache-device/' target='_blank'>Ref: Repair-ClusterStorageSpacesDirect -RecoverUnboundDrives -Verbose</a></b></h5>"
            IF((($PhysicalDisks.mediatype | sort -Unique).count -eq 1) -and ($PhysicalDisks.mediatype -ne "HDD")){
$html+="<h5><b>&nbsp;&nbsp;&nbsp;&nbsp; NOTE: No cache to bind because we only have a single media type of "+($PhysicalDisks.mediatype | sort -Unique)+" </b></h5>"
}
            $html=$html `
                -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html        
            $ResultsSummary | Export-Clixml -Path "$env:temp\ClusterResultSummary.xml" -Confirm:$false -Force
        $htmlout+=$html
            $htmlout
            $html=""
            $Name=""
}) | out-null
 $Job = $PowerShell.BeginInvoke($Inputs,$Outputs)
} 

 Write-Host "S2D Validation"
 #Write-Host "*******************************************************************************"
 #Write-Host "*                                                                             *"
 #Write-Host "*                             S2D Validation                                  *"
 #Write-Host "*                                                                             *"
 #Write-Host "*******************************************************************************"

$htmlout+='<H1 id="S2DValidation">S2D Validation</H1>'

# Find Storage NICs
    $Name="Storage Network Cards"
    Write-Host "    Gathering $Name..." 
    If($SDDCFiles.ContainsKey("ClusterNetworkLiveMigration")){
        $GetSmbMultichannelConnection=Foreach ($key in ($SDDCFiles.keys -like "*GetSmbMultichannelConnection" )) {$SDDCFiles."$key"}
        $StorageNicsFriendlyName=$GetSmbMultichannelConnection | Where-Object{$_.smbinstance -eq 2} | Select-Object @{L="StorageNicFriendlyName";E={$_.ClientInterfaceFriendlyName}},PSComputerName,ClientInterfaceIndex
        #$GetSmbMultichannelConnection |select PSComputername,ClientInterfaceFriendlyName,ClientInterfaceIndex,ServerInterfaceIndex
        $StorageNics=@()
        ForEach($StorageNic in $StorageNicsFriendlyName){
            $StorageNics+=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapter" )) {$SDDCFiles."$key" |`
            Where-Object{(($StorageNic).StorageNicFriendlyName -ieq $_.Name -and ($StorageNic).PSComputerName -ieq $key.Replace("GetNetAdapter","") -and ($StorageNic).ClientInterfaceIndex -eq $_.ifIndex )}| Select-Object @{L="ComputerName";E={($StorageNic).PSComputerName}},Name,InterfaceDescription,ifIndex,MacAddress,LinkSpeed    
        }
        }
        $StorageNics = $StorageNics| Select-Object @{L="PSComputerName";E={$_.ComputerName}},Name,InterfaceDescription,ifIndex,MacAddress,`
            @{L='LinkSpeed';E={$LinkSpeed=$_.LinkSpeed;IF((($LinkSpeed -split '\s')[0]) -lt 10){"RREEDD$LinkSpeed"}Else{$LinkSpeed}}},@{L='Combo';E={$_.ComputerName + $_.Name}}
        $StorageNics = $StorageNics |sort Combo -Unique | select-object PSComputerName,Name,InterfaceDescription,ifIndex,MacAddress,LinkSpeed
        $StorageNicsUnique=$StorageNics | Sort Name -Unique
    }
    $html+='<H2 id="StorageNetworkCards">Storage Network Cards</H2>'
    IF(-not($StorageNics)){
        $html+='<h5 style="background-color: #ffff00><b>WARNING: Missing data in SDDC. Please verify/re-run SDDC as per:</b></h5>'
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://dellservices.lightning.force.com/lightning/r/Lightning_Knowledge__kav/ka02R000000Y5fSQAS/view' target='_blank'>How to Collect Diagnostic Logs for Azure Stack HCI(S2D)</a></h5>"
    }
    IF($StorageNics){
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$StorageNics|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='Name';E={[string]$_.name}},
                @{L='InterfaceDescription';E={[string]$_.InterfaceDescription}},
                @{L='ifIndex';E={[string]$_.ifIndex}},
                @{L='MacAddress';E={[string]$_.MacAddress}},
                @{L='LinkSpeed';E={[string]$_.LinkSpeed}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=tuIBDHvdj20DngkfHw4aj3BwtFTsrJHmb9vIvEy1dtk%3D&tn=CluChkStorageNetworkCards'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}
        # HTML Report
            $html+="<h5><b>Should be:</b></h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Storage NICs 10 Gbps or faster</h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://docs.microsoft.com/en-us/windows-server/storage/storage-spaces/storage-spaces-direct-hardware-requirements#networking' target='_blank'>Ref: Storage Spaces Direct hardware requirements</a></h5>"
            $html+=$StorageNics | Sort-Object PSComputerName,Name | ConvertTo-html -Fragment
    }
    $html=$html `
      -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
      -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
    $ResultsSummary+=Set-ResultsSummary -name $name -html $html
    $htmlout+=$html
    $html=""
    $Name=""

    <# Check to see if any of the Storage NICs Mac Addresses are Virtual Nics which will tell us that its Fully Converged else NonConverged
        $FullyConverged=$False
        $NonConverged=$False
        $GetVMNetworkAdapter=Get-ChildItem -Path $SDDCPath -Filter "GetVMNetworkAdapter.xml" -Recurse -Depth 1 | import-clixml -ErrorAction Continue
        ForEach($StorageNic in $StorageNics){
            $VMStorageNics+=$GetVMNetworkAdapter| Where-Object{$_.MacAddress -imatch ($StorageNic.MacAddress -replace '-','')}
        }
        IF($VMStorageNics){$FullyConverged=$True}Else{$NonConverged=$True}
        #>

#Storage Nic Node to Node Map
     $Name="Storage Nic Node to Node Map"
     Write-Host "    Gathering $Name..."
    #$SDDCPath="C:\Users\Jim_Gandy\OneDrive - Dell Technologies\Documents\SRs\157759339\HealthTest-HCI-CLU01-20220913-1418"
    #$SDDCPath="C:\Users\Jim_Gandy\OneDrive - Dell Technologies\Documents\SRs\155793939\HealthTest-sphwdhcic1-20221130-1045"
    #$GetNetAdapter=Get-ChildItem -Path $SDDCPath -Filter GetNetadapter.xml -Recurse -Depth 2 | Import-Clixml | Where-Object{($_.InterfaceDescription -imatch "QLogic") -or ($_.InterfaceDescription -imatch "Mellanox")} 
    $GetNetAdapter=$StorageNics | sort Macaddress -Unique
    $GetIPAddress=Foreach ($key in ($SDDCFiles.keys -like "*GetNetIpAddress")) { $SDDCFiles."$key" | Where-Object {$GetNetAdapter.ifIndex -eq $_.ifIndex}}
    $GetNetNeighbor=Foreach ($key in ($SDDCFiles.keys -like "*GetNetNeighbor")) { $SDDCFiles."$key" |Where-Object {$_.state -eq 2 -or $_.state -eq 4 -or $_.state -eq 5}}
    $Table=@()
    foreach($Neighbor in $GetNetNeighbor){
        foreach($Adapter in $GetNetAdapter){
            IF($Neighbor.LinkLayerAddress -eq $Adapter.MacAddress){
                $Table  += [PSCustomObject]@{
    LocalName = $Adapter.Name
                    LocalMask = (($GetIPAddress | ?{($_.PSComputerName -eq $Adapter.PSComputerName) -and ($_.ifIndex -eq $Adapter.ifIndex) -and ($_.AddressFamily -eq "2")}).PrefixLength)
                    LocalMac =$Adapter.MacAddress
                    LocalIP = (($GetIPAddress | ?{($_.PSComputerName -eq $Adapter.PSComputerName) -and ($_.ifIndex -eq $Adapter.ifIndex) -and ($_.AddressFamily -eq "2")}).IPAddress)
                    Local = $Adapter.PSComputerName
                    Remote = $Neighbor.PSComputerName
                    RemoteIP = (($GetIPAddress | ?{($_.PSComputerName -eq $Neighbor.PSComputerName) -and ($_.ifIndex -eq $Neighbor.ifIndex) -and ($_.AddressFamily -eq "2")}).IPAddress)
                    RemoteMac = (($GetNetAdapter | ?{($_.PSComputerName -eq $Neighbor.PSComputerName) -and ($_.ifIndex -eq $Neighbor.ifIndex)}).MacAddress)
                    RemoteMask = (($GetIPAddress | ?{($_.PSComputerName -eq $Neighbor.PSComputerName) -and ($_.ifIndex -eq $Neighbor.ifIndex) -and ($_.AddressFamily -eq "2")}).PrefixLength)
                    RemoteName = $Neighbor.InterfaceAlias
                
    }
            }
        }
    }
    $StorageN2NMapOut=$Table | Sort-Object LocalMac -Unique |Sort-Object Local,remote 
    #$GetNetAdapter.Count
    # HTML Report
    $html+='<H2 id="StorageNicNodetoNodeMap">Storage Nic Node to Node Map</H2>'
    $html+= $StorageN2NMapOut| ConvertTo-html -Fragment
    $html=$html `
      -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
      -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
    $ResultsSummary+=Set-ResultsSummary -name $name -html $html
    $htmlout+=$html
    $html=""
    $Name=""

# Find the Managment NIC
    $MgmtNics=@()
    $MgmtNics=Foreach ($key in ($SDDCFiles.keys -like "*GetNetRoute")) {$SDDCFiles."$key" | `
    Where-Object{($_.DestinationPrefix -eq '0.0.0.0/0') -and (($_.RouteMetric -eq '256') -or ($_.RouteMetric -eq '1'))}
    }
    #$MgmtNics
    $MgmtNetAdapters=@()
    ForEach($MgmtNic in $MgmtNics){
        $MgmtNetAdapters+=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapter")) {$SDDCFiles."$key" |`
        Where-Object{($_.ifIndex -eq $MgmtNic.ifIndex) -and ($_.PsComputername -eq $MgmtNic.PSComputerName ) }
    }
    }
    
# Configuring the host management network as a lower-priority network for live migration

    

        $Name="Live Migration Network Priorities"
        Write-Host "    Gathering $Name..."
        #$GetClusterResource=Get-ChildItem -Path $SDDCPath -Filter "GetClusterResource.XML" -Recurse -Depth 1 | Import-Clixml
        $GetClusterNetwork=$SDDCFiles."GetClusterNetwork" | Sort Metric
        $GetNetIpAddress=Foreach ($key in ($SDDCFiles.keys -like "*GetNetIpAddress")) {$SDDCFiles."$key" | Where-Object{$_.IPv4Address}}
        $MgmtNetAdaptersIPs=@()
        # Find the Mgmt Cluster Network
            ForEach($NetIp in $GetNetIpAddress){
                $MgmtNetAdaptersIPs+=$MgmtNetAdapters | Where-Object{($_.IfIndex -eq $NetIp.IfIndex) -and ($_.PsComputername -eq $NetIp.PSComputerName )}|Select-Object *,@{L='IPv4Address';E={$NetIp.IPv4Address}},@{L='ClusterNetworkAddress';E={
                [System.Net.IPAddress]$IPAddress=$NetIp.IPv4Address
                [System.Net.IPAddress]$SubnetMask=[ipaddress]([math]::pow(2, 32) -1 -bxor [math]::pow(2, (32 - $NetIp.PrefixLength))-1)
                ([ipaddress]($IPaddress.Address -band $SubnetMask.Address)).IPAddressToString
            }}}
            #$MgmtNetAdaptersIPs | select IPv4Address,ClusterNetworkAddress
            #$CNet=@()
            $ClusterNetworkAddress=($MgmtNetAdaptersIPs|Where-Object{$_.IPv4Address.length -gt 3}|Sort-Object ClusterNetworkAddress -Unique).ClusterNetworkAddress
            $MgmtClusterNetworkId=$GetClusterNetwork | Where-Object{$ClusterNetworkAddress -eq $_.Address}
            $ClusterNetworkLiveMigration=$SDDCFiles."ClusterNetworkLiveMigration"
            #Create live migration network order

            $NetworksforLiveMigration=@()
            $NetworksforLiveMigration=($ClusterNetworkLiveMigration | Where-Object{$_.Name -eq 'MigrationNetworkOrder'}).Value -split ';'
            $NetworksforLiveMigration+=($ClusterNetworkLiveMigration | Where-Object{$_.Name -eq 'MigrationExcludeNetworks'}).Value -split ';'
            $LiveMigrationNetworkPriorities=@()
            $LiveMigrationNetworkPrioritiesOut=@()
                foreach($LMN in $NetworksforLiveMigration){
                ForEach($GCN in $GetClusterNetwork){
                    IF($LMN -eq $GCN.Id){
                        $LiveMigrationNetworkPriorities+=$GCN| Select-Object Name,@{L='LiveMigrationNetwork';E={
                            Switch($GCN.ID){
                                {(($ClusterNetworkLiveMigration | Where-Object{$_.Name -eq 'MigrationExcludeNetworks'}|Select-Object Value) -split ';') | Where-Object{$_ -iMatch $GCN.ID}}{
                                   


                                   
                                   
                                    "Excluded"
                                }
                                {(($ClusterNetworkLiveMigration | Where-Object{$_.Name -eq 'MigrationNetworkOrder'}|Select-Object Value) -split ';') | Where-Object{$_ -iMatch $GCN.ID}}{
                                    "Included"
                                }
            }}},Address}}}
            #$LiveMigrationNetworkPriorities | FT
# check if mgmt is last in MigrationNetworkOrder
            $IsMgmtInMigrationNetworkOrder=$LiveMigrationNetworkPriorities | Where-Object{$_.Name -imatch $MgmtClusterNetworkId.name}
  IF($IsMgmtInMigrationNetworkOrder.LiveMigrationNetwork -eq "Included" -and -not($MgmtClusterNetworkId.ID -ieq (($ClusterNetworkLiveMigration | Where-Object{$_.Name -eq 'MigrationNetworkOrder'}|Select-Object -ExpandProperty Value) -split ';')[-1])){
 $LiveMigrationNetworkPrioritiesOut+=$LiveMigrationNetworkPriorities | Select-Object Name,@{L='LiveMigrationNetwork';E={
 IF($_.Name -imatch $MgmtClusterNetworkId.name){
 "RREEDD"*($SysInfo[0].SysModel -notmatch "^PowerEdge")+$_.LiveMigrationNetwork } else{$_.LiveMigrationNetwork}
 }}
 }Else{$LiveMigrationNetworkPrioritiesOut=$LiveMigrationNetworkPriorities}

  $LiveMigrationNetworkPrioritiesOut=$LiveMigrationNetworkPrioritiesOut | Select-Object Name,@{L='LiveMigrationNetwork';E={
  IF($_.Name -inotmatch $MgmtClusterNetworkId.name -and $_.LiveMigrationNetwork -eq "Excluded"){$curraddr=$_.address;If (($GetClusterNetwork | ? Address -match $curraddr).role.value -eq 'Cluster'){
  "RREEDD"*($SysInfo[0].SysModel -notmatch "^PowerEdge")+$_.LiveMigrationNetwork } else{$_.LiveMigrationNetwork}}
 }},Address



            #$LiveMigrationNetworkPrioritiesOut
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$LiveMigrationNetworkPrioritiesOut|Select-Object *,
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=e9DBnTajdhCb19EXtDam4hsC6sLTLnV8W0uYZ1gZgGY%3D&tn=CluChkLiveMigrationNetworkPriorities'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}
        # HTML Report
    $html+='<H2 id="LiveMigrationNetworkPriorities">Live Migration Network Priorities</H2>'
    $html+="<h5><b>Should be:</b></h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Managment Network LiveMigrationNetwork = Excluded or Last</h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Cluster only networks are usually LiveMigrationNetwork = Included</h5>"
If($LiveMigrationNetworkPrioritiesOut.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No LiveMigration Network Priorities Entries found</span></h5>'}
    $html+=$LiveMigrationNetworkPrioritiesOut | ConvertTo-html -Fragment
    $html=$html `
      -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
      -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
    $ResultsSummary+=Set-ResultsSummary -name $name -html $html
    $htmlout+=$html
    $html=""
    $Name=""
        

#Gathering Support Matrix for Dell EMC Solutions for Microsoft Azure Stack HCI
    $SMFWDRVRData=@()

# parse Network adapter table
    $SMFWDRVRTable=@()
    $SMFWDRVRTable=$SupportMatrixtableData['Network Adapters']
    #$SMFWDRVRTable.count

    $resultObject=@()
    $SMNicData=@()
$previousVersion = ''
$previousOS = ''
$previousPlatform = ''
$previousComponent= ''
$previousPN=''
    ForEach($SMNicData in $SMFWDRVRTable[0]){
if (!($SMNicData | gm | ? Name -match 'Component')) {$SMNicData | Add-Member -NotePropertyName Component -NotePropertyValue ''}
if (!($SMNicData | gm | ? Name -match 'Part Number')) {$SMNicData | Add-Member -NotePropertyName 'Part Number' -NotePropertyValue ''}
if ($SMNicData.'Driver Minimum Supported Version'.length -eq 0) { $SMNicData.'Driver Minimum Supported Version' = $previousVersion }
if ($SMNicData.'Supported OS'.length -eq 0) { $SMNicData.'Supported OS' = $previousOS }
if ($SMNicData.'Supported Platforms'.length -eq 0) { $SMNicData.'Supported Platforms' = $previousPlatform }
if ($SMNicData.Component.length -eq 0) { $SMNicData.Component = $previousComponent }
if ($SMNicData.'Part Number'.length -eq 0) { $SMNicData.'Part Number' = $previousPN }


<#
        # manual table fixes, the support site sometimes does not have the correct driver version listed
# Mellanox package MXTKX, 03.00.01 -> Driver 3.0.25668.0 (package version listed, not driver version)
# BroadCom package VH7RD, 22.31.11.31 -> Driver 223.0.157.0 (package version listed, not driver version)
# BroadCom package 5K4FP, 223.116.1 -> Driver 221.0.157.0 (wrong version listed)
switch ($SMNicData['Driver Software Bundle']) {
'MXTKX' { $SMNicData.'Driver Minimum Supported Version' = '3.0.25668.0' } 
'VH7RD' { $SMNicData.'Driver Minimum Supported Version' = '223.0.157.0' } 
'5K4FP' { $SMNicData.'Driver Minimum Supported Version' = '221.0.5.0' } 
} # switch #>

        $resultObject += [PSCustomObject] @{
                Component                       = $SMNicData.Component
                PartNumber                      = $SMNicData.'Part Number'
                SDDCAQforWindowsServer2022      = $SMNicData.'SDDC AQ for Windows Server 2022'
                RDMAProtocol                    = $SMNicData.'RDMA Protocol'
                FirmwareSoftwareBundle          = $SMNicData.'Firmware Software Bundle'
                FirmwareMinimumSupportedVersion = $SMNicData.'Firmware Minimum Supported Version'
                DriverSoftwareBundle            = $SMNicData.'Driver Software Bundle'
                DriverMinimumSupportedVersion   = $SMNicData.'Driver Minimum Supported Version'
                SupportedOS                     = $SMNicData.'Supported OS'
                SupportedPlatforms              = $SMNicData.'Supported Platforms'
        }

# save previous row value
$previousVersion = $SMNicData.'Driver Minimum Supported Version'
$previousOS = $SMNicData.'Supported OS'
$previousPlatform = $SMNicData.'Supported Platforms'
$previousComponent = $SMNicData.Component
$previousPN = $SMNicData.'Part Number'

    }

    $SMFWDRVRData+=$resultObject

# parse storage controllers
    $SMFWDRVRTable=@()
    $SMFWDRVRTable=$SupportMatrixtableData['Storage Controllers']
    #$SMFWDRVRTable.count

    $resultObject=@()
    $SMStorageCtrlData=@()
$previousVersion = ''
$previousOS = ''
$previousPlatform = ''
    ForEach($SMStorageCtrlData in $SMFWDRVRTable[0]){

if ($SMStorageCtrlData.'Driver Minimum Supported Version'.length -eq 0) { $SMStorageCtrlData.'Driver Minimum Supported Version' = $previousVersion }
if ($SMStorageCtrlData.'Supported OS'.length -eq 0) { $SMStorageCtrlData.'Supported OS' = $previousOS }
if ($SMStorageCtrlData.'Supported Platforms'.length -eq 0) { $SMStorageCtrlData.'Supported Platforms' = $previousPlatform }

        $resultObject += [PSCustomObject] @{
                Component                       = $SMStorageCtrlData.Component
                PartNumber                      = $SMStorageCtrlData.'Dell Part Number'
                FirmwareSoftwareBundle          = $SMStorageCtrlData.'Firmware Software Bundle'
                FirmwareMinimumSupportedVersion = $SMStorageCtrlData.'Firmware Minimum Supported Version'
                DriverSoftwareBundle            = $SMStorageCtrlData.'Driver Software Bundle'
                DriverMinimumSupportedVersion   = $SMStorageCtrlData.'Driver Minimum Supported Version'
                SupportedOS                     = $SMStorageCtrlData.'Supported OS'
                SupportedPlatforms              = $SMStorageCtrlData.'Supported Platforms'
                
        }
$previousVersion = $SMStorageCtrlData.'Driver Minimum Supported Version'
$previousOS = $SMStorageCtrlData.'Supported OS'
$previousPlatform = $SMStorageCtrlData.'Supported Platforms'
    }

    $SMFWDRVRData+=$resultObject

# parse base components
    $SMFWDRVRTable=@()
    $SMFWDRVRTable=$SupportMatrixtableData['Base Components']
    #$SMFWDRVRTable.count

    $resultObject=@()
    $SMBaseData=@()
$previousVersion = ''
$previousOS = ''
$previousPlatform = ''
    ForEach($SMBaseData in $SMFWDRVRTable[0]){

if ($SMBaseData.'Driver Minimum Supported Version'.length -eq 0) { Add-Member -Force -InputObject $SMBaseData -MemberType NoteProperty -Name 'Driver Minimum Supported Version' -Value $previousVersion }
if ($SMBaseData.'Supported OS'.length -eq 0) { Add-Member -Force -InputObject $SMBaseData -MemberType NoteProperty -Name 'Supported OS' -Value $previousOS }
if ($SMBaseData.'Supported Platforms'.length -eq 0) { Add-Member -Force -InputObject $SMBaseData -MemberType NoteProperty -Name 'Supported Platforms' -Value $previousPlatform }

        $resultObject += [PSCustomObject] @{
                Component                       = $SMBaseData.Component
                DriverSoftwareBundle            = $SMBaseData.'Software Bundle'
                DriverMinimumSupportedVersion   = $SMBaseData.'Minimum Supported Version'
                SupportedOS                     = $SMBaseData.'Supported OS'
                SupportedPlatforms              = $SMBaseData.'Supported Platforms'
                
        }
$previousVersion = $SMBaseData.'Driver Minimum Supported Version'
$previousOS = $SMBaseData.'Supported OS'
$previousPlatform = $SMBaseData.'Supported Platforms'
    }

    $SMFWDRVRData+=$resultObject

if ($debug) { $SMFWDRVRData| select Component, DriverMinimumSupportedVersion, SupportedPlatforms, SupportedOS | ft}

function get-DriverVersion {
    [CmdletBinding()] 
        param(
            [Parameter(Mandatory = $true)]
            [string] $DriverName,

            [Parameter(Mandatory = $true)]
            [string] $OSversion
        )
$OSversion = ('*'+$OSversion+'*')

# there is no check on supported platforms

return ($SMFWDRVRData |Where-Object{$_.SupportedOS -like $OSversion} | Where-Object{$_.Component -like $DriverName} | Select-Object DriverMinimumSupportedVersion | Sort-Object -Unique).DriverMinimumSupportedVersion

}

#Update Out of Box drivers
    #Add baseline based on matrix or s2d catalog. Maybe add this to DriFT if the SDDC is in the same folder as the TSRs then use it for driver info
    $Name="Update Out of Box drivers"
    Write-Host "    Gathering $Name..."  
    #$UpdateOutofBoxdrivers=@()
    #$UpdateOutofBoxdrivers | gm
    $UpdateOutofBoxdriversOut=@()
    $DriverSuites=@()
    $DriverSuites=Foreach ($key in ($SDDCFiles.keys -like "*GetDriverSuiteVersion")) {$SDDCFiles."$key" |`
     Select-Object @{Label="PSComputerName";Expression={$key -replace "GetDriverSuiteVersion",""}},`
     @{Label='Qlogic';Expression={($_ | Where PSPath -like "*Qlogic*Display*").'(default)'}},`
     @{Label='Intel';Expression={($_ | Where PSPath -like "*Intel*Product_Version").'(default)'}},`
     @{Label='Mellanox';Expression={($_ | Where PSChildName -like "*MLNX*").'(default)'}},`
     @{Label='Broadcom';Expression={$bcomver=($_ | Where PSPath -like "*BroadcomNXE*Product_Version").'(default)';If ($bcomver.indexof("$([char]0)") -eq -1) {$bcomver} else {$bver.substring(0,$bcomver.indexof("$([char]0)"))}}},`
     @{Label='B1GB';Expression={($_ | Where PSPath -like "*Broadcom\*Product_Version").'(default)'}}
    }
    $DriverSuites=$DriverSuites | Select PSCOmputerName,@{Label='Qlogic';Expression={foreach ($a in $DriverSuites) {if ($a.PSComputerName -eq $_.PSComputerName -and ($a.Qlogic)) {$a.Qlogic}}}},`
    @{Label='Intel';Expression={foreach ($a in $DriverSuites) {if ($a.PSComputerName -eq $_.PSComputerName -and ($a.Intel)) {$a.Intel}}}},`
    @{Label='Mellanox';Expression={foreach ($a in $DriverSuites) {if ($a.PSComputerName -eq $_.PSComputerName -and ($a.Mellanox)) {$a.Mellanox}}}},`
    @{Label='B1GB';Expression={foreach ($a in $DriverSuites) {if ($a.PSComputerName -eq $_.PSComputerName -and ($a.B1GB)) {$a.B1GB}}}},`
    @{Label='Broadcom';Expression={foreach ($a in $DriverSuites) {if ($a.PSComputerName -eq $_.PSComputerName -and ($a.Broadcom)) {$a.Broadcom}}}} |`
    Sort PSComputerName,Qlogic,Intel -Unique
    $DriverSuites=$DriverSuites | Select PSComputerName,Qlogic,Intel,Mellanox,Broadcom,B1GB,` #Fix 06072023
     @{Label='Chipset';Expression={($SDDCFiles."$($_.PSComputerName)GetChipsetVersion").Version | Select -First 1}}
 #debug tests
#get-DriverVersion -DriverName "Broadcom 57414*" -OSVersion $OSVersionNodes
#$SMFWDRVRData | Where-Object{$_.Component -like "Broadcom*25*LOM*"} 

    $UpdateOutofBoxdrivers=Foreach ($key in ($SDDCFiles.keys -like "*GetDrivers")) {$SDDCFiles."$key" |`
    Where-Object{$_.Manufacturer -notmatch "Microsoft"}|`
    Where-Object{$_.DeviceName -match "HBA"`
-or($_.DeviceName -match "Mellanox")`
-or($_.DeviceName -match "Marvell")`
-or(($_.DeviceName -match "Chipset") -and ($_.DeviceName -match "Controller"))`
-or($_.DeviceName -match "Ethernet")`
-or(($_.DeviceName -match "FastLinQ") -and ($_.DeviceName -match "VBD"))`
-or($_.DeviceName -match "giga")`
-or($_.DeviceName -match "PERC")`
-or($_.DeviceName -match "Matrox")}|`
    Select-Object PSComputerName, DeviceName,`
    @{Label='DriverVersion';Expression={If (-not $DriverSuites) {$_.DriverVersion} else {foreach ($a in $DriverSuites) {If ($a.PSComputerName -eq $_.PSComputerName) {
        if ($_.DeviceName -like "Qlogic FastLinQ*41*" -and $a.Qlogic){
            $a.Qlogic} else {
            if (($_.DeviceName -like "Intel*1Gb*Ethernet*NDC*" -or $_.DeviceName -like "*X710*") -and $a.Intel){
            $a.Intel} else {
            if (($_.DeviceName -like "Mellanox ConnectX*") -and $a.Mellanox){
            $a.Mellanox} else {
            if (($_.DeviceName -like "Broadcom*574*" -or $_.DeviceName -like "Broadcom NetXtreme E*"-or $_.DeviceName -like "Broadcom*OCP*") -and $GenerationNodes -eq "14g" -and $a.Broadcom){
            $a.Broadcom} else {
            if (($_.DeviceName -like "Broadcom NetXtreme Gigabit*") -and $a.B1GB){
            $a.B1GB} else {
            if ($_.DeviceName -like "Intel*C620*Chipset*" -and $a.Chipset){ #Fix 06072023
            $a.Chipset} else {
            $_.DriverVersion
        }}}}}}
    }}}}},`
    @{Label='AvailableVersion';Expression={Switch -Wildcard ($_.DeviceName){
    #Driver not on support matrix web page
    'QLogic 57800*'               {'7.13.171.0'}`

    'QLogic FastLinQ*Adapter' {get-DriverVersion -DriverName "Qlogic FastLinQ 41262*SFP28*" -OSVersion $OSVersionNodes}`
    'QLogic FastLinQ*VBD*'    {get-DriverVersion -DriverName "Qlogic FastLinQ 41262*SFP28*" -OSVersion $OSVersionNodes}`

    'Mellanox ConnectX-4*'    {get-DriverVersion -DriverName "Mellanox ConnectX-4*" -OSVersion $OSVersionNodes}`
    'Mellanox Connectx-5*'    {get-DriverVersion -DriverName "Mellanox Connectx-5*" -OSVersion $OSVersionNodes}`
    'Mellanox ConnectX-6*'    {get-DriverVersion -DriverName "Mellanox ConnectX-6*" -OSVersion $OSVersionNodes}`

    'Intel*Gigabit*I350*'     {get-DriverVersion -DriverName "Intel*1Gb*Ethernet*NDC*" -OSVersion $OSVersionNodes}`
    'Intel*Ethernet*X710*'    {get-DriverVersion -DriverName "*X710*" -OSVersion $OSVersionNodes}`

    'Broadcom*BCM57412*'      {get-DriverVersion -DriverName "Broadcom 57412*" -OSVersion $OSVersionNodes}`
    'Broadcom BCM5720*'       {get-DriverVersion -DriverName "Broadcom 5720*" -OSVersion $OSVersionNodes}`
    'Broadcom*57414*'         {get-DriverVersion -DriverName "Broadcom 57414*" -OSVersion $OSVersionNodes}`
    'Broadcom*57416*'         {get-DriverVersion -DriverName "Broadcom 57416*" -OSVersion $OSVersionNodes}`
    'Broadcom NetXtreme Gigabit*' {get-DriverVersion -DriverName "Broadcom 5720*" -OSVersion $OSVersionNodes}`
'Broadcom*E-Series*10Gb*'      {get-DriverVersion -DriverName "Broadcom*10*LOM*" -OSVersion $OSVersionNodes}`
'Broadcom*E-Series*25Gb*'      {get-DriverVersion -DriverName "Broadcom*25*LOM*" -OSVersion $OSVersionNodes}`

'*HBA330*'                {get-DriverVersion -DriverName "HBA330*" -OSVersion $OSVersionNodes}`
'*HBA355*'                {get-DriverVersion -DriverName "HBA355*" -OSVersion $OSVersionNodes}`

    '*BOSS*'                  {get-DriverVersion -DriverName "BOSS-S1*" -OSVersion $OSVersionNodes}`
    '*Unify*'                 {get-DriverVersion -DriverName "BOSS-S2*" -OSVersion $OSVersionNodes}`
      
    '*AMD*Chipset*'           {get-DriverVersion -DriverName "Chipset*AMD*" -OSVersion $OSVersionNodes}`
    'Intel*C620*Chipset*'  {
    $searchTxt = "Chipset*driver*14G*Intel*"
# newer AX models use a different package, ero different search string
if (($ClusterNodes[0].model -eq "AX-650") -or ($ClusterNodes[0].model -eq "AX-750")) {
$searchTxt = ("*Chipset*Driver*15G*Intel*") #Fix 06072023
}
get-DriverVersion -DriverName $searchTxt -OSVersion $OSVersionNodes
  }` #Intel*C620*Chipset*
    }}}, Manufacturer
    }

if ($debug) { $UpdateOutofBoxdrivers | ft }

    # Check if DriverVersion is less than or greater than the AvailableVersion
$DriverVersionCheck=@()
ForEach($Device in $UpdateOutofBoxdrivers){
$DriverVersionCheck+=$Device | Select-Object PSComputerName, DeviceName, `
@{Label='DriverVersion';Expression={
$DriverVersion=$_.DriverVersion        
# skip check on inbox driver, this is always correct
IF(($_.AvailableVersion -ne $null) -and ($_.AvailableVersion -notmatch 'inbox')){
Switch($_.AvailableVersion){
# DriverVersion < AvailableVersion
{[System.Version]$DriverVersion -lt [System.Version]$_}{"RREEDD$DriverVersion"}
# DriverVersion > AvailableVersion
{[System.Version]$DriverVersion -gt [System.Version]$_}{"YYEELLLLOOWW$DriverVersion"}
Default{$DriverVersion}
}
}Else{$DriverVersion}
}
},@{Label='AvailableVersion';Expression={
IF($_.AvailableVersion -ne $Null){$_.AvailableVersion+" *"}Else{"Not Available"}
}
},Manufacturer
}
if ($debug) { $DriverVersionCheck | FT -AutoSize }
    # Remove dups
    #$DriverVersionCheck | sort PSComputerName,DeviceName 
    $UpdateOutofBoxdrivers=$DriverVersionCheck

    # Add new output format
    $UpdateOutofBoxdriverstbl = New-Object System.Data.DataTable "OOBDrivers"
    $UpdateOutofBoxdriverstbl.Columns.add((New-Object System.Data.DataColumn("Manufacturer")))
    $UpdateOutofBoxdriverstbl.Columns.add((New-Object System.Data.DataColumn("DeviceName")))
    $UpdateOutofBoxdriverstbl.Columns.add((New-Object System.Data.DataColumn("AvailableVersion"))) 

ForEach ($a in ($UpdateOutofBoxdrivers.PSComputerName | Sort-Object -Unique)){
$UpdateOutofBoxdriverstbl.Columns.Add((New-Object System.Data.DataColumn([string]$a)))
}

ForEach ($a in ($UpdateOutofBoxdrivers | Sort-Object DeviceName -unique)){
IF (($a.DeviceName.length -gt 2) -and ($a.DeviceName -inotmatch 'System.__ComObject')) {
$row=$UpdateOutofBoxdriverstbl.NewRow()
            $row["Manufacturer"]=$a.Manufacturer
            $row["DeviceName"]=$a.DeviceName
            $row["AvailableVersion"]=$a.AvailableVersion

ForEach($b in ($UpdateOutofBoxdrivers | where-object {$_.DeviceName -eq $a.DeviceName})){ 
$row["$($b.PSComputerName)"] = $b.DriverVersion
}
$UpdateOutofBoxdriverstbl.rows.add($row)
}
}

    #$UpdateOutofBoxdriverstbl |Format-Table 
    $UpdateOutofBoxdriversOut = $UpdateOutofBoxdriverstbl|Where-Object{$_.Manufacturer -inotmatch 'System.__ComObject'}|Sort-Object Manufacturer, DeviceName | Select-object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$UpdateOutofBoxdrivers|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='DeviceName';E={[string]$_.DeviceName}},
                @{L='DriverVersion';E={[string]$_.DriverVersion}},
                @{L='AvailableVersion';E={[string]$_.AvailableVersion}},
                @{L='Manufacturer';E={[string]$_.Manufacturer}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=Wt%2FZ7xUDpvgons3ycAeY5VoBEv%2FjkwVy%2FZmkHMmTmPo%3D&tn=CluChkUpdateOutofBoxdrivers'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
    $html+='<H2 id="UpdateOutofBoxdrivers">Update Out of Box drivers</H2>'
    $html+="&nbsp;&nbsp;&nbsp;&nbsp;Drivers should be listed once else drivers not same on all nodes"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;<a href='$SMRevHistLatest' target='_blank'>Ref: Support Matrix for Dell EMC Solutions for Microsoft Azure Stack HCI</a></h5>"
    
    $html+=$UpdateOutofBoxdriversOut | sort DeviceName -Unique | ConvertTo-html -Fragment
    $html=$html `
      -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
      -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
    $html+="&nbsp;*Available versions from Support Matrix Revision: "+($SMRevHistLatest.split('/')[-2])
    IF($SysInfo[0].SysModel -match "^APEX" -or $OSVersionNodes -eq "23H2") {
        $html+="<h5><b>&nbsp;!!! This script does not check against SBE package versions !!!</b></h5>"
    }
    $ResultsSummary+=Set-ResultsSummary -name $name -html $html
    $htmlout+=$html
    $html=""
    $Name=""

    # Clean up Chipset files
    $Destination="$ENV:TEMP\IntelChipsetDrvr"
# Remove only when folder is exist
If(Test-Path $Destination){
Get-ChildItem -Path $Destination -Recurse | Remove-Item -force -recurse -ErrorAction SilentlyContinue
Remove-Item $Destination -Force -ErrorAction SilentlyContinue
}
    
 Write-Host "Total Time here is $(((Get-Date)-$dstartSDDC).totalmilliseconds)"
#region Recommended updates and hotfixes for Windows Server 
        $dstart=Get-Date
        $Name="Recommended Updates and Hotfixes"
        Write-Host "    Gathering $Name..."  
        $OSVersion=$OSVersionNodes
        If ($SysInfo[0].AzureLocalVersion -gt "") {
           Invoke-WebRequest -Uri (Invoke-WebRequest -Uri "https://aka.ms/AzureEdgeUpdates" -MaximumRedirection 5 -ErrorAction Stop -UseBasicParsing).BaseResponse.ResponseUri.AbsoluteUri -OutFile $env:temp\outfile.xml
           $SUVersions=(([xml](Get-Content $env:temp\outfile.xml)).ASZSolutionBundleUpdates.ApplicableUpdate | sort version) | select-object version,type,family,@{L='RequiredSBE';E={$s=$_.validatedconfigurations.requiredpackages.package | ? Type -eq SBE;($s.version | sort -Unique) -join ","}},@{L='RequiredSolution';E={$s=$_.validatedconfigurations.requiredpackages.package | ? Type -eq Solution;($s.version | sort -Unique) -join ","}} 
           Invoke-WebRequest -Uri "https://aka.ms/AzureStackSBEUpdate/DellEMC" -UseBasicParsing -OutFile $env:temp\outfile.xml
           $SBEVersions=(([xml](Get-Content $env:temp\outfile.xml)).SBEUpdatesManifest.ApplicableUpdate | sort version) | select-object version,type,family,@{L='RequiredSBE';E={$s=$_.validatedconfigurations.requiredpackages.package | ? Type -eq SBE;($s.version | sort -Unique) -join ","}},@{L='RequiredSolution';E={$s=$_.validatedconfigurations.requiredpackages.package | ? Type -eq Solution;($s.version | sort -Unique) -join ","}} | ? Family -match $GenerationNodes
           $CurrentUpdatesandHotfixes=Foreach ($key in ($SDDCFiles.keys -like "*GetStampInformation")) { $SDDCFiles."$key" |`
              Select-Object @{Label="PSComputerName";Expression={$key.replace("GetStampInformation","")}},@{Label='SBE Version';Expression={$sbever=$_.OemVersion;"YYEELLLLOOWW"*([Version]$SBEVersions[-1].Version -ne [version]$sbever)+$sbever}},
             @{Label='MS Solution';Expression={$suver=$_.StampVersion;"YYEELLLLOOWW"*([Version]$SUVersions[-1].Version -ne [version]$suver)+$suver}},@{Label='Deployed Version';Expression={$_.InitialDeployedVersion}},
             @{Label='Next SBE';Expression={if ([Version]$SBEVersions[-1].Version -ne [version]$_.OemVersion) {($SBEVersions | ? RequiredSBE -match $_.OemVersion | ? RequiredSolution -match (([version]$_.StampVersion).major.tostring()+"."+([version]$_.StampVersion).minor))[-1].Version}}},
             @{Label='Next MS SU';Expression={if ([Version]$SUVersions[-1].Version -ne [version]$_.StampVersion) {($SUVersions | ? RequiredSolution -match (([version]$_.StampVersion).major.tostring()+"."+([version]$_.StampVersion).minor))[-1].Version}}}


           }
           
           
           
           $html+='<H2 id="RecommendedUpdatesandHotfixes">Recommended Updates and Hotfixes</H2>'
           $html+="<h5><b>Should be:</b></h5>"
           $html+="<h5>&nbsp;&nbsp;Latest Solution Update from MS: $($SUVersions[-1].Version)</h5>"
           $html+="<h5>&nbsp;&nbsp;Latest SBE Update from Dell: $($SBEVersions[-1].Version)</h5>"
           $html+="<h5><br>&nbsp;&nbsp;Currently installed:</h5>"
           $html+=$CurrentUpdatesandHotfixes | sort-object -Property @{Expression={$_.PSComputerName}; Ascending=$true} | ConvertTo-html -Fragment
        } else {
        
#Returns the download link of KB from Microsoft catalog
Add-Type @"
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
//using System; //for console writeline

public static class GetKBDLLink
{
    public static string GetDownloadLink(string KBNumber, string Product)
    {
        string kbGUID = "";
        string kbDLUriSource = "";

        // Search for URI to retrieve the latest KB information
        // Extracting the KBGUID from the KBPage
        var webRequest = WebRequest.Create("https://www.catalog.update.microsoft.com/Search.aspx?q=" + KBNumber);
        webRequest.Method = "GET";
        var webResponse = webRequest.GetResponse();
        var responseStream = webResponse.GetResponseStream();
        var streamReader = new StreamReader(responseStream);
        string responseContent = streamReader.ReadToEnd();
        //Console.WriteLine("Content is "+responseContent);
        var kbMatches=Regex.Matches(responseContent, @"id=(?:""|')(.*?)(?=_link)([^\/]*)");
        //Console.WriteLine(kbMatches.Count);
        foreach (Match ItemMatch in kbMatches)
        {
        //Console.WriteLine(Product);
        //Console.WriteLine(ItemMatch.Groups[2].Value);
            if (ItemMatch.Groups[2].Value.Contains(Product))
            {
                 kbGUID = ItemMatch.Groups[1].Value;
            }
        }
        //Console.WriteLine(kbGUID);

        // Use the KBGUID to find the actual download link for the KB
        string post1 = "https://www.catalog.update.microsoft.com/DownloadDialog.aspx?updateIDs=[{%22size%22%3A0%2C%22languages%22%3A%22%22%2C%22uidInfo%22%3A%22";
        string post2 = "%22%2C%22updateID%22%3A%22";
        string post3 = "%22}]&updateIDsBlockedForImport=&wsusApiPresent=&contentImport=&sku=&serverName=&ssl=&portNumber=&version=";
        string postText = post1 + kbGUID + post2 + kbGUID + post3;
        //Console.WriteLine(postText);
        webRequest = WebRequest.Create(postText);
        webRequest.Method = "GET";
        webResponse = webRequest.GetResponse();
        responseStream = webResponse.GetResponseStream();
        streamReader = new StreamReader(responseStream);
        responseContent = streamReader.ReadToEnd();
        kbDLUriSource = Regex.Match(responseContent, @"(?<=downloadInformation\[0\].files\[0\].url = '|"")(.*?)(?='|"";)").Groups[1].Value;

        return kbDLUriSource;
    }
}
"@ 



#$KBNumber = "123456"
#$Product = "Windows 10"
#$downloadLink = [GetKBDLLink]::GetDownloadLink($KBNumber, $Product)
#Write-Host "Download link: $downloadLink"


        $KBLatest=''
$KBList=''
$KBItemsToShow = 6

        #Lastest hotfix for Windows Server from the respective KB pages
            If($OSVersion -imatch '2008r2'-or $OSVersion -imatch '2008 r2'){\
                $OSVersion ='2008 r2'
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/4009469"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Load the HTML into an HtmlDocument object
                #$htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
                #$htmlDoc.LoadHtml($htmlpage)


                # Find all elements with the "supLeftNavLink" class
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")[^>]*>(.*?)(KB\d{7})(.*?)<\/a>')

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -split "&#x2014;")[0]}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[4].Value}},
                    @{L="OS Build";E={"6.1.7601"}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            }

            If($OSVersion -imatch '2012r2'-or $OSVersion -imatch '2012 r2'){
                $OSVersion ='2012 r2'
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/4009470"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Load the HTML into an HtmlDocument object
                #$htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
                #$htmlDoc.LoadHtml($htmlpage)

                # Find all elements with the "supLeftNavLink" class
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")[^>]*>(.*?)(KB\d{7})(.*?)<\/a>')

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -split "&#x2014;")[0]}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[4].Value}},
                    @{L="OS Build";E={"6.3.9600"}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            }
    
            If($OSVersion -imatch '2016'){
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/4000825"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Load the HTML into an HtmlDocument object
                #$htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
                #$htmlDoc.LoadHtml($htmlpage)

                # Find all elements with the "supLeftNavLink" class
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")>(.*?)(KB\d{7})\D+((?:(?!Preview).)14393.*?)\)(?:(?!Preview).)*<\/a>')

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -replace "&#x2014;"," ")}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[3].Value+$_.Groups[4].Value}},
                    @{L="OS Build";E={$_.Groups[4].Value.Trim()}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            
                <#$links = $htmlDoc.DocumentNode.SelectNodes("//*[@class='supLeftNavLink']")

                $KBList  = $Links | Where-Object{($_.InnerText -imatch "KB") -and ($_.InnerText -imatch '14393') -and ($_.InnerText -notmatch 'Preview')} | sort InnerStartIndex | Select -first $KBItemsToShow `
                    @{L='KBNumber';E={(((($_.InnerText -split "&#x2014;")[-1]) -split '\(')[0].Trim() -split '\s')[0]}},`
                    @{L='Date';E={($_.InnerText -split "&#x2014;")[0]}},`
                    @{L="Description";E={($_.InnerText -replace "&#x2014;"," ")}},
                    @{L="OS Build";E={(((($_.InnerText -split "OS Build")[-1]) -replace '\)') -replace 'OS' -replace 's' -replace ' Preview' -replace ' Out-of-band' -replace ' Update for Windows 10 Mobile').trim()}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.OuterHtml -split 'href="')[-1] -split '">')[0]}},
                    @{L="DownloadLink";E={[GetKBDLLink]::GetDownloadLink((((($_.InnerText -split "&#x2014;KB")[-1]) -split '\(')[0].Trim() -split '\s')[0]),$OSVersion}}#>
            }

            If($OSVersion -imatch '2019'){
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/4464619/windows-10-update-history"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Load the HTML into an HtmlDocument object
                #$htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
                #$htmlDoc.LoadHtml($htmlpage)

                # Find all elements with the "supLeftNavLink" class
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")>(.*?)(KB\d{7})\D+((?:(?!Preview).)17763.*?)\)(?:(?!Preview).)*<\/a>')

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -replace "&#x2014;"," ")}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[3].Value+$_.Groups[4].Value}},
                    @{L="OS Build";E={$_.Groups[4].Value.Trim()}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            }

            <#    $links = $htmlDoc.DocumentNode.SelectNodes("//*[@class='supLeftNavLink']")

                $KBList  = $Links | Where-Object{($_.InnerText -imatch "KB") -and ($_.InnerText -imatch '17763') -and ($_.InnerText -notmatch 'Preview')} | sort InnerStartIndex | Select -first $KBItemsToShow `
                    @{L='KBNumber';E={(((($_.InnerText -split "&#x2014;")[-1]) -split '\(')[0].Trim() -split '\s')[0]}},`
                    @{L='Date';E={($_.InnerText -split "&#x2014;")[0]}},`
                    @{L="Description";E={($_.InnerText -replace "&#x2014;"," ")}},
                    @{L="OS Build";E={(((($_.InnerText -split "OS Build")[-1]) -replace '\)') -replace 'OS' -replace 's' -replace ' Preview' -replace ' Out-of-band' -replace ' Update for Windows 10 Mobile').trim()}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.OuterHtml -split 'href="')[-1] -split '">')[0]}},
                    @{L="DownloadLink";E={[GetKBDLLink]::GetDownloadLink((((($_.InnerText -split "&#x2014;KB")[-1]) -split '\(')[0].Trim() -split '\s')[0]),$OSVersion}}#>
            


            <#If($OSVersion -imatch '20H2'){
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/4595086"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Load the HTML into an HtmlDocument object
                #$htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
                #$htmlDoc.LoadHtml($htmlpage)

                # Find all elements with the "supLeftNavLink" class
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")[^>]*>((?:(?!preview).)*?)(KB\d{7})(.*?)<\/a>')

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -replace "&#x2014;"," ")}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[3].Value+$_.Groups[4].Value}},
                    @{L="OS Build";E={"20348"}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            }

            If($OSVersion -imatch '21H2'){
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/5004047"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Load the HTML into an HtmlDocument object
                #$htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
                #$htmlDoc.LoadHtml($htmlpage)

                # Find all elements with the "supLeftNavLink" class
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")[^>]*>((?:(?!preview).)*?)(KB\d{7})(.*?)<\/a>')

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -replace "&#x2014;"," ")}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[3].Value+$_.Groups[4].Value}},
                    @{L="OS Build";E={"20348"}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            }
            #>
            If($OSVersion -imatch '22H2|21H2|20H2'){
                # Download the HTML content
                #$url = "https://support.microsoft.com/en-us/help/5018894"
                $url = "https://support.microsoft.com/en-us/topic/release-notes-for-azure-stack-hci-version-23h2-018b9b10-a75b-4ad7-b9d1-7755f81e5b0b"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Load the HTML into an HtmlDocument object
                #$htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
                #$htmlDoc.LoadHtml($htmlpage)

                # Find all elements with the "supLeftNavLink" class
                <#$links = $htmlDoc.DocumentNode.SelectNodes("//*[@class='supLeftNavLink']")

                $KBList  = $Links | Where-Object{($_.InnerText -imatch '\(KB.+\)') -and ($_.InnerText -notmatch 'Preview')} | sort InnerStartIndex | Select -first $KBItemsToShow `
                    @{L='KBNumber';E={(((($_.InnerText -split "\(KB")[-1]) -split '\(')[0].Trim() -split '\)')[0]}},`
                    @{L='Date';E={$It=$_.InnerText -replace ',',';';($It -split '\s')[0,1,2] -join ',' -replace '\,'," " -replace ';',','  }},`
                    @{L="Description";E={($_.InnerText -replace "&#x2014;"," ")}},
                    @{L="OS Build";E={"20348"}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.OuterHtml -split 'href="')[-1] -split '">')[0]}},
                    @{L="DownloadLink";E={[GetKBDLLink]::GetDownloadLink(((($_.innerText -split "\(KB")[-1]) -split '\)')[0]),$OSVersion}}#>
                           # Find all elements with the "supLeftNavLink" class
                $divs=[regex]::Matches($htmlpage,'(?s)supLeftNavCategory((?:.*?)(<\/div>)){2}')
                Foreach ($match in $divs) {
                    If ($match.Groups[1].Value -match $OSVersion) {
                        $links=[regex]::Matches($match.Groups[1].Value,'supLeftNavLink.*?(href=\".*?\")[^>]*>((?:(?!preview).)*?)(KB\d{7})(.*?)<\/a>')
                    }
                }

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -split "\s")[0..2] -Join " "}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[3].Value+$_.Groups[4].Value}},
                    @{L="OS Build";E={"20349"}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}

            }

            If($OSVersion -imatch '2022'){
                $OSVersion = Switch ($SysInfo[0].OSBuildNumber) {
                    '19042' {'20H2'}
                    '20348' {'21H2'}
                    '20349' {'22H2'}
                }
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/5005454"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Load the HTML into an HtmlDocument object
                #$htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
                #$htmlDoc.LoadHtml($htmlpage)

                # Find all elements with the "supLeftNavLink" class
                #$links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")[^>]*>((?:(?!preview).)*?)(KB\d{7})(.*?)<\/a>')
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")>(.*?)(KB\d{7})\D+((?:(?!Preview).)20348.*?)\)(?:(?!Preview).)*<\/a>')

               <# $KBList  = $Links | Where-Object{($_.InnerText -imatch "KB") -and ($_.InnerText -imatch '20348') -and ($_.InnerText -notmatch 'Preview')} | sort InnerStartIndex | Select -first $KBItemsToShow `
                    @{L='KBNumber';E={(((($_.InnerText -split "&#x2014;")[-1]) -split '\(')[0].Trim() -split '\s')[0]}},`
                    @{L='Date';E={($_.InnerText -split "&#x2014;")[0]}},`
                    @{L="Description";E={($_.InnerText -replace "&#x2014;"," ")}},
                    @{L="OS Build";E={(((($_.InnerText -split "OS Build")[-1]) -replace '\)') -replace 'OS' -replace 's' -replace ' Preview' -replace ' Out-of-band' -replace ' Update for Windows 10 Mobile').trim()}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.OuterHtml -split 'href="')[-1] -split '">')[0]}},
                    @{L="DownloadLink";E={[GetKBDLLink]::GetDownloadLink((((($_.InnerText -split "&#x2014;KB")[-1]) -split '\(')[0].Trim() -split '\s')[0]),$OSVersion}}
                                    $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")[^>]*>((?:(?!preview).)*?)(KB\d{7})(.*?)<\/a>')#>

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -replace "&#x2014;"," ")}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[3].Value+$_.Groups[4].Value}},
                    @{L="OS Build";E={"2022"}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            }


#$KBList = $KBList | ? DownloadLink -like "https*"
$KBLatest = $KBList[0]

        $CurrentUpdatesandHotfixes=Foreach ($key in ($SDDCFiles.keys -like "*GetHotFix")) { $SDDCFiles."$key" |`
Sort-Object HotFixID,CSName | Select-Object @{Label="PSComputerName";Expression={$_.CSName}},Description,`
@{Label='HotFixID';Expression={($_.HotFixID -replace "$KBLatest.KBNumber","GGRREEEENN$KBLatest.KBNumber")}},InstalledOn
}
        $HostAnyFixFound = @{}
ForEach ($hostnameObj in ($CurrentUpdatesandHotfixes.PSComputerName | Sort-Object -Unique)){
$HostAnyFixFound[$hostnameObj] = $false
}

$CurrentOSBuild=@()
# display kb list against kb's found on hosts
ForEach($HotFix in $KBList){

$CurrentOSBuildTmp=@{}
$CurrentOSBuildTmp = $HotFix|select-object @{Label='KBNumber';Expression={"<a href='$($_.InfoLink)' target='_blank'>$($_.KBNumber)</a>"}},@{Label='Released/Description';Expression={$HotFix.Date}},@{Label='MSCatalogLink';Expression={"NA"}}
ForEach ($hostnameObj in ($CurrentUpdatesandHotfixes.PSComputerName | Sort-Object -Unique)){
if ($HostAnyFixFound[$hostnameObj]) {
Add-Member -force -InputObject $CurrentOSBuildTmp -MemberType NoteProperty -Name $hostnameObj -Value "N/A (cumulative)"
} else {
Add-Member -force -InputObject $CurrentOSBuildTmp -MemberType NoteProperty -Name $hostnameObj -Value "Not Installed"
}
}
ForEach ($hostObj in ($CurrentUpdatesandHotfixes | where-object {$_.HotFixID -like ("*" +$HotFix.KBNumber)})) {
if ($HotFix.KBNumber -match $KBLatest.KBNumber) {
Add-Member -force -InputObject $CurrentOSBuildTmp -MemberType NoteProperty -Name $hostObj.PSComputerName -Value ("GGRREEEENN" + $hostObj.InstalledOn)
} else {
Add-Member -force -InputObject $CurrentOSBuildTmp -MemberType NoteProperty -Name $hostObj.PSComputerName -Value ("YYEELLLLOOWW" + $hostObj.InstalledOn)
If ($KBList.KBNumber -notcontains $HotFix.KBNumber) {
    If ($HotFix.DownloadLink -eq "") {
        $HotFix.DownloadLink = [GetKBDLLink]::GetDownloadLink($HotFix.KBNumber,$OSVersion)
    }
    $CurrentOSBuildTmp.MSCatalogLink="<a href='$($HotFix.DownloadLink)'>DownloadLink</a>"
}
}
$HostAnyFixFound[$hostObj.PSComputerName] = $true
} 
try {$null=$HostAnyFixFound[$hostObj.PSComputerName]} catch {
     $HotFix.DownloadLink = [GetKBDLLink]::GetDownloadLink($HotFix.KBNumber,$OSVersion)
     $CurrentOSBuildTmp.MSCatalogLink="<a href='$($HotFix.DownloadLink)'>DownloadLink</a>"
}
$CurrentOSBuild+=$CurrentOSBuildTmp
}
        ForEach ($hostnameObj in ($CurrentUpdatesandHotfixes.PSComputerName | Sort-Object -Unique)) {
            if (-not $HostAnyFixFound[$hostnameObj]) {
                ForEach($HotFix in $KBList) {
                    If ($HotFix.DownloadLink -eq "") {
                        $HotFix.DownloadLink = [GetKBDLLink]::GetDownloadLink($HotFix.KBNumber,$OSVersionNodes)
                    }
                    $CurrentOSBuild | ? KBNumber -match $HotFix.KBNumber | %{$_.MSCatalogLink="<a href='$($HotFix.DownloadLink)'>DownloadLink</a>"}
                }
                $CurrentOSBuild[-1].$hostnameObj="RREEDD"+$CurrentOSBuild[-1].$hostnameObj
            }
        }
#$CurrentOSBuild
    
        $html+='<H2 id="RecommendedUpdatesandHotfixes">Recommended Updates and Hotfixes</H2>'
        $html+="<h5><b>Should be:</b></h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Lastest $($KBLatest.KBNumber)</h5><br>"
        $html+="<h5>&nbsp;&nbsp;Current KB list from Microsoft: (showing last $KBItemsToShow)</h5>"
        $html+=($CurrentOSBuild | ConvertTo-html -Fragment) -replace '&gt;','>' -replace '&lt;','<' -replace '&#39;',"'"
        $html+="<h5><br>&nbsp;&nbsp;Currently installed:</h5>"
        $html+=$CurrentUpdatesandHotfixes | sort-object -Property @{Expression={$_.PSComputerName}; Ascending=$true}, @{Expression={$_.HotFixID} ;Descending=$true} | ConvertTo-html -Fragment
    }    
        
        $html=$html -replace '<td>GGRREEEENN','<td style="background-color: #40ff00">'`
                    -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                    -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""
        $dstop=Get-Date
        #Write-Host "Total time taken is $(($dstop-$dstart).totalmilliseconds)"
#endregion Recommended updates and hotfixes for Windows Server

#Solution and SBE Updates
        $Name="Solution and SBE Updates"
        Write-Host "    Gathering $Name..."  
        If ($SysInfo[0].AzureLocalVersion -gt "") {
           $Name="Solution and SBE Updates"
           #Write-Host "    Gathering $Name..."
           $SolutionUpdates=$SDDCFiles."$($SDDCFiles.keys | ?{$_ -match 'GetSolutionUpdate' }| Select-Object -First 1)" |`
           Sort-Object ResourceId | Select-Object ResourceId,Version,@{L="HealthState";E={"RREEDD"*(@("Unknown","Success") -notcontains $_.HealthState)+$_.HealthState}},@{L="State";E={"RREEDD"*(@("NotApplicableBecauseAnotherUpdateIsInProgress","Installed","Ready","ReadyToInstall","Obsolete") -notcontains $_.State)+$_.State}},MinVersionRequired,MinSBEVersionRequied,ComponentVersions
        
            #Azure Table
               $AzureTableData=@()
               $AzureTableData=$ClusterHeartbeatConfigurationXml | Select *,
                   @{L='ReportID';E={$CReportID}}
               $PartitionKey=$Name -replace '\s'
               $TableName="CluChk$($PartitionKey)"
               $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=f69DyaUGlwcc%2BUujAbpnJ%2B4VPk3PigwCgjIIa0DQCQY%3D&tn=CluChkClusterHeartbeatConfiguration'
               $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

           #HTML Report
           $html+='<H2 id="SolutionandSBEUpdates">Solution and SBE Updates</H2>'
           $html+=$SolutionUpdates | ConvertTo-html -Fragment
           $html=$html `
           -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
           -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
           $html+="<h5>&nbsp;&nbsp;<a href='https://learn.microsoft.com/en-us/azure/azure-local/upgrade/about-upgrades-23h2' target='_blank'>Ref: https://learn.microsoft.com/en-us/azure/azure-local/upgrade/about-upgrades-23h2</a></h5>"
           $ResultsSummary+=Set-ResultsSummary -name $name -html $html
           $htmlout+=$html
           $html=""
           $Name=""
        }
#end region Solution and SBE Updates

#Latest Action Plan Failure
If ((Get-ChildItem $SDDCPath -Filter "GetActionplanInstanceToComplete.txt" -Recurse).count) {
        $Name="Latest Action Plan Failure"
        Write-Host "    Gathering $Name..."
        $ap=Get-Content (Get-ChildItem $SDDCPath -Filter "GetActionplanInstanceToComplete.txt" -Recurse | select -first 1).fullname
        $derr=$ap | select-string -SimpleMatch 'ERROR:' -Context 10,1

        $resultObject=@()
        foreach ($ErrorMessage in $derr) {

            $resultObject += [PSCustomObject] @{
                PSComputerName                  = "RREEDD" + (($ErrorMessage.Context.PreContext | select-string 'Value') -split ': On ')[-1].trim(':')
                TimeStamp                       = (Get-Date (($ErrorMessage.Context.PreContext | select-string 'TimeStamp') -split '  ').trim(':')[-1] -Format "MM/dd/yyyy HH:mm tt")
                Message                         = $ErrorMessage.Line.Trim()
                
            }
        }
        $derr=$ap | select-string -SimpleMatch 'Test-EnvironmentReadiness: the following results were not successful:' -Context 0,20

        foreach ($ErrorMessage in $derr) {

            $resultObject += [PSCustomObject] @{
                PSComputerName                  = "RREEDD" + (($ErrorMessage.Context.PostContext | select-string 'TargetResourceName') -split ':')[-1].trim()
                TimeStamp                       = (Get-Date (($ErrorMessage.Context.PostContext | select-string 'TimeStamp') -split '  ').trim(':')[-1] -Format "MM/dd/yyyy HH:mm tt")
                Message                         = ((($ErrorMessage.Context.PostContext | select-string 'TargetResourceName' -Context 0,5).Context.PostContext ) -replace "Description        : ","").Trim() -join ","
                
            }
        }
        $ActionPlanErrors=$resultObject | sort PSComputerName,TimeStamp,Message -Unique
        If ($ActionPlanErrors) {
           #HTML Report
           $html+='<H2 id="LatestActionPlanFailure">Latest Action Plan Failure</H2>'
           $html+=$ActionPlanErrors | ConvertTo-html -Fragment
           $html=$html `
           -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
           -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
           #$html+="<h5>&nbsp;&nbsp;<a href='https://learn.microsoft.com/en-us/azure/azure-local/upgrade/about-upgrades-23h2' target='_blank'>Ref: https://learn.microsoft.com/en-us/azure/azure-local/upgrade/about-upgrades-23h2</a></h5>"
           $ResultsSummary+=Set-ResultsSummary -name $name -html $html
           $htmlout+=$html
           $html=""
           $Name=""
        }
}

#end region Latest action Plan Failure

# Firewall Profile
        $Name="Firewall Profile"
        Write-Host "    Gathering $Name..."  
        $FirewallProfile=Foreach ($key in ($SDDCFiles.keys -like "*GetNetFirewallProfile")) {$SDDCFiles."$key" |`
        Sort-Object Profile | Select-Object PSComputerName,Profile,`
        @{Label='Enabled';Expression={Switch($_.Enabled){
          '0'{'False'}
          '1'{'True'}}}}
        }
        # Check all nodes configured the same
        IF((($FirewallProfile|Sort-Object Enabled -Unique)|Measure-Object).count -gt 1){
            $FirewallProfileMin=$FirewallProfile|Group-Object Enabled| Where-Object{$_.Count -eq ($FirewallProfile|Group-Object Enabled|Measure-Object -Property Count -Minimum).Minimum} |Select-Object Name
            $FirewallProfile=$FirewallProfile|Select-Object PSComputerName,Profile,@{Label='Enabled';Expression={$FPMC=$_.Enabled;IF($_.Enabled -imatch ($FirewallProfileMin.name)){"YYEELLLLOOWW$FPMC"}Else{$_.Enabled}}}
        }

        #$FirewallProfile | FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$FirewallProfile|Select-Object *,@{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=0GiGoMRnI7ao4MoOdOkG%2BfGK3MplVO57hiZ2XofkWx4%3D&tn=CluChkFirewallProfile'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="FirewallProfile">Firewall Profile</H2>'
        $html+=""
        $html+="<h5><b>Should be:</b></h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;All nodes should have the same value</h5>"
If($FirewallProfile.count -eq 0){$html+='<h5><span style="color: #ffffff; background-color: #ff0000">&nbsp;&nbsp;&nbsp;&nbsp;No Firewall Profile Entries found</span></h5>'}
        $html+=$FirewallProfile | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'                
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $fix=""
        $Fix+="<h5><b>Solution:</b></h5>"
        $Fix+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True</h5>"
        $htmlout+=$html+$fix
        $html=""
        $Name=""


# Network ATC
    IF($SDDCFiles.ContainsKey("GetNetIntent")){
        $Name="NetworkATC Status"
        Write-Host "    Gathering $Name..."
        $GetNetAdapterAll=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapter")) {$SDDCFiles."$key"}
        #$GetNetIntentStatusXml=""
        #$GetNetIntentStatusXml=$SDDCFiles."GetNetIntentStatus"
        #$GetNetIntentStatusGlobalOverridesXml=$SDDCFiles."GetNetIntentStatusGlobalOverrides"
        $GetNetIntentStatusXmlOut=@()
        $GetNetIntentStatusXmlOut=$SDDCFiles."GetNetIntentStatus" | Select-Object Host,IntentName,LastSuccess,@{L="Progress";E={
            # Use a regular expression to extract the numbers from the match string
                $numbers = [regex]::Matches($_.progress, "\d+")
            # Convert the extracted numbers to integers
                $number1 = [int]$numbers[0].Value
                $number2 = [int]$numbers[1].Value
            # Check if the numbers are equal
                if ($number1 -ne $number2) {"RREEDD"+$_.progress} else {$_.progress}}},
        @{L="ConfigurationStatus";E={IF($_.ConfigurationStatus -inotmatch "Success"){"RREEDD"+$_.ConfigurationStatus} else {$_.ConfigurationStatus}}},
        @{L="ProvisioningStatus";E={IF($_.ProvisioningStatus -inotmatch "Completed"){"RREEDD"+$_.ProvisioningStatus} else {$_.ProvisioningStatus}}}

        $GetNetIntentStatusXmlOut+=$SDDCFiles."GetNetIntentStatusGlobalOverrides" | Select-Object Host,@{L="IntentName";E={"Global"}},LastSuccess,Error,@{L="Progress";E={
            # Use a regular expression to extract the numbers from the match string
                $numbers = [regex]::Matches($_.progress, "\d+")
            # Convert the extracted numbers to integers
                $number1 = [int]$numbers[0].Value
                $number2 = [int]$numbers[1].Value
            # Check if the numbers are equal
                if ($number1 -ne $number2) {"RREEDD"+$_.progress} else {$_.progress}}},
        @{L="ConfigurationStatus";E={IF($_.ConfigurationStatus -inotmatch "Success"){"RREEDD"+$_.ConfigurationStatus} else {$_.ConfigurationStatus}}},
        @{L="ProvisioningStatus";E={IF($_.ProvisioningStatus -inotmatch "Completed"){"RREEDD"+$_.ProvisioningStatus} else {$_.ProvisioningStatus}}}
        
        $GetNetIntentStatusXmlOut = $GetNetIntentStatusXmlOut | Where-Object{$_.host -ne $null} | Sort-Object IntentName,Host
        #$GetNetIntentStatusXmlOut| ft

            # HTML Report
            $html+='<H2 id="NetworkATCStatus">NetworkATC Status</H2>'
            $html+=$GetNetIntentStatusXmlOut | ConvertTo-html -Fragment
            $html=$html `
             -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
             -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'                
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name=""
        
        If ($OSVersionNodes -ne "23H2") {
           $Name="NetworkATC Overrides"
           Write-Host "    Gathering $Name..."
           #$GetNetIntentXml=""
           #$GetNetIntentGlobalOverridesXml=""
           #$GetNetIntentXml=Get-ChildItem -Path $SDDCPath -Filter "GetNetIntent.xml" -Recurse -Depth 1 | import-clixml
           #$GetNetIntentGlobalOverridesXml=Get-ChildItem -Path $SDDCPath -Filter "GetNetIntentGlobalOverrides.xml" -Recurse -Depth 1 | import-clixml
           #GlobalOverrides
           $ClusterOverrides=@()
           $ClusterOverrides=$SDDCFiles."GetNetIntentGlobalOverrides" | Select-Object IntentType,`
                @{L="OverrideType";E={"ClusterSettings"}},
                @{L="EnableNetworkNaming";E={$_.ClusterOverride.EnableNetworkNaming}},
                @{L="EnableVirtualMachineMigrationPerformanceSelection";E={$CEVMMPS=$_.ClusterOverride.EnableVirtualMachineMigrationPerformanceSelection;IF($CEVMMPS -inotmatch 'False' -and $SysInfo[0].SysModel -notmatch "^APEX"){"RREEDD"+$CEVMMPS}Else{$CEVMMPS}}},
                @{L="VirtualMachineMigrationPerformanceOption";E={$CVMMPO=$_.ClusterOverride.VirtualMachineMigrationPerformanceOption;IF($CVMMPO -inotmatch "SMB" -and $SysInfo[0].SysModel -notmatch "^APEX"){"RREEDD"+$CVMMPO}else{$CVMMPO}}},
                @{L="MaximumVirtualMachineMigrations";E={$MVMM=$_.ClusterOverride.MaximumVirtualMachineMigrations;IF($MVMM -ne "2" -and $SysInfo[0].SysModel -notmatch "^APEX"){"RREEDD"+$MVMM}Else{$MVMM}}},
                @{L="MaximumSMBMigrationBandwidthInGbps";E={$_.ClusterOverride.MaximumSMBMigrationBandwidthInGbps}}
           $ProxyOverrides=@()
           $ProxyOverrides=$SDDCFiles."GetNetIntentGlobalOverrides" | Select-Object IntentType,`
                @{L="OverrideType";E={"WinHttpAdvProxy"}},
                @{L="ProxyServer";E={$_.ProxyOverride.ProxyServer}},
                @{L="ProxyBypass";E={$_.ProxyOverride.ProxyBypass}},
                @{L="AutoConfigUrl";E={$_.ProxyOverride.AutoConfigUrl}},
                @{L="AutoDetect";E={$_.ProxyOverride.AutoDetect}}
           #Managment and Compute Overrides
           $ManagmentandComputeOverrides=@()
           $ManagmentandComputeOverrides=$SDDCFiles."GetNetIntent" | Where-Object{$_.IntentType -eq 10} | Select-Object IntentName,NetAdapterNamesAsList,`
                @{L="NetworkDirect";E={$AAPOND=$_.AdapterAdvancedParametersOverride.NetworkDirect;IF($AAPOND -ne '0' -and $SysInfo[0].SysModel -notmatch "^APEX"){"RREEDD"+$AAPOND}Else{$AAPOND}}},
                @{L="JumboPacket";E={$AAPOND=$_.AdapterAdvancedParametersOverride.JumboPacket;$AAPOND}}
           #Storage Overrides
           $StorageOverrides=@()
           $StorageOverrides=$SDDCFiles."GetNetIntent" | Where-Object{$_.IntentType -imatch "Storage"} | Select-Object IntentName,NetAdapterNamesAsList,`
                @{L="JumboPacket";E={$sAAPOND=$_.AdapterAdvancedParametersOverride.JumboPacket;IF($sAAPOND -ne '9014' -and $SysInfo[0].SysModel -notmatch "^APEX"){"RREEDD"+$sAAPOND}Else{$sAAPOND}}},
                @{L="NetworkDirectTechnology";E={
                    #Make sure the support version of NetworkDirectTechnology is found
                        $sAAPOND=$_.AdapterAdvancedParametersOverride.NetworkDirectTechnology
                        $GetNetIntentXmlStorage=$SDDCFiles."GetNetIntent" | Where-Object{$_.IntentType -imatch "Storage"} 
                        switch -Regex ((($GetNetAdapterAll| ?{$_.name -imatch ($GetNetIntentXmlStorage.NetAdapterNamesAsList -split ',' )[0]}).InterfaceDescription)[0]){
                            "X710" {
                                #iWARP = 1
                                IF($sAAPOND -inotmatch '1'){"RREEDD"+$sAAPOND}Else{$sAAPOND}
                                }
                            "QLogic" {
                                #iWARP = 1
                                IF($sAAPOND -inotmatch '1'){"RREEDD"+$sAAPOND}Else{$sAAPOND}
                                }
                            "E810"{
                                #iWARP = 1 or Rocev2 = 4
                                IF(($sAAPOND -inotmatch '1') -and ($sAAPOND -inotmatch '4')){"RREEDD"+$sAAPOND}Else{$sAAPOND}
                                }
                            "Mellanox"{
                                #Rocev2 = 4
                                IF($sAAPOND -inotmatch '4'){"RREEDD"+$sAAPOND}Else{$sAAPOND}
                                }
                            default {$sAAPOND}
                        }
                }},
                @{L="BandwidthPercentage_Cluster";E={$sAAPOND=$_.QosPolicyOverride.BandwidthPercentage_Cluster;IF($sAAPOND -inotmatch '2' -and $SysInfo[0].SysModel -notmatch "^APEX"){"RREEDD"+$sAAPOND}Else{$sAAPOND}}},
                @{L="PriorityValue8021Action_Cluster";E={$sAAPOND=$_.QosPolicyOverride.PriorityValue8021Action_Cluster;IF($sAAPOND -inotmatch '5' -and $sAAPOND -inotmatch '7'){"RREEDD"+$sAAPOND}Else{$sAAPOND}}},
                @{L="EnableAutomaticIPGeneration";E={$SEAIG=$_.IPOverride.EnableAutomaticIPGeneration;IF($SEAIG -inotmatch 'False' -and $SysInfo[0].SysModel -notmatch "^APEX"){"RREEDD"+$SEAIG}Else{$SEAIG}}}

           # HTML Report
           $html+='<H2 id="NetworkATCOverrides">NetworkATC Overrides</H2>'
           $html+=" Please note that a blank entry indicates that the override has not been configured."
           $html+=$ClusterOverrides | ConvertTo-html -Fragment
           $html+='<br>'
           $html+=$ProxyOverrides | ConvertTo-html -Fragment
           $html+='<br>'
           $html+=$ManagmentandComputeOverrides | ConvertTo-html -Fragment
           $html+='<br>'
           $html+=$StorageOverrides | ConvertTo-html -Fragment
           $html=$html `
            -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
            -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'                
           $ResultsSummary+=Set-ResultsSummary -name $name -html $html
           $htmlout+=$html
           $html=""
           $Name=""

        }
    }


#Fully Converged VM Switch and Adapter configuration
    #VM Switch info
        $Name="VM Switch and Adapter Configuration"
        Write-Host "    Gathering $Name..."
        $NetIfDes=""
        #$Note=@()
        #$SDDCFiles.AZLCL02GetVMNetworkAdapter.switchname
        $GetVMNetworkAdapter=Get-ChildItem -Path $SDDCPath -Filter "GetVMNetworkAdapter.XML" -Recurse -Depth 1 | import-clixml

        $GetNetAdapterXml=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapter")) {$SDDCFiles."$key"} 
        $VMSwitchandAdapterconfiguration=Foreach ($key in ($SDDCFiles.keys -like "*GetVMSwitch")) {$SDDCFiles."$key" |`
        Sort-Object EmbeddedTeamingEnabled| Select-Object ComputerName,@{Label='Name';Expression={If ($GetVMNetworkAdapter.switchname -notcontains $_.Name -and $SysInfo[0].AzureLocalVersion -gt "") {"RREEDD"+$_.Name} else {$_.Name}}},SoftwareRscEnabled,EmbeddedTeamingEnabled,`
        @{Label='BandwidthReservationMode';Expression={
            IF(-not($SDDCFiles.ContainsKey("GetNetIntent"))){
                IF(($_.EmbeddedTeamingEnabled -match 'True') -and ($_.BandwidthReservationMode -inotmatch 'Weight') -and $_.IOVEnabled -ne $true){"RREEDD"+$_.BandwidthReservationMode}Else{$_.BandwidthReservationMode}
            }Else{$_.BandwidthReservationMode}}},`
        @{Label='BandwidthPercentage';Expression={
            IF(-not($SDDCFiles.ContainsKey("GetNetIntent"))){
                IF(($_.EmbeddedTeamingEnabled -match 'True') -and ($_.BandwidthPercentage -lt 100 -and $IOVEnabled -ne $true)){"RREEDD"+$_.BandwidthPercentage}Else{$_.BandwidthPercentage}
            }Else{$_.BandwidthPercentage}}},`
        @{Label='NetAdapterInterfaceDescriptions';Expression={
            $NetIfDes=$_.NetAdapterInterfaceDescriptions -replace [regex]::match($_.NetAdapterInterfaceDescriptions,"\\d+")
            IF($NetIfDes -match "Multiplexor"){
                "YYEELLLLOOWW$NetIfDes"
                }
                Else{"$NetIfDes"}
            }}
        $VMSwitchandAdapterconfiguration=$VMSwitchandAdapterconfiguration | Sort-Object EmbeddedTeamingEnabled| Select-Object ComputerName,Name,EmbeddedTeamingEnabled,`
        @{Label='BandwidthReservationMode';Expression={Switch($_.BandwidthReservationMode){
          '0'{'Default'}
          'RREEDD0'{'RREEDDDefault'}
          '1'{'Weight'}
          '2'{'Absolute'}
          'RREEDD2'{'RREEDDAbsolute'}
          '3'{'None'}
          'RREEDD3'{'RREEDDNone'}
          'Weight'{'Weight'}
          default{Expression={"RREEDD"+$_.BandwidthReservationMode}}
        }}},BandwidthPercentage,SoftwareRscEnabled,NetAdapterInterfaceDescriptions,`
        @{Label='Note';Expression={IF($_.NetAdapterInterfaceDescriptions -match "Multiplexor"){"Found LBFO Teaming (Multiplexor). We should NOT use LBFO Teaming for Virtual Switch. Convert to SET"}}}
        }
        # Check for 1 gig NICs in SET switch or no vm network adapter attached
            IF(($VMSwitchandAdapterconfiguration.NetAdapterInterfaceDescriptions -match "giga")`
                -or ($VMSwitchandAdapterconfiguration.NetAdapterInterfaceDescriptions -match "1GB")){
                $VMSwitchandAdapterconfiguration=$VMSwitchandAdapterconfiguration|`
                Where-Object{$_.EmbeddedTeamingEnabled -eq $True}|`
                Select-Object ComputerName,Name,EmbeddedTeamingEnabled,BandwidthReservationMode,BandwidthPercentage,`
                @{Label='NetAdapterInterfaceDescriptions';`
                Expression={
                    IF(($_.NetAdapterInterfaceDescriptions -icontains "giga")`
                     -or ($_.NetAdapterInterfaceDescriptions -icontains "1GB")){"YYEELLLLOOWW"+$_.NetAdapterInterfaceDescriptions}
                Else{$_.NetAdapterInterfaceDescriptions}}},`
                @{Label='Note';Expression={IF($_.NetAdapterInterfaceDescriptions -match "YYEELLLLOOWW"){"Found 1 gig NICs in Virtual Switch. We should NOT use 1 gig NICs in Management Virtual Switch."}
                IF($_.NetAdapterInterfaceDescriptions -match "RREEDD"){"This switch should have at least one VM or external adapter attached. Updates will fail health check"}}}
            }
        $VMSwitchandAdapterconfiguration=$VMSwitchandAdapterconfiguration|`
                Select-Object ComputerName,Name,EmbeddedTeamingEnabled,BandwidthReservationMode,BandwidthPercentage,NetAdapterInterfaceDescriptions,`
                @{Label='Note';Expression={IF($_.NetAdapterInterfaceDescriptions -match "YYEELLLLOOWW"){"Found 1 gig NICs in Virtual Switch. We should NOT use 1 gig NICs in Management Virtual Switch."}
                IF($_.Name -match "RREEDD"){"This switch should have at least one VM or external virtual adapter attached. Health check will fail."}}}

        #$VMSwitchandAdapterconfiguration|FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$VMSwitchandAdapterconfiguration|Select-Object -Property `
                @{L='ComputerName';E={[string]$_.ComputerName}},
                @{L='Name';E={[string]$_.Name}},
                @{L='EmbeddedTeamingEnabled';E={[string]$_.EmbeddedTeamingEnabled}},
                @{L='BandwidthReservationMode';E={[string]$_.BandwidthReservationMode}},
                @{L='BandwidthPercentage';E={[string]$_.BandwidthPercentage}},
                @{L='SoftwareRscEnabled';E={[string]$_.SoftwareRscEnabled}},
                @{L='NetAdapterInterfaceDescriptions';E={[string]$_.NetAdapterInterfaceDescriptions}},
                @{L='Note';E={[string]$_.Note}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=s2anhT6xKk5UpSFpTmtQDoy97bDp4j3JTXUSxm7m%2BTY%3D&tn=CluChkVMSwitchandAdapterConfiguration'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="VMSwitchandAdapterConfiguration">VM Switch and Adapter Configuration</H2>'
        $VMSwitchandAdapterconfigurationkey=""
        $VMSwitchandAdapterconfigurationkey+="<h5><b>Should be:</b></h5>"
        $VMSwitchandAdapterconfigurationkey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-EmbeddedTeamingEnabled=True</h5>"
        IF(-not(Get-ChildItem -Path $SDDCPath -Filter "GetNetIntent.xml")){
            $VMSwitchandAdapterconfigurationkey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-BandwidthReservationMode=Weight</h5>"
            $VMSwitchandAdapterconfigurationkey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-BandwidthPercentage=100</h5>"
        }Else{
            $VMSwitchandAdapterconfigurationkey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-BandwidthReservationMode=None</h5>"
            $VMSwitchandAdapterconfigurationkey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-BandwidthPercentage=0</h5>"
        }
        $VMSwitchandAdapterconfigurationkey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-NetAdapterInterfaceDescriptions=two NICs</h5>"
        $VMSwitchandAdapterconfigurationkey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://www.dell.com/support/kbdoc/en-us/000200052/dell-azure-stack-hci-non-converged-network-configuration' target='_blank'>Ref: https://www.dell.com/support/kbdoc/en-us/000200052/dell-azure-stack-hci-non-converged-network-configuration</a></h5>"
        $html+=$VMSwitchandAdapterconfigurationkey
        $html+=$VMSwitchandAdapterconfiguration | sort-object PSComputerName | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'                
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $fix=""
        $Fix+="<h5><b>Solution:</b></h5>"
        $Fix+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;You must remove and recreate the VMSwitch to change MinimumBandwidthMode</h5>"
        $Fix+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;New-VMSwitch -Name VMSwitchName -AllowManagementOS 0 -NetAdapterName MgmtNic1,MgmtNic2 -MinimumBandwidthMode Weight -Verbose</h5>"
        $htmlout+=$html
        $html=""
        $Name=""
                          
    #VM Network Adapters
        $Name="VM Network Adapters vLANs"
        Write-Host "    Gathering $Name..."  
        $VMNetworkAdaptersvLANs=@()
        $VMNetworkAdaptersvLANs=Get-ChildItem -Path $SDDCPath -Filter "GetVMNetworkAdapter.XML" -Recurse -Depth 1 | import-clixml |`
        Where-Object{$_.Name -ine 'Network Adapter'}|Sort-Object Name,ComputerName |Select-Object ComputerName,@{Label='AdapterName';Expression={"vEthernet ("+$_.Name+")"}},`
        @{Label='OperationMode';Expression={$_.VlanSetting.OperationMode}},`
        @{Label='VlanId';Expression={$_.VlanSetting.AccessVlanId}},`
        Band*
<#        $VMNetworkAdapterNames=
        IF($VMNetworkAdaptersvLANs| group AdapterName,VlanId | group count)
        ForEach($VMNA in  $VMNetworkAdapterNames){
            
            ForEach ($VMNetworkAdaptersvLAN in $VMNetworkAdaptersvLANs){
                IF($VMNA.name -eq $VMNetworkAdaptersvLAN.AdapterName){
                
                }
            }
        }
#>
        $VMNetworkAdaptersvLANs=$VMNetworkAdaptersvLANs|Select-Object ComputerName,AdapterName,OperationMode,VlanId,Band*
        #$VMNetworkAdaptersvLANs | FT
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$VMNetworkAdaptersvLANs|Select-Object -Property `
                @{L='ComputerName';E={[string]$_.ComputerName}},
                @{L='AdapterName';E={[string]$_.AdapterName}},
                @{L='OperationMode';E={[string]$_.OperationMode}},
                @{L='VlanId';E={[string]$_.VlanId}},
                @{L='BandwidthSetting';E={[string]$_.BandwidthSetting}},
                @{L='BandwidthPercentage';E={[string]$_.BandwidthPercentage}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=fD4KPdbo1A35dJZSoXvLBk2Gs2YXllKNc9U0Ftv7Tfw%3D&tn=CluChkVMNetworkAdaptersvLANs'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="VMNetworkAdaptersvLANs">VM Network Adapters vLANs</H2>'
        $VMNetworkAdaptersvLANsKey=""
        $VMNetworkAdaptersvLANsKey+="<h5><b>Should be:</b></h5>"
        $VMNetworkAdaptersvLANsKey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Storage NICs should have seperate vLANs</h5>"
        $VMNetworkAdaptersvLANsKey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-BandwidthSetting = NULL or VMNetworkAdapterBandwidthSetting</h5>"
        $VMNetworkAdaptersvLANsKey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-BandwidthPercentage = 0 unless BandwidthSetting = VMNetworkAdapterBandwidthSetting</h5>"
        $html+=$VMNetworkAdaptersvLANsKey
        $html+=$VMNetworkAdaptersvLANs | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'                
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""

    #Configure Host Management VLAN
        $Name="Configure Host Management VLAN"
        Write-Host "    Gathering $Name..."
        $ConfigureHostManagementVLANOut=@()
        #Used for NetworkATC implementations 
        IF($SDDCFiles.ContainsKey("GetNetIntent")){
            $GetVMNetworkAdapterIsolation=Foreach ($key in ($SDDCFiles.keys -like "*GetVMNetworkAdapterIsolation")) {$SDDCFiles."$key" | Where-Object{$_.IsolationMode -imatch "Vlan"}|`
            Sort-Object ComputerName,AdapterName | Select-Object @{L="PSComputerName";E={$_.ComputerName}},@{L="AdapterName";E={$_.ParentAdapter -replace 'VMInternalNetworkAdapter, Name = ' -replace "\'"}},@{L="vLANID";E={$_.DefaultIsolationID}}
            }
            
            $ConfigureHostManagementVLAN=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterAdvancedProperty")) {$SDDCFiles."$key" | Where-Object{$_.DisplayName -eq "VLAN ID"}|`
            Sort-Object PSComputerName,Name | Select-Object PSComputerName, @{L="AdapterName";E={$_.Name}},@{L="vLANID";E={$_.DisplayValue}}
            }
            $ConfigureHostManagementVLANOut= $ConfigureHostManagementVLAN + $GetVMNetworkAdapterIsolation
        }Else{
            $ConfigureHostManagementVLAN=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterAdvancedProperty")) {$SDDCFiles."$key" | Where-Object{$_.DisplayName -eq "VLAN ID"}|`
        Sort-Object PSComputerName,Name | Select-Object PSComputerName,@{L="AdapterName";E={$_.Name}},DisplayName,DisplayValue
        }
        $ConfigureHostManagementVLANOut = $ConfigureHostManagementVLAN
        }
        
        #$ConfigureHostManagementVLAN|FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$ConfigureHostManagementVLAN|Select-Object *,@{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=I6FOwle68jKwVb9jLk6ON1hAHJlhyp89lutE68Nz%2B%2Fk%3D&tn=CluChkConfigureHostManagementVLAN'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="ConfigureHostManagementVLAN">Configure Host Management VLAN</H2>'
        $ConfigureHostManagementVLANkey=""
        $ConfigureHostManagementVLANkey+="<h5><b>Should be:</b></h5>"
        #$ConfigureHostManagementVLANkey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-DisplayName=VLAN ID</h5>"
        $ConfigureHostManagementVLANkey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-vLANID=VLAN the customer uses</h5>"
        $ConfigureHostManagementVLANkey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Same Number of NICs per Node</h5>"
        $html+=$ConfigureHostManagementVLANkey
        $html+=$ConfigureHostManagementVLANOut | Sort-Object PSComputerName, AdapterName  | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'                
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""  

    #Assign Host Management IP address
        $Name="Assign Host Management IP address"
        Write-Host "    Gathering $Name..."  
        $AssignHostManagementIPaddress=Foreach ($key in ($SDDCFiles.keys -like "*GetNetIpAddress")) {$SDDCFiles."$key" |Where-Object{($_.InterfaceAlias -inotmatch 'isatap') -and ($_.InterfaceAlias -inotmatch 'Pseudo')}|`
        Sort-Object PSComputerName,ifIndex | Select-Object PSComputerName,InterfaceAlias,ifIndex,IPAddress,`
        @{Label='AddressState';Expression={Switch($_.AddressState){`
          '0'{'Invalid'}`
          '1'{'Tentative'}`
          '2'{'Duplicate'}`
          '3'{'Deprecated'}`
          '4'{'Preferred'}`
        }}}
        }
        $AssignHostManagementIPaddress=$AssignHostManagementIPaddress|Select-Object PSComputerName,InterfaceAlias,ifIndex,IPAddress,AddressState
        #@{Label='AddressState';Expression={If($_.AddressState -inotmatch 'Preferred'){"YYEELLLLOOWW"+$_.AddressState}Else{$_.AddressState}}}
        #$AssignHostManagementIPaddress|FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$AssignHostManagementIPaddress|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='InterfaceAlias';E={[string]$_.InterfaceAlias}},
                @{L='ifIndex';E={[string]$_.ifIndex}},
                @{L='IPAddress';E={[string]$_.IPAddress}},
                @{L='AddressState';E={[string]$_.AddressState}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=lyCFsoyd6I220KJly4hIMBpFGJkXSvjwXLi%2BLmQuDac%3D&tn=CluChkAssignHostManagementIPaddress'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="AssignHostManagementIPaddress">Assign Host Management IP address</H2>'
        $AssignHostManagementIPaddressSB=""
        $AssignHostManagementIPaddressSB+="<h5><b>Should be:</b></h5>"
        $AssignHostManagementIPaddressSB+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-AddressState=Preferred</h5>"
        $html+=$AssignHostManagementIPaddressSB
        $html+=$AssignHostManagementIPaddress | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        $AssignHostManagementIPaddresskey=""
        $AssignHostManagementIPaddresskey+="<h5><b>AddressState Key:</b></h5>"
        $AssignHostManagementIPaddresskey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-- Invalid. IP address configuration information for addresses that are not valid and will not be used.</h5>"
        $AssignHostManagementIPaddresskey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-- Tentative. IP address configuration information for addresses that are not used for communication, as the uniqueness of those IP addresses is being verified.</h5>"
        $AssignHostManagementIPaddresskey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-- Duplicate. IP address configuration information for addresses for which a duplicate IP address has been detected and the current IP address will not be used.</h5>"
        $AssignHostManagementIPaddresskey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-- Deprecated. IP address configuration information for addresses that will no longer be used to establish new connections, but will continue to be used with existing connections.</h5>"
        $AssignHostManagementIPaddresskey+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-- Preferred. IP address configuration information for addresses that are valid and available for use.</h5>"
        $AssignHostManagementIPaddresskey+="&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://docs.microsoft.com/en-us/powershell/module/nettcpip/get-netipaddress?view=win10-ps' target='_blank'>Ref: Microsoft Docs - Get-NetIPAddress -AddressState</a>"                
        $html+=$AssignHostManagementIPaddresskey
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""        

    # iDRAC ISM IP address
        $Name="iDRAC ISM IP address"
        Write-Host "    Gathering $Name..."  
		$idracNics = Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapter")) {
			$SDDCFiles."$key" | Where-Object {($_.InterfaceDescription -imatch 'Remote NDIS Compatible Device')}|`
			Sort-Object PSComputerName, ifIndex | Select-Object PSComputerName,InterfaceDescription,ifIndex,@{L='AdminStatus';E={IF($_.AdminStatus -ne "Up") {'RREEDD'+$_.AdminStatus} else {$_.AdminStatus}}},@{L='MediaConnectionState';E={IF($_.MediaConnectionState -ne "Connected") {'RREEDD'+$_.MediaConnectionState} else {$_.MediaConnectionState}}},Speed
		}

		$GetNetIpAddress=Foreach ($key in ($SDDCFiles.keys -like "*GetNetIpAddress")) {$SDDCFiles."$key" | Where-Object{$_.IPv4Address}}

		$iDRACNetAdaptersIPs=@()
		# Find the Mgmt Cluster Network
			ForEach($NetIp in $GetNetIpAddress){
				$iDRACNetAdaptersIPs+=$idracNics | Where-Object{($_.IfIndex -eq $NetIp.IfIndex) -and ($_.PsComputername -eq $NetIp.PSComputerName )}|Select-Object *,@{L='IPv4Address';E={
					if (($SysInfo[0].SysModel -match "^APEX") -and ($NetIp.IPv4Address -ne "169.254.0.1")) {'RREEDD'+$NetIp.IPv4Address}
					elseif (($SysInfo[0].SysModel -notmatch "^APEX") -and (($GetNetIpAddress | where-object {$_.IPv4Address -eq $NetIp.IPv4Address} | sort -Unique).count -gt 1)) {'RREEDD'+$NetIp.IPv4Address}
					else {$NetIp.IPv4Address}
					}}
		}
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$AssignHostManagementIPaddress|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='InterfaceDescription';E={[string]$_.InterfaceDescription}},
                @{L='ifIndex';E={[string]$_.ifIndex}},
                @{L='AdminStatus';E={[string]$_.AdminStatus}},
                @{L='MediaConnectionState';E={[string]$_.MediaConnectionState}},
                @{L='Speed';E={[string]$_.Speed}},
                @{L='IPv4Address';E={[string]$_.IPv4Address}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=lyCFsoyd6I220KJly4hIMBpFGJkXSvjwXLi%2BLmQuDac%3D&tn=CluChkiDRACISMIPaddress'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}
		
		
        # HTML Report
        $html+='<H2 id="iDRACISMIPaddress">iDRAC ISM IP address</H2>'
        $iDRACIPaddress=""
        $iDRACIPaddress+="<h5><b>Should be:</b></h5>"
        $iDRACIPaddress+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Adminstatus=Up</h5>"
		$iDRACIPaddress+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-MediaConnectedState=Connected</h5>"
		IF($SysInfo[0].SysModel -notmatch "^APEX"){
			$iDRACIPaddress+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-IPv4Address=unique per host</h5>"
		} else {
			$iDRACIPaddress+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-IPv4Address=169.254.0.1</h5>"
		}
        $html+=$iDRACIPaddress
        $html+=$iDRACNetAdaptersIPs | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'

        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""       

    # Check for VMQ
        $Name="VMQ information"
        Write-Host "    Gathering $Name..." 
        $FoundGetNetAdaptervmq=@()
        $EnabledNICS=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapter")) {$SDDCFiles."$key" |Where-Object{$_.state -eq "2"}}
        #$EnabledNICS | FL
        $GetNetAdaptervmq=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterVmq")) {$SDDCFiles."$key"}
        ForEach($VMQ in $GetNetAdaptervmq){
            $FoundGetNetAdaptervmq+=$VMQ|Where-Object{$_.Name -cne $EnabledNICS.Name}|`
                Select-Object PSComputerName,Name,InterfaceDescription,`
                    @{L='Enabled';E={$VMQEnabled=$_.Enabled
                        IF($StorageNicsUnique -imatch $_.Name){
                           IF($VMQEnabled -eq 1){"RREEDD$VMQEnabled"}Else{$VMQEnabled}
                        }Else{$VMQEnabled}
                    }},`
                BaseProcessorNumber,MaxProcessors,NumberOfReceiveQueues
        }

        #$FoundGetNetAdaptervmq | Sort-Object PSComputerName,Name|FT -AutoSize
        $FoundGetNetAdaptervmq=$FoundGetNetAdaptervmq | Sort-Object PSComputerName,Name
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$FoundGetNetAdaptervmq|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='Name';E={[string]$_.Name}},
                @{L='InterfaceDescription';E={[string]$_.InterfaceDescription}},
                @{L='Enabled';E={[string]$_.Enabled}},
                @{L='BaseProcessorNumber';E={[string]$_.BaseProcessorNumber}},
                @{L='MaxProcessors';E={[string]$_.MaxProcessors}},
                @{L='NumberOfReceiveQueues';E={[string]$_.NumberOfReceiveQueues}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=UZ82JjKZbXWjDphBlCuufhJLt9GVu6NSsQmyINclcuw%3D&tn=CluChkVMQinformation'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="VMQinformation">VMQ Information</H2>'
        $GetNetAdaptervmqSB=""
        $GetNetAdaptervmqSB+="<h5><b>Should be:</b></h5>"
        $GetNetAdaptervmqSB+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Enabled=True (For NICs being used by the Virtual Switch the VMs connect)</h5>"
        $GetNetAdaptervmqSB+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Enabled=False (for storage NICs)</h5>"
        $html+=$GetNetAdaptervmqSB
        $html+=$FoundGetNetAdaptervmq | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""
    
    #Intel NIC should NOT have DCM enabled
        IF($GetNetAdapterXml.InterfaceDescription -like "Intel*10G*X710*"){
            $Name="Intel DCB Check"
            Write-Host "    Gathering $Name..." 
            $GetNetAdapterQos=@()
            $GetNetAdapterQosXml=@()
            $GetNetAdapterQosXml=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterQos")) {$SDDCFiles."$key" | Where-Object{$_.InterfaceDescription -like "Intel*10G*X710*"}}
            $GetNetAdapterQos=$GetNetAdapterQosXml | Select-Object PSComputerName,Name,InterfaceDescription,@{L='Enabled';E={IF($_.Enabled -eq $True){"RREEDD"+$_.Enabled}Else{$_.Enabled}}}
            #$GetNetAdapterQos | Sort-Object PSComputerName,InterfaceDescription | ft
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$GetNetAdapterQos|Select-Object *,@{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=6sXtfGfRsQnBGgJfEssdnAAZusq%2FaD9cw00xbgq0ZNg%3D&tn=CluChkIntelDCBCheck'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
            $html+='<H2 id="IntelDCBCheck">Intel DCB Check</H2>'
            $GetNetAdapterQosSB=""
            $GetNetAdapterQosSB+="<h5><b>Should be:</b></h5>"
            $GetNetAdapterQosSB+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Enabled=False</h5>"
            IF($GetNetAdapterQos -imatch "RREEDD"){
                $GetNetAdapterQosFix=""
                $GetNetAdapterQosFix+='<h5>Fix:</h5>'
                $GetNetAdapterQosFix+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;1. Suspend-ClusterNode</h5>'
                $GetNetAdapterQosFix+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;2. Get-StorageFaultDomain -type StorageScaleUnit | Where-Object {$_.FriendlyName -eq "$($Env:ComputerName)"} | Enable-StorageMaintenanceMode -ErrorAction Stop</h5>'
                $GetNetAdapterQosFix+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;3. Start-Process "C:\Program Files\Intel\Umb\Winx64\PROSETDX\DxSetup.exe" -ArgumentList "DMIX=0 /qn" -Wait</h5>'
                $GetNetAdapterQosFix+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;4. Resume-ClusterNode</h5>'
                $GetNetAdapterQosFix+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;5. Get-StorageFaultDomain -type StorageScaleUnit | Where-Object {$_.FriendlyName -eq "$($Env:ComputerName)"} | Disable-StorageMaintenanceMode -ErrorAction Stop</h5>'
                $GetNetAdapterQosFix+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;6. Get-StorageJobs #Wait for stortage jobs to complete.</h5>'
                $GetNetAdapterQosFix+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;7. Repeat steps on each node as needed</h5>'
                $GetNetAdapterQosFix+=""
            }

            $html+=$GetNetAdapterQosSB
            $html+=$GetNetAdapterQos | ConvertTo-html -Fragment
            $html+=$GetNetAdapterQosFix
            $html=$html `
             -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
             -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name=""
        }

    #Change RDMA mode on QLogic NICs -iWARP only
        $FoundGetNetAdapterAdvancedProperty=@()
        $e810=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapter")) {$SDDCFiles."$key" | Where-Object{($_.InterfaceDescription -like "*QLogic*") -or ($_.InterfaceDescription -like "*E810*")}}
        IF($e810.count) {
            $Name="RDMA mode"
            Write-Host "    Gathering $Name..."  
            $GetNetAdapterAdvancedProperty=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterAdvancedProperty")) {$SDDCFiles."$key" | `
                Where-Object{($_.InterfaceDescription -Match 'QLogic') -or ($_.InterfaceDescription -Match 'E810') }|`
                Where-Object{(($_.DisplayName -eq "RDMA Mode") `
                -or($_.DisplayName -eq "NetworkDirect Technology"))}|`
                Sort-Object PSComputerName,Name,InterfaceDescription | Select-Object PSComputerName,Name,InterfaceDescription,DisplayName,DisplayValue
            }
            ForEach($NAAP in $GetNetAdapterAdvancedProperty){
                $FoundGetNetAdapterAdvancedProperty+=$NAAP|Where-Object{$_.Name -cne $EnabledNICS.Name}
            }
            #$FoundGetNetAdapterAdvancedProperty | FT             
            $GetNetAdapterAdvancedProperty=$FoundGetNetAdapterAdvancedProperty|Select-Object PSComputerName,Name,InterfaceDescription,DisplayName,`
            @{Label='DisplayValue';Expression={
                IF($StorageNicsUnique -imatch $_.Name){
                    If($_.DisplayValue -inotmatch 'iWARP' -and $_.InterfaceDescription -notlike "*E810*" ){"RREEDD"+$_.DisplayValue}Else{$_.DisplayValue}
                }Else{$_.DisplayValue}
            }}
            #Azure Table
                $AzureTableData=@()
                $AzureTableData=$GetNetAdapterAdvancedProperty|Select-Object -Property `
                    @{L='PSComputerName';E={[string]$_.PSComputerName}},
                    @{L='Name';E={[string]$_.Name}},
                    @{L='InterfaceDescription';E={[string]$_.InterfaceDescription}},
                    @{L='DisplayName';E={[string]$_.DisplayName}},
                    @{L='DisplayValue';E={[string]$_.DisplayValue}},
                    @{L='ReportID';E={$CReportID}}
                $PartitionKey=$Name -replace '\s'
                $TableName="CluChk$($PartitionKey)"
                $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=W7OTdErChu%2BQ7IoF5UIB%2Bf7QAjnHvgMiOYV25dr61%2B0%3D&tn=CluChkRDMAmodeonQLogicNICs'
                $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

            # HTML Report
            $html+='<H2 id="RDMAmodeonQLogicNICs">RDMA mode</H2>'
            $html+=""
            $html+="<h5><b>Should be:</b></h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-DisplayName=RDMA Mode or NetworkDirect Technology</h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-DisplayValue=iWARP</h5>"
            $html+=$GetNetAdapterAdvancedPropertySB
            $html+=$GetNetAdapterAdvancedProperty | ConvertTo-html -Fragment
            $html=$html `
             -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
             -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name="" 
        }
        $Mellanox=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapter")) {$SDDCFiles."$key" | Where-Object{($_.InterfaceDescription -like "*Mellanox*")}}
        IF($Mellanox.count){
            $Name="RDMA mode on Mellanox NICs"
            Write-Host "    Gathering $Name..."  
            $GetNetAdapterAdvancedProperty=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterAdvancedProperty")) {$SDDCFiles."$key" | `
                Where-Object{$_.InterfaceDescription -Match 'Mellanox'}|`
                Where-Object{($_.DisplayName -eq "NetworkDirect Technology")}|`
                Sort-Object PSComputerName,Name,InterfaceDescription | Select-Object PSComputerName,Name,InterfaceDescription,DisplayName,DisplayValue
            }
            $FoundGetNetAdapterAdvancedProperty=@()
            ForEach($NAAP in $GetNetAdapterAdvancedProperty){
                $FoundGetNetAdapterAdvancedProperty+=$NAAP|Where-Object{$_.Name -cne $EnabledNICS.Name}
            }
            #$FoundGetNetAdapterAdvancedProperty | FT             
            $GetNetAdapterAdvancedProperty=$FoundGetNetAdapterAdvancedProperty|Select-Object PSComputerName,Name,InterfaceDescription,DisplayName,`
            @{Label='DisplayValue';Expression={
                IF($StorageNicsUnique -imatch $_.Name){
                    If($_.DisplayValue -inotmatch 'RoCEv2' -and $_.DisplayValue -ne $null){"RREEDD"+$_.DisplayValue}Else{$_.DisplayValue}
                }Else{$_.DisplayValue}
            }}
            #Azure Table
                $AzureTableData=@()
                $AzureTableData=$GetNetAdapterAdvancedProperty|Select-Object -Property `
                    @{L='PSComputerName';E={[string]$_.PSComputerName}},
                    @{L='Name';E={[string]$_.Name}},
                    @{L='InterfaceDescription';E={[string]$_.InterfaceDescription}},
                    @{L='DisplayName';E={[string]$_.DisplayName}},
                    @{L='DisplayValue';E={[string]$_.DisplayValue}},
                    @{L='ReportID';E={$CReportID}}
                $PartitionKey=$Name -replace '\s'
                $TableName="CluChk$($PartitionKey)"
                $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=W7OTdErChu%2BQ7IoF5UIB%2Bf7QAjnHvgMiOYV25dr61%2B0%3D&tn=CluChkRDMAmodeonMellanoxNICs'
                $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

            # HTML Report
            $html+='<H2 id="RDMAmodeonMellanoxNICs">RDMA mode on Mellanox NICs</H2>'
            $html+=""
            $html+="<h5><b>Should be:</b></h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-DisplayName=NetworkDirect Technology</h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-DisplayValue=RoCEv2</h5>"
            $html+=$GetNetAdapterAdvancedPropertySB
            $html+=$GetNetAdapterAdvancedProperty | ConvertTo-html -Fragment
            $html=$html `
             -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
             -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name="" 
        }

    #RDMA configuration
        $Name="RDMA configuration"
        Write-Host "    Gathering $Name..."  
        $GetNetAdapterRdma=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterRdma")) {$SDDCFiles."$key" |`
            Sort-Object PSComputerName,Name | Select-Object PSComputerName,Name,Description,Enabled
        }
        $GetNetAdapterRdma=$GetNetAdapterRdma|Select-Object PSComputerName,Name,Description,`
        @{Label='Enabled';Expression={
            IF($StorageNicsUnique -imatch $_.Name){
                If($_.Enabled -inotmatch 'True'){"RREEDD"+$_.Enabled}Else{$_.Enabled}
                }Else{$_.Enabled}
            }}
        #$GetNetAdapterRdma | FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$GetNetAdapterRdma|Select-Object *,@{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=P3gRpp9R4wktEyZdNHwSJM4%2FleADwTYc4zaKw18PSXM%3D&tn=CluChkRDMAconfiguration'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="RDMAconfiguration">RDMA configuration</H2>'
        $GetNetAdapterRdmaSB=""
        $GetNetAdapterRdmaSB+="<h5><b>Should be:</b></h5>"
        $GetNetAdapterRdmaSB+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Enabled=True</h5>"
        $GetNetAdapterRdmaSB+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;NOTE: Enabled for physical NICs used for Fully-converged or storage NICs only</h5>"
        $html+=$GetNetAdapterRdmaSB
        $html+=$GetNetAdapterRdma | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""

    # VM Network Adapter Team Mapping
        $Name="VM Network Adapter Team Mapping"
        Write-Host "    Gathering $Name..." 
        
        # IF Fully Converged then we need to check for VM Network Adapter Team Mapping
            IF($FullyConverged -eq $True){
                $GetVMNetworkAdapterTeamMapping=Foreach ($key in ($SDDCFiles.keys -like "*GetVMNetworkAdapterTeamMapping")) {$SDDCFiles."$key" |`
                Sort-Object ComputerName,NetAdapterName | Select-Object ComputerName,NetAdapterName,ParentAdapter
                }
                $GetVMNetworkAdapterTeamMapping=$GetVMNetworkAdapterTeamMapping|Select-Object NetAdapterName,`
                @{Label='ParentAdapter';Expression={If($_.ParentAdapter -inotmatch 'VMInternalNetworkAdapter'){"RREEDD"+$_.ParentAdapter}Else{$_.ParentAdapter}}}
                #$GetVMNetworkAdapterTeamMapping | FT -AutoSize
                #Azure Table
                    $AzureTableData=@()
                    $AzureTableData=$GetNetAdapterAdvancedProperty|Select-Object -Property `
                        @{L='ComputerName';E={[string]$_.ComputerName}},
                        @{L='NetAdapterName';E={[string]$_.NetAdapterName}},
                        @{L='ParentAdapter';E={[string]$_.ParentAdapter}},
                        @{L='ReportID';E={$CReportID}}
                    $PartitionKey=$Name -replace '\s'
                    $TableName="CluChk$($PartitionKey)"
                    $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=UTwCcXsxAhk8VKwC96bXWx1ohE3mGepFM%2FGq23k5bS0%3D&tn=CluChkVMNetworkAdapterTeamMapping'
                    $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

                # HTML Report
                $html+='<H2 id="VMNetworkAdapterTeamMapping">VM Network Adapter Team Mapping</H2>'
                $html+=""
                $html+="<h5><b>Should be:</b></h5>"
                $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Each Storage vNIC should be mapped to pNIC</h5>"
                $html+=$GetVMNetworkAdapterTeamMappingSB
                $html+=$GetVMNetworkAdapterTeamMapping | ConvertTo-html -Fragment
                $html=$html `
                 -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                 -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
                $ResultsSummary+=Set-ResultsSummary -name $name -html $html
                $htmlout+=$html
            }
        $html=""
        $Name=""
       If ($SysInfo[0].SysModel -notmatch "^APEX") {
    #Set-VMHost " VirtualMachineMigrationPerformanceOption SMB
        $Name="Virtual Machine Migration Performance Option"
        Write-Host "    Gathering $Name..." 
        $GetVMHost=Foreach ($key in ($SDDCFiles.keys -like "*GetVMHost")) {$SDDCFiles."$key" |`
        Sort-Object ComputerName | Select-Object ComputerName,VirtualMachineMigrationEnabled,VirtualMachineMigrationPerformanceOption
        }
        $GetVMHost=$GetVMHost|Select-Object ComputerName,`
        @{Label='VirtualMachineMigrationEnabled';Expression={If($_.VirtualMachineMigrationEnabled -inotmatch 'True'){"RREEDD"+$_.VirtualMachineMigrationEnabled}Else{$_.VirtualMachineMigrationEnabled}}},`
        @{Label='VirtualMachineMigrationPerformanceOption';Expression={If($_.VirtualMachineMigrationPerformanceOption -inotmatch 'SMB'){"RREEDD"+$_.VirtualMachineMigrationPerformanceOption}Else{$_.VirtualMachineMigrationPerformanceOption}}}
        #$GetVMHost | FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$GetVMHost|Select-Object -Property `
                @{L='ComputerName';E={[string]$_.ComputerName}},
                @{L='VirtualMachineMigrationEnabled';E={[string]$_.VirtualMachineMigrationEnabled}},
                @{L='VirtualMachineMigrationPerformanceOption';E={[string]$_.VirtualMachineMigrationPerformanceOption}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=X88Vowz2E2nOF1mtTSBzfypc%2BW31CzQ1Z6uIAFlN294%3D&tn=CluChkVirtualMachineMigrationPerformanceOption'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="VirtualMachineMigrationPerformanceOption">Virtual Machine Migration Performance Option</H2>'
        $GetVMHostSB=""
        $GetVMHostSB+="<h5><b>Should be:</b></h5>"
        $GetVMHostSB+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-VirtualMachineMigrationEnabled=True</h5>"
        $GetVMHostSB+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-VirtualMachineMigrationPerformanceOption=SMB</h5>"
        $html+=$GetVMHostSB
        #####$GetVMHost
        $html+=$GetVMHost | sort-object PSComputerName | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""
       }
    #***ROCE ONLY***
 IF(($AllNVMe -eq $True) -or ($Mellanox.count) -or $SysInfo[0].SysModel -match "^APEX"){
    If($AllNVMe -eq $True){
        $Name="DCB and QOS Configuration"
        $html+='<H2 id="DCBandQOSConfiguration">DCB and QOS Configuration</H2>'
        $html+=""
        $html+='<h5 style="background-color: #ffff00"><b>Storage Pool contains all NVMe disks or Mellanox NICs. DCB and QoS configurations recommeneded. Link below.</b></h5>'
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://infohub.delltechnologies.com/l/reference-guide-network-integration-and-host-network-configuration-options-1/qos-policy-configuration-4' target='_blank'>QoS policy configuration</a></h5>"
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""
    }
    <#
    If(Get-ChildItem -Path $SDDCPath -Filter "GetNetAdapter.xml" -Recurse -Depth 1| import-clixml | Where-Object{$_.InterfaceDescription -match "Mellanox"}){
        $Name="DCB and QOS Configuration"
        $html+='<H2 id="DCBandQOSConfiguration">DCB and QOS Configuration</H2>'
        $html+=""
        $html+='<h2 style="background-color: #ffff00"><b>Mellanox Storage Nics found. DCB and QoS configurations REQUIRED. Links below.</b></h2>'
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://infohub.delltechnologies.com/l/reference-guide-network-integration-and-host-network-configuration-options-1/qos-policy-configuration-4' target='_blank'>QoS policy configuration</a></h5>"
        $htmlout+=$html
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""
    }#>
    
    #QoS Policy configuration
        $GetNetQosPolicyOut=@()
        $Name="QoS Policy configuration"
        Write-Host "    Gathering $Name..." 
        $GetNetQosPolicy=Foreach ($key in ($SDDCFiles.keys -like "*GetNetQosPolicy")) {$SDDCFiles."$key"}
        $GetNetQosPolicySMB=$GetNetQosPolicy| Where-Object{$_.Name -imatch "SMB"} |`
            Select-Object PSComputerName,Name,`
                @{Label='NetDirectPortMatchCondition';Expression={If($_.NetDirectPortMatchCondition -inotmatch '445'){"RREEDD"+$_.NetDirectPortMatchCondition}Else{$_.NetDirectPortMatchCondition}}},`
                @{Label='PriorityValue8021Action';Expression={If($_.PriorityValue8021Action -inotmatch '3'){"RREEDD"+$_.PriorityValue8021Action}Else{$_.PriorityValue8021Action}}}
        $GetNetQosPolicyCluster=$GetNetQosPolicy| Where-Object{$_.Name -eq "Cluster"} |`
            Select-Object PSComputerName,Name,`
                @{Label='NetDirectPortMatchCondition';Expression={$_.NetDirectPortMatchCondition}},`
                @{Label='PriorityValue8021Action';Expression={If($_.PriorityValue8021Action -inotmatch '5' -and $_.PriorityValue8021Action -inotmatch '7'){"RREEDD"+$_.PriorityValue8021Action}Else{$_.PriorityValue8021Action}}}
        
        $GetNetQosPolicyOut=$GetNetQosPolicySMB+$GetNetQosPolicyCluster
        #$GetNetQosPolicyOut | Format-Table -AutoSize
        #$html+="<h2>QoS Policy configuration</h2>"
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$GetNetQosPolicyOut|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='Name';E={[string]$_.Name}},
                @{L='NetDirectPortMatchCondition';E={[string]$_.NetDirectPortMatchCondition}},
                @{L='PriorityValue8021Action';E={[string]$_.PriorityValue8021Action}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=l%2FxctB2uM6sDhH%2BFEG0HpRhvTj8NMFIQ%2Fbg8lBjDF40%3D&tn=CluChkQoSPolicyconfiguration'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="QoSPolicyconfiguration">QoS Policy configuration</H2>'
        $html+=""
        $html+="<h5><b>Should be:</b></h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Name=Cluster NetDirectPortMatchCondition=0 PriorityValue8021Action=5 or 7</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Name=SMB NetDirectPortMatchCondition=445 PriorityValue8021Action=3</h5>"
        $html+=$GetNetQosPolicyOut |Sort-Object PSComputerName,Name | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""

    #NetQosTrafficClass
        $Name="Net Qos Traffic Class"
        Write-Host "    Gathering $Name..." 
        $GetNetQosTrafficClass=Foreach ($key in ($SDDCFiles.keys -like "*GetNetQosTrafficClass")) {$SDDCFiles."$key" | Select *,@{L='ComputerName';E={$key -replace "GetNetQosTrafficClass","" }}}
        $GetNetQosTrafficClassSMB=$GetNetQosTrafficClass| Where-Object{$_.Name -eq "SMB"}| Select-Object ComputerName,Name,`
            @{Label='Priority';Expression={If($_.Priority -inotmatch '3'){"RREEDD"+$_.Priority}Else{$_.Priority}}},`
            @{Label='BandwidthPercentage';Expression={If($_.BandwidthPercentage -ne '50'){"RREEDD"+$_.BandwidthPercentage}Else{$_.BandwidthPercentage}}},`
            @{Label='Algorithm';Expression={If($_.Algorithm -inotmatch '2' -and $_.Algorithm -inotmatch 'ETS'){"RREEDD"+$_.Algorithm}Else{$_.Algorithm}}}
        $GetNetQosTrafficClassCluster=$GetNetQosTrafficClass| Where-Object{$_.Name -eq "Cluster"}| Select-Object ComputerName,Name,`
            @{Label='Priority';Expression={If($_.Priority -inotmatch '5' -and $_.Priority -inotmatch '7'){"RREEDD"+$_.Priority}Else{$_.Priority}}},`
@{Label='BandwidthPercentage';Expression={
                $StorageNicsLinkSpeed = ([regex]::Matches($StorageNics[0].linkspeed,'\d+')).value
                If($StorageNicsLinkSpeed -eq "10"){
                    IF($_.BandwidthPercentage -lt '2'){"RREEDD"+$_.BandwidthPercentage}}
                ElseIf($StorageNicsLinkSpeed -ge "25"){
                    IF($_.BandwidthPercentage -lt '1'){"RREEDD"+$_.BandwidthPercentage}Else{$_.BandwidthPercentage}}}},`
            @{Label='Algorithm';Expression={If($_.Algorithm -inotmatch '2' -and $_.Algorithm -inotmatch 'ETS'){"RREEDD"+$_.Algorithm}Else{$_.Algorithm}}}
        $GetNetQosTrafficClassRest=$GetNetQosTrafficClass| Where-Object{$_.Name -ne "SMB" -and $_.Name -ne "Cluster"}| Select-Object ComputerName,Name,@{Label='Priority';Expression={$_.Priority}},BandwidthPercentage,Algorithm

        $GetNetQosTrafficClassOut=$GetNetQosTrafficClassSMB+$GetNetQosTrafficClassCluster+$GetNetQosTrafficClassRest
        #$GetNetQosTrafficClassOut | FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$GetNetQosTrafficClassOut|Select-Object -Property `
                @{L='ComputerName';E={[string]$_.ComputerName}},
                @{L='Name';E={[string]$_.NamePriority}},
                @{L='Priority';E={[string]$_.NetDirectPortMatchCondition}},
                @{L='BandwidthPercentage';E={[string]$_.BandwidthPercentage}},
                @{L='Algorithm';E={[string]$_.Algorithm}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=25Sw3P7K8SVETmDtm%2BnpUQOUU7nYdK%2FXMaPN1KD6EmY%3D&tn=CluChkNetQosTrafficClass'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="NetQosTrafficClass">Net Qos Traffic Class</H2>'
        $html+=""
        $html+="<h5><b>Should be:</b></h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;SMB:</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-Priority=3 BandwidthPercentage=50 Algorithm=2</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;Cluster:</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-Priority = 5 or 7 BandwidthPercentage: 10GbE = 2 25GbE or higher = 1 Algorithm=2</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-BandwidthPercentage:</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;10GbE = 2</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;25GbE or higher = 1</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Algorithm=2</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://github.com/MicrosoftDocs/azure-stack-docs/blob/main/azure-stack/hci/concepts/host-network-requirements.md#cluster-traffic-class' target='_blank'>Ref: Host network requirements for Azure Stack HCI</a></h5>"
        $html+=$GetNetQosTrafficClassOut |Sort-Object ComputerName,Name | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""

    #NetQosFlowControl   
        $Name="Net Qos Flow Control"
        Write-Host "    Gathering $Name..." 
        $getNetQosFlowControl=Foreach ($key in ($SDDCFiles.keys -like "*GetNetQosFlowControl")) {$SDDCFiles."$key"}
        $getNetQosFlowControlOut=$getNetQosFlowControl|`
            Sort-Object PSComputerName,Priority|Select-Object PSComputerName,Priority,`
                @{Label='Enabled';Expression={
                    If(($_.Priority -eq "3") -and ($_.Enabled -ne 'True')){"RREEDD"+$_.Enabled}`
                    ElseIF(($_.Priority -eq "0") -and ($_.Enabled -eq 'True')){"RREEDD"+$_.Enabled}`
                    ElseIF(($_.Priority -eq "1") -and ($_.Enabled -eq 'True')){"RREEDD"+$_.Enabled}`
                    ElseIF(($_.Priority -eq "2") -and ($_.Enabled -eq 'True')){"RREEDD"+$_.Enabled}`
                    ElseIF(($_.Priority -eq "4") -and ($_.Enabled -eq 'True')){"RREEDD"+$_.Enabled}`
                    ElseIF(($_.Priority -eq "5") -and ($_.Enabled -eq 'True')){"RREEDD"+$_.Enabled}`
                    ElseIF(($_.Priority -eq "6") -and ($_.Enabled -eq 'True')){"RREEDD"+$_.Enabled}`
                    ElseIF(($_.Priority -eq "7") -and ($_.Enabled -eq 'True')){"RREEDD"+$_.Enabled}`
                    Else{$_.Enabled}}}
        #$getNetQosFlowControlOut | FT -AutoSize

        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$getNetQosFlowControlOut|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='Priority';E={[string]$_.NetDirectPortMatchCondition}},
                @{L='Enabled';E={[string]$_.Enabled}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=Qk4ld4Bp8zv0TARtCnbf2Rl1V5GgwbUkwDUtkl4%2B7j0%3D&tn=CluChkNetQosFlowControl'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="NetQosFlowControl">Net Qos Flow Control</H2>'
        $html+=""
        $html+="<h5><b>Should be:</b></h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Priority=3 Enabled=True</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Priority=0,1,2,4,5,6,7 Enabled=False</h5>"
        $html+=$getNetQosFlowControlOut |Sort-Object PSComputerName,Priority | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""

    #NetAdapterQos
        $Name="Net Adapter Qos"
        $FoundQualityOfService=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterAdvancedProperty")) {$SDDCFiles."$key" | Where-Object{$_.DisplayName -eq "Quality Of Service"}}
        IF($FoundQualityOfService.count -gt 0){
        Write-Host "    Gathering $Name..." 
        $GetNetAdapterAdvancedProperty=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterAdvancedProperty")) {$SDDCFiles."$key" | Where-Object{$_.DisplayName -eq "Quality Of Service"}|`
            Sort-Object PSComputerName,Name | Select-Object PSComputerName,Name,DisplayName,DisplayValue
        }
        $GetNetAdapterAdvancedProperty=$GetNetAdapterAdvancedProperty|Select-Object PSComputerName,Name,DisplayName,`
        @{Label='DisplayValue';Expression={If(($_.name -imatch $StorageNics.name | Sort-Object -Unique ) -and ($_.DisplayValue -inotmatch 'Enabled')){"RREEDD"+$_.DisplayValue}Else{$_.DisplayValue}}}
        #$GetNetAdapterAdvancedProperty | FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$GetNetAdapterAdvancedProperty|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='Name';E={[string]$_.Name}},
                @{L='DisplayName';E={[string]$_.DisplayName}},
                @{L='DisplayValue';E={[string]$_.DisplayValue}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=xdc%2BNsof05Le2uFBgusk2VGETzbffsCTk9D3X1PKo%2Fg%3D&tn=CluChkNetAdapterQos'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="NetAdapterQos">Net Adapter Qos</H2>'
        $html+=""
        $html+="<h5><b>Should be:</b></h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-DisplayName=Quality Of Service</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-DisplayValue=Enabled</h5>"
        $html+=$GetNetAdapterAdvancedProperty | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""
        }

    #NetQosDcbxSetting
        $Name="Net Qos Dcbx Setting"
        Write-Host "    Gathering $Name..." 
        $GetNetQosDcbxSetting=""
        #Used for NetworkATC implementations 
        IF($SDDCFiles.ContainsKey("GetNetIntent")){
            $GetNetQosDcbxSettingOut=@()
            # Initialize an array to store the parsed results
                $results = @()
            (Get-ChildItem -Path $SDDCPath -Filter "GetNetQosDcbxSettingPerNic.txt" -Recurse -Depth 1).FullName | %{
                # Read the content of the text file
                $filename=$_
                $fileContent = Get-Content -Path $_

                # Split the output into lines
                $lines = $fileContent -split "`n"

                # Iterate over the lines starting from the 5th line (index 4)
                for ($i = 4; $i -lt $lines.Count; $i++) {
                    $line = $lines[$i].Trim()
                    if ($line -match '^\s*(\w+)\s+(\w+)\s+(\d+)\s+(.+)\s+(\w+)$') {
                        $willing = $matches[1]
                        $policySet = $matches[2]
                        $ifIndex = $matches[3]
                        $ifAlias = $matches[4]
                        $psComputerName = $matches[5]
        
                        $result = [PSCustomObject]@{
                            'Willing' = $willing
                            'PolicySet' = $policySet
                            'IfIndex' = $ifIndex
                            'IfAlias' = $ifAlias
                            'PSComputerName' = (Split-Path $filename).Split("\")[-1].replace("Node_","")
                        }
        
                        $results += $result
                    }
                }
            }

            # Print the parsed results
            #$GetNetQosDcbxSettingOut | ft 
            ForEach($NQDS in $results){
                ForEach($StorageNic in $StorageNics){
                    IF($NQDS.PSComputerName -ieq $StorageNic.PSComputerName -and $NQDS.IfIndex -eq $StorageNic.IfIndex ){
                        #Check for willing 
                        $GetNetQosDcbxSettingOut+=$NQDS|Select-Object PSComputerName,IfAlias,`
                            @{Label='Willing';Expression={If($_.Willing -inotmatch 'False'){"RREEDD"+$_.Willing}Else{$_.Willing}}}
                    }
                }
            }
            
            $GetNetQosDcbxSetting = $GetNetQosDcbxSettingOut
        }ElseIF(-not($SDDCFiles.ContainsKey("GetNetIntent"))){
            # Used for non-network ATC
            $GetNetQosDcbxSetting=Foreach ($key in ($SDDCFiles.keys -like "*GetNetQosDcbxSetting")) {$SDDCFiles."$key" |`
                Sort-Object PSComputerName | Select-Object PSComputerName,Willing
            }
            $GetNetQosDcbxSetting=$GetNetQosDcbxSetting|Select-Object PSComputerName,`
            @{Label='Willing';Expression={If($_.Willing -inotmatch 'False'){"RREEDD"+$_.Willing}Else{$_.Willing}}}
        }
        
        #$GetNetQosDcbxSetting | FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$GetNetQosDcbxSetting|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='Willing';E={[string]$_.Willing}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=PSLjnO7SZl8M9VutCu2rvDXlKBAzG0ShdVG5oz3Ts3s%3D&tn=CluChkNetQosDcbxSetting'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="NetQosDcbxSetting">Net Qos Dcbx Setting</H2>'
        $html+=""
        $html+="<h5><b>Should be:</b></h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Willing=False</h5>"
        $html+=$GetNetQosDcbxSetting | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""

    #NetAdapterAdvancedProperty 
        $FoundDcbxMode=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterAdvancedProperty")) {$SDDCFiles."$key" | Where-Object{$_.DisplayName -eq "DcbxMode"}}
        If($FoundDcbxMode.count -gt 0){
            $Name="Net Qos Dcbx Property"
            Write-Host "    Gathering $Name..." 
            $GetNetAdapterAdvancedProperty=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterAdvancedProperty")) {$SDDCFiles."$key" | Where-Object{($_.DisplayName -match 'DcbxMode')}|` # ?{($_.DisplayName -match 'RDMA Mode') -or ($_.DisplayName -match 'NetworkDirect Technology')}|`
                Sort-Object PSComputerName,Name | Select-Object PSComputerName,Name,DisplayName,DisplayValue
            }
            $GetNetAdapterAdvancedProperty=$GetNetAdapterAdvancedProperty|Select-Object PSComputerName,Name,DisplayName,`
            @{Label='DisplayValue';Expression={If((($StorageNics.name | Sort-Object -Unique) -imatch $_.name ) -and ($SysInfo[0].SysModel -notmatch "^APEX" -and $_.DisplayValue -inotmatch 'Host In Charge')){"RREEDD"+$_.DisplayValue}Else{$_.DisplayValue}}}
            #$GetNetAdapterAdvancedProperty | FT -AutoSize
            #Azure Table
                $AzureTableData=@()
                $AzureTableData=$GetNetAdapterAdvancedProperty|Select-Object -Property `
                    @{L='PSComputerName';E={[string]$_.PSComputerName}},
                    @{L='Name';E={[string]$_.Name}},
                    @{L='DisplayName';E={[string]$_.DisplayName}},
                    @{L='DisplayValue';E={[string]$_.DisplayValue}},
                    @{L='ReportID';E={$CReportID}}
                $PartitionKey=$Name -replace '\s'
                $TableName="CluChk$($PartitionKey)"
                $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=SkZ99DtzpMzHK6s5kMVpQ1NTG0jQia%2BDEliu8WQCVtQ%3D&tn=CluChkNetQosDcbxProperty'
                $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

            # HTML Report
            $html+='<H2 id="NetQosDcbxProperty">Net Qos Dcbx Property</H2>'
            $html+=""
            $html+="<h5><b>Should be:</b></h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-DisplayName=DcbxMode</h5>"
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-DisplayValue=Host In Charge</h5>"
            $html+=$GetNetAdapterAdvancedProperty | ConvertTo-html -Fragment
            $html=$html `
             -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
             -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name=""
        }
  }
    $SystemInfoContent=@((Get-Content -Path ((Get-ChildItem -Path $SDDCPath -Filter "SystemInfo.TXT" -Recurse -Depth 1).FullName[0]))[0..5])

    #Disable SMB Signing
        $Name="Disable SMB Signing"
        Write-Host "    Gathering $Name..." 
        $GetSmbClientConfiguration=Foreach ($key in ($SDDCFiles.keys -like "*GetSmbClientConfiguration")) {$SDDCFiles."$key"}
        $DisableSMBSigningGetSmbClientConfiguration = $GetSmbClientConfiguration | Sort-Object PSComputerName | Select-Object PSComputerName,`
        @{L='ClientEnableSecuritySignature';E={$CEnableSecuritySignature=$_.EnableSecuritySignature;IF($CEnableSecuritySignature -eq 0){"RREEDD$CEnableSecuritySignature"}Else{$CEnableSecuritySignature}}},`
        @{L='ClientRequireSecuritySignature';E={$RequireSecuritySignature=$_.RequireSecuritySignature;IF(($SysInfo[0].SysModel -notmatch "^APEX" -and $OSVersionNodes -ne "23H2" -and $RequireSecuritySignature -eq 1) -or (($SysInfo[0].SysModel -match "^APEX" -or $OSVersionNodes -eq "23H2") -and $RequireSecuritySignature -eq 0)){"RREEDD$RequireSecuritySignature"}Else{$RequireSecuritySignature}}},`
        @{L='ServerEnableSecuritySignature';E={}},`
        @{L='ServerEncryptData';E={}},`
        @{L='ServerRequireSecuritySignature';E={}}

        $GetSmbServerConfiguration=Foreach ($key in ($SDDCFiles.keys -like "*GetSmbServerConfiguration")) {$SDDCFiles."$key"}
        $DisableSMBSigningGetSmbServerConfiguration = $GetSmbServerConfiguration | Sort-Object PSComputerName | Select-Object PSComputerName,`
        @{L='ClientEnableSecuritySignature';E={}},`
        @{L='ClientRequireSecuritySignature';E={}},`
        @{L='ServerEnableSecuritySignature';E={$SEnableSecuritySignature=$_.EnableSecuritySignature;IF($SEnableSecuritySignature -eq $True){"RREEDD$SEnableSecuritySignature"}Else{$SEnableSecuritySignature}}},`
        @{L='ServerEncryptData';E={$EncryptData=$_.EncryptData;IF($EncryptData -eq $True){"RREEDD$EncryptData"}Else{$EncryptData}}},`
        @{L='ServerRequireSecuritySignature';E={$RequireSecuritySignature=$_.RequireSecuritySignature;IF(($SysInfo[0].SysModel -notmatch "^APEX" -and $OSVersionNodes -ne "23H2" -and $RequireSecuritySignature -eq 1) -or (($SysInfo[0].SysModel -match "^APEX" -or $OSVersionNodes -eq "23H2") -and $RequireSecuritySignature -eq 0)){"RREEDD$RequireSecuritySignature"}Else{$RequireSecuritySignature}}}
        $DisableSMBSigning=$DisableSMBSigningGetSmbClientConfiguration+$DisableSMBSigningGetSmbServerConfiguration
        #$DisableSMBSigning | FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$DisableSMBSigning|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='ClientEnableSecuritySignature';E={[string]$_.ClientEnableSecuritySignature}},
                @{L='ServerEnableSecuritySignature';E={[string]$_.ServerEnableSecuritySignature}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=cRyZoHr9trp1Ux816ouwBGd%2BgjruqO9ZK2VHwawhRy4%3D&tn=CluChkDisableSMBSigning'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="DisableSMBSigning">Disable SMB Signing</H2>'
        $html+=""
        $html+="<h5><b>Should be:</b></h5>"
        IF(($SysInfo[0].SysModel -notmatch "^APEX") -and ($OSVersionNodes -ne "23H2")){
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-ClientEnableSecuritySignature=True(1)</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-ClientRequireSecuritySignature=False(0)</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-ServerEnableSecuritySignature=False(0)</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-ServerEncryptData=False(0)</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-ServerRequireSecuritySignature=False(0)</h5>"
        }
        IF(($SysInfo[0].SysModel -match "^APEX") -or ($OSVersionNodes -eq "23H2")){
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-ClientEnableSecuritySignature=True(1)</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-ClientRequireSecuritySignature=True(1)</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-ServerEnableSecuritySignature=False(0)</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-ServerEncryptData=False(0)</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-ServerRequireSecuritySignature=True(1)</h5>"
        }
        IF($SystemInfoContent[2] -inotmatch "HCI"){
            $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://docs.microsoft.com/en-US/troubleshoot/windows-server/networking/reduced-performance-after-smb-encryption-signing' target='_blank'>Ref: Reduced networking performance after you enable SMB Encryption or SMB Signing</a></h5>"
        }Else{$html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://learn.microsoft.com/en-us/windows-server/storage/file-server/smb-direct#smb-encryption-with-smb-direct' target='_blank'>Ref: SMB Encryption with SMB Direct</a></h5>"}
        $html+=$DisableSMBSigning | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOoWW','<td style="background-color: #ffff00">'
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""

    #Update hardware timeout for Spaces port
        #Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Services\spaceport\Parameters -Name HwTimeout -Value 0x00002710 " Verbose
        $Name="Update hardware timeout for Spaces port"
        Write-Host "    Gathering $Name..." 
        #$GetItemProperty=Get-ChildItem -Path $SDDCPath -Filter "GetItemProperty.xml" -Recurse -Depth 1 | import-clixml  -ErrorAction Continue
        $GetItemProperty=Foreach ($ThisHost in (Get-ChildItem -Path $SDDCPath -Filter "GetRegSpacePortParameters.xml" -Recurse -Depth 1)) {import-clixml -Path $ThisHost.fullname -ErrorAction Continue | Select-Object @{L="ComputerName";E={$ThisHost.Directory.Name -replace "Node_",""}},*}
        $GetItemPropertyOut =$GetItemProperty|Where-Object{$_.HwTimeout -ne $Null} | Select-Object ComputerName,@{L='HwTimeout';E={$HwTimeout=$_.HwTimeout;IF($HwTimeout -lt 10000){"RREEDD"}ElseIF($HwTimeout -gt 10000){"YYEELLLLOOWW$HwTimeout"}Else{$HwTimeout}}} | Sort-Object ComputerName
        #$GetItemPropertyOut | FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$GetItemPropertyOut|Select-Object -Property `
                @{L='ComputerName';E={[string]$_.ComputerName}},
                @{L='HwTimeout';E={[string]$_.HwTimeout}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=3HzUWGRpxC1olR52fqdn%2FwPmvmfS2%2FGx8J5RTEcnecY%3D&tn=CluChkUpdatehardwaretimeoutforSpacesport'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="UpdatehardwaretimeoutforSpacesport">Update hardware timeout for Spaces port</H2>'
        $html+=""
        $html+="<h5><b>Should be:</b></h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-HwTimeout=0x00002710(10000)</h5>"
        $html+=$GetItemPropertyOut | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOoWW','<td style="background-color: #ffff00">'
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name=""

    #OEM Information Support Provider
        # Only check for this is Azure Stack HCI OS 20H2 or 21H2, check on first node OS
If($SystemInfoContent[2] -imatch 'HCI'){
        #If($Sys.OSName -imatch 'HCI'){
            $Name="OEM Information Support Provider"
            Write-Host "    Gathering $Name..." 
            $GetRegOEMInformation=Foreach ($key in ($SDDCFiles.keys -like "*GetRegOEMInformation")) {$SDDCFiles."$key" | Select *,@{L="ComputerName";E={$key.Replace("GetRegOEMInformation","")}}}
            #Skip if no 
                $GetRegOEMInformationOutMissing=@()
                $GetRegOEMInformationOutAll=@()
                $GetRegOEMInformationOut=@()
                $GetRegOEMInformationOut+=$GetRegOEMInformation|Where-Object{$_.SupportProvider -ne $Null} | Sort-Object ComputerName | Select-Object ComputerName,@{L='SupportProvider';E={$SupportProvider=$_.SupportProvider;IF($SupportProvider -inotmatch 'dell'){"RREEDD$SupportProvider"}Else{$SupportProvider}}}
                #Checking/Adding missing nodes
#check when none of the nodes have an entry, just list them all then
if($GetRegOEMInformationOut.PSComputerName.count -eq 0 -and $SysInfo[0].SysModel -notmatch "^APEX") {
ForEach($Node in $ClusterNodes){
                        $GetRegOEMInformationOut+=[PSCustomObject]@{
                            PSComputerName = $Node.name
                            SupportProvider = "RREEDDMissing"
                        }
                    }
} else {
IF($GetRegOEMInformationOut.PSComputerName.count -le $ClusterNodeCount){
$MissingNodes=(Compare-Object $GetRegOEMInformationOut.PSComputerName $ClusterNodes.name).InputObject
ForEach($Node in $MissingNodes){
$GetRegOEMInformationOut+=[PSCustomObject]@{
PSComputerName = $Node
SupportProvider = "RREEDDMissing"
}
}
}
}
            #$GetRegOEMInformationOut | FT -AutoSize
            #Azure Table
                $AzureTableData=@()
                $AzureTableData=$GetRegOEMInformationOut|Select-Object -Property `
                    @{L='PSComputerName';E={[string]$_.PSComputerName}},
                    @{L='SupportProvider';E={[string]$_.SupportProvider}},
                    @{L='ReportID';E={$CReportID}}
                $PartitionKey=$Name -replace '\s'
                $TableName="CluChk$($PartitionKey)"
                $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=dXnVLJdVlRklepyKrNgKHL412vvZQzR5vGzm9oBhNrg%3D&tn=CluChkOEMInformationSupportProvider'
                $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

            # HTML Report
                $html+='<H2 id="OEMInformationSupportProvider">OEM Information Support Provider</H2>'
                $html+=""
                $html+="<h5><b>Should be:</b></h5>"
                $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation\SupportProvider=Dell</h5>"
                $html+=$GetRegOEMInformationOut | ConvertTo-html -Fragment
                $html=$html `
                 -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                 -replace '<td>YYEELLLLOoWW','<td style="background-color: #ffff00">'
                $ResultsSummary+=Set-ResultsSummary -name $name -html $html
                $htmlout+=$html
                $html=""
                $Name=""
        }


    #Enabling jumbo frames
        $Name="Jumbo Frames"
        Write-Host "    Gathering $Name..."  
        $GetNetAdapterAdvancedProperty=@()
        $GetNetAdapterAdvancedProperty=Foreach ($key in ($SDDCFiles.keys -like "*GetNetAdapterAdvancedProperty")) {$SDDCFiles."$key" |`
          Where-Object{($_.DisplayName -eq "Jumbo Packet") -or ($_.DisplayName -eq "Jumbo MTU")}|`
          Where-Object{(($_.Name -like "*Port*")`
          -or($_.Name -like "vEthernet*")`
          -or($_.ifDesc -imatch "Intel")`
          -or($_.ifDesc -imatch "QLogic")`
          -or($_.ifDesc -imatch "Mellanox")`
          -and($_.ifDesc -inotmatch 'Gigabit'))}|`
        Sort-Object PSComputerName,Name | Select-Object PSComputerName,Name,DisplayName,DisplayValue 
        }
        $GetNetAdapterAdvancedProperty=$GetNetAdapterAdvancedProperty|`
            Select-Object PSComputerName,Name,DisplayName,`
                @{Label='DisplayValue';Expression={
                    If($StorageNicsUnique -imatch $_.name){
                        If($_.DisplayValue -inotmatch '9000' -and $_.DisplayValue -inotmatch '9014' -and $SysInfo[0].SysModel -notmatch "^APEX"){"RREEDD"+$_.DisplayValue}
				        Else{$_.DisplayValue}
                    }
                    # None Storage NICs should be 1514
                    ElseIF($_.DisplayValue -inotmatch 'Disabled' -and $_.DisplayValue -inotmatch '1514' -and $SysInfo[0].SysModel -notmatch "^APEX"){"YYEELLLLOOWW"+$_.DisplayValue}
                    Else{$_.DisplayValue}
                }
            }

        #$GetNetAdapterAdvancedProperty|FT -AutoSize
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$GetNetAdapterAdvancedProperty|Select-Object -Property `
                @{L='PSComputerName';E={[string]$_.PSComputerName}},
                @{L='Name';E={[string]$_.Name}},
                @{L='DisplayName';E={[string]$_.DisplayName}},
                @{L='DisplayValue';E={[string]$_.DisplayValue}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=AZOwLfHrqbTgoPiRHtlRjlqb6fPphTTiCRMgJRszzYc%3D&tn=CluChkJumboFrames'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
        $html+='<H2 id="JumboFrames">Jumbo Frames</H2>'
        $html+=""
        $html+="<h5><b>Should be:</b></h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Storage NICs DisplayValue=9014 Bytes</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;-Switchless Storage NICs DisplayValue=9014 Bytes</h5>"
        $html+=$GetNetAdapterAdvancedProperty | ConvertTo-html -Fragment
        $html=$html `
         -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
         -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
        $htmlout+=$html
        $html=""
        $Name="" 
#}
# Microsoft-Windows-StorageSpaces-Driver/Diagnostic Events
    #xml events ([xml]$myevents.Objects.Object[4].'#text').event.System.eventid.'#text'
    $Name="StorageSpaces Driver Diagnostic Events"
    Write-Host "    Gathering $Name..."
    $SSDDE=Get-ChildItem -Path $SDDCPath -Filter "Microsoft-Windows-StorageSpaces-Driver-Diagnostic.EVTX" -Recurse -Depth 1
    $LogPath=""
    $LogPath=$SSDDE.FullName
$LogName="Microsoft-Windows-StorageSpaces-Driver/Diagnostic"
    $LogID=""
    $LogID='1017','1018'
    Write-Host "        Checking Event Log $LogName for ID $LogID..."
    $SSDDEvents=@()
    $SSDDEventsOut=@()
    $SSDDEvents = Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable @{Path=$LogPath;Id=$LogID}
    #$SSDDEvents|Select-Object -First 1 | FL *
    $SSDDEventsOut=$SSDDEvents|Sort-Object MachineName|Sort-Object -Descending TimeCreated | Select-Object MachineName,LogName,TimeCreated,Id,LevelDisplayName,Message
    If ($Null -eq $SSDDEvents.count) { Write-Host "            No such EventId $LogId exists" -ForegroundColor Yellow } Else { Write-Host "            Found "($SSDDEvents).Count" Events for ID $LogId"}
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$SSDDEventsOut|Select-Object -Property `
                @{L='MachineName';E={[string]$_.MachineName}},
                @{L='LogName';E={[string]$_.LogName}},
                @{L='TimeCreated';E={[string]$_.TimeCreated}},
                @{L='Id';E={[string]$_.Id}},
                @{L='LevelDisplayName';E={[string]$_.LevelDisplayName}},
                @{L='Message';E={[string]$_.Message}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=YKz%2F0%2BLMBvGU97hc0sH%2FVf0%2F1bi9gyr6vb6r2sjxbHg%3D&tn=CluChkStorageSpacesDriverDiagnosticEvents'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
    $html+='<H2 id="StorageSpacesDriverDiagnosticEvents">Storage Spaces Driver Diagnostic Events</H2>'
    $html+="<h5><b>Key:</b></h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;These events indicate possible resource contention:</h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ID: 1017 - Took more than 30 seconds to aquire global lock</h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ID: 1018 - An exclusive lock was held for more than 30 seconds</h5>"
    $html+=""
    $html+=$SSDDEventsOut | ConvertTo-html -Fragment
If($SSDDEventsOut.count -eq 0){$html+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;No Storage Spaces Driver Diagnostic Events found</h5>'}
    $html=$html `
            -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
            -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
    $ResultsSummary+=Set-ResultsSummary -name $name -html $html
$htmlout+=$html
$html=""
    $Name=""


# Microsoft-Windows-StorageSpaces-Driver/Diagnostic Events
    #xml events ([xml]$myevents.Objects.Object[4].'#text').event.System.eventid.'#text'
    $Name="Hyper-V Network Events"
    Write-Host "    Gathering $Name..."
    $SSDDE=Get-ChildItem -Path $SDDCPath -Filter "System.EVTX" -Recurse -Depth 1
    $LogPath=""
    $LogPath=$SSDDE.FullName
$LogName="System"
    $LogID=""
    $LogID='252','253'
    Write-Host "        Checking Event Log $LogName for ID $LogID..."
    $SSDDEvents=@()
    $SSDDEventsOut=@()
    $SSDDEvents = Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable @{Path=$LogPath;Id=$LogID}
    #$SSDDEvents|Select-Object -First 1 | FL *
    $SSDDEventsOut=$SSDDEvents|Sort-Object MachineName|Sort-Object -Descending TimeCreated | Select-Object MachineName,LogName,TimeCreated,@{L="Id";E={"YYEELLOOWW$($_.Id)"}},LevelDisplayName,Message
    If ($Null -eq $SSDDEvents.count) { Write-Host "            No such EventId $LogId exists" -ForegroundColor Yellow } Else { Write-Host "            Found "($SSDDEvents).Count" Events for ID $LogId"}
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$SSDDEventsOut|Select-Object -Property `
                @{L='MachineName';E={[string]$_.MachineName}},
                @{L='LogName';E={[string]$_.LogName}},
                @{L='TimeCreated';E={[string]$_.TimeCreated}},
                @{L='Id';E={[string]$_.Id}},
                @{L='LevelDisplayName';E={[string]$_.LevelDisplayName}},
                @{L='Message';E={[string]$_.Message}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=YKz%2F0%2BLMBvGU97hc0sH%2FVf0%2F1bi9gyr6vb6r2sjxbHg%3D&tn=CluChkStorageSpacesDriverDiagnosticEvents'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
    $html+='<H2 id="Hyper-VNetworkEvents">Hyper-V Network Events</H2>'
    $html+="<h5><b>Key:</b></h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;These events indicate possible network resource congestion:</h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ID: 252 - Memory allocated for packets in a vRss queue on switch <> due to low resource on the physical NIC has increased</h5>"
    $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ID: 253 - Memory allocated for packets in a vRss queue on switch <> due to low resource on the physical NIC has reduced to 0MB</h5>"
    $html+=""
    $html+=$SSDDEventsOut | ConvertTo-html -Fragment
If($SSDDEventsOut.count -eq 0){$html+='<h5>&nbsp;&nbsp;&nbsp;&nbsp;No Hyper-V Network Events found</h5>'}
    $html=$html `
            -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
            -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
    $ResultsSummary+=Set-ResultsSummary -name $name -html $html
$htmlout+=$html
$html=""
    $Name=""


    # System Page File - Add by JG on 12/16/2022
    $Name="System Page File"
    Write-Host "    Gathering $Name..."
    #$GetComputerInfo=Foreach ($key in ($SDDCFiles.keys -like "*GetComputerInfo")) {$SDDCFiles."$key"}
    #Check if GetComputerInfo exisits
        IF($GetComputerInfo){
            $ClusterName=$SDDCFiles."GetCluster" | Select-Object Name,BlockCacheSize
            $PageFileInfoOut=""
            $PageFileInfoOut=$GetComputerInfo | Select-Object `
                @{L="ComputerName";E={$_.CsName}},`
                @{L="PageFileSize(MB)";E={$PFS=$_.OsFreeSpaceInPagingFiles/1KB;
                  IF($PFS -ne (51200 + $ClusterName.BlockCacheSize) -and $SysInfo[0].SysModel -notmatch "^APEX"){
                    IF($PFS -gt (51200 + $ClusterName.BlockCacheSize)){
                      "YYEELLLLOOWW"+$PFS
                    } else {
                      "RREEDD"+$PFS
                    }
                  }Else{$PFS}
		}},`
                @{L="BlockCacheSize";E={$ClusterName.BlockCacheSize}}
         }
      # HTML Report
        $html+='<H2 id="SystemPageFile">System Page File</H2>'
        $html+="<h5><b>Should be:</b></h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;51200MB + BlockCacheSize</h5>"
        $html+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://www.dell.com/support/manuals/en-us/ax-740xd/ashci_deployment_option_guide_switchless/updating-the-page-file-settings?guid=guid-6b016ead-a14f-46de-89ef-e313733ee4c1&lang=en-us' target='_blank'>Ref:https://www.dell.com/support/manuals/en-us/ax-740xd/ashci_deployment_option_guide_switchless/updating-the-page-file-settings?guid=guid-6b016ead-a14f-46de-89ef-e313733ee4c1&lang=en-us</a></h5>"
        $html+=""
        IF($PageFileInfoOut){$html+=$PageFileInfoOut | ConvertTo-html -Fragment}Else{$html+="GetComputerInfo.xml missing from SDDC. Nothing to display"}
        $html=$html `
                -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">' 
        $ResultsSummary+=Set-ResultsSummary -name $name -html $html
    $htmlout+=$html
    $html=""
        $Name=""
    
}
If($ProcessSDDC -ieq 'y'){
    Write-Host "Start to S2D validation Complete is $(((Get-Date)-$dstartSDDC).totalmilliseconds)"
    $htmlout2=""
    While (-not $Job.IsCompleted) {Sleep -Milliseconds 200;Write-Host "." -NoNewline}
    $PowerShell.EndInvoke($Job)
    $htmlout2=$Outputs
    $Runspace.Close()
    $Runspace.Dispose()
    IF(Test-Path "$env:temp\ClusterResultSummary.xml"){
        $ClusterResultSummary=Import-Clixml "$env:temp\ClusterResultSummary.xml"
        $ResultsSummary=$ClusterResultSummary+$ResultsSummary
        Remove-Item "$env:temp\ClusterResultSummary.xml" -Force -ErrorAction SilentlyContinue
    }
    $htmlout="$htmlout2$htmlout"
        $htmlout2=""
        if ($Job3) {
        if (-not $Job3.IsCompleted) {Write-Host "`nWaiting on Recent Cluster Events`n"}
        While (-not $Job3.IsCompleted) {Sleep -Milliseconds 500;Write-Host "." -NoNewline}
        $PowerShell3.EndInvoke($Job3)
        $htmlout2=$Outputs3
        $Runspace3.Close()
        $Runspace3.Dispose()
        IF(Test-Path "$env:temp\ClusterResultSummary2.xml"){
            $ClusterResultSummary=Import-Clixml "$env:temp\ClusterResultSummary2.xml"
            $ResultsSummary+=$ClusterResultSummary
            Remove-Item "$env:temp\ClusterResultSummary2.xml" -Force -ErrorAction SilentlyContinue
        }
        $htmlout="$htmlout$htmlout2"
    }
}
#endregion End Process SDDC
Write-Host "Complete time is $(((Get-Date)-$dstartSDDC).totalmilliseconds)"
#region Process TSR(s) with Drift
IF($ProcessTSR -ieq "y"){
 #Write-Host "*******************************************************************************"
 #Write-Host "*                                                                             *"
 #Write-Host "*                                 DriFT                                       *"
 #Write-Host "*                                                                             *"
 #Write-Host "*******************************************************************************"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $CurrentLoc=$ENV:TEMP
    Write-Host "Downloading latest version..."
    #Dev
    #$url = 'https://raw.githubusercontent.com/DellProSupportGse/internaltools/main/driftdev.ps1'
    #Prod
    $url = 'https://raw.githubusercontent.com/DellProSupportGse/source/main/drift.ps1'
    $output = "$env:TEMP\DriFT.ps1"
    $start_time = Get-Date
    Remove-Item $output -Force -ErrorAction SilentlyContinue
    Try{invoke-webrequest -Uri $url -UseDefaultCredentials -TimeoutSec 30 | Out-Null #-OutFile $output 
    Write-Output "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"}
    Catch{Write-Host "ERROR: Source location NOT accessible. Please try again later"-foregroundcolor Red
    }
    Finally{
        $TSRsToProcesswithDrift = $TSRLOC -join ","
        $params="-Cluchk $CluChkGuid -Input $TSRsToProcesswithDrift "
        #Import-Module $output
        Invoke-Expression('$module="RunDriFT";$repo="PowershellScripts"'+(new-object net.webclient).DownloadString($url))
        Invoke-RunDriFT $params 
    }
#}
}

IF($ProcessTSR -ieq "y"){
# Creates CluChk Html Report
    # Import Drift xml data
$name="Sel Log Errors and Warnings"
Write-Host "Checking for BIOS and NIC Configuration..."

$BIOSandiDRACCfg=@()
#$SourcePath=""
#$SourcePath=([regex]::match($TSRLoc,'(.*\\).*')).Groups[1].value
$BIOSandiDRACCfg=Get-ChildItem -Path ((Split-Path -Path $TSRLOC) | sort -Unique) -Filter "$($CluChkGuid)_BIOSandNICCFG.xml" -Recurse -Depth 1  | Sort-Object LastWriteTime | Select-Object -Last 1 | import-clixml
        

    IF($BIOSandiDRACCfg.count -gt 0){ 

    # Sel Logs
$SelOut2Report=@()
$Name="Sel Log Errors and Warnings"
Write-Host "    Gathering $Name..."
IF(($BIOSandiDRACCfg | Where-Object{$_.Record -ne $null}).count -gt 0){
#$BIOSandiDRACCfg|ft
$SelOut=$BIOSandiDRACCfg |Where-Object{$_.Record -ne $Null}
IF($SelOut.count -gt 0){
ForEach($Sel in $SelOut){
$SelOut2Report+= $Sel|Select-Object node,Record,DateTime,`
@{L='Severity';E={IF($_.Severity -eq 3){"YYEELLLLOOWW"+$_.Severity}`
ElseIF($_.Severity -eq 4){"RREEDD"+$_.Severity}}},Description
}
}
}
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$SelOut|Select-Object -Property `
                @{L='Node';E={[string]$_.Node}},
                @{L='Record';E={[string]$_.Record}},
                @{L='DateTime';E={[string]$_.DateTime}},
                @{L='Severity';E={[string]$_.Severity}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=leVgQ3%2FAL7zd0ev0ikrM4s84OxmLmLhy0LSimv17zyU%3D&tn=CluChkSelLogErrorsandWarnings'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
            $html+='<H2 id="SelLogErrorsandWarnings">Sel Log Errors and Warnings</H2>'
            $html+='<H5>&nbsp;&nbsp;&nbsp;&nbsp;Only showing the last 30 days</H5>'
            $html+=""
            IF(($BIOSandiDRACCfg | Where-Object{$_.Record -eq $null}).count -gt 0){
                $html+='<H5>&nbsp;&nbsp;&nbsp;&nbsp;No records found.</H5>'
            }
            $html+=$SelOut2Report|Where-Object{$_ -iNotMatch 'System.__ComObject'}|Sort-Object node,DateTime  | ConvertTo-html -Fragment
            $html=$html `
                -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
                -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name="" 
        

    # Switch to Host map
        IF(($BIOSandiDRACCfg | Where-Object{$_.SwitchMacAddress -ne $Null}).count -gt 0){
            $Name="Switch to Host Map"
            Write-Host "    Gathering $Name..."
            $SwMap=@()
            $FoundSwMap=@()
            $SwMap=$BIOSandiDRACCfg |Where-Object{$_.SwitchMacAddress -ne $Null}
            IF($SwMap.count -gt 0){
                ForEach($SWP in $SwMap){
                    $SWPMac=$swp.HostNICMacAddress -replace ":","-"
                    $FoundSwMap+=$SWP|Select-Object HostName,SwitchMacAddress,SwitchPortConnectionID,HostNicSlotPort,`
                    @{Label='Name';Expression={($GetNetAdapterXml|Where-Object{$_.MacAddress.startswith($SWPMac)}).Name}},`
                    @{Label='HostNICMacAddress';Expression={IF(($GetNetAdapterXml|Where-Object{$_.MacAddress.startswith($SWPMac)}).MacAddress -eq $Null){$SWPMac -replace "-",":"}Else{($GetNetAdapterXml|Where-Object{$_.MacAddress.startswith($SWPMac)}).MacAddress -replace "-",":"}}}
                }
            } 
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$FoundSwMap|Select-Object -Property `
                @{L='HostName';E={[string]$_.HostName}},
                @{L='SwitchMacAddress';E={[string]$_.SwitchMacAddress}},
                @{L='SwitchPortConnectionID';E={[string]$_.SwitchPortConnectionID}},
                @{L='HostNicSlotPort';E={[string]$_.HostNicSlotPort}},
                @{L='Name';E={[string]$_.Name}},
                @{L='HostNICMacAddress';E={[string]$_.HostNICMacAddress}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=Gm2fK%2BBG%2B%2Fo4ZrVmjPYnn45LI4Rj03OqyZMTBdSZI4I%3D&tn=CluChkSwitchtoHostMap'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report            
            $html+='<H2 id="SwitchtoHostMap">Switch to Host map</H2>'
            $SwMapSB=""
            $html+=$SwMapSB
            $html+=$FoundSwMap|Sort-Object HostName  | ConvertTo-html -Fragment
            $html=$html `
              -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name=""
        }

    # BIOS and iDRAC configuration output
            $Name="BIOS and iDRAC configuration"
            Write-Host "    Gathering $Name..."
            $BIOSandiDRACCfg=$BIOSandiDRACCfg | Where-Object{$_.DesiredValue -ne $null} | Sort-Object "Setting Name",Hostname | Select-Object HostName,ServiceTag,Type,@{Label='SettingName';E={$_.'Setting Name'}},Device,"Setting Name",`
            @{Label='CurrentValue';Expression={$_.CurrentValue -replace [Regex]::Escape("***"),"RREEDD"}},DesiredValue
            # Add new output format
            $BIOSandiDRACCfgtbl = New-Object System.Data.DataTable "Compare"
            $BIOSandiDRACCfgtbl.Columns.add((New-Object System.Data.DataColumn("Type")))
            $BIOSandiDRACCfgtbl.Columns.add((New-Object System.Data.DataColumn("Category")))
            $BIOSandiDRACCfgtbl.Columns.add((New-Object System.Data.DataColumn("Device")))
            $BIOSandiDRACCfgtbl.Columns.add((New-Object System.Data.DataColumn("SettingName")))
            $BIOSandiDRACCfgtbl.Columns.add((New-Object System.Data.DataColumn("DesiredValue")))

            ForEach ($a in ($BIOSandiDRACCfg.ServiceTag | Sort-Object -Unique)){
                $BIOSandiDRACCfgtbl.Columns.Add((New-Object System.Data.DataColumn([string]$a)))}
                $a=$null
                ForEach($b in ($BIOSandiDRACCfg | Sort-Object SettingName )){
                    IF($b.SettingName.length -gt 2 -and $b.SettingName.length -notmatch 'System.__ComObject'){
                        if ($b.SettingName -ne $a) {
                            $a=$b.SettingName
                            if ($Null -ne $a) {
                                IF($row.constructor -inotmatch 'System.__ComObject'){
                                $BIOSandiDRACCfgtbl.rows.add($row)}}
                            $row=$BIOSandiDRACCfgtbl.NewRow()
                            $row["Type"]=$b.Type 
                            $row["Category"]=$b.Category
                            $row["Device"]=$b.Device
                            $row["SettingName"]=$b.SettingName
                            $row["DesiredValue"]=$b.DesiredValue
                        }
                        $row["$($b.ServiceTag)"] = $b.CurrentValue
            
                    }
                }
            #$BIOSandiDRACCfgtbl |Format-Table SettingName,AvailableVersion,???????
            $BIOSandiDRACCfgOut=$BIOSandiDRACCfgtbl|Where-Object{$_ -iNotMatch 'System.__ComObject'}|Sort-Object Type,SettingName | Select-object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors
            #$BIOSandiDRACCfgOut|ft
        #Azure Table
            $AzureTableData=@()
            $AzureTableData=$BIOSandiDRACCfg|Select-Object -Property `
                @{L='HostName';E={[string]$_.HostName}},
                @{L='ServiceTag';E={[string]$_.ServiceTag}},
                @{L='Type';E={[string]$_.Type}},
                @{L='SettingName';E={[string]$_.SettingName}},
                @{L='CurrentValue';E={[string]$_.CurrentValue}},
                @{L='Device';E={[string]$_.Device}},
                @{L='ReportID';E={$CReportID}}
            $PartitionKey=$Name -replace '\s'
            $TableName="CluChk$($PartitionKey)"
            $SasToken='?sv=2019-02-02&si=CluChkUpdate&sig=GdjiCg7GVltbnYdPPt1ZySwgEhgHunJUZ4FS3nmYkgw%3D&tn=CluChkBIOSandiDRACconfiguration'
            $AzureTableData | %{add-TableData -TableName $TableName -PartitionKey $PartitionKey -RowKey (new-guid).guid -data $_ -SasToken $SasToken}

        # HTML Report
            $html+='<H2 id="BIOSandiDRACconfiguration">iDRAC, BIOS and QLogic NIC configuration</H2>'
            $iDRACandBIOSconfigurationSB=""
            $iDRACandBIOSconfigurationSB+="<h5><b>Should be:</b></h5>"
            $iDRACandBIOSconfigurationSB+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;CurrentValue=DesiredValue</h5>"
            $iDRACandBIOSconfigurationSB+="<h5>&nbsp;&nbsp;&nbsp;&nbsp;<a href='https://www.dell.com/support/kbdoc/en-us/000135856/bios-and-idrac-configuration-recommendations-for-servers-in-a-dell-emc-solutions-for-microsoft-azure-stack-hci' target='_blank'>Ref: https://www.dell.com/support/kbdoc/en-us/000135856/bios-and-idrac-configuration-recommendations-for-servers-in-a-dell-emc-solutions-for-microsoft-azure-stack-hci</a></h5>"
            $html+=$iDRACandBIOSconfigurationSB
            $html+=$BIOSandiDRACCfgOut | ConvertTo-html -Fragment
            $html=$html `
              -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'
            $ResultsSummary+=Set-ResultsSummary -name $name -html $html
            $htmlout+=$html
            $html=""
            $Name=""
        }
        Else{Write-host "    No BIOSandNICCFG.xml found. Nothing to do." -ForegroundColor Yellow }
    
    # Cleanup BIOSandNICCFG.xml
        #$CleanupBIOSandNICCFG=""
        $GUIDFIle2Delete=$CluChkGuid+"_BIOSandNICCFG.xml"
        $GUIDFIle2DeletePath=Get-ChildItem -Path (Split-Path -Path $TSRLOC) -Filter $GUIDFIle2Delete -Recurse -File -ErrorAction SilentlyContinue
        Remove-Item -Path $GUIDFIle2DeletePath.FullName -Force -ErrorAction SilentlyContinue
}
#endregion Process TSR with Drift


#region Create CluChk Html Report
IF($selection -ne "4"){
    $htmloutReport = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml" lang="en">
    <head>'
    $htmloutReport+=$htmlStyle
    $htmloutReport+='<title>CluChk Report</title>'
    $htmloutReport+='<meta charset="UTF-8">'
    $htmloutReport+="</head>"
    $htmloutReport+="<body>"

    $html='<h1>CluChk Configuration Report</h1>'
    $html+='<h3>&nbsp;Version: '+ $CluChkVer +' </h3>'
    #$RunDate=Get-Date
    $html+='<h3>&nbsp;Run Date: '+ $RunDate +' </h3>'
    $html+=''
    $html+='<h1>Results Summary</h1>'
    $html+= $ResultsSummary | Sort-Object Name | Select-Object `
    @{Label='Name';Expression={
    $Part1='<A href="#'
    $Part2=$_.Name -replace '\s',""
    $Part3='">'
    $Part4=$_.Name
    $Part5='</A>'
    $Part1+$Part2+$Part3+$Part4+$Part5
    }},`
    @{Label='Warnings';Expression={IF($_.Warnings -gt 0){"YYEELLLLOOWW"+$_.Warnings}Else{$_.Warnings}}},`
    @{Label='Errors';Expression={IF($_.Errors -gt 0){"RREEDD"+$_.Errors}Else{$_.Errors}}}| ConvertTo-html -Fragment
    $html=$html `
     -replace '&gt;','>' -replace '&lt;','<' -replace '&quot;','"'`
     -replace '<td>RREEDD','<td style="color: #ffffff; background-color: #ff0000">'`
     -replace '<td>YYEELLLLOOWW','<td style="background-color: #ffff00">'
    $htmloutReport+=$html

    #$htmlout=$htmlout -replace '<table>','<table class="Sort-Objectable">'
    $htmlout=$htmlout `
    -replace [Regex]::Escape("***"),''`
    -replace '&amp;gt;','>'`
    -replace '&amp;lt;','<'`
    -replace '&quot;','"'`
    -replace '&lt;br&gt;','<br>'
    $htmloutReport+=$htmlout

    # Close body
    $htmloutReport+='</body></html>'

    # Generate HTML Report
    $SDDCFileName=($SDDCPath -split '\\')[-1]
    If($CluChkReportLoc.Count -gt 1) {$CluChkReportLoc = $CluchkReportLoc[0]}
    $HtmlReport= Join-Path -Path $CluChkReportLoc -ChildPath CluChkReport_v$CluChkVer-$DTString$SDDCFileName.html
    Write-Host ("Report Output location: " + $HtmlReport)
    if (Test-Path "$HtmlReport") {Remove-Item $HtmlReport}
    Out-File -FilePath $HtmlReport -InputObject $htmloutReport -Encoding ASCII
    # open HTML file
    if (test-path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice) {Invoke-Item($HtmlReport)}
}
#endregion  Create CluChk Html Report

#region Create CluChk Performance Html Report
IF($SDDCPerf -ieq "YES"){
    If (-not $Job2.IsCompleted) {Write-Host "`nProcessing HCI Performance data" -NoNewline}
    While (-not $Job2.IsCompleted) {Sleep -Milliseconds 200;Write-Host "." -NoNewline}
    Write-Host "`n"
    $Outputs2
    $PowerShell2.EndInvoke($Job2)
    $Runspace2.Close()
    $Runspace2.Dispose()
    $SDDCFileName=($SDDCPath -split '\\')[-1]
    If($CluChkReportLoc.Count -gt 1) {$CluChkReportLoc = $CluchkReportLoc[0]}
    $HtmlReport= Join-Path -Path $CluChkReportLoc -ChildPath CluChkPerfReport_v$CluChkVer-$DTString$SDDCFileName.html
    Write-Host ("Report Output location: " + $HtmlReport)
    # open HTML file
    if (test-path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice) {Invoke-Item($HtmlReport)}
}
#endregion  Create CluChk Html Report


TLogCleanup 
$allArrayout=@()
$allArray=@()
$ServiceTagList=@()
try {Stop-Transcript -ErrorAction SilentlyContinue} catch {}


break
$GetVMProcCount=Foreach ($key in ($SDDCFiles.keys -like "*GetVM" )) {
    $VMProcCount=$SDDCFiles."$key" | ? State -match "Running" | Select ComputerName,ProcessorCount
    $VMProcCount | Sort ComputerName -Unique | Select ComputerName,@{L="VMProcessorCount";E={($VMProcCount | Measure-Object ProcessorCount -Sum).Sum}},`
    @{L="RunningVMs";E={$VMProcCount.count}},@{L="LogicalProcessor";E={($SDDCFiles."$($key)Host").LogicalProcessorCount}}

}
}
