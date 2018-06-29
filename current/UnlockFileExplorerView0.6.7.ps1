##############################################

#UnlockFileExplorerView

#Version           = "0.6.7"

#Copyright:         Free to use, please leave this header intact
#Author:            Marek Obuchowski (mobuchowski.pl)
#Credits:           EMEA Territory Services Team for all the bits and pieces that allowed this script to be created
#Credits:           Jos Lieben (http://www.ogd.nl) OneDriveMapper creator for sharing his awesome work
#Purpose:           To atomate opening SharePoint Online library in IE so the user is autoamtically authenticated and mapped drives unlocked

##############################################

####Changelog
<#
####0.1
#1 First release
#2 Testing access 
#3 Opening library in File Explore mode
#4 Closing windows explorer window
#5 Closing IE processes


####0.2
#1 Testing connectivity to SharePoint and DC (for ensuring we are on corporate network)
#2 Auto discovery of mapped SharePoint libraries now replaced hardcoded paths

####0.3
#hmm... not sure


####0.4
#1 Added logic to check if no SharePoint mapping, stopping if not mapped
#2 Little bit of code cleanup
#3 Updated info when access already enabled

####0.5
#1 Windows 10 v1709 compatibility added (open blank IE page first, kill all iexplore processes at the end)
#2 Change drive discovery from PSDrive to Net Use

####0.6
#1 Hardcoded DC replaced with current logon server
#2 BaseURL building code updated - script now extract it from the mapped library, this mean that script can be used with any SharePoint tenant now
#3 Typos in commnets now corrected.. oh yeah!!

####0.6.1
#hmm... not sure

####0.6.2
#1 Added logging to file
#2 Added extended debug mode
#Startign rewiriting IE handling

####0.6.3
#1 Rewriteen IE handling
#   -Checking if IE is snadboxed by Protected View
#   -Added fucntion to kill IE if sandboxed

####0.6.4
#1 Added support for two mapping fomrats: UNC and URL
#2 Code cleanup
#   -moved part of the URLs building code to getMappedDrives function so it sits with within one function

#0.6.5
#1 Fixed the issue with building File explorer URL for mapped drives with library sub-folders

#0.6.6
#1 Extended Debug output to the console and file
#2 Updated Error messages handling
#3 Added code to add tenant to the IE pop-up blocker whitelist

#0.6.7
#1 Updated network and sharepoint connectivity test and result logging
#2 Removed unnecesary clutter form Error messages

#>


####To-do:
<#
#1 Update drive discovery to check for the best mapped drive
#2 Create function for building URLs and move code to it
#3 Write Protected Mode discovery code
#4 Add timout fr waiting for Library window in Windows Explorer
#5 Native authenitcation rather than IE?
#6 Add support for hiding the console entirely
####>

param(
    [Switch]$hideConsole,                                                            #Show or hide output in the console based on the script parameter
    [Switch]$debugon                                                                 #Enable debug via parameter        
    )


$version           = "0.6.7"

if($debugon){$debugMode      = $True}                                                #Enable debug based on parameter
if($hideConsole){$debugMode  = $false}                                               #Force DebugMode Off if script run without Consloe output                                                 
#$debugMode         = $True                                                          #Use to overvrite parameters .Set to $True if you want the script to ignore current state of the access and go ahead with all actions. Set to $False for normal operation.

#Variable definitions - do not change
$unlocked          = $null                                                           #Variable for holding current state of access to mapped libraries: $True = access already unlocked, $False = access locked
$done              = $null                                                           #Variable for holding for the state of unlocking process: $True = unlock process completed, $False = unlock process in progress
$IE                = $null                                                           #Variable for storing Internet Explorer object, for closing once unlocking process is completed
$urlOptions        = "Forms/AllItems.aspx?ExplorerWindowUrl="                        #Variable for storing part of the URL responsible for opening library in File Explorer
$baseURL           = $null                                                           #Variable for storing SharePoint tenant URL
$spoConnection     = $null                                                           #Variable for storing SharePoint connectivity results
$dcConnection      = $null                                                           #Variable for storing DC connectivity results
$dcName            = $env:LOGONSERVER.Substring(2)                                   #Setting up DC variable for testing connectivity
$lob               = $null                                                           #Variable for storing LOB name
$siteName          = $null                                                           #Variable for storing Site name
$libraryName       = $null                                                           #Variable for storing Library name
$lob2              = $null                                                           #Variable for building URLS
$siteName2         = $null                                                           #Variable for building URLS
$libraryName2      = $null                                                           #Variable for building URLS
$subfolders        = $null                                                           #Variable for building URLS
$logfile           = ($env:APPDATA + "\UnlockFileExplorerView_$version.log")         #Logfile to log to                                                         
$mappingFormat     = $null                                                           #Variable for storing mapping format
$i_MaxLocalLogSize = 2                                                               #Set the Max local log size in MB
$protectedMode     = $True                                                           #Variable for storing info re Protected Mode on or off in IE
$showConsoleOutput = $True                                                           #Show console output by default

#Hide console it parameter found
if($hideConsole){
    $showConsoleOutput     = $False
}

if($showConsoleOutput -eq $False){
    $t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
    try{
        add-type -name win -member $t -namespace native
        [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
    }catch{$Null}
}

function log{
    <#
    -------------------------------------------------------------------------------------------
    Manage the local log file size
    Always keep a backup
    #credits to OneDriveMapper creator
    -------------------------------------------------------------------------------------------
    #>
    param (
        [Parameter(Mandatory=$true)][String]$text,
        [Switch]$fout,
        [Switch]$warning,
        [Switch]$debugg
    )
    if($fout){
        $text = "ERROR   | $text"
    }
    elseif($warning){
        $text = "WARNING | $text"
    }
    elseif($debugg){
        $text = "DEBUG   | $text"
    }
    else{
        $text = "INFO    | $text"
    }

    try{
        Add-Content $logfile "$(Get-Date) | $text"
    }catch{$Null}
    if($showConsoleOutput){
        if($fout){
            Write-Host $text -ForegroundColor Red
        }elseif($warning){
            Write-Host $text -ForegroundColor Yellow
        }elseif($debugg){
            Write-Host $text -ForegroundColor Cyan
        }else{
            Write-Host $text -ForegroundColor Green
        }
    }
}

function ResetLog{
    <#
    -------------------------------------------------------------------------------------------
    Manage the local log file size
    Always keep a backup
    #credits to Steven Heimbecker
    -------------------------------------------------------------------------------------------
    #>
    #Restart the local log file if it exists and is bigger than $i_MaxLocalLogSize MB as defined below
    [int]$i_LocalLogSize
    if ((Test-Path $logfile) -eq $True){
        #The log file exists
        try{
            $i_LocalLogSize=(Get-Item $logfile).Length
            if($i_LocalLogSize / 1Mb -gt $i_MaxLocalLogSize){
                #The log file is greater than the defined maximum.  Let's back it up / rename it
                #Blank line in the old log
                
                log -text "******** End of log - maximum size ********"
                #Save the current log as a .old.  If one already exists, delete it.
                if ((Test-Path ($logfile + ".old")) -eq $True){
                    #Already a backup file, delete it
                    Remove-Item ($logfile + ".old") -Force -Confirm:$False
                }
                #Now lets rename 
                Rename-Item -path $logfile -NewName ($logfile + ".old") -Force -Confirm:$False
                #Start a new log
                log -text "******** Log file reset after reaching maximum size ********`n"
            }
        }catch{
            $ErrorMessage = $_.Exception.Message
            log -text "there was an issue resetting the logfile!" -fout
            log -text $ErrorMessage
        }
    }
}

function checkDrive ($URL) {                                             #Function for testing current state of access
    log -text "Testing access state"                                     #Inform user   
    $result = test-path -path $URL                                       #Test if path accessible
    if($DebugMode){log -text ("Testing access state: "+ $URL) -debugg }

   if ($result) {
    log -text "Access already enabled"                                   #Inform user
    
    } else {
        log -text "Access locked, unlock process started" -warning       #Inform user
        }
   return $result                                                        #Return $True if access already unlocked, $False if access locked
   }

function openIE {                                                        #Function for opening IE and navigating to View in File Explorer URL         
        
    try {
        
        if($debugMode){log -text ("Testing if IE popup blocker conffigured to allow tenant URL: " + $baseURL) -debugg}

        $tmp = TestRegistryValue -path "hkcu:\Software\Microsoft\Internet Explorer\New Windows\Allow" -value "$baseURL"  #Check if popup blocker already configured

        if(!$tmp){                                                                                                       #Add tenant URL to IE popup blocker whitelist
            if($debugMode){log -text ("Adding IE popup blocker exception: " + $baseURL) -debugg}
            Set-Location "hkcu:\Software\Microsoft\Internet Explorer\New Windows\Allow"                   
            New-ItemProperty -Path . -Name "$baseURL" -Type Binary                                                       
            if($debugMode){log -text "Added popup blocker exception" -debugg}
             }else {if($debugMode){log -text "IE popup blocker already configured" -debugg}}

        log -text "Opening library in View in File Explorer mode"                                            #Inform user
        
        $IE=new-object -com internetexplorer.application                                                     #Create IE object
        $IE.visible = $debugMode                                                                             #Hide IE window if not in debug mode
        
        $IE.navigate2("about:blank")                                                                         #This is added for compatibility with Windows 10 v1709 - for some reason, IE fails to open FileExplrorer URL when there is no IE window opened before
        sleep -s 1                                                                                           #Give time to open IE                                                                  
        
        $IE.navigate2($URL)                                                                                  #Navigate to View in File Explorer URL
        $IE.visible = $debugMode                                                                             #Hide IE window if not in debug mode
   
        do  {$done  = closeFileExplorerWindow} while (!$done)                                                #Loop for checking if library already open. Check for the library window, set process state. Repeat if process state not done, quit loop if done
        

        if($IE.HWND -eq $null){                                                                                     
            if ($debugMode) {log -text "IE.HWND not found - killing IE" -debugg}                             #Write dubg log
            killIE                                                                                           #Kill IE processes if HWDN not found for the IE object  
        } else {
            if ($debugMode) {log -text "IE.HWND found - Quitting IE gracefully" -debugg}                     #Write dubg log
            $IE.Parent.Quit()
            
            if($protectedMode){
                if ($debugMode) {log -text "Kiling IE anyway as URL in Zone with Protected Mode on" -debugg} #Write dubg log
                killIE                                                                                       #Kill IE despite of callin Quite method on the COM object (Quit does not work if URL is in the Zone with protected mode on). This can be removed once base URL is added to Trusted Sites where protected mode is off.
                }                                                             
        }

    }catch{
        $ErrorMessage = $_.Exception.Message
        log -text ("Failed to manage Internet Explorer:") -fout
        log -text $ErrorMessage -fout
        killIE      
        ExitScript
    }
}

function killIE {
    Stop-Process -Name iexplore                                           #Kill IE processes
    sleep -s 1
    $IE         =new-object -com internetexplorer.application             #Open blank IE to get rid off the warning message regarding previous session that was killed above
    $IE.visible = $debugMode                                              #Hide IE window if not in debug mode  
    while ($IE.busy) {sleep -m 500}                                       #Give time to open IE
    $IE.Quit()                                                            #Quit IE gracefully
}

function closeFileExplorerWindow {                                        #Function for testing if Library already open, closing File Explorer window and changing the state of the unlocking state

    try {
        $fs = (New-Object -comObject Shell.Application).Windows()|`       #Create object to store list of File Explorer windows containing SharePoint library
            where-object { ($_.LocationName -like "*$libraryName*")`
            -and  ($_.Name -like "*File Explorer*") }
    
        if ($fs) {                                                        #Check if File Explorer window is open.
               if ($debugMode) {
                log -text ("Library open`n" + $fs) -debugg                #Wait for keystroke if in Debug mode
                pause
                write-host ""                
                }

               log -text "Closing File Explorer Window"                   #Inform user
               $fs | ForEach-Object {$_.Quit()}                           #Window open: Close and return $true.
               return $true

        } else {                                                          #Window not open, return $false
                return $false}
    }catch{
    $ErrorMessage = $_.Exception.Message
    log -text ("Failed to manage Library in File Explorer:") -fout
    log -text $ErrorMessage -fout
    exitScript
    }
}

#Function for testing registry values - returns $True if exists, $Flase if does not exist
function TestRegistryValue {

param (

 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Path,

[parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Value
)

try {

Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
 return $true
 }

catch {

return $false

}

}


function testConnection ($a, $type) {                                            #Function to test connection to the SharePoint server
    log -text "Testing connectivity to $type"                                #Inform user
    $test = Test-Connection -computer $a -quiet -count 1                  #Test connection to the SharePoint server
    if ($test) {
        log -text "Connection to $type succesfull"                        #Inform user
        } else {
            log -text "Connection to $type failed" -fout                  #Inform user
            }
    return $test                                                          #Return results
    }

function getMappedDrives {                                                #Function for getting mapped Sharepoint drives and setting up library for opening in View in File Explorer mode

    #Get-PSDrive method replaced with NET USE
    #Get-PSDrive does not return mapped drives on some machines, Net Use does not have this issue
    #$mp           = Get-PSDrive -PSProvider FileSystem | where-object {$_.DisplayRoot -like "*$baseURL*"} | select-object -first 1    #Get drives and select first entry
    
    #NET USE method
    $nu = net use                                                        #Collect Net Use output

    if ($debugMode){                                                     #Write debug log
        log -text "*****************" -debugg
        log -text ("Net Use output:") -debugg
        log -text ("$nu") -debugg
        log -text "*****************" -debugg

        }          

    foreach ($item in $nu) {                                             #Get the first instance (Get-PSDrvive method) or last instance (Net Use method) of the SharePoint Online mapping
            if($item -like "*sharepoint.com*") {
                if ($debugMode){log -text ("NET USE line with *sharepoint.com*: " + $item) -debugg}  #Write debug log
                $item = $item -split ":"                  #Split to get rid off drive letter
                if ($debugMode){log -text ("NET USE line split by colon:" + $item) -debugg}          #Write debug log

                #Extract mapped URL based on the format of the mapping
                switch ($item.length) {
                    2 {
                        $mp = $item[1].trim()
                        $mappingFormat = "UNC"           #Format: \\<tenanturl>@SSL\sites\EMEA-TSG\Shared Documents
                        if($debugMode){log -text("Mapped drive in UNC format: "+ $mp) -debugg}
                        }                                
                    3 {
                        $mp = $item[2].trim()
                        $mappingFOrmat = "URL"           #Format: https://<tenanturl>/sites/EMEAITManagement/Shared Documents
                        if($debugMode){log -text("Mapped drive in URL format: "+$mp) -debugg}
                        }
                    default {
                        log -text "Unknown mapping format found. Contact your IT support team" -fout
                        log -text $item -fout
                        }                                
                    }             
            }  
        }

    if ($mp -ne $null) {                                                  #Check if any mapping exist, proceed with extracting if yes
        #Replaced with Net Use version
        #$tempMappedURL    = $mp.displayroot                              #Set mapped URL to what already mapped
        $tempMappedURL     = $mp
        $array             = $tempMappedURL -split "sites"                #split url to base and site parts
        
        if ($debugMode){
            log -text ("tempMappedURL: " + $mp) -debugg
            log -text ("tempMappedURL split by sites word: " +$array) -debugg

        }

        #Manipulate base URL to format supportable by Test-Connection and Internet Explorer
        $tempBaseURL       = $array[0]                                                                            #extract base url
        if($DebugMode){log -text ("tempBaseURL extracted from tempMappedURL: " + $tempBaseURL) -debugg}
        $tempBaseURL       = $tempBaseURL.Substring(2)
        if($DebugMode){log -text ("tempBaseURL with removed two initial charactes: " + $tempBaseURL) -debugg}        
        $tempBaseURL       = $tempBaseURL.TrimEnd('@SSL\')
        if($DebugMode){log -text ('tempBaseURL with removed trailing "@SSL\": '+$tempBaseURL) -debugg}
        
        $tmp               = $array[1]                                                                            #extract site part
        if($DebugMode){log -text ("Temporary site part: "+$tmp) -debugg}

        $tmp = $tmp.substring(1)                                                                                  #delete initial /
        if($DebugMode){log -text ("Temporary site part without first character: " + $tmp) -debugg}
                                                                          
        $array.Clear()                                                                                            #clear array for future use
        
        
        switch ($mappingFormat) {
            "UNC" {
                $array = $tmp.Split('\')
                $tempBaseURL = $tempBaseURL.TrimEnd('\')         #split site part into an array if delimiter found
                }                              
            "URL" {
                $array = $tmp.Split('/')
                $tempBaseURL = $tempBaseURL.TrimEnd('/')         #split site part into an array if delimiter found
                }                              
        
        }
        if ($array.lengt -gt 1) {$array = $array | ? {$_}}       #delete empty elements if more than 1 element exist
    
            switch ($array.length){                              #Extract variables based on the array length
        
                1 { #Entire site mapped, no LOB and library
                    $tempLOB   = $nul
                    $tempSite  = $array[0]
                    $tempLib   = $null
                    }

                2 { #Entire site mapped, no library OR library mapped but no LOB
                    $tempLOB   = $null #$array[0]
                    $tempSite  = $array[0]
                    $tempLib   = $array[1]
                    }
        
                3 {#Library mapped, no subfolders
                    $tempLOB   = $array[0]
                    $tempSite  = $array[1]
                    $tempLib   = $array[2]
                    }

                default {#Library mapped with subfodlers
                    $tempLOB  = $array[0]
                    $tempSite = $array[1]
                    $tempLib  = $array[2]
                    $array    = $array[3..($array.Length-1)]
                    foreach ($element in $array){
                        if ($tempLibWithSubFold -eq $null) {
                        $tempLibWithSubFold = $element
                        }else{
                            $tempLibWithSubFold = $tempLibWithSubFold + "2%F" + $element
                            }
                        }
                    }
                }
    }else{}
    

    #Return values
    $tempMappedURL
    $tempBaseURL
    $tempLOB
    $tempSite
    $tempLib
    $mp
    $mappingFOrmat
    $tempLibWithSubFold
}

function ExitScript{
    log -text "******** End log ********"
    Start-Sleep -s 5         #Give time to read the script
    exit
    }

#### Start script ####

ResetLog #Reset log :D :P

log -text "******** Start log ********"

if($debugMode){log -text "Debug Mode enabled" -warning}


#Check connections and current access state
if (testConnection $dcName "Corporate network") {                                #Check connection to the logon server, abort if fails
#proceed
}else{
    log -text ("Not connected to the Corporate network") -fout
    exitScript
    }  

#getMappedDrives
try {
    $return        = getMappedDrives
    }
catch {
    $ErrorMessage = $_.Exception.Message
    log -text "Unable to obtain fileds from the mapped drives:" -fout 
    log -text $ErrorMessage -fout
    exitScript
}

$mappedURL     = $return[0]
$baseURL       = $return[1]
$lob           = $return[2]
$siteName      = $return[3]
$libraryName   = $return[4]
$driveMapped   = $Return[5]
$mappingFOrmat = $Return[6]
$subfolders    = $Return[7]


if (testConnection $baseURL "SharePoint Online Servers") {                               #Check connection to the SharePoint Online server, abort if fails
#proceed    
}else{
    log -text ("No connection to the SharePoint Online Servers") -fout
    exitScript
    }



#Check if SharePoint mapping present, exit if not
try {
    if ($driveMapped -eq $null) {                                              
            log -text "No SharePoint mapping found. Terminating" -warning
            exitScript                                                               #Exit script
        }else{

            #Display collected information
            if($mappedURL)    {log -text ("Mapped url: " + $mappedURL)}   else{log -text ("Mapped url: NOT FOUND") -warning}
            if ($baseURL)     {log -text ("Base URL: "   + $baseURL)}     else{log -text ("Base URL: NOT FOUND"  ) -warning}
            if ($lob)         {log -text ("LOB: "        + $lob)}         else{log -text ("LOB: NOT FOUND"       ) -warning}
            if ($siteName)    {log -text ("Site: "       + $siteName)}    else{log -text ("Site: NOT FOUND"      ) -warning}
            if ($libraryName) {log -text ("Library: "    + $libraryName)} else{log -text ("Library: NOT FOUND"   ) -warning}
            
            if ($lob -ne $null) {$lob2 = $lob + "/"}                           #If LOB contains data, add / for building URL
            if ($siteName -ne $null) {$siteName2 = $siteName + "/"}            #If Site contains data, add / for building URL
            if ($libraryName -ne $null) {                                      #If Library contains data, prepare for building URL
                
            $libraryName2 = $libraryName -replace " ","%20"                         
            $libraryName2 = $libraryName + "/"
                }
            if ($lob -ne $null) {$lob = "%2F" + $lob}

        }
}catch{
    $ErrorMessage = $_.Exception.Message
    log -text ("Unable to build URL components based on fileds extracted from mapped drives:") -fout
    log -text $ErrorMessage -fout
    exitScript
}


#Proceed if libary found in the mapped drive

if ($libraryName -ne $null) {                                              #Stop if no SharePoint or no library found
    
    #Build File Explorer View URL.
    $URL       = "https://" + $baseURL + "/sites/" + $lob2 + $siteName2 + $libraryName2 + $urlOptions + "%2Fsites" + ($lob -replace '-','%2D') + "%2F" + $siteName +"%2F" + ($libraryName.TrimEnd('/') -replace " ", "%20")   #Build URL for opening in View in File Explorer mode
    if($DebugMode) {log -text ("View in File Explorer URL: " + $URL) -debugg}
    
    $unlocked = checkDrive ($mappedURL)                                    #Check the current state of access and store in the $Unlocked variable                          
    
        if ($debugMode) {
            $unlocked = $false                                                 #Set the access state to $False if in DEbug mode. This allows testing entire scritpt
            log -text 'Debug mode: setting up $unlocked to $False for testing' -warning
            }                                   
            


        if ($unlocked) {                                                   #Check if access already enabled, terminate if yes
            log -text "You can now use mapped drives. Terminating"         #Inform user
            } else {
                openIE                                                     #Open library in IE, with View in File Explorer URL                 
                log -text "You can now use mapped drives. Terminating"     #Inform user
                } 
                         
} else {
                    log -text "No Library mapping found. Terminating" -warning
                    }
    
          

ExitScript


####End Script 


