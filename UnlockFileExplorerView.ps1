##############################################

#UnlockFileExplorerView

#Version           = "0.7.0"

#Copyright:         Free to use, please leave this header intact
#Author:            Marek Obuchowski (http://www.mobuchowski.pl)
#Credits:           EMEA Territory Services Team for all the bits and pieces that allowed this script to be created
#Credits:           Jos Lieben (http://www.lieben.nu http://www.ogd.nl) OneDriveMapper creator for sharing his excellent work
#Purpose:           To automate opening SharePoint Online library in IE, so the user is automatically authenticated and mapped drives unlocked

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
#1 Testing connectivity to SharePoint and DC (for ensuring we are on a corporate network)
#2 Auto-discovery of mapped SharePoint libraries now replaced hardcoded paths

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
#2 BaseURL building code updated - script now extract it from the mapped library, this means that script can be used with any SharePoint tenant now
#3 Typos in comments now corrected.. oh yeah!!

####0.6.1
#hmm... not sure

####0.6.2
#1 Added logging to file
#2 Added extended debug mode
#Startign rewiriting IE handling

####0.6.3
#1 Rewritten IE handling
#   -Checking if Protected View sandboxes IE
#   -Added function to kill IE if sandboxed

####0.6.4
#1 Added support for two mapping formats: UNC and URL
#2 Code cleanup
#   -moved part of the URLs building code to getMappedDrives function, so it sits with within one function

#0.6.5
#1 Fixed the issue with building File explorer URL for mapped drives with library sub-folders

#0.6.6
#1 Extended Debug output to the console and file
#2 Updated Error messages handling
#3 Added code to add tenant to the IE pop-up blocker whitelist
#4 Added support for hiding the console entirely

#0.6.7
#1 Updated network and SharePoint connectivity test and result logging
#2 Removed unnecessary clutter form Error messages

#0.6.8
#1 Fixed typos
#2 Registry path variables defined as global now
#3 Added code to check the presence of the registry keys for setting up popup blocker
#4 Added timeout for waiting for Library window in Windows Explorer
#5 Re-order testing connection and discover drives code, drive first, test later

#0.7.0
#1 Added functions for checking if given process is running
#2 Added code to manage Protected Mode settings - needed for future
#3 Waiting for IE to be closed added with a timeout
#4 Add code to evaluate the output of IE browsing
#5 Add login code if user not logged in automatically
#>


####To-do:
<#
#1 Update drive discovery to check for the best mapped drive
#2 Create function for building URLs and move code to it
#3 Implement Native authenitcation rather than IE
#4 Implement 2fa login in IE mode
#5 Implement graphical status bar or window
####>

param(  
    [Switch]$hideConsole,                                                                                           #Show or hide output in the console based on the script parameter
    [Switch]$debugon                                                                                                #Enable debug via parameter        
    )


$version                     = "0.7.0"

if($debugon){$debugMode      = $True}                                                                               #Enable debug based on parameter
if($hideConsole){$debugMode  = $false}                                                                              #Force DebugMode Off if script run without Consloe output                                                 
#$debugMode                   = $True                                                                                #Use to overvrite parameters .Set to $True if you want the script to ignore current state of the access and go ahead with all actions. Set to $False for normal operation.

#region:Variable definitions
$unlocked                    = $null                                                                                #Variable for holding current state of access to mapped libraries: $True = access already unlocked, $False = access locked
$done                        = $null                                                                                #Variable for holding for the state of unlocking process: $True = unlock process completed, $False = unlock process in progress
$IE                          = $null                                                                                #Variable for storing Internet Explorer object, for closing once unlocking process is completed
$urlOptions                  = "Forms/AllItems.aspx?ExplorerWindowUrl="                                             #Variable for storing part of the URL responsible for opening library in File Explorer
$baseURL                     = $null                                                                                #Variable for storing SharePoint tenant URL
$spoConnection               = $null                                                                                #Variable for storing SharePoint connectivity results
$dcConnection                = $null                                                                                #Variable for storing DC connectivity results
$dcName                      = $env:LOGONSERVER.Substring(2)                                                        #Setting up DC variable for testing connectivity
$lob                         = $null                                                                                #Variable for storing LOB name
$siteName                    = $null                                                                                #Variable for storing Site name
$libraryName                 = $null                                                                                #Variable for storing Library name
$lob2                        = $null                                                                                #Variable for building URLS
$siteName2                   = $null                                                                                #Variable for building URLS
$libraryName2                = $null                                                                                #Variable for building URLS
$subfolders                  = $null                                                                                #Variable for building URLS
$logfile                     = ($env:APPDATA + "\UnlockFileExplorerView_$version.log")                              #Logfile to log to                                                         
$mappingFormat               = $null                                                                                #Variable for storing mapping format
$i_MaxLocalLogSize           = 2                                                                                    #Set the Max local log size in MB
$GPOprotectedMode            = $null                                                                                #Variable for storing info re Protected Mode on or off in IE
$showConsoleOutput           = $True                                                                                #Show console output by default
$IEpopuppath                 = "hkcu:\Software\Microsoft\Internet Explorer\New Windows"                             #Registry path for IE popup blocker configuration
$protectedModeValues         = @{}                                                                                  #Array for storing initial Protected Mode settings
$autoProtectedMode           = $True                                                                                #Automatically temporarily disable IE Protected Mode if it is enabled
$protectedModeOverwrite      = $null
$userIEzones                 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\"           #Location of the user zones
$machineIEzones              = "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\"  #Location of the computer zones
$UPN                         = $null                                                                                #Variable for storing user's UPN needed for IE logon functions
$IE                          = $null                                                                                #Variable for storing IE object
$startTime                   = Get-Date                                                                             #Variables for logging script execution time
$endTime                     = $null                                                                                #Variables for logging script execution time
#endregion

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
    #credits to the OneDriveMapper creator
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

function checkDrive ($URL) {                                               #Function for testing current state of access
    log -text "Testing access state"                                       #Inform user   
    $result = test-path -path $URL                                         #Test if path accessible
    if($DebugMode){log -text ("Testing access state: "+ $URL) -debugg }

   if ($result) {
    log -text "Access enabled"                                             #Inform user
    
    } else {
        log -text "Access locked, unlock process started" -warning         #Inform user
        }
   return $result                                                          #Return $True if access already unlocked, $False if access locked
   }

function Get-ProcessWithOwner{

    param( 
        [parameter(mandatory=$true,position=0)]$ProcessName 
    )
     
    $ComputerName=$env:COMPUTERNAME 
    $UserName=$env:USERNAME 
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($(New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$('ProcessName','UserName','Domain','ComputerName','handle')))) 
    
    try { 
        $Processes = Get-wmiobject -Class Win32_Process -ComputerName $ComputerName -Filter "name LIKE '$ProcessName%'"    #Pull the processes
    } catch { 
        return -1 
    } 
    if ($Processes -ne $null) {
        $OwnedProcesses = @()                                                                                              #Define empty array
        foreach ($Process in $Processes) {                                                                                 #Loop through processes
            if($Process.GetOwner().User -eq $UserName){                                                                    #Check if process owned by the user, add properties if yes
                $Process |  
                Add-Member -MemberType NoteProperty -Name 'Domain' -Value $($Process.getowner().domain) 
                $Process | 
                Add-Member -MemberType NoteProperty -Name 'ComputerName' -Value $ComputerName  
                $Process | 
                Add-Member -MemberType NoteProperty -Name 'UserName' -Value $($Process.GetOwner().User)  
                $Process |  
                Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers                        #Keep our defiend members set only
                $OwnedProcesses += $Process 
            } 
        } 
        return $OwnedProcesses 
    } else { 
        return 0 
    } 
}

function Get-ProcessAll{

    param( 
        [parameter(mandatory=$true,position=0)]$ProcessName 
    )
     
    $ComputerName=$env:COMPUTERNAME 
    
    try { 
        $Processes = Get-wmiobject -Class Win32_Process -ComputerName $ComputerName -Filter "name LIKE '$ProcessName%'"    #Pull the processes
    } catch { 
        return -1 
    }

    if($Processes){return $true} else {return $flase}   
}

#region IE functions

function openIE {                                                                       #Function for opening IE and navigating to View in File Explorer URL         
    param (
        [Switch]$unprotected                                                            #Parametr to switch between default and unprotected mode                                     
    )    
    try {
        
        if(!$unprotected){                                                              #Execute with or without changing user's Protected Mode settings based on parameter
            
            if($debugMode){log -text "No changes to the IE protected mode" -debugg}
            log -text "Opening library in View in File Explorer mode"                   #Inform user
            
            #Start IE instance
            $script:IE = new-object -com internetexplorer.application                   #Create IE object
            $script:IE.visible = $debugMode                                             #Hide IE window if not in debug mode
            
            $script:IE.navigate2("about:blank")                                         #This is added for compatibility with Windows 10 v1709 and protected mode - for some reason, IE fails to open file explorer URL when there is no IE window opened before
            sleep -s 2                                                                  #Give time to open IE                                                                  

            $script:IE.navigate2($URL)                                                  #Navigate to View in File Explorer URL
            sleep -s 2

            if(closeFileExplorerWindow) {return $true}else{return $false}

        }else{
              disableProtectedMode                                                      #disable protected mode temporarily
              
              $script:IE = new-object -com internetexplorer.application                 #Create IE object
              $script:IE.visible = $debugMode                                           #Hide IE window if not in debug mode
                                                                                  
              $script:IE.navigate2("about:blank")                                       #This is added for compatibility with Windows 10 v1709 and protected mode - for some reason, IE fails to open FileExplrorer URL when there is no IE window opened before
              while($script:IE.busy()){sleep -m 100}                                    #Give time to open IE                                                                  

              $script:IE.navigate2($URL)                                                #Navigate to View in File Explorer URL      
              while($script:IE.busy()){sleep -m 100}   
                        
              logonProcess                                                              #Initiate automated logon process

              $test=closeFileExplorerWindow
              if($test) {
                return $true
                }else{
                return $false}
              
        }    

    }catch{
        $ErrorMessage = $_.Exception.Message
        log -text ("Failed to manage Internet Explorer:") -fout
        log -text $ErrorMessage -fout
        if($unprotected){revertProtectedMode}
        closeIE      
        Return $false
    }
}

function closeIE{
    log -text "Closing Internet Explorer processes"
    if($script:protectedModeOverwrite){
        if($debugmode){log -text ("protectedModeOverwrite set to: " + $script:protectedModeOverwrite) -debugg}
        if($debugmode){log -text ("Reverting Protected Mode settings") -debugg}
        revertProtectedMode                                                                                          #Ensure we always revert Protected Mode change
        }else{
            if($debugmode){log -text ("protectedModeOverwrite set to: " + $script:protectedModeOverwrite) -debugg}
            if($debugmode){log -text ("No need to revert Protected Mode settings") -debugg}

        }

    if($script:IE.HWND -ne $null){                                                                                    
        if ($debugMode) {log -text "IE.HWND found - Quitting IE gracefully" -debugg}                                 #Write dubg log
        $script:IE.Parent.Quit()                                                                                     #Quit IE processes if HWDN not found for the IE object 
        $script:IE = $null 
    } else {
        if ($debugMode) {log -text "IE.HWND not found - killing IE" -debugg}                                         #Write dubg log
        killIE                                   
    }
    sleep -s 1
    $test = Get-ProcessAll iexplore                                                                                  #Double check if IE not running, kill if Quit did not work
    if($test){
        if($debugmode){log -text "IE still running - killing the process" -debugg}    
        killIE
        }                                                                                                            #Kill IE if still working
}

function killIE {
    try{
        sleep -s 1
        Stop-Process -Name iexplore -ErrorAction SilentlyContinue                                                    #Kill IE processes
        sleep -m 500
        $IE         = new-object -com internetexplorer.application                                                   #Open blank IE to get rid off the warning message regarding previous session that was killed above
        $IE.visible = $debugMode                                                                                     #Hide IE window if not in debug mode  
        while ($IE.busy) {sleep -m 100}                                                                              #Give time to open IE
        $IE.Quit()                                                                                                   #Quit IE gracefully
        $script:IE = $null
        sleep -Seconds 1
    }catch{
        log -text "Failed to force close IE"
        }
}

function IEpopup {
#Function to test and configure Internet Explroer popup whitelist

try{
    if($debugMode){log -text ("Testing if IE popup blocker conffigured to allow tenant URL: " + $baseURL) -debugg}

        if(Test-Path ($IEpopuppath+"\Allow")){                                                                      #Test if Allow key already exists
            if((TestRegistryValue -path ($IEpopuppath+"\Allow") -value "$baseURL")){                                #Test if Values already exist
                if($debugMode){log -text ("IE popup blocker already configured") -debugg}
                    }else{
                        New-ItemProperty -Path ($IEpopuppath+"\Allow") -Name "$baseURL" -Type Binary                #Create value                                                                 
                        if($debugMode){log -text 'Added popup blocker exception to existing "Allow" key' -debugg}
                    }
        }else{                                                                                                      #If Allow key does not exits
            if($debugMode){log -text '"Allow" key does not exist, creating it now' -debugg}
            New-Item -path $IEpopuppath -Name "Allow" >$null 2>&1                                                   #Create Allow key
            if($debugMode){log -text 'Added "Allow" key' -debugg}            
            New-ItemProperty -Path ($IEpopuppath+"\Allow") -Name "$baseURL" -Type Binary >$null 2>&1                #Create value                                                                
            if($debugMode){log -text 'Added popup blocker exception to the "Allow" key' -debugg}
            }       
    }catch{
    $ErrorMessage = $_.Exception.Message
    log -text ("Failed to manage Internet Explorer Pop-up blocker:") -fout
    log -text $ErrorMessage -fout
    exitScript

    }
}

function logonProcess{

    $timeout = New-TimeSpan -Seconds 30                                   #Set timout
    $sw      = [diagnostics.stopwatch]::StartNew()                        #Start stop watch
    
    log -text "Waiting for O365 logon page. Timeout: $timeout"    
    while (($sw.elapsed -lt $timeout) -and ($script:IE.LocationName -ne "Sign in to your account" -and $script:IE.LocationURL -ne $URL)){   #wait for IE page to load
        write-host -NoNewline "." -ForegroundColor Green
        sleep -s 1
        }
    $sw.stop()
    write-host ""
    if($sw.elapsed -gt $timeout){log -text "Timout reached, analyzing loaded page" -warning}

    #Function to browse the portal pages and login without user interaction
    if($script:IE.LocationName -eq "Sign in to your account"){
      if($debugMode){log -text "IE landed on logon page. Logon process activated" -debugg}  
      
      $divs = $script:IE.Document.body.getElementsByClassName("row tile")                             #Get div containig accounts names
    
        #obtain user logon and abort if not found
        getUPN  
        if(!$UPN){
            log -text "Cannot proceed without O365 logon name (UPN). Terminating script" -fout
            closeIE
            ExitScript
            }

        foreach ($div in $divs) {
            if ($div.innerText -like "*$UPN*"){                                                       #Search for entry containing our UPN 
                log -text "User account found on the logon page"
                if($debugmode){log -text ($div.innerText) -debugg}
                $button = $div.getElementsByClassName("table")[0]                                     #Get the button
                $button.click()                                                                       #Click button
                while ($script:IE.busy) {sleep -m 100}                                                #Wait for IE
                break
            }
        
            if($script:IE.LocationURL -ne $URL){
                log -text "Unknown page opened after logging to the O365" -fout
                if($debugMode){
                    log -text ("IE landed on this Location Name: " + $script:IE.locationname) -debugg
                    log -text ("IE landed on this Location URL: " + $script:IE.locationurl) -debugg
                }
                log -text "Terminating script" -fout

                closeIE
                exitScript
            }else{log -text "View in File Explorer URL loaded`nWaiting for library to open in Windows Explorer"}
        }

    }elseif($script:IE.LocationURL -eq $URL){
        log -text "Logged in to O365 automaticaly"
        }else{
            log -text "IE landed on unexpected page" -fout
            if($debugMode){
                log -text ("Failed to open O365 logon page") -debugg
                log -text ("IE landed on this Location Name: " + $script:IE.locationname) -debugg
                log -text ("IE landed on this Location URL: " + $script:IE.locationurl) -debugg
                closeIE
                ExitScript
            }
        }
}

function getElementById{
    Param(
        [Parameter(Mandatory=$true)]$id
    )
    $localObject = $Null
    try{
        $localObject = $script:ie.document.getElementById($id)
        if($localObject.tagName -eq $Null){Throw "The element $id was not found (1) or had no tagName"}
        return $localObject
    }catch{$localObject = $Null}
    try{
        $localObject = $script:ie.document.IHTMLDocument3_getElementById($id)
        if($localObject.tagName -eq $Null){Throw "The element $id was not found (2) or had no tagName"}
        return $localObject
    }catch{
        Throw
    }
}

function revertProtectedMode(){ 
    if($debugmode){log -text "autoProtectedMode is set to True, reverting to old settings" -debugg}
    try{ 
        for($i=0; $i -lt 4; $i++){ 
            if($protectedModeValues[$i] -ne $Null){ 
                if($DebugMode) {log -text "Setting zone $i back to $($protectedModeValues[$i])" -debugg}
                Set-ItemProperty -Path "$($userIEzones)\$($i)\" -Name "2500"  -Value $protectedModeValues[$i] -Type Dword -ErrorAction SilentlyContinue 
            } 
        }
        $script:protectedModeOverwrite = $false 
        if($debugmode){log -text ("protectedModeOverwrite set to: " + $script:protectedModeOverwrite) -debugg}
    } 
    catch{ 
        if($debugmode){log -text "Failed to modify registry keys to change ProtectedMode back to the original settings: $($Error[0])" -fout}
    } 
} 

function checkProtectedModeGPO(){
#check if any zones are configured with Protected Mode through group policy (which we can't modify) 
    $res= $false
    for($i=0; $i -lt 4; $i++){ 
        $curr = Get-ItemProperty -Path "$($machineIEzones)\$($i)\" -Name "2500" -ErrorAction SilentlyContinue | select -exp 2500 
        if($curr -ne $Null -and $curr -ne 3){ 
            log -text "IE Zone $i protectedmode is enabled through group policy, autoprotectedmode cannot disable it. This will likely cause the script to fail." -fout
            $res = $true
        }
    }

   return $res
}

function disableProtectedMode(){
        if($autoProtectedMode){ 
            if($debugmode) {log -text "autoProtectedMode is set to True, disabling ProtectedMode temporarily" -debugg}
     
            #store old values and change new ones 
            try{ 
                for($i=0; $i -lt 4; $i++){ 
                    $curr = Get-ItemProperty -Path "$($userIEzones)\$($i)\" -Name "2500" -ErrorAction SilentlyContinue| select -exp 2500 
                    if($curr -ne $Null){ 
                        $protectedModeValues[$i] = $curr 
                        if($debugmode){log -text "Zone $i was set to $curr, setting it to 3" -debugg}
                    }else{
                        $protectedModeValues[$i] = 0 
                        if($Debugmode){log -text "Zone $i was not yet set, setting it to 3" -debugg}
                    }
                    Set-ItemProperty -Path "$($userIEzones)\$($i)\" -Name "2500"  -Value "3" -Type Dword -ErrorAction Stop
                }
                $script:protectedModeOverwrite = $True
                if($debugmode){log -text ("protectedModeOverwrite set to: " + $script:protectedModeOverwrite) -debugg}
            } 
            catch{ 
                log -text "Failed to prepare IE (PM) $($error[0])" -fout
            } 
        }
}

#endregion

function closeFileExplorerWindow {                                        #Function for testing if Library already open, closing File Explorer window and changing the state of the unlocking state

    try {
    
    $timeout = New-TimeSpan -Seconds 30                                   #Set timout
    $sw      = [diagnostics.stopwatch]::StartNew()                        #Start stop watch
        
    while ($sw.elapsed -lt $timeout){       
        
        $fs = (New-Object -comObject Shell.Application).Windows()|`       #Create object to store list of File Explorer windows containing SharePoint library
            where-object { ($_.LocationName -like ("*$libraryName*"))`
            -and  ($_.Name -like "*File Explorer*") }
    
        if ($fs) {                                                        #Check if File Explorer window is open.
               if ($debugMode) {
                log -text ("Library found open") -debugg                  
                pause                                                     #Wait for keystroke if in Debug mode
                write-host ""                
                }

               log -text "Closing File Explorer Window"                   #Inform user
               $fs | ForEach-Object {$_.Quit()}
               return $true                                               #Window open: Close and return $true.
        } 
    }

    log -text "Timout: Failed to find open library" -fout
    return $false


    }catch{
    $ErrorMessage = $_.Exception.Message
    log -text ("Failed to manage Library in File Explorer:") -fout
    log -text $ErrorMessage -fout
    return $false
    }
}

function TestRegistryValue {
#Function for testing registry values - returns $True if exists, $Flase if does not exist
param (

 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Path,

[parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Value
)


$c = 0
Get-Item -Path $Path | Select-Object -ExpandProperty property | % { if ($_ -match $Value) { $c=1 ; return $true} }
if ($c -eq 0) {return $false}

}

function testConnection ($a, $type) {                                     #Function to test connection to the SharePoint server
    log -text "Testing connectivity to $type"                             #Inform user
    $test = Test-Connection -computer $a -quiet -count 1                  #Test connection to the SharePoint server
    if ($test) {
        log -text "Connection to $type succesfull"                        #Inform user
        } else {
            log -text "Connection to $type failed" -fout                  #Inform user
            }
    return $test                                                          #Return results
    }

function getMappedDrives {                                                #Function for getting mapped Sharepoint drives and setting up the library for opening in View in File Explorer mode

    #Get-PSDrive method replaced with NET USE
    #Get-PSDrive does not return mapped drives on some machines, Net Use does not have this issue
    #$mp           = Get-PSDrive -PSProvider FileSystem | where-object {$_.DisplayRoot -like "*$baseURL*"} | select-object -first 1    #Get drives and select first entry
    
    #NET USE method
    $nu = net use                                                         #Collect Net Use output

    if ($debugMode){                                                      #Write debug log
        log -text "*****************" -debugg
        log -text ("Net Use output:") -debugg
        log -text ("$nu") -debugg
        log -text "*****************" -debugg

        }          

    foreach ($item in $nu) {                                                                         #Get the first instance (Get-PSDrvive method) or last instance (Net Use method) of the SharePoint Online mapping
            if($item -like "*sharepoint.com*") {
                if ($debugMode){log -text ("NET USE line with *sharepoint.com*: " + $item) -debugg}  #Write debug log
                $item = $item -split ":"                  #Split to get rid off drive letter
                if ($debugMode){log -text ("NET USE line split by colon:" + $item) -debugg}          #Write debug log

                #Extract mapped URL based on the format of the mapping
                switch ($item.length) {
                    2 {
                        $mp = $item[1].trim()
                        $mappingFormat = "UNC"                                                       #Format: \\<tenanturl>@SSL\sites\EMEA-TSG\Shared Documents
                        if($debugMode){log -text("Mapped drive in UNC format: "+ $mp) -debugg}
                        }                                
                    3 {
                        $mp = $item[2].trim()
                        $mappingFOrmat = "URL"                                                       #Format: https://<tenanturl>/sites/EMEAITManagement/Shared Documents
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
        if ($array.length -gt 1) {$array = $array | ? {$_}}      #delete empty elements if more than 1 element exist
    
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

                default {#Library mapped with subfolders
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
    Start-Sleep -s 5         #Give time to read the script
    $endTime = get-date
    if($debugMode){log -text ("Script execution time: " + ($endTime - $startTime)) -debugg}
    log -text "******** End script: $endTime ********"
    exit
    }

function GetUPN{
#Credits: https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-powershell-1.0/ff730963(v=technet.10)
    try{  
        $strName = $env:username                                                   #obtain username
        $strFilter = "(&(objectCategory=User)(samAccountName=$strName))"           #construct LDAP search filter
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher       #create ad object
        $objSearcher.Filter = $strFilter                                           #set the ldap search filter
        $objPath = $objSearcher.FindOne()                                          #find our user
        $objUser = $objPath.GetDirectoryEntry()                                    #bind to our AD user
        $script:UPN = $objUser.userprincipalname
        if($script:UPN){log -text ("O365 logon name (UPN) found: " + $script:UPN)}
        $script:UPN=$objUser.userprincipalname                                     #return UPN
        }catch{
            log -text "Failed to obtain user's logon name for O365 (UPN)" -fout
            return $false
        }
}


#### Start script ####


ResetLog #Reset log :P

log -text "******** Start script: $startTime ********"

if($debugMode){log -text "Debug Mode enabled" -warning}


#Check connections and current access state
if (testConnection $dcName "Corporate network") {                                  #Check connection to the logon server, abort if fails
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


if (testConnection $baseURL "SharePoint Online Servers") {                         #Check connection to the SharePoint Online server, abort if fails
#proceed    
}else{
    log -text ("No connection to the SharePoint Online Servers") -fout
    exitScript
    }



#Check if SharePoint mapping present, exit if not
try {
    if ($driveMapped -eq $null) {                                              
            log -text "No SharePoint mapping found. Terminating" -warning
            exitScript                                                             #Exit script
        }else{

            #Display collected information
            if($mappedURL)    {log -text ("Mapped url: " + $mappedURL)}   else{log -text ("Mapped url: NOT FOUND") -warning}
            if ($baseURL)     {log -text ("Base URL: "   + $baseURL)}     else{log -text ("Base URL: NOT FOUND"  ) -warning}
            if ($lob)         {log -text ("LOB: "        + $lob)}         else{log -text ("LOB: NOT FOUND"       ) -warning}
            if ($siteName)    {log -text ("Site: "       + $siteName)}    else{log -text ("Site: NOT FOUND"      ) -warning}
            if ($libraryName) {log -text ("Library: "    + $libraryName)} else{log -text ("Library: NOT FOUND"   ) -warning}
            
            if ($lob -ne $null) {$lob2 = $lob + "/"}                                                                                   #If LOB contains data, add / for building URL
            if ($siteName -ne $null) {$siteName2 = $siteName + "/"}                                                                    #If Site contains data, add / for building URL
            if ($libraryName -ne $null) {$libraryName2 = ($libraryName -replace " ","%20") + "/"}                                      #If Library contains data, prepare for building URL
            if ($lob -ne $null) {$lob = "%2F" + $lob}
        }
}catch{
    $ErrorMessage = $_.Exception.Message
    log -text ("Unable to build URL components based on fileds extracted from mapped drives:") -fout
    log -text $ErrorMessage -fout
    exitScript
}


#Proceed if library found in the mapped drive
if ($libraryName -ne $null) {                                                      #Stop if no SharePoint or no library found
    
    #Build File Explorer View URL.
    $URL       = "https://" + $baseURL + "/sites/" + $lob2 + $siteName2 + $libraryName2 + $urlOptions + "%2Fsites" + ($lob -replace '-','%2D') + "%2F" + ($siteName -replace " ","%2D") +"%2F" + ($libraryName.TrimEnd('/') -replace " ", "%20")   #Build URL for opening in View in File Explorer mode
    
    if($DebugMode) {log -text ("View in File Explorer URL: " + $URL) -debugg}
    


} else {
                    log -text "No Library mapping found. Terminating" -fout
                    ExitScript
                    }


$unlocked = checkDrive ($mappedURL)                                                #Check the current state of access and store in the $Unlocked variable                          

#Overwirte access test result if in Debug mode    
if ($debugMode) {
    $unlocked = $false                                                             #Set the access state to $False if in DEbug mode. This allows testing entire scritpt
    log -text 'Debug mode: setting up $unlocked to $False for testing' -warning
    }                                   
            
if ($unlocked) {                                                                   #Check if access already enabled, terminate if yes
    log -text "You can now use mapped drives. Terminating"                         #Inform user
    ExitScript
    }

#Ensure no IE processes are running before starting our instance
if(Get-ProcessAll iexplore){                                                                          #Check if IE already running
    $timeout = New-TimeSpan -seconds 120                                                              #Set timout
    log -text "Internet Explorer already open, close all windows Internet Explorer windows" -fout     #Prompt user for closing IE
    log -text "Script will terminate in $timeout otherwise" -fout
    $sw      = [diagnostics.stopwatch]::StartNew()                                                    #Start stop watch

    while((Get-ProcessAll iexplore) -and ($sw.elapsed -lt $timeout)){                                                                            
        write-host -nonewline -foreground red "."
        sleep -Seconds 1
    }
    $sw.stop()                                                                                         #stop the timer

    if($sw.elapsed -gt $timeout){                                                                      #exit script if timed out
        log -text "Timeout reached. Terminating script for user data safety reasons" -fout
        log -text "You need to login to SharePoint and View in File Explorer manually" -fout
        exitscript
    }
}

#Configrue IE popup blocker
IEpopup                                                               

#First attempt with no changes to the protected mode
log -text "First attempt to unlock the drives"
$firsttry = OpenIE                                                    
closeIE

#Overwirte first attempt result if in Debug mode
if($Debugmode){
    $firsttry = $false
    log -text "Simulating first attempt as failed to force fallback into Protected Mode overwrite functions" -debugg
    }

#Check if we can overwrite Protected Mode settings
$gpoProtectedmode = checkProtectedModeGPO                             #Check if IE Protected mode set via GPO
if(!$gpoProtectedMode){                                               #If IE Protected mode not set in GPO
    if($debugMode){log -text "Prtoected mode IS NOT set via GPO. Protected Mode override functions ARE available" -debugg}
}else{
    if($debugMode){log -text "Prtoected mode IS set via GPO. Protected Mode override functions ARE NOT available" -fout
        log -text "Unable to proceed with the second attempt due to the company policy applied to the computer" -fout
        log -text "Terminating script" -fout
        exitScript
        }
    }

#Second attempt with overwriting Protected Mode settings
if(Get-ProcessAll iexplore){                                          #Close it if IE still running after first attempt
    closeIE
    }

if(($firsttry -ne $True) -and $autoProtectedMode -and !$GPOprotectedMode){
    if($debugMode){log -text "Failed to process with current user configruation. Re-trying with protected mode disabled" -warning}
        $secondtry = openIE -unprotected                              #Check if process succeeded with current configuration. Re-try with Protected Mode override
        closeIE
        if(!$secondtry){log -text "Unexpected result of the second logon attempt, re-evaluating to obtain current access state" -warning}
}


#Double check if access unlocked
$unlocked = checkDrive ($mappedURL)                                   #Check the current state of access and store in the $Unlocked variable

if ($unlocked) {                                                      #Check if access already enabled, terminate if yes
    log -text "You can now use mapped drives. Terminating"            #Inform user
    ExitScript
    } else {
        log -text "Script failed to unlock the mapped drives" -fout
        ExitScript
        }

####End Script