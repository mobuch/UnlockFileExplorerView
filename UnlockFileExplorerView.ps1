##############################################

#UnlockFileExplorerView

#Version           = "0.7.4"

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
#3 Moved IE popup configuration to function
#4 Added code to check the presence of the registry keys for setting up popup blocker
#5 Added timeout for waiting for Library window in Windows Explorer
#6 Re-order testing connection and discover drives code, drive first, test later

#0.7.0
#1 Functions for checking if IE process is running before we start our own instances added
#2 If IE already open by user - Waiting for IE to be closed timeout added
#3 Code to evaluate the output of IE browsing added
#4 Logging in to the O365 process if user not logged in automatically added
#5 Code to manage IE Protected Mode settings (for O365 login process) added
#6 Mapped drives optimization/conversion from URL to UNC added

#0.7.1
#1 Improved O365 logon process
#2 Scanning registry entry replaced net use and Get-PSDrive as a mapped drives discovery methods
#3 Better logic for extracting URL elements from mapped drives. Now we are able to extract correct Site Collections if Site Collection pattern is provided
#4 Rebuilt logic behind buildig URL from extracted elements. Less variables needed now
#5 Simplified code order for easier flow via different functions

#0.7.2
#1 Testing path: HKCU:\Network before drives optimization to handle those rare cases where a user does not have any drives mapped and the key does not exist
#2 Updates drive optimization; it now converts %20 in URL format to " " in UNC

#0.7.3
#1 Disabled "first unlock" that assumes that the user is logged in to O365 automatically

#0.7.4
#1 Add support for DavWWWRoot keyword in the mapping format
#2 Added extraction of the drive letters
#3 Changed access state test to testign drive lettes first, hten unc paths
#4 Cosmetic changes to the exitScript functions
#5 Add variable for managing timeouts and set it to 60 seconds
#6 Update library hadnling with code that dynamically switches from UNC window to the mapped letter in order to make the unlock process more reliable and consistent
#>


####To-do:
<#
#1 Update drive discovery to check for the best mapped drive
#3 Implement Native authenitcation rather than IE
#4 Implement 2fa login in IE mode
#5 Implement graphical status bar or window
####>


<#
.SYNOPSIS
Unlocks the SharePoint mapped drives.
.DESCRIPTION
Script checks if the user has any SharePoint Online sites or libraries mapped to their computer.
    1.	Initialize logging to the file, for troubleshooting and audit
    2.	Testing corporate network connectivity.
    3.	Search for SharePoint Online mapped drives. Terminate if mappings not found. Optimize URL type mappings into UNC type mappings
    4.	Extract tenant url, site collection, site name, library name, folders and subfolder names
    5.	Build View in File Explorer View URL from previously extracted data
    6.	Test connectivity to the tenant SharePoint url. Terminate if no connection the SharePoint servers
    7.	Test mapped drives access state. Terminate if access already unlocked
    #Disabled 8.	Prompt user to close IE if already running
    9.	Configure IE popup blocker (allow tenant url)
    #Disabled 10.	The first attempt to unlock the drives. First attempt assumes the user is configured for auto-logon.
        a.	Open View in File Explorer View URL in IE
        b.	Wait for the library to open in the file explorer, close when found
        c.	Retest access state
    11.	Second attempt to unlock the drives. Execute only If the first attempt fails. 
        a.	Pull user UPN from AD
        b.	Disable IE Protected Mode temporarily
        c.	Open View in File Explorer View URL in IE
        d.	Wait for page to load. If login page found, search for login button where it equals user UPN, activate the login process if the correct user located on page
        e.	Wait for the View in File Explorer URL to load,
        f.	Wait for the library to open in file explorer, close when found
        g.	Retest access state
    12.	Exit script
.PARAMETER hideConsole  
The parameter is used for hiding the console window from the user
.PARAMETER debugon
The parameter is used for troubleshooting the script execution when activated, all processes are happening in the foreground and visible to the user, and additional pauses are added for the user to confirm script phases
.EXAMPLE
Run the script in normal mode. Console and output are displayed to the user.
.\UnlockFileExplorerView
.EXAMPLE 
Run the script in a hidden mode. Console and output are hidden from the user.
Hideconsole cannot be used in conjunction with -debugon parameter.
.\UnlockFileExplorerView -hideconsole
.EXAMPLE 
Run the script in a hidden mode. Console and output are visible to the user. Additional debug information is displayed and logged to the file. All processes are happening in the foreground and visible to the user.
Debugon cannot be used in conjuction with -hideconsole parameter.
.\UnlockFileExplorerView -debugon
#> 
param(  
    [Switch]$hideConsole,                                                                                           #Show or hide output in the console based on the script parameter
    [Switch]$debugon                                                                                                #Enable debug via parameter        
    )


$version                     = "0.7.4"

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
$siteCollection              = $null                                                                                #Variable for storing site collection name
$siteCollectionPattern       = "SPE-*"                                                                              #Variable for storing site collection pattern, this is needed for distinguishing Site from Site Collection
$siteName                    = $null                                                                                #Variable for storing Site name
$libraryName                 = $null                                                                                #Variable for storing Library name
$subfolders                  = $null                                                                                #Variable for building URLS
$logfile                     = ($env:APPDATA + "\UnlockFileExplorerView_$version.log")                              #Logfile to log to                                                         
$mappingFormat               = $null                                                                                #Variable for storing mapping format
$i_MaxLocalLogSize           = 2                                                                                    #Set the Max local log size in MB
$GPOprotectedMode            = $null                                                                                #Variable for storing info re Protected Mode on or off in IE
$showConsoleOutput           = $True                                                                                #Show console output by default
$IEpopuppath                 = "hkcu:\Software\Microsoft\Internet Explorer\New Windows"                             #Registry path for IE popup blocker configuration
$protectedModeValues         = @{}                                                                                  #Array for storing initial Protected Mode settings
$autoProtectedMode           = $True                                                                                #Automatically temporarily disable IE Protected Mode if it is enabled
$protectedModeOverwrite      = $null                                                                                #Variable for storing Protected Mode Overwriten flag
$userIEzones                 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\"           #Location of the user zones
$machineIEzones              = "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\"  #Location of the computer zones
$UPN                         = $null                                                                                #Variable for storing user's UPN needed for IE logon functions
$IE                          = $null                                                                                #Variable for storing IE object
$startTime                   = Get-Date                                                                             #Variables for logging script execution time
$endTime                     = $null                                                                                #Variables for logging script execution time
$firsttry                    = $False                                                                               #Variable for storing outcome of the first try, if False, secodn try will be actioned
$timeoutSec                  = 60                                                                                   #Timeout value in seconds (for page and library laoding wait time)

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
            $script:IE.visible = $debugMode                                             #Hide IE window if not in debug mode
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
        if ($debugMode) {log -text "IE.HWND found - quitting IE gracefully" -debugg}                                 #Write dubg log
        $script:IE.Parent.Quit()                                                                                     #Quit IE processes if HWDN not found for the IE object 
        $script:IE = $null 
    } else {
        if ($debugMode) {log -text "IE.HWND not found - unable to close IE" -debugg}                                         #Write dubg log
        #Disabled as par tof the 0.7.2
        #killIE                                   
    }
    sleep -s 1
    #Disabled as par tof the 0.7.2
    #$test = Get-ProcessAll iexplore                                                                                  #Double check if IE not running, kill if Quit did not work
    #if($test){
    #    if($debugmode){log -text "IE still running - killing the process" -debugg}    
    #    killIE
    #    }                                                                                                            #Kill IE if still working
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

    $timeout = New-TimeSpan -Seconds $timeoutSec                          #Set timout
    $sw      = [diagnostics.stopwatch]::StartNew()                        #Start stop watch
    
    log -text ("Waiting for O365 logon page. Timeout: $timeout" + " [hh:mm:ss]")    
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
                log -text "User account found on the logon page. Initiaiting logon"
                if($debugmode){log -text $($div.innerText.trim()) -debugg}
                $button = $div.getElementsByClassName("table")[0]                                     #Get the button
                $button.click()                                                                       #Click button 
                break       
            }
        }
    }


    $timeout = New-TimeSpan -Seconds $timeoutSec                                                               #Set timout
    $sw      = [diagnostics.stopwatch]::StartNew()                                                    #Start stop watch
                
    log -text ("Timeout: " + $timeout +" [hh:mm:ss]")
    while (($sw.elapsed -lt $timeout) -and ($script:IE.LocationURL -ne $URL)) {                       #Wait for IE to land on the view in file explorer url
        Write-Host -NoNewline -ForegroundColor Green "."
        sleep -s 1
        } 
    $sw.stop()
    write-host ""
    if($sw.elapsed -gt $timeout){log -text "Timout reached, process may fail if library not found open in the next stage" -warning}


    if($script:IE.LocationURL -ne $URL){
        log -text "Unknown page opened after logging to the O365" -fout
        if($debugMode){
            log -text ("IE landed on this Location Name: " + $script:IE.locationname) -debugg
            log -text ("IE landed on this Location URL: " + $script:IE.locationurl) -debugg
        }
        log -text "Terminating script" -fout
        closeIE
        exitScript
    }elseif($script:IE.LocationURL -eq $URL){
        log -text "Logged in to O365"
        log -text "View in File Explorer URL loaded"
        log -text "Waiting for library to open in Windows Explorer"
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
    
    $timeout = New-TimeSpan -Seconds $timeoutSec                          #Set timout
    $sw      = [diagnostics.stopwatch]::StartNew()                        #Start stop watch
    
    log -text ("Timout: " + $timeout + " [hh:mm:ss]")    
    while ($sw.elapsed -lt $timeout){       
        
        $fs = (New-Object -comObject Shell.Application).Windows()|`       #Create object to store list of File Explorer windows containing SharePoint library
            where-object { ($_.LocationName -like ("*$libraryName*"))`
            -and  ($_.Name -like "*File Explorer*") }

    
        if ($fs) {                                                        #Check if File Explorer window is open.
               #region:Disabled in 0.7.3
               
               #if ($debugMode) {
               # log -text ("Library found open. Press enter to continue") -debugg                  
               # pause                                                     #Wait for keystroke if in Debug mode
               # write-host ""                
               # }

               #endregion

               $fs | ForEach-Object {$_.Visible = $debugMode}             #Hide library windows
               sleep -s 1
               
               try {
                $fs.Navigate2($($driveletter+":"))
                sleep -s 1
               }catch{}

               log -text "Library found open"                             #Inform user

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

function extractMappedDrivesElements($mp) {                                                #Function for getting mapped Sharepoint drives and setting up the library for opening in View in File Explorer mode

    
    $i = $null
    foreach ($key in $mp.Keys){
            $line              = $mp[$key]
            $driveletter       = $key
            if($line -like "\\*.sharepoint.com@ssl\*"){
                $tempMappedURL = $line
                $line = $line.trimstart('\\').trim().split("\")
                switch ($line.Length)       
                {
                    4 {
                        if(($i -eq $null) -or ($i -lt 4)){
                            $i  = $line.Length
                            $tempBaseURL              = $line[0]
                            if($line[2] -like $siteCollectionPattern){
                                $tempsiteCollection   = $line[2]
                                $tempSite             = $line[3]
                                $tempLib              = $null
                                $tempFolder           = $null
                            }else{
                                $tempsiteCollection   = $null
                                $tempSite             = $line[2]
                                $tempLib              = $line[3]
                                $tempFolder           = $null
                            }
                            $tempMappedURL2       = $tempMappedURL
                            }
                        }
                    5 { 
                        if(($i -eq $null) -or ($i -lt 5)){
                            $i  = $line.Length
                            $tempBaseURL          = $line[0]
                            if($line[2] -like $siteCollectionPattern){
                                $tempsiteCollection   = $line[2]
                                $tempSite             = $line[3]
                                $tempLib              = $line[4]
                                $tempFolder           = $null
                            }else{
                                $tempsiteCollection   = $null
                                $tempSite             = $line[2]
                                $tempLib              = $line[3]
                                $tempFolder           = $line[4]
                            }
                            $tempMappedURL2       = $tempMappedURL
                            }
                        }
                    {$_ -gt 5}{
                            if(($i -eq $null) -or ($i -lt 7)){
                            $i  = $line.Length
                            $tempBaseURL          = $line[0]
                            if($line[2] -like $siteCollectionPattern){
                                $tempsiteCollection   = $line[2]
                                $tempSite             = $line[3]
                                $tempLib              = $line[4]
                                $array    = $line[5..($line.Length-1)]
                                foreach ($element in $array)
                                {
                                    if ($tempFolder -eq $null) {
                                        $tempFolder = $element
                                    }else{
                                        $tempFolder = $tempFolder + "2%F" + $element
                                    }
                                }
                                #best match
                            }else{
                                $tempsiteCollection   = $null
                                $tempSite             = $line[2]
                                $tempLib              = $line[3]
                                $array    = $line[4..($line.Length-1)]
                                foreach ($element in $array)
                                {
                                    if ($tempFolder -eq $null) {
                                        $tempFolder = $element
                                    }else{
                                        $tempFolder = $tempFolder + "2%F" + $element
                                    }
                                }                            }
                            $tempMappedURL2       = $tempMappedURL
                            }
                        }
                    
                    default {
                            log -text ("Unsupported SharePoint mapping: " + $tempMappedURL) -fout
                            log -text "Terminating script" -fout
                            closeIE
                            exitScript
                        }
                }
            }
        }
    
  

    #Return values
    $tempMappedURL2
    $($tempBaseURL -replace("@ssl"))
    $tempsiteCollection
    $tempSite
    $tempLib
    $tempFolder
    $driveletter
}

function UpdateNetworkMappings {

log -text "Optimizing mapped drives for View in File Explorer functionality"

try{
    $mappeddrives = @{}                                                                                       #An array for storing all mapped sharepoitn drives
    if($debugmode){log -text "Trying to convert all URL mappings to UNC mappings" -debugg}
    
      if(Test-Path HKCU:\Network){ 
        Push-Location                                                                                                                             #store current location
        Get-ChildItem HKCU:\Network | ForEach-Object {                                                                                            #list all mappings in HCKU:\Network and loop throuch each object
           $driveLetter = $_.name.substring($_.name.Length -1)                                                                                    #Extract drive letter
           $temploc = $("HKCU:\Network\"+$driveLetter)                                                                                            #build registry location path for each mapping
           set-location $temploc                                                                                                                  #set location to each mapping
           $oldpath = Get-ItemProperty -Path . -Name "RemotePath" | select -ExpandProperty "RemotePath"                                           #store old path

           
            if ($oldpath -like "https:*sharepoint.com*"){                                                                                         #if old path contains https://*sharepoint.com*
                if($debugmode){log -text ("URL mappings found: " + $oldpath) -debugg}
                $newpath = ((($oldpath -replace "https:","") -replace '/','\') -replace '\\sites','@ssl\sites') -replace '%20',' '                #transform it into unc
                if($debugmode){log -text ("Converting it to the UNC mapping: " + $newpath) -debugg}
                Set-ItemProperty -path . -Name "RemotePath" -Value $newpath -type String                                                          #and set the value to new unc path
            }elseif($oldpath -like "http:*sharepoint.com*"){                                                                                      #if old path contains http://*sharepoint.com*  
                if($debugmode){log -text ("URL mappings found: " + $oldpath) -debugg}                
                $newpath = ((($oldpath -replace "http:","") -replace '/','\') -replace '\\sites','@ssl\sites') -replace '%20',' '                 #transform it into unc
                if($debugmode){log -text ("Converting it to the UNC mapping: " + $newpath) -debugg}
                Set-ItemProperty -path . -Name "RemotePath" -Value $newpath -type String                                                          #and set the value to new unc path
            }elseif($oldpath -like "\\*sharepoint.com@ssl\DavWWWRoot\*"){
                if($debugmode){log -text ("UNC DavWWWRoot mappings found: " + $oldpath) -debugg}
                $newpath = ($oldpath -replace '\\DavWWWRoot', "") -replace '%20',' '
                if($debugmode){log -text ("Converting it to the UNC mapping: " + $newpath) -debugg}
                Set-ItemProperty -path . -Name "RemotePath" -Value $newpath -type String
            }elseif($oldpath -like "\\*sharepoint.com\DavWWWRoot\*"){
                if($debugmode){log -text ("UNC DavWWWRoot mappings found: " + $oldpath) -debugg}
                $newpath = ($oldpath -replace '\\DavWWWRoot', "@ssl") -replace '%20',' '
                if($debugmode){log -text ("Converting it to the UNC mapping: " + $newpath) -debugg}
                Set-ItemProperty -path . -Name "RemotePath" -Value $newpath -type String
            }elseif ($oldpath -like "\\*sharepoint.com@ssl*"){
                if($debugmode){log -text ("UNC mapping found, no need to optimize it: " + $oldpath) -debugg}
                $newpath = $oldpath
                }
            if($newpath){$mappeddrives.add($driveLetter, $newpath)}
        }
       
        Pop-Location
        log -text "Mapped drives optimized"
        if ($mappeddrives.Count -ne 0){
            return $mappeddrives
        }else{
            return $false}
    }else{
        log -text "No mapped drives found in the system, terminating"
        exitScript
    }

    }catch{
     if($debugmode){log -text "Errors while converting URL mappings into the UNC mappins" -fout}
    }
}

function ExitScript{
    $endTime = get-date
    if($debugMode){log -text ("Script execution time: " + ($endTime - $startTime)) -debugg}
    log -text "******** End script: $endTime ********"
    Start-Sleep -s 5         #Give time to read the script
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

#find if SharePoint Online mappings exist and ensure all are in the UNC format
$mp = UpdateNetworkMappings

if($mp -ne $False){

    #Extract URL elements from the Mapped Drives Elements
    try {
        $return        = extractMappedDrivesElements $mp
        }
    catch {
        $ErrorMessage = $_.Exception.Message
        log -text "Unable to obtain fileds from the mapped drives:" -fout 
        log -text $ErrorMessage -fout
        exitScript
    }

    $mappedURL      = $return[0]
    $baseURL        = $return[1]
    $siteCollection = $return[2]
    $siteName       = $return[3]
    $libraryName    = $return[4]
    $subfolders     = $Return[5]
    $driveletter    = $Return[6]
}else {
        log -text "No SharePoint mapping found. Terminating" -warning
        exitScript
        }

#Check if SharePoint mapping present, exit if not
try {
    #Display collected information
    if($mappedURL)       {log -text ("Mapped URL: " + $mappedURL)}                    else{log -text ("Mapped url: NOT FOUND")      -warning}
    if ($baseURL)        {log -text ("Base URL: "   + $baseURL)}                      else{log -text ("Base URL: NOT FOUND"  )      -warning}
    if ($siteCollection) {log -text ("Site Collection: "        + $siteCollection)}   else{log -text ("Site Collection: NOT FOUND") -warning}
    if ($siteName)       {log -text ("Site: "       + $siteName)}                     else{log -text ("Site: NOT FOUND"      )      -warning}
    if ($libraryName)    {log -text ("Library: "    + $libraryName)}                  else{log -text ("Library: NOT FOUND"   )      -warning}


}catch{
    $ErrorMessage = $_.Exception.Message
    log -text ("Unable to build URL components based on fileds extracted from mapped drives:") -fout
    log -text $ErrorMessage -fout
    exitScript
}

#Check connections and current access state
if (testConnection $dcName "Corporate network") {                                  #Check connection to the logon server, abort if fails
#proceed
}else{
    log -text ("Not connected to the Corporate network") -fout
    exitScript
    }  


#Check connection to the SharePoint Online server, abort if fails
if (testConnection ($baseURL) "SharePoint Online Servers") {
#proceed    
}else{
    log -text ("No connection to the SharePoint Online Servers") -fout
    if($debugon){log -text ("SharePoint URL: " + $baseURL) -debugg}
    exitScript
    }


#Build File Explorer View URL 
 $URL       = $("https://" + $baseURL + "/sites/")`
     + $(if($siteCollection) {$siteCollection + "/"})`
     + $(if($siteName)       {$siteName + "/"})`
     + $(if($libraryName)    {$libraryName + "/"})`
     + $($urlOptions + "%2Fsites")`
     + $(if($siteCollection) {"%2F" + ($siteCollection -replace '-','%2D')})`
     + $(if($siteName)       {"%2F" + ($siteName       -replace '-','%2D')})`
     + $(if($libraryName)    {"%2F" + ($libraryName    -replace '-','%2D')})`
     -replace" ","%20"   #Build URL for opening in View in File Explorer mode

    
if($DebugMode) {log -text ("View in File Explorer URL: " + $URL) -debugg}
    

if($driveletter) {$unlocked = checkDrive $($driveletter + ":")}else{$unlocked = checkDrive $mappedURL}                                        #Check the current state of access and store in the $Unlocked variable                          


#Overwirte access test result if in Debug mode    
if ($debugMode) {
    $unlocked = $false                                                             #Set the access state to $False if in DEbug mode. This allows testing entire scritpt
    log -text 'Debug mode: setting up $unlocked to $False for testing' -warning
    }                                   
            
if ($unlocked) {                                                                   #Check if access already enabled, terminate if yes
    log -text "You can now use mapped drives. Terminating"                         #Inform user
    ExitScript
    }


#region:IE processes checks disabled in 0.7.3
<#
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
    Write-Host ""

    if($sw.elapsed -gt $timeout){                                                                      #exit script if timed out
        log -text "Timeout reached. Terminating script for user data safety reasons" -fout
        log -text "You need to login to SharePoint and View in File Explorer manually" -fout
        exitscript
    }
}
#>
#endregion

#Configrue IE popup blocker
IEpopup                                                               

#region:First Attempt disabled in 0.7.3
<#


#First attempt with no changes to the protected mode
log -text "First attempt to unlock the drives"
$firsttry = OpenIE                                                    
closeIE

#Overwirte first attempt result if in Debug mode
if($Debugmode){
    $firsttry = $false
    log -text "Simulating first attempt as failed to force fallback into Protected Mode overwrite functions" -debugg
    }

#Second attempt with overwriting Protected Mode settings
if(Get-ProcessAll iexplore){                                          #Close it if IE still running after first attempt
    closeIE
    }

#>
#endregion

#Check if we can overwrite Protected Mode settings
$gpoProtectedmode = checkProtectedModeGPO                             #Check if IE Protected mode set via GPO
if(!$gpoProtectedMode){                                               #If IE Protected mode not set in GPO
    if($debugMode){log -text "Prtoected mode IS NOT set via GPO. Protected Mode override functions ARE available" -debugg}
}else{
    if($debugMode){
        log -text "Prtoected mode IS set via GPO. Protected Mode override functions ARE NOT available" -fout
        log -text "Unable to proceed with the second attempt due to the company policy applied to the computer" -fout
        log -text "Terminating script" -fout
        exitScript
        }
    }

if(($firsttry -ne $True) -and $autoProtectedMode -and !$GPOprotectedMode){
    if($debugMode){log -text "SharePoint drives unlocking process started" -warning}
        $secondtry = openIE -unprotected                              #Check if process succeeded with current configuration. Re-try with Protected Mode override
        closeIE
        if(!$secondtry){log -text "Unexpected result of the unlock attempt, re-evaluating to obtain current access state" -warning}
}


#Double check if access unlocked
if($driveletter) {$unlocked = checkDrive $($driveletter + ":")}else{$unlocked = checkDrive $mappedURL}                                   #Check the current state of access and store in the $Unlocked variable

if ($unlocked) {                                                      #Check if access already enabled, terminate if yes
    log -text "You can now use mapped drives. Terminating"            #Inform user
    } else {
        log -text "Script failed to unlock the mapped drives" -fout
        }


ExitScript

####End Script