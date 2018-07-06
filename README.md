# UnlockFileExplorerView

## Usage:

1.	Download zip archive
2.	Unpack zip archive
3.	Launch one of the 3 batch files:
	* Launch UnlockFileExplorerView0.6.6 in Normal mode.bat
	* Launch UnlockFileExplorerView0.6.6 in Debug mode.bat
	* Launch UnlockFileExplorerView0.6.6 in Hidden mode.bat


## The script can be launched in 3 modes:
	1.	Normal = regular output to the console window and the log file
		This is a mode for the case where we want to show the user current state of the unlocking process
	2.	Debug   = extended output to the console window and the log file
		This mode is for troubleshooting issues.
		It will make all windows and actions visible to the user
		It will execute entire script even if the user has the mapped drive already unlocked
		The script will pause and wait for keystroke when View in File Explorer URL  is open in IE and Library accessible in Windows Explorer. This to allow to see and troubleshoot the URLs and windows’ states.
	3.	Hidden = hidden console, no output to the user
		This is a mode for the case where we want to hide the console output from the user but keep logging information to the log file

## Log file location:
	%appdata%\UnlockFileExplorerView_<version>.log	
