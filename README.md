# ResTool NXT
ResTool is a tool used by the Residential Technology Helpdesk at Northern Illinois University. It automates many of the common troubleshooting tasks performed at the helpdesk.
## Getting ResTool
ResTool can be acquired using [ResTool Rebuilder](https://github.com/kmorgan2/ResTool-Rebuilder) ([Download](https://github.com/kmorgan2/ResTool-Rebuilder/releases/download/v1.0/ResTool.Rebuilder.exe)).
ResTool Rebuilder can place a full ResTool installation onto a flash drive (>256MB). ResTool NXT itself can then be launched from the root of that drive by running 'ResToolNXT.exe' or 'ResToolNXT_x64.exe'.
##Using ResTool
After launching ResTool NXT for the first time, you will be asked to enter the ResTech Trouble Ticket number associated with the machine. When the ticket number is recorded, the main GUI will load, which displays all possible options
The possible programs are broken down by function into tab-groups, and further grouped on each tab. When programs should be executed in an order,
said programs will be placed such that they should be performed top-down, that is the first item at the top of the group should be executed first, then the one below that, and so on and so forth.

Clicking on a program or task will start its respective script. Progress if available will be noted using the progress bar at the bottom of the window.
ResTool will by nature be unresponsive during the execution of a script. Since none of the programs in ResTool should be launched concurrently, this is a non-issue.
If for whatever reason the script becomes unresponsive (i.e. it does not react to a change in state of the target program within 60 seconds), ResTool can be force-closed by right-clicking on its icon in the taskbar and selecting 'Exit'.

Tasks that are fully automated should not be interrupted by the user. ResTool will perform all necessary actions, and making any adjustments or otherwise modifying the execution of a program may cause undesired results.
ResTool may be unable to complete execution of a sctipt if the expected state of the target program has been disrupted in any way. If for any reason you need to modify the execution of the program, you must first exit ResTool.

### Programs and Tasks
- Virus Removal
  - Antivirus Scans
    - ComboFix
    - Malwarebytes Anti-Malware (Full Scan)
    - ESET Online Scanner
    - Spybot Search & Destroy
    - SuperAntiSpyware (Full Scan)
    - TrendMicro HouseCall (Full Scan)
  - System Cleanup
    - Piriform CCleaner (File & Registry scans)
    - Open Programs and Features in Control Panel
    - Windows Disk Cleanup
  - Rootkit Removal
    - Malwarebytes Anti-Rootkit
    - Kaspersky Labs TDSSKiller
  - Miscellaneous
    - Install Microsoft Security Essentials or Enable Windows Defender
    - Open the ResTech Trouble Ticket
- OS Troubleshooting
  - Network
    - Windows IP Configuration Release/Renew
    - Winsock & Network Interface Reset
    - Hidden Network Adapter Removal
    - NIUwireless Profile Import
    - NIU DoIT Speed Test
    - Network Adapter Properties
  - System
    - System File Check
    - Deployment Image Servicing and Management Image Health Restoration
    - Tweaking.com Windows All-in-One Repair
    - Disk Defragment
    - Disk Check
    - Open Device Management
  - Miscellaneous
    - Uninstall Cisco NAC Agent
    - Open Control Panel
    - Open RegEdit
    - Toggle Safe Mode With Networking using MSConfig
    - Install Pharos Client for NIU Anywhere Prints
- Removal Tools
  - None Currently Implemented
- Auto
  - Auto Mode Not Implemented
  
## Interpreting Log Data
ResTool's standard log, located in the 'Logs' folder on the ResTool flash drive, contains a log of all procedures, errors, and results captured by the program. Currently, log entries, which are separated by linebreaks, will resemble:  
`YYYY/MM/DD hh:mm:ss Log Message`  
For Example, the logfile entry for failing to connect the computer to the internet before running an ESET scan looks like:  
`2015/01/01 12:34:56 Error occurred while running ESET Online Scanner. The program also generated the following message: No Internet connection. Scan aborted.`

The log is built, as you can see, for human readability. However, it only has five classes of entry to increase consistency:
- Program Start:  `Starting [Program]`
- Error: `Error occurred while running [Program]`
- Results: `[Program] generataed the following results [Message]` 
- Program Completion: `[Program] completed successfully.`
- Other Message: `[Program] generated the following message: [Message]`

Additionally, if a message is provided to a log entry class that normally does not accept one, the following will be appended to the entry:
`The program also generated the following message: [Message]`

