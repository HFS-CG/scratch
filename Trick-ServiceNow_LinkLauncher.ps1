## Link-Launcher: ServiceNow
## Developed by BTYB.
## 
## USAGE
## ======
## 1) Save this file.
## 2) Create Shortcut to the file
##    Shortcut properties:
##      Target:
##        C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File C:\Launch_SN_link.ps1
##      Shortcut Key:
##        <Set to what you like or put it in the first 10 spots of your Quick Launch.>
##      Run:
##         Minimized
## 
## 3) Copy a servicenow item to the clipboard, hit the hotkey, should open the item in the default browser.
##
##
## 
## HISTORY
## =======
## 1.0.0  2022-06-08  BTYB  Initial writing
## 
##

do {
#Get contents of the clipboard - force to text format (gets rid of rich text characters)
[string]$CurrentClip = Get-Clipboard -Format Text -TextFormatType Text

#Get rid of leading and/or trailing spaces
$CurrentClip = $CurrentClip.Trim()

#Get the first three characters for comparison to known SN item types below
[string]$FirstThree = ($CurrentClip.SubString(0,3))

#Confirm that the clipboard contents end with 7 digits
$LastSevenNumbers = $CurrentClip -match '\d{7}$'

#write-host "Are the last 7 characters numbers?" $LastSevenNumbers
if (!$LastSevenNumbers) {
	Write-Host "Exiting... the clipboard contents don't end with 7 digits, so it isn't a valid SN item number!"
	Exit
}

#Compare the first 3 characters from the clipboard with known SN item prefixes, and if a match, create a variable with the URL to be launched
switch ( $FirstThree )
 {
    "CHG" { $SNType  ="https://capitalgroup.service-now.com/nav_to.do?uri=change_request.do?sysparm_query=number=$CurrentClip"}
	"INC" { $SNType = "https://capitalgroup.service-now.com/nav_to.do?uri=incident.do?sysparm_query=number=$CurrentClip"}
	"SCT" { $SNType = "https://capitalgroup.service-now.com/nav_to.do?uri=sc_task.do?sysparm_query=number=$CurrentClip"}
	"CTA" { $SNType = "https://capitalgroup.service-now.com/nav_to.do?uri=%2Fchange_task.do?sysparm_query=number=$CurrentClip"}
	"RIT" { $SNType = "https://capitalgroup.service-now.com/nav_to.do?uri=%2Fsc_req_item.do?sysparm_query=number=$CurrentClip"}
	"REQ" { $SNType = "https://capitalgroup.service-now.com/nav_to.do?uri=%2Fsc_request.do?sysparm_query=number=$CurrentClip"}
	"PRB" { $SNType = "https://capitalgroup.service-now.com/nav_to.do?uri=%2Fproblem.do?sysparm_query=number=$CurrentClip"}
	default { $SNType = 'NOGOOD' }
 }

#If the first 3 characters don't match a known SN item prefix, then just exit the script (comment out the "Exit" to get a prompt to allow you to paste in manually)
 If ($SNType -eq "NOGOOD")
 {
	Exit
	 Add-Type -AssemblyName Microsoft.VisualBasic
	 $title = "$CurrentClip is not a supported ServiceNow item number"
	 $msg   = 'Enter CHG, INC, CTASK, or SCTASK, RITM, REQ, or PRB item number:'
	 $text = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
	

	if ($text -eq "") {
		$SNType = ""
	}
 }
 # Launch the dynamically created URL path with the default browser
 If ($SNType -ne "") {
	Start $SNType
	$SNType = ""
 }
 Else {
  Write-Host "Exiting..."
 }


} while ($SnType -ne "")

