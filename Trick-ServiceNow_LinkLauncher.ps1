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
## 
## 1.0.1  2022-06-08  BTYB  Clipboard contents now converted to HTML
##                          Known Issue: 
##                            once the clipboard is modified, once my script exists,
##                            can't paste into Slack.
##                            Can still paste into Teams, Outlook, notepad, etc.
## 
## 1.0.0  2022-06-08  BTYB  Initial writing
## 
## 
## 
## 
##
# New version with HTML hyperlink of SN item number on clipboard
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
	"CTA" { $SNType = "https://capitalgroup.service-now.com/nav_to.do?uri=change_task.do?sysparm_query=number=$CurrentClip"}
	"RIT" { $SNType = "https://capitalgroup.service-now.com/nav_to.do?uri=sc_req_item.do?sysparm_query=number=$CurrentClip"}
	"REQ" { $SNType = "https://capitalgroup.service-now.com/nav_to.do?uri=sc_request.do?sysparm_query=number=$CurrentClip"}
	"PRB" { $SNType = "https://capitalgroup.service-now.com/nav_to.do?uri=problem.do?sysparm_query=number=$CurrentClip"}
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

$HTMLSNLink = @"
Version:0.9
StartHTML:0000000000
EndHTML:0000000000
StartFragment:0000000000
EndFragment:0000000000
SourceURL:$SNType
<html>
<body>
<!--StartFragment--><a class="linked formlink" aria-label="Open record: $CurrentClip" href="$SNType" style="box-sizing: border-box; text-decoration: underline; padding: 0px; cursor: pointer; color: rgb(46, 46, 46); background: transparent; white-space: normal; outline: 0px; font-family: SourceSansPro, &quot;Helvetica Neue&quot;, Arial; font-size: 13.3333px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px;">$CurrentClip</a><!--EndFragment-->
</body>
</html>
"@

<#
$HTMLSNLinkTest = @"
Version:0.9
StartHTML:0000000000
EndHTML:0000000000
StartFragment:0000000000
EndFragment:0000000000
SourceURL:https://capitalgroup.service-now.com/sc_request.do?sysparm_query=number=REQ0676208
<html>
<body>
<!--StartFragment--><a class="linked formlink" aria-label="Open record: RITM0680585" href="https://capitalgroup.service-now.com/sc_req_item.do?sys_id=ff391f8e1be33854371277741a4bcb9a&amp;sysparm_record_target=sc_req_item&amp;sysparm_record_row=1&amp;sysparm_record_rows=1&amp;sysparm_record_list=request%3D7b391f8e1be33854371277741a4bcb9a%5EORDERBYDESCnumber" style="box-sizing: border-box; text-decoration: underline; padding: 0px; cursor: pointer; color: rgb(46, 46, 46); background: transparent; white-space: normal; outline: 0px; font-family: SourceSansPro, &quot;Helvetica Neue&quot;, Arial; font-size: 13px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px;">RITM0680585</a><!--EndFragment-->
</body>
</html>
"@
#>

	$data = New-Object System.Windows.Forms.DataObject
	$data.SetData([System.Windows.Forms.DataFormats]::Html, $HTMLSNLink)
	$data.SetData([System.Windows.Forms.DataFormats]::Text, $CurrentClip)
	[System.Windows.Forms.Clipboard]::SetDataObject($data)
	Start $SNType
	Start-Sleep -Seconds 30
	$SNType = ""

 }
 Else {
  Write-Host "Exiting..."
 }


} while ($SnType -ne "")
