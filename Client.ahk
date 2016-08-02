#SingleInstance force
#Persistent
#NoEnv
#Include Notify.ahk

mushVersion := "3.0.6"

;Tray menu for program
Menu, Tray, Nostandard
;~ Menu, Tray, Add, IT Support, ITSupport
;~ Menu, Tray, Disable, IT Support
Menu, Tray, Add, Help Desk, HelpDesk
Menu, Tray, Add, Email IT Support, EmailIT
Menu, Tray, Add, Computer Information, CompInfo
Menu, Tray, Default, Computer Information
Menu, Tray, Add, Show Alerts, ResetAlerts

StringMid, NFDPC, A_ComputerName, 1, 4
if (NFDPC = "NFDM") {
	Menu, Tray, Add
	Menu, Tray, Add, Launch VPN, StartVPN
	Menu, Tray, Add
}

Menu, Tray, Add, About, CompAbout
;~ Menu, Tray, Add, Close, CloseApp
Menu, Tray, Tip, MUSH - %A_ComputerName%

SeenAnnouncements := 
;Start the notify timer
Goto, BeginAlerts

ping(host) {
    colPings := ComObjGet( "winmgmts:" ).ExecQuery("Select * From Win32_PingStatus where Address = '" host "'")._NewEnum

    While colPings[objStatus]
        Return ((oS:=(objStatus.StatusCode="" or objStatus.StatusCode<>0)) ? "0" : "" ) ( oS ? "" : "1" )
}

MUSH_DisplayAlerts() {
	if (ping("notify.server.tld") = 1) {
		IniRead, ShowAlert, %A_WorkingDir%\settings.ini, ClientOptions, ShowAlert
		IniRead, AlertAddress, %A_WorkingDir%\settings.ini, ClientOptions, AlertAddress
		if (ShowAlert = "1") {
			global SeenAnnouncements
			whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
			whr.Open("GET", AlertAddress , true)
			;Uncomment and set below for AD authentication if required by webserver
			; whr.SetCredentials("USERNAME@DOMAIN.COM", "PASSWORD", 0)
			whr.Send()
			; Using 'true' above and the call below allows the script to remain responsive.
			; Try and Catch added so if webserver is pingable but the web server isn't full up yet it won't throw an error
			try whr.WaitForResponse()
			catch
				return 0
			Announcements := whr.ResponseText
			Loop, parse, Announcements, ¿
			{
				StringSplit, AlertArray, A_LoopField, |
				If(InStr(AlertArray4,A_ComputerName) or InStr(AlertArray4,"All"))
				{
					TempSeenAlert := AlertArray1 . AlertArray2
					If(InStr(SeenAnnouncements,TempSeenAlert))
					{
						
					}
					else {
						SeenAnnouncements .= TempSeenAlert
						StringReplace, AlertArray2, AlertArray2, `;, `r`n, All
						Notify(AlertArray1,AlertArray2,0,AlertArray3)
					}
				}
			}
		}
	}
}

;Gui for double click on tray icon or Windows key plus I

ShowCompInfo:
;Set reg view to 64 bit
SetRegView, 64
;Get Asset ID
RegRead, Asset_Tag, HKEY_LOCAL_MACHINE, SOFTWARE\CompanyName, Asset_Tag
;Get Last Boot Time
objWMIService := ComObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
colItems := objWMIService.ExecQuery("Select * from Win32_OperatingSystem")._NewEnum
while colItems[objItem]
{
	LastBootTime := objItem.LastBootUpTime
}
StringMid, LastBootYear, LastBootTime, 1, 4
StringMid, LastBootMonth, LastBootTime, 5, 2
StringMid, LastBootDay, LastBootTime, 7, 2
StringMid, LastBootHour, LastBootTime, 9, 2
StringMid, LastBootMinute, LastBootTime, 11, 2
StringMid, LastBootSecond, LastBootTime, 13, 2
BootDate = %LastBootYear%-%LastBootMonth%-%LastBootDay% %LastBootHour%:%LastBootMinute%:%LastBootSecond%

;Create a gui that is ontop of everything missing the maximize and minimize buttons
Gui,+AlwaysOnTop -MaximizeBox -MinimizeBox
Gui,Color,FAFAFA
Gui,font,s10 bold c3C3C3C
Gui Add, Text,cFF0000, Help Desk:`t x8005
Gui Add, Text,, Computer:`t %A_ComputerName%
Gui Add, Text,, Username:`t %A_UserName%
Gui Add, Text,, Last Boot:`t %BootDate%
Gui Add, Text,, Asset Tag:`t %Asset_Tag%

if (A_IPAddress1 <> "0.0.0.0") {
	IP1 := A_IPAddress1
} else {
	IP1 := ""
}
if (A_IPAddress2 <> "0.0.0.0") {
	IP2 := A_IPAddress2
} else {
	IP2 := ""
}
if (A_IPAddress3 <> "0.0.0.0") {
	IP3 := A_IPAddress3
} else {
	IP3 := ""
}
if (A_IPAddress4 <> "0.0.0.0") {
	IP4 := A_IPAddress4
} else {
	IP4 := ""
}

Gui Add, Text,, IP Addresses:`t %IP1%`n`t`t %IP2%`n`t`t %IP3%`n`t`t %IP4%

Gui Show,,MUSH - Computer Information
return

;Persistent loop to check for alerts and other procssesing
BeginAlerts:
MUSH_DisplayAlerts()
IniRead, AlertTimer, %A_WorkingDir%\settings.ini, ClientOptions, AlertTimer
SetTimer, DisplayAlerts, %AlertTimer%
return

ResetAlerts:
SetTimer, DisplayAlerts, Off
SeenAnnouncements := 

Sleep 2000
MUSH_DisplayAlerts()
IniRead, AlertTimer, %A_WorkingDir%\settings.ini, ClientOptions, AlertTimer
SetTimer, DisplayAlerts, %AlertTimer%
return

DisplayAlerts:
MUSH_DisplayAlerts()
return

;Hotkey Windows+i to show computer information
#i::
Gui Destroy
Goto, ShowCompInfo
return

;Open the helpdesk in Chrome
HelpDesk:
IniRead, HelpdeskUrl, %A_WorkingDir%\settings.ini, ClientOptions, HelpdeskUrl
Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" %HelpdeskUrl%
return

;Open Outlook to send an email to IT Support
EmailIT:
SetRegView, 64
RegRead, Asset_Tag, HKEY_LOCAL_MACHINE, SOFTWARE\CompanyName, Asset_Tag
MailMessage := "ITSUPPORT@domain.com&subject=Assistance%20Required&body=Computer%3A%20"
MailMessage .= A_ComputerName
MailMessage .= "%0AAsset%20Tag%3A%20"
MailMessage .= Asset_Tag
MailMessage .= "%0AIP%20Address%3A%20"
MailMessage .= A_IPAddress1
MailMessage .= "%0ACurrent%20User%3A%20"
MailMessage .= A_UserName
MailMessage .= "%0A-------------%0A%0APlease%20place%20your%20question%20here."

Run, "OUTLOOK.EXE" /c ipm.note /m %MailMessage%
Return

;Blank label for disabled menu
ITSupport:
return

;Label to show CompInfo
CompInfo:
Gui Destroy
Goto, ShowCompInfo
return

;About Box
CompAbout:
Gui Destroy
Gui,+AlwaysOnTop -MaximizeBox -MinimizeBox
Gui, font,bold
Gui, Add, Text, x60 y21 w280 h20 , MUSH
Gui, font
Gui, Add, Text, x60 y41 w280 h20 , Monitoring and Utilization of Software and Hardware
Gui, Add, Text, x60 y61 w280 h20 , Version:`t`t`t%mushVersion%
Gui, Add, Text, x60 y81 w280 h20 , Northfield Helpdesk:`tx8005
Gui, Add, Text, x60 y101 w280 h20 , Written By:`t`tJoe Williams
Gui, Add, Button, x280 y134 w88 h26 Default, Close
;Gui, Add, Button, x8 y134 w88 h26 gUpdate, Update
Gui, Show, w376 h170, About
return

;Destroy memory of the gui so over writting does not happen
GuiClose:
Gui Destroy
return

#IfWinActive, MUSH About
Esc::Gui Destroy

#IfWinActive, MUSH - Computer Information
Esc::Gui Destroy

ButtonClose:
Gui Destroy
Return

StartVPN:
IniRead, VPNShortcut, %A_WorkingDir%\settings.ini, ClientOptions, VPNShortcut
Run, %VPNShortcut%

Loop {
	if (ping("notify.server.tld") = 1) {
		Notify("Network Drives","Your network drives`nare being`nmapped currently.",10,"GC=FFFFFF GR=3 TC=Black MC=Black BC=000000 BW=4 BF=600 SC=600 SI=600 Image=" 276)
		IniRead, LogonScript, %A_WorkingDir%\settings.ini, ClientOptions, LogonScript
		RunWait  %comspec% /c %LogonScript%,, Hide
		break
	}
	Sleep, 5000
}

return

CloseApp:
ExitApp
return
