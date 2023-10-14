On Error Resume Next
Randomize

Set oWSH = CreateObject("WScript.Shell")
Set oAPP = CreateObject("Shell.Application")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")
Set oWEB = CreateObject("MSXML2.ServerXMLHTTP")
strUser = CreateObject("WScript.Network").UserName

currentVersionST = "3.5.0 "
versionNameST = " DFX WinTweaks 3.5.0 "
currentFolder  = oFSO.GetParentFolderName(WScript.ScriptFullName)

Call checkNTonStart()
Call forceConsole()
Call runElevated()
versionDate = " September 24, 2023 "
currentVersionST = "3.5.0 "
versionNameST = " DFX WinTweaks 3.5.0 "
textf "Before using DFX WinTweaks, you should do a Restore Point manually."
textf "Please wait..."
wait 5
Call startMenu()

Function startMenu()
	cls
	textf " "
	textf "   ____  _______  __ __        ___     _____                    _        "
	textf "  |  _ \|  ___\ \/ / \ \      / (_)_ _|_   _|_      _____  __ _| | _____ " & currentVersionST
	textf "  | | | | |_   \  /   \ \ /\ / /| | '_ \| | \ \ /\ / / _ \/ _` | |/ / __|"
	textf "  | |_| |  _|  /  \    \ V  V / | | | | | |  \ V  V /  __/ (_| |   <\__ \"
	textf "  |____/|_|   /_/\_\    \_/\_/  |_|_| |_|_|   \_/\_/ \___|\__,_|_|\_\___/"
      	textf "     Created by ivandfx							"
	textf " "
	textf "  Welcome, " & strUser
	textf "  Select the option you want to use:"
	textf " "
	textf "  1 = Tweak Settings                      						66 = My Windows version"
	textf " "
	textf "  2 = Experimental Tweaks"
	textf " "
	textf "  3 = Safe Mode Settings"
	textf " "
	textf "  4 = Quick Settings"
	textf " "
	textf "  5 = WAST: Shutdown settings"
	textf " "
	textf "  6 = Presets Beta"
	textf " "
	textf "  88 = Check for updates (Online)					      If you find any issues, type '55'"
	textf "  99 = Open DFX WinTweaks GitHub						(This will open GitHub on your browser)"
	textf "  44 = Open DFX WinTweaks Web"
	textf " "
	textf "  0 = Close								        		10 = Credits"
	textf " "
	textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
		wait 1
		Call startMenu()
		Exit Function
	End If
	Select Case RP
		Case 1	
			Call dfxmain()	
		Case 88
			cls
			Call updatedl()
		Case 99
			cls
			Call dfxgithub()
			wait 1
			Call startMenu
		Case 55
			cls
			Call reportIssue()
			wait 1
			Call startMenu
		Case 2
			cls
			Call expTweaks()
		Case 3
			cls
			Call safemodesettings()
		Case 4
			cls
			Call quicksettings()
		Case 5
			Call WAST()
		Case 6
			result = MsgBox ("Please keep in mind that Presets is on a Beta state and some things may not work properly. If you find something weird, submit an Issue on GitHub", vbOkOnly, "Presets Beta")
			Call presetsMenu()
		Case 44
			cls
			wait 1
			Call dfxtweakerweb()
		Case 66
			oWSH.Run "winver.exe"
			Call startMenu()
		Case 10
			Call dfxCredits()
		Case 0
			Call tweakerexit()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 1
			Call startMenu()
			Exit Function
	End Select
End Function

Function dfxtweakerweb()
		Dim url
		Set url= CreateObject("WScript.Shell")
		url.Run "https://ivandfx.github.io/DFXWinTweaks", 9
		Call startMenu()
	Exit Function
End Function

Function dfxgithub()
		Dim url
		Set url= CreateObject("WScript.Shell")
		url.Run "https://github.com/ivandfx/dfxwintweaks", 9
		Call startMenu()
	Exit Function
End Function

Function dfxrelease()
		Dim url
		Set url= CreateObject("WScript.Shell")
		url.Run "https://github.com/ivandfx/dfxwintweaks/releases", 9
		Call startMenu()
	Exit Function
End Function

Function reportIssue()
		Dim url
		Set url= CreateObject("WScript.Shell")
		url.Run "https://github.com/ivandfx/dfxwintweaks/issues/new", 9
		Call startMenu()
	Exit Function
End Function

Function safemodesettings()
	cls
	textf " "
	textf "   ____         __        __  __           _        ____       _   _   _                 "
	textf "  / ___|  __ _ / _| ___  |  \/  | ___   __| | ___  / ___|  ___| |_| |_(_)_ __   __ _ ___ "
	textf "  \___ \ / _` | |_ / _ \ | |\/| |/ _ \ / _` |/ _ \ \___ \ / _ \ __| __| | '_ \ / _` / __|"
	textf "   ___) | (_| |  _|  __/ | |  | | (_) | (_| |  __/  ___) |  __/ |_| |_| | | | | (_| \__ \"
	textf "  |____/ \__,_|_|  \___| |_|  |_|\___/ \__,_|\___| |____/ \___|\__|\__|_|_| |_|\__, |___/"
	textf "                                                                              |___/      "
	textf " "
	textf " "
	textf "  Select an option:"
	textf " "
	textf " "
	textf "  1 = Restart in Safe Mode (Normal)"
	textf " "
	textf "  2 = Restart in Safe Mode (Networking)"
	textf " "
	textf "  3 = Reboot to Standard Windows"
	textf " "
	textf " "
	textf "  0 = Return to Start Menu"
	textf " "
	textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
		wait 2
		Call safemodesettings()
		Exit Function
	End If
	Select Case RP
	Case 1	
		MsgBox "Your PC will reboot right after you close this window, make sure you saved all your data", vbInformation + vbOkOnly, "DFX WinTweaks Safe Mode"
		Set objShell = WScript.CreateObject("WScript.Shell")
		oWSH.Run "bcdedit /set {current} safeboot minimal"
		wait 1
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
	Case 2
		MsgBox "Your PC will reboot right after you close this window, make sure you saved all your data", vbInformation + vbOkOnly, "DFX WinTweaks Safe Mode"
		Set objShell = WScript.CreateObject("WScript.Shell")
		oWSH.Run "bcdedit /set {current} safeboot network"
		wait 1
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
	Case 3
		MsgBox "Your PC will reboot right after you close this window, make sure you did all your changes", vbInformation + vbOkOnly, "DFX WinTweaks Safe Mode"
		Set objShell = WScript.CreateObject("WScript.Shell")
		oWSH.Run "bcdedit /deletevalue {current} safeboot"
		wait 1
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
	Case 0
		cls
		Call startMenu()
		Exit Function
	End Select
End Function

Function quicksettings()
	cls
	textf " "
	textf "    ___        _      _               _   _   _                 "  
	textf "   / _ \ _   _(_) ___| | __  ___  ___| |_| |_(_)_ __   __ _ ___ "  
	textf "  | | | | | | | |/ __| |/ / / __|/ _ \ __| __| | '_ \ / _` / __|"  
	textf "  | |_| | |_| | | (__|   <  \__ \  __/ |_| |_| | | | | (_| \__ \"  
	textf "   \__\_\\__,_|_|\___|_|\_\ |___/\___|\__|\__|_|_| |_|\__, |___/"  
	textf "                                                      |___/     "  
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf "  1 = Disable Windows Update"
	textf " "
	textf "  2 = Disable Windows Defender (Safe Mode)"
	textf " "
	textf "  3 = Show file extensions"
	textf " "
	textf "  4 = Show Windows license status"
	textf " "
	textf "  5 = Open Additional Windows features"
	textf " "
	textf " "
	textf "  0 = Return to Start Menu"
	textf " "
	textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
		wait 2
		Call quicksettings()
		Exit Function
	End If
	Select Case RP
	Case 1
		oWSH.Run "sc stop wuauserv"
		oWSH.Run "sc config wuauserv start=disabled"
	cls
	textf ""
	textf "  Windows Update is now disabled"
	wait 1
		Call quicksettings()
	Case 2
		oWSH.Run "sc stop WdNisSvc"
		oWSH.Run "sc stop WinDefend"
		oWSH.Run "sc config WdNisSvc start=disabled"
		oWSH.Run "sc config WinDefend start=disabled"	
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cache Maintenance" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cleanup" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Scheduled Scan" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Verification" & chr(34) & " /DISABLE"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableBehaviorMonitoring", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableOnAccessProtection", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableScanOnRealtimeEnable", 1, "REG_DWORD"
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\NOC_GLOBAL_SETTING_TOASTS_ENABLED", 0, "REG_DWORD"
		textf " "
		textf " INFO: Windows Defender has been disabled"
		wait 1
		Call quicksettings()
	Case 3
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", 0, "REG_DWORD"
		textf ""
		textf "Files extensions will now be shown"
		wait 1
		Call quicksettings()
	Case 4
		textf " Your license status will appear in a few seconds..."
		wait 0.2
		textf " Collecting license data..."
		wait 0.5
		oWSH.Run "slmgr.vbs /dli"
		oWSH.Run "slmgr.vbs /xpr"
		Call quicksettings()
	Case 5
		oWSH.Run "optionalfeatures.exe"
		Call quicksettings()
	Case 0
		cls
		Call startMenu()
		Exit Function
	End Select
End Function

Function dfxmain()
	cls
	textf "  Wait..."
	wait 0.1
	Call mainMenu()
	Exit Function
End Function

Function presetsMenu()
	cls
	textf " "
	textf "   ____                     _       "
	textf "  |  _ \ _ __ ___  ___  ___| |_ ___  BETA"
	textf "  | |_) | '__/ _ \/ __|/ _ \ __/ __|"
	textf "  |  __/| | |  __/\__ \  __/ |_\__ \"
	textf "  |_|   |_|  \___||___/\___|\__|___/"
      	textf " "
	textf "  You don't know what a Preset does? Just select it to see if it works for you."
	textf " "
	textf " "
	textf "  1 = Basic Tweaking"
	textf " "
	textf "  2 = Standard Tweaking"
	textf " "
	textf "  3 = Advanced Tweaking"
	textf " "
	textf " "
	textf " "
	textf "  10 = Community Feedback Presets"
	textf " "
	textf "  0 = Return to Start Menu"
	textf " "
	textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
		wait 1
		Call startMenu()
		Exit Function
	End If
	Select Case RP
		Case 1
result = MsgBox ("This Preset disables Windows Update, shows file extensions on Windows Explorer, disables MS OneDrive and disables MS Cortana. Do you want to apply it?", vbYesNo, "Presets: Basic Tweaking")
Select Case result
    Case vbYes
		oWSH.Run "sc stop wuauserv"
		oWSH.Run "sc config wuauserv start=disabled"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", 0, "REG_DWORD"
		oWSH.Run "taskkill.exe /F /IM OneDrive.exe /T"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableLibrariesDefaultSaveToOneDrive", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableMeteredNetworkFileSync", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\Onedrive\DisableLibrariesDefaultSaveToOneDrive", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\Onedrive\DisableMeteredNetworkFileSync", 1, "REG_DWORD"
		oWSH.RegWrite "HKCR\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegWrite "HKCR\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\Software\Classes\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegDelete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\OneDrive"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\AllowCortana", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\CortanaEnabled", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\SearchboxTaskbarMode", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\BingSearchEnabled", 0, "REG_DWORD"
		textf ""
		textf " >> Restarting Windows Explorer..."
		oWSH.Run "taskkill.exe /F /IM explorer.exe"
		wait 5
		oWSH.Run "explorer.exe"
    Case vbNo
		Call presetsMenu()
	End Select
		Case 2
result = MsgBox ("This Preset disables Windows Update, shows file extensions on Windows Explorer, disables MS OneDrive, disables MS Cortana, enables Dark Mode, disables Tracking, disables MS Defender (if this doesn't work you'll have to disable it manually) Do you want to apply it?", vbYesNo, "Presets: Standard Tweaking")
Select Case result
    Case vbYes
		oWSH.Run "sc stop wuauserv"
		oWSH.Run "sc config wuauserv start=disabled"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", 0, "REG_DWORD"
		oWSH.Run "taskkill.exe /F /IM OneDrive.exe /T"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableLibrariesDefaultSaveToOneDrive", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableMeteredNetworkFileSync", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\Onedrive\DisableLibrariesDefaultSaveToOneDrive", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\Onedrive\DisableMeteredNetworkFileSync", 1, "REG_DWORD"
		oWSH.RegWrite "HKCR\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegWrite "HKCR\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\Software\Classes\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegDelete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\OneDrive"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\AllowCortana", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\CortanaEnabled", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\SearchboxTaskbarMode", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\BingSearchEnabled", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
		oWSH.Run "sc stop DiagTrack"
		oWSH.Run "sc config DiagTrack start= disabled"
		oWSH.Run "sc stop dmwappushservice"
		oWSH.Run "sc config dmwappushservice start= disabled"
		oWSH.Run "sc stop WdNisSvc"
		oWSH.Run "sc stop WinDefend"
		oWSH.Run "sc config WdNisSvc start=disabled"
		oWSH.Run "sc config WinDefend start=disabled"	
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cache Maintenance" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cleanup" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Scheduled Scan" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Verification" & chr(34) & " /DISABLE"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableBehaviorMonitoring", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableOnAccessProtection", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableScanOnRealtimeEnable", 1, "REG_DWORD"
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\NOC_GLOBAL_SETTING_TOASTS_ENABLED", 0, "REG_DWORD"
		textf ""
		textf " >> Restarting Windows Explorer..."
		oWSH.Run "taskkill.exe /F /IM explorer.exe"
		wait 5
		oWSH.Run "explorer.exe"
    Case vbNo
		Call presetsMenu()
	End Select
		Case 3
result = MsgBox ("This Preset disables Windows Update, shows file extensions on Windows Explorer, disables MS OneDrive, disables MS Cortana, enables Dark Mode, disables Tracking, disables MS Defender (if this doesn't work you'll have to disable it manually), enables all System Bandwith, disables BitLocker, Encryption and OfflineFiles and disables CPU Core Parking (this will require a reboot) Do you want to apply it?", vbYesNo, "Presets: Advanced Tweaking")
Select Case result
    Case vbYes
		oWSH.Run "sc stop wuauserv"
		oWSH.Run "sc config wuauserv start=disabled"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", 0, "REG_DWORD"
		oWSH.Run "taskkill.exe /F /IM OneDrive.exe /T"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableLibrariesDefaultSaveToOneDrive", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableMeteredNetworkFileSync", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\Onedrive\DisableLibrariesDefaultSaveToOneDrive", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\Onedrive\DisableMeteredNetworkFileSync", 1, "REG_DWORD"
		oWSH.RegWrite "HKCR\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegWrite "HKCR\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\Software\Classes\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
		oWSH.RegDelete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\OneDrive"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\AllowCortana", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\CortanaEnabled", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\SearchboxTaskbarMode", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\BingSearchEnabled", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
		oWSH.Run "sc stop DiagTrack"
		oWSH.Run "sc config DiagTrack start= disabled"
		oWSH.Run "sc stop dmwappushservice"
		oWSH.Run "sc config dmwappushservice start= disabled"
		oWSH.Run "sc stop WdNisSvc"
		oWSH.Run "sc stop WinDefend"
		oWSH.Run "sc config WdNisSvc start=disabled"
		oWSH.Run "sc config WinDefend start=disabled"	
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cache Maintenance" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cleanup" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Scheduled Scan" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Verification" & chr(34) & " /DISABLE"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableBehaviorMonitoring", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableOnAccessProtection", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableScanOnRealtimeEnable", 1, "REG_DWORD"
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\NOC_GLOBAL_SETTING_TOASTS_ENABLED", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Psched\Psched", 0, "REG_DWORD"
		oWSH.Run "sc config BDESVC start=disabled"
		oWSH.Run "sc config EFS start=disabled"
		oWSH.Run "sc config CscService start=disabled"
		oWSH.Run "sc stop BDESVC"
		oWSH.Run "sc stop EFS"
		oWSH.Run "sc stop CscService"
		oWSH.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\PowerSettings\54533251-82be-4824-96c1-47b60b740d00\0cc5b647-c1df-4637-891a-dec35c318583\ValueMax", 0, "REG_DWORD"
		textf ""
		textf " >> The system will reboot in 5 seconds..."
		textf " >> The system will reboot in 4 seconds..."
		textf " >> The system will reboot in 3 seconds..."
		textf " >> The system will reboot in 2 seconds..."
		textf " >> The system will reboot in 1 seconds..."
		textf " >> The system will reboot NOW!"
		Set objShell = WScript.CreateObject("WScript.Shell")
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
    Case vbNo
		Call presetsMenu()
	End Select
		Case 10
			MsgBox "Community Feedback Presets are not available at the moment. You can post your ideas on GitHub Issues", vbInformation + vbOkOnly, "DFX WinTweaks: CF Presets"
			Call presetsMenu()
		Case 0
			Call startMenu()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 1
			Call presetsMenu()
			Exit Function
	End Select
End Function

Function mainMenu()
	cls
	textf " "
	textf "   ____  _______  __ __        ___     _____                    _        "
	textf "  |  _ \|  ___\ \/ / \ \      / (_)_ _|_   _|_      _____  __ _| | _____ " & currentVersionST
	textf "  | | | | |_   \  /   \ \ /\ / /| | '_ \| | \ \ /\ / / _ \/ _` | |/ / __|"
	textf "  | |_| |  _|  /  \    \ V  V / | | | | | |  \ V  V /  __/ (_| |   <\__ \"
	textf "  |____/|_|   /_/\_\    \_/\_/  |_|_| |_|_|   \_/\_/ \___|\__,_|_|\_\___/"
      textf "     Created by ivandfx"
	textf " "
	textf " "
	textf "  Select an option:"
	textf " "
	textf "  1 = Customization"
	textf "  2 = Performance Tweaks"
	textf "  3 = UWP Debloater"
	textf "  4 = Uninstall certain UWP apps/services"
	textf ""
	textf ""
	textf "  5 = Tracking"
	textf "  6 = Microsoft OneDrive"
	textf "  7 = Microsoft Cortana"
	textf "  8 = Microsoft Defender (Safe Mode)"
	textf "  9 = Windows Update settings"
	textf ""
	textf ""
	textf "  10 = Show Windows license status"
	textf ""
	textf ""
	textf "  0 = Return to Start Menu							 44 = Safe Mode Settings"
	textf ""
	textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
		wait 1
		Call mainMenu()
		Exit Function
	End If
	Select Case RP
		Case 1
			Call systemTweaks()
		Case 2
			Call perfTweaks()
		Case 3
			Call uwpDebloat()
		Case 4
			Call uwpUninst()
		Case 5
			Call menuTracking()
		Case 6
			Call onedriveConf()
		Case 7
			Call menuCortana()
		Case 8
			Call defenderConf()
		Case 9
			Call wupdateConf()
		Case 10
			Call licenseView()
		Case 44
			Call safemoConf()
		Case 0
			Call startMenu()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 2
			Call mainMenu()
			Exit Function
	End Select
End Function

Function licenseView()
	cls
	On Error Resume Next
	textf ""
	textf " Your license status will appear in a few seconds..."
	wait 0.2
	textf " Collecting license data..."
	wait 0.3
	oWSH.Run "slmgr.vbs /dli"
	oWSH.Run "slmgr.vbs /xpr"
	Call mainMenu
End Function

Function expTweaks()
	cls
	textf " "
	textf "  _____                      _                      _        _   _____                    _        "
	textf " | ____|_  ___ __   ___ _ __(_)_ __ ___   ___ _ __ | |_ __ _| | |_   _|_      _____  __ _| | _____ "
	textf " |  _| \ \/ / '_ \ / _ \ '__| | '_ ` _ \ / _ \ '_ \| __/ _` | |   | | \ \ /\ / / _ \/ _` | |/ / __|"
	textf " | |___ >  <| |_) |  __/ |  | | | | | | |  __/ | | | || (_| | |   | |  \ V  V /  __/ (_| |   <\__ \"
	textf " |_____/_/\_\ .__/ \___|_|  |_|_| |_| |_|\___|_| |_|\__\__,_|_|   |_|   \_/\_/ \___|\__,_|_|\_\___/"
	textf "            |_|                                                                                    "
      	textf " "
	textf " "
	textf " "
	textf "  Select an option:"
	textf " "
	textf "  1 = Disable CPU Core Parking (Reboot required)"
	textf "  2 = Enable CPU Core Parking (Reboot required)"
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf "  0 = Return to Main Menu"
	textf " "
	textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
		wait 1
		Call startMenu()
		Exit Function
	End If
	Select Case RP
		Case 1
		MsgBox "This will improve your system performance, but it may spike up the power consumption.", vbInformation + vbOkOnly, "DFX WinTweaks: CPU Core Park (Enable)"
		oWSH.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\PowerSettings\54533251-82be-4824-96c1-47b60b740d00\0cc5b647-c1df-4637-891a-dec35c318583\ValueMax", 0, "REG_DWORD"
		wait 2
		MsgBox "CPU Core Parking has been disabled and Windows will now reboot. Make sure you saved your data.", vbInformation + vbOkOnly, "DFX WinTweaks: CPU Core Park (Disabled)"
			Set objShell = WScript.CreateObject("WScript.Shell")
			objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
		Case 2
		MsgBox "This will get your system performance back to normal, lowering power consumption.", vbInformation + vbOkOnly, "DFX WinTweaks: CPU Core Park (Enable)"
		oWSH.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\PowerSettings\54533251-82be-4824-96c1-47b60b740d00\0cc5b647-c1df-4637-891a-dec35c318583\ValueMax", 100, "REG_DWORD"
		wait 2
		MsgBox "CPU Core Parking has been enabled and Windows will now reboot. Make sure you saved your data.", vbInformation + vbOkOnly, "DFX WinTweaks: CPU Core Park (Enabled)"
			Set objShell = WScript.CreateObject("WScript.Shell")
			objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
		Case 0
			Call mainMenu()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 1
			Call startMenu()
			Exit Function
	End Select
End Function

Function uwpUninst()
	cls
	textf " "
	textf "   _   ___        ______    _   _       _           _        _ _           "
	textf "  | | | \ \      / /  _ \  | | | |_ __ (_)_ __  ___| |_ __ _| | | ___ _ __ "
	textf "  | | | |\ \ /\ / /| |_) | | | | | '_ \| | '_ \/ __| __/ _` | | |/ _ \ '__|"
	textf "  | |_| | \ V  V / |  __/  | |_| | | | | | | | \__ \ || (_| | | |  __/ |   "
	textf "   \___/   \_/\_/  |_|      \___/|_| |_|_|_| |_|___/\__\__,_|_|_|\___|_|   "
      textf "                   		UWP App/Service uninstaller (WIP)"
	textf " "
	textf " "
	textf "  Select an option:"
	textf " "
	textf "  1 = Remove Widgets (News) app"
	textf ""
	textf ""
	textf ""
	textf ""
	textf ""
	textf ""
	textf ""
	textf ""
	textf ""
	textf ""
	textf ""
	textf ""
	textf ""
	textf ""
	textf ""
	textf "  0 = Return to Main Menu							 44 = Safe Mode Settings"
	textf ""
	textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
		wait 1
		Call mainMenu()
		Exit Function
	End If
	Select Case RP
		Case 1
			oWSH.Run "powershell Get-AppxPackage *WebExperience* | Remove-AppxPackage"
			Call uwpUninst()
		Case 0
			Call mainMenu()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 2
			Call uwpUninst()
			Exit Function
	End Select
End Function

Function showBannerWAST()
	textf "  __        ___    ____ _____   _____           _              _     _          _ "
	textf "  \ \      / / \  / ___|_   _| | ____|_ __ ___ | |__   ___  __| | __| | ___  __| |"
	textf "   \ \ /\ / / _ \ \___ \ | |   |  _| | '_ ` _ \| '_ \ / _ \/ _` |/ _` |/ _ \/ _` |"
	textf "    \ V  V / ___ \ ___) || |   | |___| | | | | | |_) |  __/ (_| | (_| |  __/ (_| |"
	textf "     \_/\_/_/   \_\____/ |_|   |_____|_| |_| |_|_.__/ \___|\__,_|\__,_|\___|\__,_|"
	textf "                          		Windows Advanced Shutdown Tool Embedded 1.4"
	textf " "
	End Function

Function WAST()
	cls
	Call showBannerWAST()
	textf " "
	textf "  Loading WAST for DFX WinTweaks..."
	wait 0.3
	cls
	On Error Resume Next
	Call showBannerWAST()
	textf " "
	textf " "
	textf " "
	textf " "
	textf " "
	textf "  Select an option:                         55 = Restart Windows Explorer"
	textf "                                            66 = About Reboot to UEFI"
	textf ""
	textf "  1 = Shut down the PC                      6 = Disable Hyper-V and reboot (Pro Only)"
	textf " "
	textf "  2 = Restart the PC                        7 = Enable Hyper-V and reboot (Pro Only)"
	textf " "
	textf "  3 = Log off from this user"
	textf " "
	textf "  4 = Go to advanced options"
	textf ""
	textf "  5 = Reboot to UEFI (?)"
	textf " "
	textf " "
	textf "  0 = Return to Start Menu"
	textf ""
	textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
		Call WAST()
		Exit Function
	End If
Select Case RP
		Case 1
		result = MsgBox ("Shut down?", vbYesNo, "WAST Shutdown")
		Select Case result
    		Case vbYes
		Set objShell = WScript.CreateObject("WScript.Shell")
		objShell.Run "%WINDIR%\system32\shutdown.exe -s -t 0"
        	Dim objShell
    		Case vbNo
		cls
		textf = "  Wait..."
		wait 0.2
		Call WAST()
End Select
		Case 2
		result = MsgBox ("Restart?", vbYesNo, "WAST Restart")
		Select Case result
    		Case vbYes
		Set objShell = WScript.CreateObject("WScript.Shell")
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
    		Case vbNo
		cls
		textf = "  Wait..."
		wait 0.2
		Call WAST()
End Select
		Case 3
		result = MsgBox ("Log off? Unsaved data will be lost.", vbYesNo, "WAST Logoff")
		Select Case result
   		Case vbYes
		Set objShell = WScript.CreateObject("WScript.Shell")
		objShell.Run "%WINDIR%\system32\shutdown.exe -l"
    		Case vbNo
		cls
		textf = "  Wait..."
		wait 0.2
		Call WAST()
End Select
		Case 4
		result = MsgBox ("Go to advanced options menu? This will close all active user sessions.", vbYesNo, "WAST Advanced")
		Select Case result
   		Case vbYes
		Set objShell = WScript.CreateObject("WScript.Shell")
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -o -t 0"
		wait 1
		Call WAST()
    		Case vbNo
		cls
		textf = "  Wait..."
		wait 0.2
		Call WAST()
End Select
		Case 5
		result = MsgBox ("Reboot to BIOS/UEFI? Make sure you saved your data.", vbYesNo, "WAST Reboot to UEFI")
		Select Case result
    		Case vbYes
		Set objShell = WScript.CreateObject("WScript.Shell")
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -fw -t 0"
    		Case vbNo
		cls
		textf = "  Wait..."
		wait 0.2
		Call WAST()
End Select
		Case 6
		MsgBox "This will reboot the PC and turn off Hyper-V.", vbInformation + vbOkOnly, "WAST: Disable Hyper-V"
		Set objShell = WScript.CreateObject("WScript.Shell")
		oWSH.Run "bcdedit /set hypervisorlaunchtype off"
		wait 1
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
		Case 7
		MsgBox "This will reboot the PC and turn on Hyper-V.", vbInformation + vbOkOnly, "WAST: Enable Hyper-V"
		Set objShell = WScript.CreateObject("WScript.Shell")
		oWSH.Run "bcdedit /set hypervisorlaunchtype on"
		wait 1
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
		Case 55
		result = MsgBox ("Restart Windows Explorer?", vbYesNo, "WAST Explorer")
		Select Case result
    		Case vbYes
		textf = "  Wait..."
		textf "  Restarting Windows Explorer..."
		oWSH.Run "taskkill.exe /F /IM explorer.exe"
		wait 5
		oWSH.Run "explorer.exe"
		Call WAST()
    		Case vbNo
		cls
		textf = "  Wait..."
		wait 0.2
		Call WAST()
End Select
		Case 66
		result = MsgBox ("Reboot to UEFI only works on a UEFI Windows install and is tested and known to work on Windows 10 1809 and later,, please let me know if it works on an older version of Windows 10.", vbOkOnly, "WAST UEFI (About)")
		Call WAST()
		Case 0
		cls
		textf "  Going back to DFX WinTweaks..."
		wait 0.3
		textf "  Wait..."
		wait 0.2
		Call startMenu()
	End Select
End Function

Function systemTweaks()
	cls
	On Error Resume Next
	textf ""
	textf "    ____          _                  _          _   _             "
	textf "   / ___|   _ ___| |_ ___  _ __ ___ (_)______ _| |_(_) ___  _ __  "
	textf "  | |  | | | / __| __/ _ \| '_ ` _ \| |_  / _` | __| |/ _ \| '_ \ "
	textf "  | |__| |_| \__ \ || (_) | | | | | | |/ / (_| | |_| | (_) | | | |"
	textf "   \____\__,_|___/\__\___/|_| |_| |_|_/___\__,_|\__|_|\___/|_| |_|"                                               
	textf ""
	textf "  Select an option:"
	textf ""
	textf " "
	textf "  1 = Enable Dark mode"
	textf "  2 = Create a 'God Mode' icon on the Desktop"
	textf "  3 = Enable 'Quick Access' on Windows Explorer"
	textf "  4 = Show file extensions" 
	textf "  5 = Enable 'Classic View' on the Control Panel"
	textf "  6 = Enable Classic Volume slider"
	textf "  7 = Enable/Disable User Account Control"
	textf "  8 = Enable/Disable login without password"
	textf "  9 = Use Windows 11 default desktop icon spacing"
	textf " "
	textf ""
	textf "  0 = Back to menu		99 = Restore"
	textf ""
	textl "  > "
	Select Case scanf
		Case 1
			textf ""
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call systemTweaks()
		Case 2
			textf ""
		godFolder = oWSH.SpecialFolders("Desktop") & "\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
		If oFSO.FolderExists(godFolder) = False Then oFSO.CreateFolder(godFolder)
			textf ""
			textf ""
			wait 0.2
			Call systemTweaks()
		Case 3
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\LaunchTo", 1, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call systemTweaks()
		Case 4
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", 0, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call systemTweaks()
		Case 5
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\ForceClassicControlPanel", 1, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call systemTweaks()
		Case 6
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\MTCUVC\EnableMtcUvc", 0, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call systemTweaks()
		Case 9
		oWSH.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\IconVerticalSpacing", 75, "REG_DWORD"
		oWSH.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\IconSpacing", 75, "REG_DWORD"
			MsgBox "You will need to log off to apply the changes.", vbInformation + vbOkOnly, "DFX WinTweaks Icon Spacing"
			Call systemTweaks()
		Case 7
			cls
			textf "  Wait..."
			wait 0.2
			textf ""
			oWSH.Run "UserAccountControlSettings.exe"
			MsgBox "After changing this setting, you must restart the PC. Do you want to do it now?", vbInformation + vbYesNo, "DFX WinTweaks UAC"
	Select Case result
  	  Case vbYes
		Set objShell = WScript.CreateObject("WScript.Shell")
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
  	  Case vbNo
		cls
		textf = "  Wait..."
			Call systemTweaks()
		Case 99
			Call restoreSysTweaks()
		Case 0
			Call mainMenu
		Case 8
			cls
			textf " Uncheck the option: Users must enter their name and password to use the PC"
			textf " Accept changes and restart your PC"
			wait 0.2
			oWSH.Run "control userpasswords2"
			wait 0.2
			Call systemTweaks()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 1
			Call systemTweaks()
		End Select
	End Select
End Function

Function restoreSysTweaks()
	cls
	On Error Resume Next	
	textf ""
	textf "    ____          _                  _          _   _   RESTORE   "
	textf "   / ___|   _ ___| |_ ___  _ __ ___ (_)______ _| |_(_) ___  _ __  "
	textf "  | |  | | | / __| __/ _ \| '_ ` _ \| |_  / _` | __| |/ _ \| '_ \ "
	textf "  | |__| |_| \__ \ || (_) | | | | | | |/ / (_| | |_| | (_) | | | |"
	textf "   \____\__,_|___/\__\___/|_| |_| |_|_/___\__,_|\__|_|\___/|_| |_|"  
	textf ""
	textf "  Select an option:"
	textf ""
	textf " "
	textf "  1 = Disable Dark mode"
	textf "  2 = Remove the 'God Mode' icon on the Desktop"
	textf "  3 = Disable 'Quick Access' on Windows Explorer"
	textf "  4 = Stop showing file extensions" 
	textf "  5 = Disable 'Classic View' on the Control Panel"
	textf "  6 = Disable Classic Volume slider"
	textf "  7 = Disable CMD on pressing Win+U (Safe Mode)"
	textf " "
	textf ""
	textf "  0 = Back to previous menu"
	textf ""
	textl "  > "
	Select Case scanf
		Case 1
			textf ""
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call restoreSysTweaks()
		Case 2
			textf ""
		godFolder = oWSH.SpecialFolders("Desktop") & "\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
		If oFSO.FolderExists(godFolder) = True Then oFSO.DeleteFolder(godFolder)
			textf ""
			textf ""
			wait 0.2
			Call restoreSysTweaks()
		Case 3
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\LaunchTo", 2, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call restoreSysTweaks()
		Case 4
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", 1, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call restoreSysTweaks()
		Case 5
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\ForceClassicControlPanel", 0, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call restoreSysTweaks()
		Case 6
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\MTCUVC\EnableMtcUvc", 1, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call restoreSysTweaks()
		Case 7
		oWSH.RegDelete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\utilman.exe\Debugger"
			textf ""
			textf ""
			wait 0.2
			Call restoreSysTweaks()
		Case 0
			Call mainMenu()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 1
			Call restoreSysTweaks()
			Exit Function
		End Select
End Function

Function onedriveConf()
	cls
	On Error Resume Next	
	textf "   __  __ ____     ___             ____       _           "
	textf "  |  \/  / ___|   / _ \ _ __   ___|  _ \ _ __(_)_   _____ "
	textf "  | |\/| \___ \  | | | | '_ \ / _ \ | | | '__| \ \ / / _ \"
	textf "  | |  | |___) | | |_| | | | |  __/ |_| | |  | |\ V /  __/"
	textf "  |_|  |_|____/   \___/|_| |_|\___|____/|_|  |_| \_/ \___|"                                                               
	textf ""
	textf "  Select an option:"
	textf ""
	textf "  1 = Disable MS OneDrive"
	textf "  2 = Enable MS OneDrive"
	textf ""
	textf "  0 = Return to menu"
	textf ""
	textl "  > "
	Select Case scanf
		Case "1"
			textf ""
			textf " Disabling OneDrive..."
			wait 1
				oWSH.Run "taskkill.exe /F /IM OneDrive.exe /T"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 1, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 1, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableLibrariesDefaultSaveToOneDrive", 1, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableMeteredNetworkFileSync", 1, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\Onedrive\DisableLibrariesDefaultSaveToOneDrive", 1, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\Onedrive\DisableMeteredNetworkFileSync", 1, "REG_DWORD"
				oWSH.RegWrite "HKCR\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
				oWSH.RegWrite "HKCR\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
				oWSH.RegWrite "HKCU\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
				oWSH.RegWrite "HKCU\Software\Classes\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 0, "REG_DWORD"
				oWSH.RegDelete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\OneDrive"
			textf ""
			textf " INFO: OneDrive has been disabled"
			wait 1
		Case "2"
			textf ""
			textf " Enabling OneDrive..."
			wait 1
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableLibrariesDefaultSaveToOneDrive", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableMeteredNetworkFileSync", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\Onedrive\DisableLibrariesDefaultSaveToOneDrive", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\Onedrive\DisableMeteredNetworkFileSync", 0, "REG_DWORD"
				oWSH.RegWrite "HKCR\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 1, "REG_DWORD"
				oWSH.RegWrite "HKCR\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 1, "REG_DWORD"
				oWSH.RegWrite "HKCU\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 1, "REG_DWORD"
				oWSH.RegWrite "HKCU\Software\Classes\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\System.IsPinnedToNameSpaceTree", 1, "REG_DWORD"
			textf ""
			textf " INFO: OneDrive is now enabled"
			wait 1
		Case "0"
			Call mainMenu()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 1
			Call onedriveConf()
	End Select
	Call onedriveConf()
End Function

Function menuCortana()
	cls
	On Error Resume Next
	textf "   __  __ ____     ____           _                    "
	textf "  |  \/  / ___|   / ___|___  _ __| |_ __ _ _ __   __ _ "
	textf "  | |\/| \___ \  | |   / _ \| '__| __/ _` | '_ \ / _` |"
	textf "  | |  | |___) | | |__| (_) | |  | || (_| | | | | (_| |"
	textf "  |_|  |_|____/   \____\___/|_|   \__\__,_|_| |_|\__,_|"                                                         
	textf " "
	textf "  Select an option:"
	textf ""
	textf "  1 = Disable MS Cortana"
	textf "  2 = Enable MS Cortana"
	textf "  3 = Remove Cortana App (Windows 10 20H1 and later)"
	textf ""
	textf "  0 = Return to menu"
	textf ""
	textl "  > "
	Select Case scanf
		Case "1"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\AllowCortana", 0, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\CortanaEnabled", 0, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\SearchboxTaskbarMode", 0, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\BingSearchEnabled", 0, "REG_DWORD"
			textf ""
			textf " >> Restarting Windows Explorer..."
			oWSH.Run "taskkill.exe /F /IM explorer.exe"
			wait 5
			oWSH.Run "explorer.exe"
		Case "2"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\AllowCortana", 1, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\CortanaEnabled", 1, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\SearchboxTaskbarMode", 1, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\BingSearchEnabled", 1, "REG_DWORD"
			textf ""
			textf " >> Restarting Windows Explorer..."
			oWSH.Run "taskkill.exe /F /IM explorer.exe"
			wait 5
			oWSH.Run "explorer.exe"
		Case 3
			oWSH.Run "powershell Get-AppxPackage Microsoft.549981C3F5F10 | Remove-AppxPackage"
		Case "0"
			Call mainMenu()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 1
			Call menuCortana()
	End Select
	Call menuCortana()
End Function

Function menuTracking()
	cls
	On Error Resume Next
	textf "   _____               _    _             "
	textf "  |_   _| __ __ _  ___| | _(_)_ __   __ _ "
	textf "    | || '__/ _` |/ __| |/ / | '_ \ / _` |"
	textf "    | || | | (_| | (__|   <| | | | | (_| |"
	textf "    |_||_|  \__,_|\___|_|\_\_|_| |_|\__, |"
	textf "                                     |___/" 
	textf ""
	textf "  Select an option:"
	textf ""
	textf "  1 = Disable tracking"
	textf ""
	textf "  2 = Enable tracking"
	textf " "
	textf " "
	textf "  0 = Return to menu"
	textf ""
	textl "  > "
	Select Case scanf
		Case 1
			textf ""
			textf " Disabling tracking services..."
			oWSH.Run "sc stop DiagTrack"
			oWSH.Run "sc config DiagTrack start= disabled"
			oWSH.Run "sc stop dmwappushservice"
			oWSH.Run "sc config dmwappushservice start= disabled"
			wait 0.2
			Call menuTracking()
		Case 2
			textf ""
			textf " Enabling tracking services..."
			oWSH.Run "sc start DiagTrack"
			oWSH.Run "sc config DiagTrack start= enabled"
			oWSH.Run "sc start dmwappushservice"
			oWSH.Run "sc config dmwappushservice start= enabled"
			wait 0.2
			Call menuTracking()
		Case 0
			Call mainMenu()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 1
			Call menuTracking()
	End Select
	Call menuTracking()
End Function

Function defenderConf()
	cls
	On Error Resume Next
	textf "   __  __ ____    ____        __                _           "
	textf "  |  \/  / ___|  |  _ \  ___ / _| ___ _ __   __| | ___ _ __ "
	textf "  | |\/| \___ \  | | | |/ _ \ |_ / _ \ '_ \ / _` |/ _ \ '__|"
	textf "  | |  | |___) | | |_| |  __/  _|  __/ | | | (_| |  __/ |   "
	textf "  |_|  |_|____/  |____/ \___|_|  \___|_| |_|\__,_|\___|_|   "
	textf ""
	textf "  In some versions of Windows 10 and 11, MS Defender have to be disabled in Safe Mode, because"
	textf "  doing so in normal mode will not work."
	textf ""
	textf "  Select an option:"
	textf ""
	textf "  1 = Disable MS Defender"
	textf "  2 = Enable MS Defender"
	textf " "
	textf "  3 = Safe Mode Settings"
	textf ""
	textf "  0 = Return to menu"
	textf ""
	textl "  > "
	Select Case scanf
		Case "1"
			textf ""
			textf " Disabling MS Defender..."
			wait 1
		oWSH.Run "sc stop WdNisSvc"
		oWSH.Run "sc stop WinDefend"
		oWSH.Run "sc config WdNisSvc start=disabled"
		oWSH.Run "sc config WinDefend start=disabled"	
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cache Maintenance" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cleanup" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Scheduled Scan" & chr(34) & " /DISABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Verification" & chr(34) & " /DISABLE"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableBehaviorMonitoring", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableOnAccessProtection", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableScanOnRealtimeEnable", 1, "REG_DWORD"
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\NOC_GLOBAL_SETTING_TOASTS_ENABLED", 0, "REG_DWORD"	
			textf ""
			textf " MS Defender has been disabled"
			wait 1
		Case "2"
			textf ""
			textf " Enabling MS Defender..."
			wait 1
		oWSH.Run "sc config WdNisSvc start=auto"
		oWSH.Run "sc config WinDefend start=auto"	
		oWSH.Run "sc start WdNisSvc"
		oWSH.Run "sc start WinDefend"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cache Maintenance" & chr(34) & " /ENABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cleanup" & chr(34) & " /ENABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Scheduled Scan" & chr(34) & " /ENABLE"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Verification" & chr(34) & " /ENABLE"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableBehaviorMonitoring", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableOnAccessProtection", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection\DisableScanOnRealtimeEnable", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\NOC_GLOBAL_SETTING_TOASTS_ENABLED", 1, "REG_DWORD"
			textf ""
			textf " MS Defender is now enabled"
			wait 1
		Case 3
			Call safemoConf()
		Case "0"
			Call mainMenu()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 1
			Call defenderConf()
	End Select
	Call defenderConf()
End Function

Function wupdateConf()
	cls
	On Error Resume Next
	textf "  __        ___           _                     _   _           _       _       "
	textf "  \ \      / (_)_ __   __| | _____      _____  | | | |_ __   __| | __ _| |_ ___ "
	textf "   \ \ /\ / /| | '_ \ / _` |/ _ \ \ /\ / / __| | | | | '_ \ / _` |/ _` | __/ _ \"
	textf "    \ V  V / | | | | | (_| | (_) \ V  V /\__ \ | |_| | |_) | (_| | (_| | ||  __/"
	textf "     \_/\_/  |_|_| |_|\__,_|\___/ \_/\_/ |___/  \___/| .__/ \__,_|\__,_|\__\___|"
	textf "                                                     |_|     		BETA    "
	textf " "
	textf "  Windows Update on some releases of Windows 10 and Windows 11 may ignore this option and continue working."
	textf "  I'm currently trying to fix this issue."
	textf " "
	textf "  Select an option:"
	textf " "
	textf " "
	textf "  1 = Disable Windows Update"
	textf " "
	textf "  2 = Enable Windows Update" 
	textf " "
	textf " "
	textf "  0 = Return to menu"
	textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
	Call wupdateConf()
		Exit Function
	End If
	Select Case RP
		Case 1
		oWSH.Run "sc stop wuauserv"
		oWSH.Run "sc config wuauserv start=disabled"
	cls
	textf ""
	textf "  Windows Update is now disabled"
	wait 1
		Call wupdateConf()
		Case 2
		oWSH.Run "sc config wuauserv start=auto"
		oWSH.Run "sc start wuauserv"
	cls
	textf ""
	textf "  Windows Update is now enabled"
	wait 1
		Call wupdateConf()		
		Case 0
	Call mainMenu()
End Select
End Function

Function perfTweaks()
	cls
	On Error Resume Next	
	textf "   ____            __                                             _                      _        "
	textf "  |  _ \ ___ _ __ / _| ___  _ __ _ __ ___   __ _ _ __   ___ ___  | |___      _____  __ _| | _____ "
	textf "  | |_) / _ \ '__| |_ / _ \| '__| '_ ` _ \ / _` | '_ \ / __/ _ \ | __\ \ /\ / / _ \/ _` | |/ / __|"
	textf "  |  __/  __/ |  |  _| (_) | |  | | | | | | (_| | | | | (_|  __/ | |_ \ V  V /  __/ (_| |   <\__ \"
	textf "  |_|   \___|_|  |_|  \___/|_|  |_| |_| |_|\__,_|_| |_|\___\___|  \__| \_/\_/ \___|\__,_|_|\_\___/"                                                             
	textf ""
	textf ""
	textf ""
	textf "  Select an option:"
	textf ""
	textf ""
	textf "  1 = Disable BitLocker, Encryption and OfflineFiles"
	textf ""
	textf "  2 = Disable WiFi services"
	textf ""
	textf "  3 = Open Windows disk cleaner"
	textf ""
	textf "  4 = Additional Windows Features"
	textf ""
	textf "  5 = Enable all system bandwith"
	textf ""
	textf ""
	textf "  0 = Return to menu			99 = Restore"
	textf ""
	textl "  > "
	Select Case scanf
		Case 1
			textf ""
		oWSH.Run "sc config BDESVC start=disabled"
		oWSH.Run "sc config EFS start=disabled"
		oWSH.Run "sc config CscService start=disabled"
		oWSH.Run "sc stop BDESVC"
		oWSH.Run "sc stop EFS"
		oWSH.Run "sc stop CscService"
			textf ""
			textf ""
			wait 0.2
			Call perfTweaks()
		Case 2
			textf ""
		oWSH.Run "sc config WlanSvc start=disabled"
		oWSH.Run "sc stop WlanSvc"
			textf ""
			textf ""
			wait 0.2
			Call perfTweaks()
		Case 3
		oWSH.Run "cleanmgr.exe"
			textf ""
			textf ""
			wait 0.2
			Call perfTweaks()
		Case 4
		oWSH.Run "optionalfeatures.exe"
			textf ""
			textf ""
			wait 0.2
			Call perfTweaks()
		Case 5
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Psched\Psched", 0, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call perfTweaks()
		Case 99
			Call restorePerformanceEN()
		Case 0
			Call mainMenu()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 1
			Call perfTweaks()
			Exit Function
		End Select
End Function

Function restorePerformanceEN()
	cls
	On Error Resume Next	
	textf "   ____            __                                             _       Restore        _        "
	textf "  |  _ \ ___ _ __ / _| ___  _ __ _ __ ___   __ _ _ __   ___ ___  | |___      _____  __ _| | _____ "
	textf "  | |_) / _ \ '__| |_ / _ \| '__| '_ ` _ \ / _` | '_ \ / __/ _ \ | __\ \ /\ / / _ \/ _` | |/ / __|"
	textf "  |  __/  __/ |  |  _| (_) | |  | | | | | | (_| | | | | (_|  __/ | |_ \ V  V /  __/ (_| |   <\__ \"
	textf "  |_|   \___|_|  |_|  \___/|_|  |_| |_| |_|\__,_|_| |_|\___\___|  \__| \_/\_/ \___|\__,_|_|\_\___/"                                                            
	textf ""
	textf ""
	textf ""
	textf "  Select an option:"
	textf ""
	textf ""
	textf "  1 = Enable BitLocker, Encryption and OfflineFiles"
	textf ""
	textf "  2 = Enable WiFi services"
	textf ""
	textf "  3 = Disable all system bandwith"
	textf ""
	textf ""
	textf "  0 = Return to previous menu"
	textf ""
	textl "  > "
	Select Case scanf
		Case 1
			textf ""
		oWSH.Run "sc config BDESVC start=auto"
		oWSH.Run "sc config EFS start=auto"
		oWSH.Run "sc config CscService start=auto"
		oWSH.Run "sc start BDESVC"
		oWSH.Run "sc start EFS"
		oWSH.Run "sc start CscService"
			textf ""
			textf ""
			wait 0.2
			Call restorePerformanceEN()
		Case 2
			textf ""
		oWSH.Run "sc config WlanSvc start=auto"
		oWSH.Run "sc start WlanSvc"
			textf ""
			textf ""
			wait 0.2
			Call restorePerformanceEN()
		Case 3
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Psched\Psched", 20, "REG_DWORD"
			textf ""
			textf ""
			wait 0.2
			Call restorePerformanceEN()
		Case 0
			Call perfTweaks()
		Case Else
			textf ""
			textf "  This option does not exist."
			wait 1
			Call restorePerformanceEN()
			Exit Function
		End Select
End Function


Function uwpDebloat()
	cls
	On Error Resume Next
	textf "   _   ___        ______    ____       _     _             _            "
	textf "  | | | \ \      / /  _ \  |  _ \  ___| |__ | | ___   __ _| |_ ___ _ __ "
	textf "  | | | |\ \ /\ / /| |_) | | | | |/ _ \ '_ \| |/ _ \ / _` | __/ _ \ '__|"
	textf "  | |_| | \ V  V / |  __/  | |_| |  __/ |_) | | (_) | (_| | ||  __/ |   "
	textf "   \___/   \_/\_/  |_|     |____/ \___|_.__/|_|\___/ \__,_|\__\___|_|   "
	textf " "
	textf "  This will uninstall EVERY UWP APP detected in your Windows install. (excluding Settings, of course)"
	textf "  Doing this will not crazily increase your system performance."
	textf "  THIS OPTION IS NOT REVERSIBLE (at least offline), get that in mind."
	textf " "
	textf " "
	textf "  1 = Debloat now!"
	textf " "
	textf " "
	textf "  0 = Return to menu"
	textf " "
	textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
		wait 1
		Call uwpDebloat()
		Exit Function
	End If
	Select Case RP
	Case 1
		textf "  Debloating your Windows install..."
		wait 0.2
		textf "  This will take a while..."
		wait 0.2
		oWSH.Run "powershell Get-AppxPackage -AllUsers | Remove-AppxPackage"
		textf ""
		textf "  All apps have been successfully uninstalled..."
		Call mainMenu()
	Case 0
	Call mainMenu()
End Select
End Function

Function safemoConf()
	cls
	textf " "
	textf "   ____         __        __  __           _        ____       _   _   _                 "
	textf "  / ___|  __ _ / _| ___  |  \/  | ___   __| | ___  / ___|  ___| |_| |_(_)_ __   __ _ ___ "
	textf "  \___ \ / _` | |_ / _ \ | |\/| |/ _ \ / _` |/ _ \ \___ \ / _ \ __| __| | '_ \ / _` / __|"
	textf "   ___) | (_| |  _|  __/ | |  | | (_) | (_| |  __/  ___) |  __/ |_| |_| | | | | (_| \__ \"
	textf "  |____/ \__,_|_|  \___| |_|  |_|\___/ \__,_|\___| |____/ \___|\__|\__|_|_| |_|\__, |___/"
	textf "                                                                              |___/      "
	textf " "
	textf " "
	textf "  Select an option:"
	textf " "
	textf " "
	textf "  1 = Restart in Safe Mode (Normal)"
	textf " "
	textf "  2 = Restart in Safe Mode (Networking)"
	textf " "
	textf "  3 = Reboot to Standard Windows"
	textf " "
	textf " "
	textf "  0 = Return to menu"
	textf " "
	textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
		wait 1
		Call safemoConf()
		Exit Function
	End If
	Select Case RP
	Case 1	
		MsgBox "Your PC will reboot right after you close this window, make sure you saved all your data", vbInformation + vbOkOnly, "DFX WinTweaks Safe Mode"
		Set objShell = WScript.CreateObject("WScript.Shell")
		oWSH.Run "bcdedit /set {current} safeboot minimal"
		wait 1
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
	Case 2
		MsgBox "Your PC will reboot right after you close this window, make sure you saved all your data", vbInformation + vbOkOnly, "DFX WinTweaks Safe Mode"
		Set objShell = WScript.CreateObject("WScript.Shell")
		oWSH.Run "bcdedit /set {current} safeboot network"
		wait 1
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
	Case 3
		MsgBox "Your PC will reboot right after you close this window, make sure you did all your changes", vbInformation + vbOkOnly, "DFX WinTweaks Safe Mode"
		Set objShell = WScript.CreateObject("WScript.Shell")
		oWSH.Run "bcdedit /deletevalue {current} safeboot"
		wait 1
		objShell.Run "%WINDIR%\system32\shutdown.exe -r -t 0"
	Case 0
		cls
		wait 0.2
		Call mainMenu()
		Exit Function
	End Select
End Function

Function dfxCredits()
cls
textf " "
textf " " & versionNameST
textf "________________________________________________________________________________________________________________________"
textf "  Special Thanks to:"
textf "  "
textf "  "
textf "  - AikonCWD, who developed AikonCWD W10 Script. Without his work, DFX WinTweaks would never have been possible."
textf "  "
textf "  - Users who made the same questions as me in 2009 on StackOverflow."
textf "  "
textf "  - GitHub, what made DFX WinTweaks distribution possible"
textf "  "
textf "  - My friend, who supported NT6 way longer than originally planned."
textf "  "
textf "  "
textf "  "
textf "  "
textf "  "
textf "  "
textf "  "
textf "  - And you, the user, for using DFX WinTweaks <3"
textf "  "
textf "  "
textf "  Licensed under a GNU General Public License v3.0"
textf "_______________________________________________________________________________________________________________________"
textf " "
textf "  Updated Aug 27, 2023	0 = Return to Start Menu					 2022 - 2023 ivandfx"
textf " "
textl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		textf ""
		textf "  This option does not exist."
		Call dfxCredits()
		Exit Function
	End If
	Select Case RP
	Case 0
		Call startMenu()
		Exit Function
	End Select
End Function

Function tweakerexit()
cls
textf " "
textf " "
textf " "
textf " "
textf "________________________________________________________________________________________________________________________"
textf " "
textf " "
textf " "
textf " "
textf " "
textf " 			  ____  _______  __ __        ___     _____                    _        "
textf " 			 |  _ \|  ___\ \/ / \ \      / (_)_ _|_   _|_      _____  __ _| | _____ "
textf " 			 | | | | |_   \  /   \ \ /\ / /| | '_ \| | \ \ /\ / / _ \/ _` | |/ / __|"
textf "			 | |_| |  _|  /  \    \ V  V / | | | | | |  \ V  V /  __/ (_| |   <\__ \"
textf "  			 |____/|_|   /_/\_\    \_/\_/  |_|_| |_|_|   \_/\_/ \___|\__,_|_|\_\___/ is closing..."
textf " "
textf " "
textf " "
textf " 								2023 ivandfx"
textf " "
textf " "
textf " "
textf "________________________________________________________________________________________________________________________"
textf " "
textf " "
textf " "
textf " "
wait 2
WScript.Quit
End Function

Function dfxBanner()
	textf" "
	textf "   ____  _______  __ __        ___     _____                    _        "
	textf "  |  _ \|  ___\ \/ / \ \      / (_)_ _|_   _|_      _____  __ _| | _____ "
	textf "  | | | | |_   \  /   \ \ /\ / /| | '_ \| | \ \ /\ / / _ \/ _` | |/ / __|"
	textf "  | |_| |  _|  /  \    \ V  V / | | | | | |  \ V  V /  __/ (_| |   <\__ \"
	textf "  |____/|_|   /_/\_\    \_/\_/  |_|_| |_|_|   \_/\_/ \___|\__,_|_|\_\___/"
	textf "     Created by ivandfx	                          	" & currentVersionST
	textf " "
	textf "  Licensed under a GNU General Public License v3.0"
	textf " "
End Function

Function seLogonExit()
	cls
	textf "  DFX WinTweaks is closing. Thank you!"
	textf "  "
	wait 2
	WScript.Quit
End Function

Function updatedl()
	On Error Resume Next
	cls
	textf" "
	textf "   ____  _______  __ __        ___     _____                    _        "
	textf "  |  _ \|  ___\ \/ / \ \      / (_)_ _|_   _|_      _____  __ _| | _____ "
	textf "  | | | | |_   \  /   \ \ /\ / /| | '_ \| | \ \ /\ / / _ \/ _` | |/ / __|"
	textf "  | |_| |  _|  /  \    \ V  V / | | | | | |  \ V  V /  __/ (_| |   <\__ \"
	textf "  |____/|_|   /_/\_\    \_/\_/  |_|_| |_|_|   \_/\_/ \___|\__,_|_|\_\___/"
	textf "     Created by ivandfx			    Update from GitHub (BETA)"
	textf " "
	textf "  Licensed under a GNU General Public License v3.0"
	textf " "
	textf " "
	textf " "
	textf "  You're running version " & currentVersionST
	oWEB.Open "GET", "https://raw.githubusercontent.com/ivandfx/DFXWinTweaks/master/update", False
	oWEB.Send
	textf "  And the newest one is " & oWEB.responseText

	If CDbl(Replace(oWEB.responseText, vbcrlf, "")) > CDbl(currentVersion) Then
		textl "  Do you want to update? (y/n): "
		res = scanf()
		If res = "y" Then
			textf ""
			textl " Downloading update... "
			oWEB.Open "GET", "https://raw.githubusercontent.com/aikoncwd/win10script/master/dfxwt_st.vbs", False
			oWEB.Send
			wait(1)
			Set F = oFSO.CreateTextFile(WScript.ScriptFullName, 2, True)
				F.Write oWEB.responseText
			F.Close
			textf "OK!"
			wait(1)
			oWSH.Run WScript.ScriptFullName
			WScript.Quit
		End If
	Else
		textf "   Download completed"
		textf "   Starting DFX WinTweaks version " & currentVersionST
	End If
End Function

Function textf(txt)
	WScript.StdOut.WriteLine txt
End Function

Function textl(txt)
	WScript.StdOut.Write txt
End Function

Function scanf()
	scanf = LCase(WScript.StdIn.ReadLine)
End Function

Function wait(n)
	WScript.Sleep Int(n * 1000)
End Function

Function cls()
	For i = 1 To 50
		textf ""
	Next
End Function

Function ForceConsole()
	If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
		oWSH.Run "cscript //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
		WScript.Quit
	End If
End Function

Function checkNTonStart()
If getNTversion < 10 Then
	result = MsgBox ("This version of DFX WinTweaks requires Windows 10 or later.", vbExclamation + vbOkOnly, "DFX WinTweaks")
		WScript.Quit
End If
End Function

Function runElevated()
	If isUACRequired Then
		If Not isElevated Then RunAsUAC
	Else
		If Not isAdmin Then
	result = MsgBox ("DFX WinTweaks needs Administrator privileges", vbCritical + vbOkOnly, "DFX WinTweaks: Administrator")
			WScript.Quit
		End If
	End If
End Function

Function isUACRequired()
	r = isUAC()
	If r Then
		intUAC = oWSH.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA")
		r = 1 = intUAC
	End If
	isUACRequired = r
End Function

Function isElevated()
	isElevated = CheckCredential("S-1-16-12288")
End Function

Function isAdmin()
	isAdmin = CheckCredential("S-1-5-32-544")
End Function
 
Function CheckCredential(p)
	Set oWhoAmI = oWSH.Exec("whoami /groups")
	Set WhoAmIO = oWhoAmI.StdOut
	WhoAmIO = WhoAmIO.ReadAll
	CheckCredential = InStr(WhoAmIO, p) > 0
End Function
 
Function RunAsUAC()
	If isUAC Then
		Call dfxBanner()
		textf ""
		textf "  DFX WinTweaks needs to be ran with Administrator privileges"
		textf "  Waiting for UAC prompt..."
		wait 0.3
		oAPP.ShellExecute "cscript", "//NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34), "", "runas", 1
		WScript.Quit
	End If
End Function
 
Function isUAC()
	Set cWin = oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	r = False
	For Each OS In cWin
		If Split(OS.Version,".")(0) > 5 Then
			r = True
		Else
			r = False
		End If
	Next
	isUAC = r
End Function

Function CheckCredential(p)
	Set oWhoAmI = oWSH.Exec("whoami /groups")
	Set WhoAmIO = oWhoAmI.StdOut
	WhoAmIO = WhoAmIO.ReadAll
	CheckCredential = InStr(WhoAmIO, p) > 0
End Function
 
Function isUAC()
	Set cWin = oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	r = False
	For Each OS In cWin
		If Split(OS.Version,".")(0) > 5 Then
			r = True
		Else
			r = False
		End If
	Next
	isUAC = r
End Function

Function getNTversion()
	Set cWin = oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	For Each OS In cWin
		getNTversion = Split(OS.Version,".")(0)
	Next
End Function
