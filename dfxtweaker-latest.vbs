On Error Resume Next
Randomize

Set oADO = CreateObject("Adodb.Stream")
Set oWSH = CreateObject("WScript.Shell")
Set oAPP = CreateObject("Shell.Application")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWEB = CreateObject("MSXML2.ServerXMLHTTP")
Set oVOZ = CreateObject("SAPI.SpVoice")
Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")

currentVersion = "1.4.2 "On Error Resume Next
Randomize

Set oADO = CreateObject("Adodb.Stream")
Set oWSH = CreateObject("WScript.Shell")
Set oAPP = CreateObject("Shell.Application")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWEB = CreateObject("MSXML2.ServerXMLHTTP")
Set oVOZ = CreateObject("SAPI.SpVoice")
Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")

currentVersion = "1.5 "
currentFolder  = oFSO.GetParentFolderName(WScript.ScriptFullName)
currentBuild   = "Build 1190 "
currentLanguage = " Spanish_ES"
currentRelease = "Release"
Call ForceConsole()
Call showBanner()
printf " "
Call checkW10orW11()
Call runElevated()
printf "  Cargando el script..."
wait 0.3
Call showMenu(1)

Function showBanner()
	printf" "
	printf"   ____  _______  __  _____ Para Windows 10 y  _  Windows 11 "
	printf"  |  _ \|  ___\ \/ / |_   _|_      _____  __ _| | _____ _ __ "
	printf"  | | | | |_   \  /    | | \ \ /\ / / _ \/ _` | |/ / _ \ '__|"
	printf"  | |_| |  _|  /  \    | |  \ V  V /  __/ (_| |   <  __/ |   "
	printf"  |____/|_|   /_/\_\   |_|   \_/\_/ \___|\__,_|_|\_\___|_| "
        printf "     Creado por ivandfx            	 	          v" & currentVersion
	printf " "
	printf "  DFX Tweaker es un fork de AikonCWD Script 5.6"
	printf "  Bajo licencia de Creative Commons 4.0"
End Function

Function showMenu(n)
	wait(n)
	cls
	Call showBanner
	printf "  "
	printf "  Selecciona una opcion:                   		     11 = Ayuda sobre (1X) y (!)"
	printf " "
	printf "  1 = Configurar tweaks de sistema            		     12 = Opciones de apagado avanzadas "
	printf "  2 = Configurar tweaks de rendimiento			     13 = Sobre mi version de Windows"
	printf "  3 = Desinstalar aplicaciones de Windows 10 (1X)"
	printf ""
	printf "  4 = Eliminar la telemetria (!)"
	printf "  5 = Configurar MS OneDrive"
	printf "  6 = Configurar MS Cortana (1X)"
	printf "  7 = Configurar Windows Defender"
	printf "  8 = Configurar Windows Update"
	printf ""
	printf "  9 = Ver el estado de la licencia de Windows"
	printf "  10 = Atajos de teclado de Windows"
	printf ""
	printf "  99 = Restauracion"
	printf "  0 = Salir"
	printf ""
	printl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		printf ""
		printf " Solo se permiten numeros."
		Call showMenu(2)
		Exit Function
	End If
	Select Case RP
		Case 1
			Call menuSysTweaks()
		Case 2
			Call menuPerfomance()
		Case 3
			Call menuCleanApps()
		Case 4
			Call menuTelemetry()
		Case 5
			Call menuOneDrive()
		Case 6
			Call menuCortana()
		Case 7
			Call menuWindowsDefender()
		Case 8
			Call menuWindowsUpdate()
		Case 9
			Call menuXPR()
		Case 10
			Call showKeyboardTips()
		Case 99
			Call restoreMenu()
		Case 11
			MsgBox "Las opciones con (1X) solo son compatibles con Windows 10. Las opciones con (!) son irrevertibles o pueden causar problemas.", vbInformation + vbOkOnly, "DFX Tweaker: Ayuda"
			Call showMenu(0)
		Case 12
			Call shutdownMenu()
		Case 13
			oWSH.Run "winver.exe"
			Call showMenu (0)
		Case 550
			Call menuPowerSSD()		
		Case 0
			cls
			printf ""
			printf"   ____  _______  __  _____                    _		    "
			printf"  |  _ \|  ___\ \/ / |_   _|_      _____  __ _| | _____ _ __ "
			printf"  | | | | |_   \  /    | | \ \ /\ / / _ \/ _` | |/ / _ \ '__|"
			printf"  | |_| |  _|  /  \    | |  \ V  V /  __/ (_| |   <  __/ |   "
			printf"  |____/|_|   /_/\_\   |_|   \_/\_/ \___|\__,_|_|\_\___|_|   "
			printf ""
			printf " 	  Gracias por utilizar DFX Tweaker :D"
			printf "                        ivandfx"
			printf " "
			printf "                 " & currentBuild & currentLanguage
			printf ""
			printf ""
			printf ""
			printf ""
			wait 1
			printf "  El script se está cerrando..."
			wait(2)
			WScript.Quit
		Case Else
			printf ""
			printf " Solo se permiten numeros."
			Call showMenu(2)
			Exit Function
	End Select
End Function

Function menuXPR()
	cls
	On Error Resume Next
	printf ""
	printf " En unos segundos aparecera el estado de tu activacion..."
	wait 0.2
	printf " Recopilando datos de la activacion..."
	wait 2
	oWSH.Run "slmgr.vbs /dli"
	oWSH.Run "slmgr.vbs /xpr"
	Call showMenu (1)
End Function

Function showBannerWAST()
	printf "  __        ___    ____ _____   _____           _              _     _          _ "
	printf "  \ \      / / \  / ___|_   _| | ____|_ __ ___ | |__   ___  __| | __| | ___  __| |"
	printf "   \ \ /\ / / _ \ \___ \ | |   |  _| | '_ ` _ \| '_ \ / _ \/ _` |/ _` |/ _ \/ _` |"
	printf "    \ V  V / ___ \ ___) || |   | |___| | | | | | |_) |  __/ (_| | (_| |  __/ (_| |"
	printf "     \_/\_/_/   \_\____/ |_|   |_____|_| |_| |_|_.__/ \___|\__,_|\__,_|\___|\__,_|"
	printf "                          Windows Advanced Shutdown Tool para DFX Tweaker 1.2"
	printf " "
	End Function

Function shutdownMenu()
	cls
	Call showBannerWAST()
	printf " "
	printf "  Cargando WAST para DFX Tweaker..."
	wait 0.4
	printf "  Espera..."
	wait 3
	cls
	On Error Resume Next
	Call showBannerWAST()
	printf " "
	printf " "
	printf " "
	printf " "
	printf " "
	printf "   ¿Que quieres hacer?:"
	printf "                                            55 = Reiniciar el Explorador de Windows"
	printf ""
	printf "  1 = Apagar el equipo"
	printf " "
	printf "  2 = Reiniciar el equipo"
	printf " "
	printf "  3 = Cerrar sesion de este usuario"
	printf " "
	printf "  4 = Ir a opciones avanzadas"
	printf ""
	printf "  5 = Causar un BSOD (Blue Screen Of Death)"
	printf " "
	printf " "
	printf "  0 = Volver al menu principal"
	printf ""
	printl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		printf ""
		printf " Solo se permiten numeros."
		Call shutdownMenu()
		Exit Function
	End If
	Select Case RP
		Case 1
			result = MsgBox ("¿Apagar?", vbYesNo, "WAST Apagado")
Select Case result
    Case vbYes
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "C:\WINDOWS\system32\shutdown.exe -s -t 0"
        Dim objShell
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
		Case 2
						result = MsgBox ("¿Reiniciar?", vbYesNo, "WAST Reinicio")
Select Case result
    Case vbYes
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 0"
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
		Case 3
						result = MsgBox ("¿Cerrar sesion? Los datos no guardados se perderan.", vbYesNo, "WAST Cierre de sesion")
Select Case result
    Case vbYes
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "C:\WINDOWS\system32\shutdown.exe -l"
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
		Case 4
						result = MsgBox ("¿Ir al menu de opciones avanzadas? Esto cerrrara todas las sesiones de todos los usuarios del equipo.", vbYesNo, "WAST Opciones avanzadas")
Select Case result
    Case vbYes
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -o -t 0"
	wait 1
		Call shutdownMenu()
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
		Case 5
						result = MsgBox ("¿Quieres causar un pantallazo azul de la muerte? Asegurate de haber guardado TODOS los datos que estuvieras usando.", vbYesNo, "WAST BSOD")
Select Case result
    Case vbYes
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "taskkill /f /im crss.exe"
	objShell.Run "taskkill /f /im winnit.exe"
	objShell.Run "taskkill /f /im winlogon.exe"
	objShell.Run "taskkill /f /im svchost.exe"
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
		Case 55
						result = MsgBox ("¿Quieres reiniciar el Explorador de Windows?", vbYesNo, "WAST Explorador")
Select Case result
    Case vbYes
	printf = "  Espera..."
		printf " >> Reiniciando el explorador de Windows... espera 5 segundos!"
		oWSH.Run "taskkill.exe /F /IM explorer.exe"
		wait(5)
		oWSH.Run "explorer.exe"
		Call shutdownMenu()
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
Case 0
		cls
		printf "  Volviendo a DFX Tweaker..."
		wait 0.3
		printf "  Espera..."
		wait 2.7
		Call showMenu(0)
	End Select
End Function

Function menuSysTweaks()
	cls
	On Error Resume Next
	printf ""
	printf "   _____                    _              _      _       _     _                       "
	printf "  |_   _|_      _____  __ _| | _____    __| | ___| |  ___(_)___| |_ ___ _ __ ___   __ _ "
	printf "    | | \ \ /\ / / _ \/ _` | |/ / __|  / _` |/ _ \ | / __| / __| __/ _ \ '_ ` _ \ / _` |"
	printf "    | |  \ V  V /  __/ (_| |   <\__ \ | (_| |  __/ | \__ \ \__ \ ||  __/ | | | | | (_| |"
	printf "    |_|   \_/\_/ \___|\__,_|_|\_\___/  \__,_|\___|_| |___/_|___/\__\___|_| |_| |_|\__,_|"
	printf ""
	printl " # Deshabilitar 'Acceso Rapido' en el Explorador de Windows? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\LaunchTo", 1, "REG_DWORD"
	Else
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\LaunchTo", 2, "REG_DWORD"
	End If
	printl " # Crear icono 'Modo Dios' en el Escritorio? (s/n) > "
	If LCase(scanf) = "s" Then
		godFolder = oWSH.SpecialFolders("Desktop") & "\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
		If oFSO.FolderExists(godFolder) = False Then oFSO.CreateFolder(godFolder)
	Else
		godFolder = oWSH.SpecialFolders("Desktop") & "\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
		If oFSO.FolderExists(godFolder) = True Then oFSO.DeleteFolder(godFolder)	
	End If
	printl " # Habilitar el tema oscuro de Windows? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"		
	End If
	printl " # ¿Mostrar icono 'Mi PC' en el Escritorio? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", 1, "REG_DWORD"
	End If
	printl " # Mostrar siempre la extension de los archivos? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", 1, "REG_DWORD"
	End If
	printl " # ¿Deshabilitar Pantalla de bloqueo? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Personalization\NoLockScreen", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\Software\Policies\Microsoft\Windows\System\DisableLogonBackgroundImage", 1, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Personalization\NoLockScreen", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\Software\Policies\Microsoft\Windows\System\DisableLogonBackgroundImage", 0, "REG_DWORD"
	End If
	printl " # ¿Forzar Vista Clasica en el Panel de Control? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\ForceClassicControlPanel", 1, "REG_DWORD"
	Else
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\ForceClassicControlPanel", 0, "REG_DWORD"
	End If
	printl " # Deshabilitar 'Reporte de Errores' de Windows? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting\Disabled", 1, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting\Disabled", 0, "REG_DWORD"
	End If
	printl " # Abrir cmd.exe al pulsar Win+U? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\utilman.exe\Debugger", "cmd.exe", "REG_SZ"
	Else
		oWSH.RegDelete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\utilman.exe\Debugger"
	End If
	printl " # Habilitar/Deshabilitar el control de cuentas de usuario UAC? (s/n) > "
	If LCase(scanf) = "s" Then
		printf ""
		printf " Ahora se abrira una ventana..."
		printf " Mueve la barra vertical hasta el nivel mas bajo"
		printf " Acepta los cambios y reinicia el PC"
		wait(2)
		printf ""
		printf " > Executing UserAccountControlSettings.exe"
		oWSH.Run "UserAccountControlSettings.exe"
		printf ""
	End If
	printl " # Habilitar/Deshabilitar el inicio de sesion sin clave? (s/n) > "
	If LCase(scanf) = "s" Then
		printf ""
		printf " Ahora se abrira una ventana..."
		printf " Desmarca la opcion: Los usuarios deben escribir su nombre y password para usar el equipo"
		printf " Acepta los cambios y reinicia el PC"
		wait(2)
		printf ""
		printf " > Executing control userpasswords2"
		oWSH.Run "control userpasswords2"
		printf ""
	End If
	printl " # Utilizar control de volumen clasico? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\MTCUVC\EnableMtcUvc", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\MTCUVC\EnableMtcUvc", 1, "REG_DWORD"
	End If
	printl " # Utilizar el centro de notificaciones clasico? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\ImmersiveShell\UseActionCenterExperience", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\ImmersiveShell\UseActionCenterExperience", 1, "REG_DWORD"
	End If
	printf ""
	printf " >> Reiniciando el explorador de Windows... espera 5 segundos!"
	oWSH.Run "taskkill.exe /F /IM explorer.exe"
	wait(5)
	oWSH.Run "explorer.exe"
	wait (3)
	printf ""
	printf " Todos los tweaks de sistema se han aplicado correctamente"
	wait (1)
	Call showMenu(2)
End Function

Function menuOneDrive()
	cls
	On Error Resume Next	
	printf "   __  __ ____     ___             ____       _           "
	printf "  |  \/  / ___|   / _ \ _ __   ___|  _ \ _ __(_)_   _____ "
	printf "  | |\/| \___ \  | | | | '_ \ / _ \ | | | '__| \ \ / / _ \"
	printf "  | |  | |___) | | |_| | | | |  __/ |_| | |  | |\ V /  __/"
	printf "  |_|  |_|____/   \___/|_| |_|\___|____/|_|  |_| \_/ \___|"                                                               
	printf ""
	printf " Selecciona una opcion:"
	printf ""
	printf "  1 = Deshabilitar MS OneDrive"
	printf "  2 = Habilitar MS OneDrive"
	printf ""
	printf "  0 = Volver al menu principal"
	printf ""
	printl "  > "
	Select Case scanf
		Case "1"
			printf ""
			printf " Deshabilitando OneDrive..."
			wait(1)
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
			printf ""
			printf " INFO: OneDrive se ha deshabilitado correctamente"
			wait(2)
		Case "2"
			printf ""
			printf " Habilitando OneDrive..."
			wait(1)
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
			printf ""
			printf " INFO: OneDrive se ha habilitado correctamente"
			wait(2)
		Case "0"
			Call showMenu(0)
		Case Else
			Call menuOneDrive()
	End Select
	Call menuOneDrive()
End Function

Function menuCortana()
	cls
	On Error Resume Next
	printf "   __  __ ____     ____           _                    "
	printf "  |  \/  / ___|   / ___|___  _ __| |_ __ _ _ __   __ _ "
	printf "  | |\/| \___ \  | |   / _ \| '__| __/ _` | '_ \ / _` |"
	printf "  | |  | |___) | | |__| (_) | |  | || (_| | | | | (_| |"
	printf "  |_|  |_|____/   \____\___/|_|   \__\__,_|_| |_|\__,_|"                                                         
	printf " "
	printf " Selecciona una opcion:"
	printf ""
	printf "  1 = Deshabilitar MS Cortana"
	printf "  2 = Habilitar MS Cortana"
	printf ""
	printf "  0 = Volver al menu principal"
	printf ""
	printl "  > "
	Select Case scanf
		Case "1"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\AllowCortana", 0, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\CortanaEnabled", 0, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\SearchboxTaskbarMode", 0, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\BingSearchEnabled", 0, "REG_DWORD"
			printf ""
			printf " >> Reiniciando el explorador de Windows... espera 5 segundos!"
			oWSH.Run "taskkill.exe /F /IM explorer.exe"
			wait(5)
			oWSH.Run "explorer.exe"
		Case "2"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\AllowCortana", 1, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\CortanaEnabled", 1, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\SearchboxTaskbarMode", 1, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\BingSearchEnabled", 1, "REG_DWORD"
			printf ""
			printf " >> Reiniciando el explorador de Windows... espera 5 segundos!"
			oWSH.Run "taskkill.exe /F /IM explorer.exe"
			wait(5)
			oWSH.Run "explorer.exe"
		Case "0"
			Call showMenu(0)
		Case Else
			Call menuCortana()
	End Select
	Call menuCortana()
End Function

Function menuTelemetry()
	cls
	On Error Resume Next
	printf "   _____    _                     _        __      "
	printf "  |_   _|__| | ___ _ __ ___   ___| |_ _ __/_/ __ _ "
	printf "    | |/ _ \ |/ _ \ '_ ` _ \ / _ \ __| '__| |/ _` |"
	printf "    | |  __/ |  __/ | | | | |  __/ |_| |  | | (_| |"
	printf "    |_|\___|_|\___|_| |_| |_|\___|\__|_|  |_|\__,_|"
	printf ""
	printf " Selecciona una opcion:"
	printf ""
	printf "  1 = Eliminar TODA la telemetria (!)"
	printf ""
	printf "  0 = Volver al menu principal"
	printf ""
	printl "  > "
	Select Case scanf
		Case "1"
			printf ""
			printf " Aplicando parches para eliminar la telemetria (5 segundos)..."
			printf " Deshabilitando la telemetria usando el registro..."
			wait 5
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection\AllowTelemetry", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\DataCollection\AllowTelemetry", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection\AllowTelemetry", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..riencehost.appxmain_31bf3856ad364e35_10.0.10240.16384_none_0ab8ea80e84d4093\f!telemetry.js", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!dss-winrt-telemetry.js", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry.js", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry-event_8ac43a41e5030538", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry-inter_58073761d33f144b", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\MRT\DontOfferThroughWUAU", 1, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat\AITEnable", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\SQMClient\Windows\CEIPEnable", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat\DisableUAR", 1, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata\PreventDeviceMetadataFromNetwork", 1, "REG_DWORD"		
			printf ""
			printf " INFO: La telemetria se ha eliminado correctamente"
			pathLOG = oWSH.ExpandEnvironmentStrings("%ProgramData%") & "\Microsoft\Diagnosis\ETLLogs\AutoLogger\AutoLogger-Diagtrack-Listener.etl"
			printf ""
			printf " Borrando DiagTrack Log..."
			wait(1)
				If oFSO.FileExists(pathLOG) Then oFSO.DeleteFile(pathLOG)
				oWSH.Run "cmd /C echo " & chr(34) & chr(34) & " > " & pathLOG
			printf ""
			printf " INFO: DiagTrack Log se ha borrado correctamente"
			printf ""
			printf " Deshabilitando servicios de seguimiento..."
			wait(1)
				oWSH.Run "sc stop TrkWks"
				oWSH.Run "sc stop DiagTrack"
				oWSH.Run "sc stop RetailDemo"
				oWSH.Run "sc stop WMPNetworkSvc"
				oWSH.Run "sc stop dmwappushservice"
				oWSH.Run "sc stop diagnosticshub.standardcollector.service"
				oWSH.Run "sc config TrkWks start=disabled"
				oWSH.Run "sc config DiagTrack start=disabled"
				oWSH.Run "sc config RetailDemo start=disabled"
				oWSH.Run "sc config WMPNetworkSvc start=disabled"
				oWSH.Run "sc config dmwappushservice start=disabled"
				oWSH.Run "sc config diagnosticshub.standardcollector.service start=disabled"
			printf ""
			printf " INFO: Servicios de seguimiento deshabilitados"
			printf ""			
			printf " Deshabilitando tareas programadas que envian datos a Microsoft..."	
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Application Experience\ProgramDataUpdater" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\KernelCeipTask" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\Uploader" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\AppID\SmartScreenSpecific" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\NetTrace\GatherNetworkInfo" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Error Reporting\QueueReporting" & chr(34) & " /DISABLE"
			printf ""
			printf " INFO: Tareas programadas de seguimiento deshabilitadas"
			'printf ""
			'printf " Descargando listado actualizado para bloquear los servidores de publicidad de Microsoft..."
			'wait(1)
			'hostsFile = oWSH.ExpandEnvironmentStrings("%WinDir%") & "\System32\drivers\etc\hosts"
			'If oFSO.FileExists(hostsFile & ".cwd") = False Then
			'	oFSO.CopyFile hostsFile, hostsFile & ".cwd"
			'	Set F = oFSO.OpenTextFile(hostsFile, 8, True)
			'		F.Write oWEB.ResponseText
			'	F.Close
			'	Set F = oFSO.OpenTextFile(hostsFile, 8, True)
			'		F.WriteLine "#Antimalware"
			'		F.WriteLine "0.0.0.0 tracking.opencandy.com.s3.amazonaws.com"
			'		F.WriteLine "0.0.0.0 media.opencandy.com"
			'		F.WriteLine "0.0.0.0 cdn.opencandy.com"
			'		F.WriteLine "0.0.0.0 tracking.opencandy.com"
			'		F.WriteLine "0.0.0.0 api.opencandy.com"
			'		F.WriteLine "0.0.0.0 api.recommendedsw.com"
			'		F.WriteLine "0.0.0.0 installer.betterinstaller.com"
			'		F.WriteLine "0.0.0.0 installer.filebulldog.com"
			'		F.WriteLine "0.0.0.0 d3oxtn1x3b8d7i.cloudfront.net"
			'		F.WriteLine "0.0.0.0 inno.bisrv.com"
			'		F.WriteLine "0.0.0.0 nsis.bisrv.com"
			'		F.WriteLine "0.0.0.0 cdn.file2desktop.com"
			'		F.WriteLine "0.0.0.0 cdn.goateastcach.us"
			'		F.WriteLine "0.0.0.0 cdn.guttastatdk.us"
			'		F.WriteLine "0.0.0.0 cdn.inskinmedia.com"
			'		F.WriteLine "0.0.0.0 cdn.insta.oibundles2.com"
			'		F.WriteLine "0.0.0.0 cdn.insta.playbryte.com"
			'		F.WriteLine "0.0.0.0 cdn.llogetfastcach.us"
			'		F.WriteLine "0.0.0.0 cdn.montiera.com"
			'		F.WriteLine "0.0.0.0 cdn.msdwnld.com"
			'		F.WriteLine "0.0.0.0 cdn.mypcbackup.com"
			'		F.WriteLine "0.0.0.0 cdn.ppdownload.com"
			'		F.WriteLine "0.0.0.0 cdn.riceateastcach.us"
			'		F.WriteLine "0.0.0.0 cdn.shyapotato.us"
			'		F.WriteLine "0.0.0.0 cdn.solimba.com"
			'		F.WriteLine "0.0.0.0 cdn.tuto4pc.com"
			'		F.WriteLine "0.0.0.0 cdn.appround.biz"
			'		F.WriteLine "0.0.0.0 cdn.bigspeedpro.com"
			'		F.WriteLine "0.0.0.0 cdn.bispd.com"
			'		F.WriteLine "0.0.0.0 cdn.bisrv.com"
			'		F.WriteLine "0.0.0.0 cdn.cdndp.com"
			'		F.WriteLine "0.0.0.0 cdn.download.sweetpacks.com"
			'		F.WriteLine "0.0.0.0 cdn.dpdownload.com"
			'		F.WriteLine "0.0.0.0 cdn.visualbee.net"
			'		F.WriteLine "#Telemetry"
			'		F.WriteLine "0.0.0.0 vortex.data.microsoft.com"
			'		F.WriteLine "0.0.0.0 vortex-win.data.microsoft.com"
			'		F.WriteLine "0.0.0.0 telecommand.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 telecommand.telemetry.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 oca.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 oca.telemetry.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 sqm.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 sqm.telemetry.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 watson.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 watson.telemetry.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 redir.metaservices.microsoft.com"
			'		F.WriteLine "0.0.0.0 choice.microsoft.com"
			'		F.WriteLine "0.0.0.0 choice.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 df.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 wes.df.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 reports.wes.df.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 services.wes.df.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 sqm.df.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 watson.ppe.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 telemetry.appex.bing.net"
			'		F.WriteLine "0.0.0.0 telemetry.urs.microsoft.com"
			'		F.WriteLine "0.0.0.0 telemetry.appex.bing.net:443"
			'		F.WriteLine "0.0.0.0 settings-sandbox.data.microsoft.com"
			'		F.WriteLine "0.0.0.0 vortex-sandbox.data.microsoft.com"
			'		F.WriteLine "0.0.0.0 survey.watson.microsoft.com"
			'		F.WriteLine "0.0.0.0 watson.live.com"
			'		F.WriteLine "0.0.0.0 watson.microsoft.com"
			'		F.WriteLine "0.0.0.0 statsfe2.ws.microsoft.com"
			'		F.WriteLine "0.0.0.0 corpext.msitadfs.glbdns2.microsoft.com"
			'		F.WriteLine "0.0.0.0 compatexchange.cloudapp.net"
			'		F.WriteLine "0.0.0.0 cs1.wpc.v0cdn.net"
			'		F.WriteLine "0.0.0.0 a-0001.a-msedge.net"
			'		F.WriteLine "0.0.0.0 statsfe2.update.microsoft.com.akadns.net"
			'		F.WriteLine "0.0.0.0 sls.update.microsoft.com.akadns.net"
			'		F.WriteLine "0.0.0.0 fe2.update.microsoft.com.akadns.net"
			'		F.WriteLine "0.0.0.0 diagnostics.support.microsoft.com"
			'		F.WriteLine "0.0.0.0 corp.sts.microsoft.com"
			'		F.WriteLine "0.0.0.0 statsfe1.ws.microsoft.com"
			'		F.WriteLine "0.0.0.0 pre.footprintpredict.com"
			'		F.WriteLine "0.0.0.0 i1.services.social.microsoft.com"
			'		F.WriteLine "0.0.0.0 i1.services.social.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 feedback.windows.com"
			'		F.WriteLine "0.0.0.0 feedback.microsoft-hohm.com"
			'		F.WriteLine "0.0.0.0 feedback.search.microsoft.com"
			'	F.Close
			'	printf ""
			'	printf " INFO: Fichero HOSTS escrito correctamente"
			'End If
			'wait(2)
		Case "0"
			Call showMenu(0)
		Case Else
			Call menuTelemetry()
	End Select
	Call menuTelemetry()
End Function

Function menuWindowsDefender()
	cls
	On Error Resume Next
	printf "   __  __ ____    ____        __                _           "
	printf "  |  \/  / ___|  |  _ \  ___ / _| ___ _ __   __| | ___ _ __ "
	printf "  | |\/| \___ \  | | | |/ _ \ |_ / _ \ '_ \ / _` |/ _ \ '__|"
	printf "  | |  | |___) | | |_| |  __/  _|  __/ | | | (_| |  __/ |   "
	printf "  |_|  |_|____/  |____/ \___|_|  \___|_| |_|\__,_|\___|_|   "
	printf ""
	printf " Selecciona una opcion:"
	printf ""
	printf "  1 = Deshabilitar Windows Defender"
	printf "  2 = Habilitar Windows Defender"
	printf ""
	printf "  0 = Volver al menu principal"
	printf ""
	printl "  > "
	Select Case scanf
		Case "1"
			printf ""
			printf " Deshabilitando Windows Defender..."
			wait(1)
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
			printf ""
			printf " INFO: Windows Defender se ha deshabilitado correctamente"
			wait(3)
		Case "2"
			printf ""
			printf " Habilitando Windows Defender..."
			wait(1)
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
			printf ""
			printf " INFO: Windows Defender se ha habilitado correctamente"
			wait(2)
		Case "0"
			Call showMenu(0)
		Case Else
			Call menuWindowsDefender()
	End Select
	Call menuWindowsDefender()
End Function

Function menuWindowsUpdate()
	cls
	On Error Resume Next
	printf "  __        ___           _                     _   _           _       _       "
	printf "  \ \      / (_)_ __   __| | _____      _____  | | | |_ __   __| | __ _| |_ ___ "
	printf "   \ \ /\ / /| | '_ \ / _` |/ _ \ \ /\ / / __| | | | | '_ \ / _` |/ _` | __/ _ \"
	printf "    \ V  V / | | | | | (_| | (_) \ V  V /\__ \ | |_| | |_) | (_| | (_| | ||  __/"
	printf "     \_/\_/  |_|_| |_|\__,_|\___/ \_/\_/ |___/  \___/| .__/ \__,_|\__,_|\__\___|"
	printf "                                                     |_|                        "
	printf ""
	printf "Para activar Windows Update, selecciona <<n>> en cada opcion"
	printf ""
	printl " # Deshabilitar 'Windows Update Service'? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\DeferUpgrade", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\AUOptions", 2, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 1, "REG_DWROD"
		oWSH.Run "sc stop wuauserv"
		oWSH.Run "sc config wuauserv start=disabled"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\DeferUpgrade", 0, "REG_DWORD"
		oWSH.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\AUOptions"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWROD"
		oWSH.Run "sc config wuauserv start=auto"
		oWSH.Run "sc start wuauserv"
	End If
	printl " # Deshabilitar 'Windows Update Sharing'? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config\DownloadMode", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config\DODownloadMode", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\SystemSettingsDownloadMode", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config\DownloadMode", 3, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config\DODownloadMode", 3, "REG_DWORD"
		oWSH.RegDelete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\SystemSettingsDownloadMode"
	End If
	printl " # Deshabilitar 'Windows Update App'? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsStore\WindowsUpdate\AutoDownload", 2, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsStore\WindowsUpdate\AutoDownload", 4, "REG_DWORD"
	End If
	printl " # Deshabilitar 'Windows Update Driver'? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\DriverSearching\DontSearchWindowsUpdate", 1, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\DriverSearching\DontSearchWindowsUpdate", 0, "REG_DWORD"
	End If
	printf ""
	printf "Todos los tweaks de Windows Update se han aplicado correctamente"
	Call showMenu(2)
End Function

Function menuPerfomance()
	cls
	On Error Resume Next
	printf "   _____                    _              _                           _ _           _            _        "
	printf "  |_   _|_      _____  __ _| | _____    __| | ___   _ __ ___ _ __   __| (_)_ __ ___ (_) ___ _ __ | |_ ___  "
	printf "    | | \ \ /\ / / _ \/ _` | |/ / __|  / _` |/ _ \ | '__/ _ \ '_ \ / _` | | '_ ` _ \| |/ _ \ '_ \| __/ _ \ "
	printf "    | |  \ V  V /  __/ (_| |   <\__ \ | (_| |  __/ | | |  __/ | | | (_| | | | | | | | |  __/ | | | || (_) |"
	printf "    |_|   \_/\_/ \___|\__,_|_|\_\___/  \__,_|\___| |_|  \___|_| |_|\__,_|_|_| |_| |_|_|\___|_| |_|\__\___/ "                                                                     
	printf ""
	printl " # Acelerar el cierre de aplicaciones y servicios? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKCU\Control Panel\Desktop\WaitToKillAppTimeout", 1000, "REG_SZ"
		oWSH.RegWrite "HKCU\Control Panel\Desktop\AutoEndTasks", 1, "REG_SZ"
		oWSH.RegWrite "HKCU\Control Panel\Desktop\HungAppTimeout", 1000, "REG_SZ"
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\WaitToKillServiceTimeout", 1000, "REG_SZ"
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize\StartupDelayInMSec", 0, "REG_DWORD"
	End If
	printl " # Deshabilitar servicios: BitLocker, Cifrado y OfflineFiles? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.Run "sc config BDESVC start=disabled"
		oWSH.Run "sc config EFS start=disabled"
		oWSH.Run "sc config CscService start=disabled"
		oWSH.Run "sc stop BDESVC"
		oWSH.Run "sc stop EFS"
		oWSH.Run "sc stop CscService"
	Else
		oWSH.Run "sc config BDESVC start=auto"
		oWSH.Run "sc config EFS start=auto"
		oWSH.Run "sc config CscService start=auto"
		oWSH.Run "sc start BDESVC"
		oWSH.Run "sc start EFS"
		oWSH.Run "sc start CscService"
	End If
	printf ""
	printf " >> No utilizar si usas un portatil o WiFi <<"
	printf ""
	printl " # Deshabilitar servicios WiFi? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.Run "sc config WlanSvc start=disabled"
		oWSH.Run "sc stop WlanSvc"
	Else
		oWSH.Run "sc config WlanSvc start=auto"
		oWSH.Run "sc start WlanSvc"
	End If
	printl " # Ejecutar limpiador de Windows. Libera espacio y borrar Windows.old (s/n) > "
	If LCase(scanf) = "s" Then	
		printf ""
		printf " Ahora se ejecutara una ventana..."
		printf " Marca las opciones deseadas de limpieza"
		printf " Acepta los cambios y reinicia el ordenador"
		wait(2)
		printf ""
		printf " > Executing cleanmgr.exe"
		oWSH.Run "cleanmgr.exe"
		printf ""
	End If
	printl " # Instalar/Desinstalar caracteristicas adicionales de Windows (s/n) > "
	If LCase(scanf) = "s" Then
		printf ""
		printf " Ahora se ejecutara una ventana..."
		printf " Marca/Desmarca las opciones deseadas"
		printf " Acepta los cambios y reinicia el ordenador"
		wait(2)
		printf ""
		printf " > Executing optionalfeatures.exe"
		oWSH.Run "optionalfeatures.exe"
		printf ""
	End If
	printl " # Cambiar la configuracion de la compresion de ficheros? (tarda un poco!) (s/n) > "
	If LCase(scanf) = "s" Then
		printl " -> Deshabilitar la compresion de ficheros en el disco duro principal? (s/n) > "
		If LCase(scanf) = "s" Then
			oWSH.Run "compact /CompactOs:never"
		Else
			oWSH.Run "compact /CompactOs:always"
		End If
		wait(3)
	End If
	printl " # Habilitar el 100% del ancho de banda para el sistema? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Psched\Psched", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Psched\Psched", 20, "REG_DWORD"
	End If
	printf ""
	printf " Todos los tweaks de sistema se han aplicado correctamente"	
	showMenu(3)
End Function

Function menuPowerSSD()
	cls
	On Error Resume Next
	printf ""
	printf ""
	printf " Felicidades, has descubierto la opcion oculta: Optimizar SSD"
	printf " "
	printf " Esta opcion se elimino porque podia causar problemas serios a usuarios con HDD"
	printf "  y causar inestabilidad en ciertos SSD"
	printf " "
	printf " Asi que utiliza esta opcion bajo TU PROPIO RIESGO, te he avisado :P"
	printf ""
	printf ""
	printf " Este script va a modificar las siguientes configuraciones:"
	printf ""
	printf "  > Habilitar TRIM"
	printf "  > Deshabilitar VSS (Shadow Copy)"
	printf "  > Deshabilitar Windows Search"
	printf "  > Deshabilitar Servicios de Indexacion"
	printf "  > Deshabilitar defragmentador de discos"
	printf "  > Deshabilitar hibernacion del sistema"
	printf "  > Deshabilitar Prefetcher + Superfetch"
	printf "  > Deshabilitar ClearPageFileAtShutdown + LargeSystemCache"
	printf ""
	printl "  # Deseas continuar y aplicar los cambios? (s/n) "	
	If scanf = "s" Then
		printf ""
		oWSH.Run "fsutil behavior set disabledeletenotify 0"
		printf " # TRIM habilitado"
		wait(1)
		oWSH.Run "vssadmin Delete Shadows /All /Quiet"
		oWSH.Run "sc stop VSS"
		oWSH.Run "sc config VSS start=disabled"
		printf " # Shadow Copy eliminada y deshabilitada"
		wait(1)
		oWSH.Run "sc stop WSearch"
		oWSH.Run "sc config WSearch start=disabled"
		printf " # Windows Search + Indexing Service deshabilitados"
		wait(1)
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\OptimizeComplete", "No"
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\Enable", "N"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Defrag\ScheduledDefrag" & chr(34) & " /DISABLE"
		printf " # Defragmentador de disco deshabilitado"
		wait(1)		
		oWSH.Run "powercfg -h off"
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power\HiberbootEnabled", 0, "REG_DWORD"
		printf " # Hibernacion deshabilitada"
		wait(1)
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnablePrefetcher", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnableSuperfetch", 0, "REG_DWORD"
		oWSH.Run "sc stop SysMain"
		oWSH.Run "sc config SysMain start=disabled"
		printf " # Prefetcher + Superfetch deshabilitados"
		wait(1)
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\ClearPageFileAtShutdown", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\LargeSystemCache", 0, "REG_DWORD"
		printf " # ClearPageFileAtShutdown + LargeSystemCache deshabilitados"
		wait(1)
		printf ""
		printf " INFO: Felicidades, acabas de prolongar la vida y el rendimiento de tu SSD"
		printf "       Se recomienda tu PC para aplicar cambios..."
	Else
		printf ""
		printf " Operacion cancelada."
	End If
	Call showMenu(3)
End Function

Function menuCleanApps()
	cls
	On Error Resume Next
	printf "      _                      _   ___        ______  "
	printf "     / \   _ __  _ __  ___  | | | \ \      / /  _ \ "
	printf "    / _ \ | '_ \| '_ \/ __| | | | |\ \ /\ / /| |_) |"
	printf "   / ___ \| |_) | |_) \__ \ | |_| | \ V  V / |  __/ "
	printf "  /_/   \_\ .__/| .__/|___/  \___/   \_/\_/  |_|    "
	printf "          |_|   |_|                                 "
	printf " "
	printf " Esto va a desinstalar las sguientes aplicaciones:"
	printf ""
	printf "  > Bing, Zune, Skype, XboxApp"
	printf "  > Getstarted, Messagin, 3D Builder"
	printf "  > Windows Maps, Phone, Camera, Alarms, People"
	printf "  > Windows Communications Apps, Sound Recorder"
	printf "  > Microsoft Office Hub, Office Sway, OneNote"
	printf "  > Solitaire Collection, CandyCrushSaga"
	printf ""
	printl " La opcion NO es reversible. Deseas continuar? (s/n) "
	
	If scanf = "s" Then
		oWSH.Run "powershell get-appxpackage -Name *Bing* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *Zune* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *XboxApp* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *OneNote* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *SkypeApp* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *3DBuilder* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *Getstarted* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *Microsoft.People* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *MicrosoftOfficeHub* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *MicrosoftSolitaireCollection* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *WindowsCamera* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *WindowsAlarms* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *WindowsMaps* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *WindowsPhone* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *WindowsSoundRecorder* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *windowscommunicationsapps* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *CandyCrushSaga* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *Messagin* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *ConnectivityStore* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *CommsPhone* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *Office.Sway* | Remove-AppxPackage", 1, True
		printf ""
		printf " > Todas las aplicaciones se han desinstalado correctamente..."
	Else
		printf ""
		printf " > Operacion cancelada."
	End If
	wait(1)
	Call showMenu(2)
End Function

Function showKeyboardTips()
	msg = msg & "WIN+A		Abre el centro de actividades" & vbcrlf
	msg = msg & "WIN+C		Activa el reconocimiento de voz de Cortana" & vbcrlf
	msg = msg & "WIN+D		Muestra el escritorio" & vbcrlf
	msg = msg & "WIN+E		Abre el explorador de Windows" & vbcrlf
	msg = msg & "WIN+G		Activa Game DVR para grabar la pantalla" & vbcrlf
	msg = msg & "WIN+H		Compartir en las apps Modern para Windows 10" & vbcrlf
	msg = msg & "WIN+I		Abre la configuracion del sistema" & vbcrlf
	msg = msg & "WIN+K		Inicia 'Conectar' para enviar datos a dispositivos" & vbcrlf
	msg = msg & "WIN+L		Bloquea el equipo" & vbcrlf
	msg = msg & "WIN+R		Ejecutar un comando" & vbcrlf
	msg = msg & "WIN+S		Activa Cortana" & vbcrlf
	msg = msg & "WIN+X		Abre el menu de opciones avanzadas" & vbcrlf
	msg = msg & "WIN+TAB		Abre la vista de tareas" & vbcrlf
	msg = msg & "WIN+Flechas	Pega una ventana a la pantalla (Windows Snap)" & vbcrlf
	msg = msg & "WIN+CTRL+D		Crea un nuevo escritorio virtual" & vbcrlf
	msg = msg & "WIN+CTRL+F4	Cierra un escritorio virtual" & vbcrlf
	msg = msg & "WIN+CTRL+Flechas	Cambia de escritorio virtual" & vbcrlf
	msg = msg & "WIN+SHIFT+Flechas	Mueve la ventana actual de un monitor a otro" & vbcrlf
	
	MsgBox msg, vbOkOnly, "DFX Tweaker: Atajos de teclado"
	Call showMenu(0)
End Function

Function restoreMenu()
	cls
	printf "   ____           _                             _   __        "
	printf "  |  _ \ ___  ___| |_ __ _ _   _ _ __ __ _  ___(_) /_/  _ __  "
	printf "  | |_) / _ \/ __| __/ _` | | | | '__/ _` |/ __| |/ _ \| '_ \ "
	printf "  |  _ <  __/\__ \ || (_| | |_| | | | (_| | (__| | (_) | | | |"
	printf "  |_| \_\___||___/\__\__,_|\__,_|_|  \__,_|\___|_|\___/|_| |_|"
	printf ""
	printf " Selecciona una opcion:"
	printf ""
	printf "   1 = Habilitar la telemetria"
	printf "   2 = Habilitar servicios DiagTrack, RetailDemo y Dmwappush"
	printf "   3 = Habilitar tareas programadas que envian datos a Microsoft"
	printf "   4 = Restaurar hosts y acceso a servidores de publicidad de Microsoft"
	printf "   5 = Habilitar Windows Defender Antivirus"
	printf "   6 = Habilitar OneDrive"
	printf ""
	printf "   7 = Habilitar Shadow Copy (VSS) e Instantaneas de Volumen"
	printf "   8 = Habilitar Windows Search + Indexing Service"
	printf "   9 = Habilitar tarea programada del Defragmentador de discos"
	printf "   10 = Habilitar la hibernacion en el sistema"
	printf "   11 = Habilitar Prefetcher + Superfetch"
	printf "   12 = Deshabilitar el tema oscuro"
	printf ""															
	printf "   13 = Habilitar Monitorizacion para Sensores de Tablets con Windows 10"
	printf "   0 = Regresar al menu"
	printf ""
	printl " > "
	RP = scanf
	If Not isNumeric(RP) = True Then
		printf ""
		printf "  Solo se permiten numeros."
		Call restoreMenu()
		Exit Function
	End If
	Select Case RP
		Case 1
			printf ""
			printf " INFO: La opcion de Telemetria se ha restaurado a su valor original"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection\AllowTelemetry", 3, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\DataCollection\AllowTelemetry", 3, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection\AllowTelemetry", 3, "REG_DWORD"
			oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..riencehost.appxmain_31bf3856ad364e35_10.0.10240.16384_none_0ab8ea80e84d4093\f!telemetry.js", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!dss-winrt-telemetry.js", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry.js", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry-event_8ac43a41e5030538", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry-inter_58073761d33f144b", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\MRT\DontOfferThroughWUAU", 0, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat\AITEnable", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\SQMClient\Windows\CEIPEnable", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat\DisableUAR", 0, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata\PreventDeviceMetadataFromNetwork", 0, "REG_DWORD"
		Case 2
			printf ""
			printf " INFO: Se han habilitado los servicios DiagTrack, RetailDemo y Dmwappush"
			oWSH.Run "sc config DiagTrack start=auto"
			oWSH.Run "sc config RetailDemo start=auto"
			oWSH.Run "sc config dmwappushservice start=auto"
			oWSH.Run "sc config WMPNetworkSvc start=auto"
			oWSH.Run "sc config diagnosticshub.standardcollector.service start=auto"
			oWSH.Run "sc start DiagTrack"
			oWSH.Run "sc start RetailDemo"
			oWSH.Run "sc start dmwappushservice"
			oWSH.Run "sc start WMPNetworkSvc"		
			oWSH.Run "sc start diagnosticshub.standardcollector.service"
		Case 3
			printf ""
			printf " INFO: Se han habilitado las tareas programadas que envian datos a Microsoft"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Application Experience\ProgramDataUpdater" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\KernelCeipTask" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\Uploader" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\AppID\SmartScreenSpecific" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\NetTrace\GatherNetworkInfo" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Error Reporting\QueueReporting" & chr(34) & " /ENABLE"			
		Case 4
			hostsFile = oWSH.ExpandEnvironmentStrings("%WinDir%") & "\System32\drivers\etc\hosts"
			If oFSO.FileExists(hostsFile & ".cwd") = True Then
				oFSO.DeleteFile	hostsFile
				oFSO.CopyFile	hostsFile & ".cwd", hostsFile
			Else
				Set F = oFSO.CreateTextFile("C:\Windows\System32\drivers\etc\hosts", True)
					F.WriteLine "127.0.0.1	localhost"
					F.WriteLine "::1		localhost"
					F.WriteLine "127.0.0.1	local"
				F.Close
			End If
			printf ""
			printf " INFO: El fichero hosts se ha restablecido correctamente"
		Case 6
			printf ""
			printf " INFO: Se ha habilitado One Drive correctamente"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 0, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 0, "REG_DWORD"
		Case 5
			printf ""
			printf " INFO: Se ha habilitado Windows Defender Antivirus correctamente"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 0, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 0, "REG_DWORD"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cache Maintenance" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cleanup" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Scheduled Scan" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Verification" & chr(34) & " /ENABLE"
			oWSH.Run "sc config WdNisSvc start=auto"
			oWSH.Run "sc config WinDefend start=auto"	
			oWSH.Run "sc start WdNisSvc"
			oWSH.Run "sc start WinDefend"
		Case 7
			printf ""
			printf " INFO: Se ha habilitado el servicio de VSS (Shadow Copy)"
			oWSH.Run "sc config VSS start=auto"
			oWSH.Run "sc start VSS"
		Case 8
			printf ""
			printf " INFO: Se ha habilitado el servicio de Windows Search + Indexing Service"
			oWSH.Run "sc config WSearch start=auto"
			oWSH.Run "sc start WSearch"
		Case 9
			printf ""
			printf " INFO: Se ha habilitado la tarea programada del defragmentador de discos de Windows"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Defrag\ScheduledDefrag" & chr(34) & " /ENABLE"
		Case 10
			printf ""
			printf " INFO: Hibernacion del sistema activada correctamente"
			oWSH.Run "powercfg -h on"
			oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power\HiberbootEnabled", 1, "REG_DWORD"
		Case 11
			printf ""
			printf " INFO. Se ha habilitado Prefetcher + Superfetch en el registro y en el servicio"
			oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnablePrefetcher", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnableSuperfetch", 1, "REG_DWORD"
			oWSH.Run "sc config SysMain start=auto"
			oWSH.Run "sc start SysMain"
		Case 12
			printf ""
			printf " INFO: Se ha deshabilitado el tema oscuro (Dark Theme)"
			oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"		
		Case 13 
			printf ""
			printf " Info: Se ha habilitado el Sensor preview, comprueba si ya funcionan los sensores de acelerometro y luz."
			oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}\SensorPermissionState", 1,  "REG_DWORD"
                        oWSH.Run "sc start SensorDataService"
			oWSH.Run "sc start SensrSvc"
		Case 0
			MsgBox "Si has restaurado alguna opcion/configuracion, te recomiendo que reinicies el sistema ahora", vbInformation + vbOkOnly, "DFX Tweaker"
			Call showMenu(0)
		Case Else
			printf ""
			printf "  Ese numero no esta disponible."
			Call restoreMenu()
			Exit Function
	End Select
	wait(2)
	Call restoreMenu()
End Function

Function printf(txt)
	WScript.StdOut.WriteLine txt
End Function

Function printl(txt)
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
		printf ""
	Next
End Function

Function ForceConsole()
	If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
		oWSH.Run "cscript //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
		WScript.Quit
	End If
End Function

Function checkW10orW11()
	If getNTversion < 10 Then
		printf "  ERROR: Necesitas ejecutar DFX Tweaker bajo Windows 10 o Windows 11"
		printf ""
		printf "  Pulsa <<Enter>> para salir"
		scanf
		WScript.Quit
	End If
End Function

Function runElevated()
	If isUACRequired Then
		If Not isElevated Then RunAsUAC
	Else
		If Not isAdmin Then
			printf "  ERROR: Necesitas ejecutar DFX Tweaker como Administrador!"
			printf ""
			printf "  Pulsa <<Enter>> para salir"
			scanf
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
		printf ""
		printf "  DFX Tweaker necesita ejecutarse como Administrador..."
		printf "  Espera..."
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

Function getNTversion()
	Set cWin = oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	For Each OS In cWin
		getNTversion = Split(OS.Version,".")(0)
	Next
End Function
currentFolder  = oFSO.GetParentFolderName(WScript.ScriptFullName)
currentBuild   = "Build 1182 "
currentRelease = "Release"
Call ForceConsole()
Call showBanner()
printf " "
Call checkW10orW11()
Call runElevated()
Call updateFinder()
printf "  Cargando el script..."
wait 0.3
Call showMenu(1)

Function showBanner()
	printf" "
	printf"   ____  _______  __  _____ Para Windows 10 y  _  Windows 11 "
	printf"  |  _ \|  ___\ \/ / |_   _|_      _____  __ _| | _____ _ __ "
	printf"  | | | | |_   \  /    | | \ \ /\ / / _ \/ _` | |/ / _ \ '__|"
	printf"  | |_| |  _|  /  \    | |  \ V  V /  __/ (_| |   <  __/ |   "
	printf"  |____/|_|   /_/\_\   |_|   \_/\_/ \___|\__,_|_|\_\___|_| v" & currentVersion
        printf "             Creado por ivandfx        " & currentBuild
	printf " "
	printf "  DFX Tweaker, un fork de AikonCWD Script 5.6"
	printf "  Bajo licencia de Creative Commons 4.0"
End Function

Function showMenu(n)
	wait(n)
	cls
	Call showBanner
	printf "  "
	printf "  Selecciona una opcion:                    44 = Ayuda del menu - ¿Que significa (1X) o (!)?"
	printf "				            55 = Opciones de apagado avanzadas - WAST para DFX Tweaker"
	printf "  1 = Ajustar tweaks de sistema             "
	printf "  2 = Ajustar tweaks de rendimiento"
	printf "  3 = Desinstalar aplicaciones UWP de Windows 10 (1X)"
	printf ""
	printf "  4 = Configurar la telemetria de Windows"
	printf "  5 = Configurar MS OneDrive"
	printf "  6 = Configurar MS Cortana (1X)"
	printf "  7 = Configurar Windows Defender"
	printf "  8 = Configurar Windows Update"
	printf ""
	printf "  9 = Mostrar estado de la activacion de Windows"
	printf "  10 = Mostrar atajos de teclado utiles para Windows"
	printf ""
	printf "  99 = Menu de restauracion"
	printf "  0 = Salir"
	printf ""
	printl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		printf ""
		printf " Solo se permiten numeros."
		Call showMenu(2)
		Exit Function
	End If
	Select Case RP
		Case 1
			Call menuSysTweaks()
		Case 2
			Call menuPerfomance()
		Case 3
			Call menuCleanApps()
		Case 4
			Call menuTelemetry()
		Case 5
			Call menuOneDrive()
		Case 6
			Call menuCortana()
		Case 7
			Call menuWindowsDefender()
		Case 8
			Call menuWindowsUpdate()
		Case 9
			Call menuXPR()
		Case 10
			Call showKeyboardTips()
		Case 99
			Call restoreMenu()
		Case 44
			MsgBox "Las opciones con (1X) solo son compatibles con Windows 10. Las opciones con (!) son irrevertibles o pueden causar problemas. Las opciones que requieren conexion a internet no estan disponibles", vbInformation + vbOkOnly, "DFX Tweaker: Ayuda del menu"
			Call showMenu(0)
		Case 55
			Call shutdownMenu()

		Case 550
			Call menuPowerSSD()		
		Case 0
			cls
			printf ""
			printf"   ____  _______  __  _____                    _		    "
			printf"  |  _ \|  ___\ \/ / |_   _|_      _____  __ _| | _____ _ __ "
			printf"  | | | | |_   \  /    | | \ \ /\ / / _ \/ _` | |/ / _ \ '__|"
			printf"  | |_| |  _|  /  \    | |  \ V  V /  __/ (_| |   <  __/ |   "
			printf"  |____/|_|   /_/\_\   |_|   \_/\_/ \___|\__,_|_|\_\___|_|   "
			printf ""
			printf " 	  Gracias por utilizar DFX Tweaker :D"
			printf "                        ivandfx"
			printf " "
			printf ""
			printf ""
			printf ""
			wait 1
			printf "  El script se está cerrando..."
			wait(2)
			WScript.Quit
		Case Else
			printf ""
			printf " Solo se permiten numeros."
			Call showMenu(2)
			Exit Function
	End Select
End Function

Function menuXPR()
	cls
	On Error Resume Next
	printf ""
	printf " En unos segundos aparecera el estado de tu activacion..."
	wait 0.2
	printf " Recopilando datos de la activacion..."
	wait 2
	oWSH.Run "slmgr.vbs /dli"
	oWSH.Run "slmgr.vbs /xpr"
	Call showMenu (1)
End Function

Function showBannerWAST()
	printf "  __        ___    ____ _____   _____           _              _     _          _ "
	printf "  \ \      / / \  / ___|_   _| | ____|_ __ ___ | |__   ___  __| | __| | ___  __| |"
	printf "   \ \ /\ / / _ \ \___ \ | |   |  _| | '_ ` _ \| '_ \ / _ \/ _` |/ _` |/ _ \/ _` |"
	printf "    \ V  V / ___ \ ___) || |   | |___| | | | | | |_) |  __/ (_| | (_| |  __/ (_| |"
	printf "     \_/\_/_/   \_\____/ |_|   |_____|_| |_| |_|_.__/ \___|\__,_|\__,_|\___|\__,_|"
	printf "                          Windows Advanced Shutdown Tool para DFX Tweaker 1.2"
	printf " "
	End Function

Function shutdownMenu()
	cls
	Call showBannerWAST()
	printf " "
	printf "  Cargando WAST para DFX Tweaker..."
	wait 0.4
	printf "  Espera..."
	wait 3
	cls
	On Error Resume Next
	Call showBannerWAST()
	printf " "
	printf " "
	printf " "
	printf " "
	printf " "
	printf "   ¿Que quieres hacer?:"
	printf "                                            55 = Reiniciar el explorador de Windows"
	printf ""
	printf "  1 = Apagar el equipo"
	printf " "
	printf "  2 = Reiniciar el equipo"
	printf " "
	printf "  3 = Cerrar sesion de este usuario"
	printf " "
	printf "  4 = Ir a opciones avanzadas"
	printf ""
	printf "  5 = Causar un BSOD (Blue Screen Of Death)"
	printf " "
	printf " "
	printf "  0 = Volver al menu principal"
	printf ""
	printl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		printf ""
		printf " Solo se permiten numeros."
		Call shutdownMenu()
		Exit Function
	End If
	Select Case RP
		Case 1
			result = MsgBox ("¿Apagar?", vbYesNo, "WAST Apagado")
Select Case result
    Case vbYes
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "C:\WINDOWS\system32\shutdown.exe -s -t 0"
        Dim objShell
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
		Case 2
						result = MsgBox ("¿Reiniciar?", vbYesNo, "WAST Reinicio")
Select Case result
    Case vbYes
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 0"
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
		Case 3
						result = MsgBox ("¿Cerrar sesion? Los datos no guardados se perderan.", vbYesNo, "WAST Cierre de sesion")
Select Case result
    Case vbYes
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "C:\WINDOWS\system32\shutdown.exe -l"
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
		Case 80
						result = MsgBox ("DFX Tweaker tendra selector de idiomas desde la 1.6 :D", vbAccept, "Nada aqui")
		Call shutdownMenu()
		Case 4
						result = MsgBox ("¿Ir al menu de opciones avanzadas? Esto cerrrara todas las sesiones de todos los usuarios del equipo.", vbYesNo, "WAST Opciones avanzadas")
Select Case result
    Case vbYes
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -o -t 0"
	wait 1
		Call shutdownMenu()
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
		Case 5
						result = MsgBox ("¿Quieres causar un pantallazo azul de la muerte? Asegurate de haber guardado TODOS los datos que estuvieras usando.", vbYesNo, "WAST BSOD")
Select Case result
    Case vbYes
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "taskkill /f /im crss.exe"
	objShell.Run "taskkill /f /im winnit.exe"
	objShell.Run "taskkill /f /im winlogon.exe"
	objShell.Run "taskkill /f /im svchost.exe"
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
		Case 55
						result = MsgBox ("¿Quieres reiniciar el Explorador de Windows?", vbYesNo, "WAST Explorador")
Select Case result
    Case vbYes
	printf = "  Espera..."
		printf " >> Reiniciando el explorador de Windows... espera 5 segundos!"
		oWSH.Run "taskkill.exe /F /IM explorer.exe"
		wait(5)
		oWSH.Run "explorer.exe"
		Call shutdownMenu()
    Case vbNo
	cls
	printf = "  Espera..."
		wait 1
		Call shutdownMenu()
End Select
Case 0
		cls
		printf "  Volviendo a DFX Tweaker..."
		wait 0.3
		printf "  Espera..."
		wait 2.7
		Call showMenu(0)
	End Select
End Function

Function menuSysTweaks()
	cls
	On Error Resume Next
	printf ""
	printf "   _____                    _              _      _       _     _                       "
	printf "  |_   _|_      _____  __ _| | _____    __| | ___| |  ___(_)___| |_ ___ _ __ ___   __ _ "
	printf "    | | \ \ /\ / / _ \/ _` | |/ / __|  / _` |/ _ \ | / __| / __| __/ _ \ '_ ` _ \ / _` |"
	printf "    | |  \ V  V /  __/ (_| |   <\__ \ | (_| |  __/ | \__ \ \__ \ ||  __/ | | | | | (_| |"
	printf "    |_|   \_/\_/ \___|\__,_|_|\_\___/  \__,_|\___|_| |___/_|___/\__\___|_| |_| |_|\__,_|"
	printf ""
	printl " # Deshabilitar 'Acceso Rapido' como opcion por defecto en Explorer? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\LaunchTo", 1, "REG_DWORD"
	Else
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\LaunchTo", 2, "REG_DWORD"
	End If
	printl " # Crear icono 'Modo Dios' en el Escritorio? (s/n) > "
	If LCase(scanf) = "s" Then
		godFolder = oWSH.SpecialFolders("Desktop") & "\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
		If oFSO.FolderExists(godFolder) = False Then oFSO.CreateFolder(godFolder)
	Else
		godFolder = oWSH.SpecialFolders("Desktop") & "\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
		If oFSO.FolderExists(godFolder) = True Then oFSO.DeleteFolder(godFolder)	
	End If
	printl " # Habilitar el tema oscuro de Windows 'Dark Theme'? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"		
	End If
	printl " # Mostrar icono 'Mi PC' en el Escritorio? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", 1, "REG_DWORD"
	End If
	printl " # Mostrar siempre la extension para archivos conocidos? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt", 1, "REG_DWORD"
	End If
	printl " # Deshabilitar 'Lock Screen'? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Personalization\NoLockScreen", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\Software\Policies\Microsoft\Windows\System\DisableLogonBackgroundImage", 1, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Personalization\NoLockScreen", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\Software\Policies\Microsoft\Windows\System\DisableLogonBackgroundImage", 0, "REG_DWORD"
	End If
	printl " # Forzar 'Vista Clasica' en el Panel de Control? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\ForceClassicControlPanel", 1, "REG_DWORD"
	Else
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\ForceClassicControlPanel", 0, "REG_DWORD"
	End If
	printl " # Deshabilitar 'Reporte de Errores' de Windows? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting\Disabled", 1, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting\Disabled", 0, "REG_DWORD"
	End If
	printl " # Abrir cmd.exe al pulsar Win+U? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\utilman.exe\Debugger", "cmd.exe", "REG_SZ"
	Else
		oWSH.RegDelete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\utilman.exe\Debugger"
	End If
	printl " # Habilitar/Deshabilitar el control de cuentas de usuario UAC? (s/n) > "
	If LCase(scanf) = "s" Then
		printf ""
		printf " Ahora se abrira una ventana..."
		printf " Mueve la barra vertical hasta el nivel mas bajo"
		printf " Acepta los cambios y reinicia el ordenador"
		wait(2)
		printf ""
		printf " > Executing UserAccountControlSettings.exe"
		oWSH.Run "UserAccountControlSettings.exe"
		printf ""
	End If
	printl " # Habilitar/Deshabilitar el inicio de sesion sin password? (s/n) > "
	If LCase(scanf) = "s" Then
		printf ""
		printf " Ahora se abrira una ventana..."
		printf " Desmarca la opcion: Los usuarios deben escribir su nombre y password para usar el equipo"
		printf " Acepta los cambios y reinicia el ordenador"
		wait(2)
		printf ""
		printf " > Executing control userpasswords2"
		oWSH.Run "control userpasswords2"
		printf ""
	End If
	printl " # Utilizar control de volumen clasico? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\MTCUVC\EnableMtcUvc", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\MTCUVC\EnableMtcUvc", 1, "REG_DWORD"
	End If
	printl " # Utilizar el centro de notificaciones clasico? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\ImmersiveShell\UseActionCenterExperience", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\ImmersiveShell\UseActionCenterExperience", 1, "REG_DWORD"
	End If
	printl " # Utilizar el visor de fotos clasico? (s/n) > "
	If LCase(scanf) = "s" Then
		oWEB.Open "GET", "ivandfxlink", False
		oWEB.Send
		wait(1)
		Set F = oFSO.CreateTextFile(currentFolder & "\photoview.reg")
			F.Write oWEB.ResponseText
		F.Close
		wait(1)
		oWSH.Run "reg import " & currentFolder & "\photoview.reg"
	End If
	printf ""
	printf " >> Reiniciando el explorador de Windows... espera 5 segundos!"
	oWSH.Run "taskkill.exe /F /IM explorer.exe"
	wait(5)
	oWSH.Run "explorer.exe"
	wait (3)
	printf ""
	printf " Todos los tweaks de sistema se han aplicado correctamente"
	wait (1)
	Call showMenu(2)
End Function

Function menuOneDrive()
	cls
	On Error Resume Next	
	printf "   __  __ ____     ___             ____       _           "
	printf "  |  \/  / ___|   / _ \ _ __   ___|  _ \ _ __(_)_   _____ "
	printf "  | |\/| \___ \  | | | | '_ \ / _ \ | | | '__| \ \ / / _ \"
	printf "  | |  | |___) | | |_| | | | |  __/ |_| | |  | |\ V /  __/"
	printf "  |_|  |_|____/   \___/|_| |_|\___|____/|_|  |_| \_/ \___|"                                                               
	printf ""
	printf " Selecciona una opcion:"
	printf ""
	printf "  1 = Deshabilitar MS OneDrive"
	printf "  2 = Habilitar MS OneDrive"
	printf ""
	printf "  0 = Volver al menu principal"
	printf ""
	printl "  > "
	Select Case scanf
		Case "1"
			printf ""
			printf " Deshabilitando OneDrive..."
			wait(1)
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
			printf ""
			printf " INFO: OneDrive deshabilitado correctamente"
			wait(2)
		Case "2"
			printf ""
			printf " Habilitando OneDrive..."
			wait(1)
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
			printf ""
			printf " INFO: OneDrive habilitado correctamente"
			wait(2)
		Case "0"
			Call showMenu(0)
		Case Else
			Call menuOneDrive()
	End Select
	Call menuOneDrive()
End Function

Function menuCortana()
	cls
	On Error Resume Next
	printf "   __  __ ____     ____           _                    "
	printf "  |  \/  / ___|   / ___|___  _ __| |_ __ _ _ __   __ _ "
	printf "  | |\/| \___ \  | |   / _ \| '__| __/ _` | '_ \ / _` |"
	printf "  | |  | |___) | | |__| (_) | |  | || (_| | | | | (_| |"
	printf "  |_|  |_|____/   \____\___/|_|   \__\__,_|_| |_|\__,_|"                                                         
	printf " "
	printf " Selecciona una opcion:"
	printf ""
	printf "  1 = Deshabilitar MS Cortana"
	printf "  2 = Habilitar MS Cortana"
	printf ""
	printf "  0 = Volver al menu principal"
	printf ""
	printl "  > "
	Select Case scanf
		Case "1"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\AllowCortana", 0, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\CortanaEnabled", 0, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\SearchboxTaskbarMode", 0, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\BingSearchEnabled", 0, "REG_DWORD"
			printf ""
			printf " >> Reiniciando el explorador de Windows... espera 5 segundos!"
			oWSH.Run "taskkill.exe /F /IM explorer.exe"
			wait(5)
			oWSH.Run "explorer.exe"
		Case "2"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\AllowCortana", 1, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\CortanaEnabled", 1, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\SearchboxTaskbarMode", 1, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\BingSearchEnabled", 1, "REG_DWORD"
			printf ""
			printf " >> Reiniciando el explorador de Windows... espera 5 segundos!"
			oWSH.Run "taskkill.exe /F /IM explorer.exe"
			wait(5)
			oWSH.Run "explorer.exe"
		Case "0"
			Call showMenu(0)
		Case Else
			Call menuCortana()
	End Select
	Call menuCortana()
End Function

Function menuTelemetry()
	cls
	On Error Resume Next
	printf "   _____    _                     _        __      "
	printf "  |_   _|__| | ___ _ __ ___   ___| |_ _ __/_/ __ _ "
	printf "    | |/ _ \ |/ _ \ '_ ` _ \ / _ \ __| '__| |/ _` |"
	printf "    | |  __/ |  __/ | | | | |  __/ |_| |  | | (_| |"
	printf "    |_|\___|_|\___|_| |_| |_|\___|\__|_|  |_|\__,_|"
	printf ""
	printf " Selecciona una opcion:"
	printf ""
	printf "  1 = Deshabilitar TODA la telemetria"
	printf "  Habilita la telemetria en el Menu de restauracion"
	printf ""
	printf "  0 = Volver al menu principal"
	printf ""
	printl "  > "
	Select Case scanf
		Case "1"
			printf ""
			printf " Aplicando parches para deshabilitar la telemetria (5 segundos)..."
			printf " Deshabilitando la telemetria usando el registro..."
			wait 5
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection\AllowTelemetry", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\DataCollection\AllowTelemetry", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection\AllowTelemetry", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..riencehost.appxmain_31bf3856ad364e35_10.0.10240.16384_none_0ab8ea80e84d4093\f!telemetry.js", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!dss-winrt-telemetry.js", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry.js", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry-event_8ac43a41e5030538", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry-inter_58073761d33f144b", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\MRT\DontOfferThroughWUAU", 1, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat\AITEnable", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\SQMClient\Windows\CEIPEnable", 0, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat\DisableUAR", 1, "REG_DWORD"
				oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata\PreventDeviceMetadataFromNetwork", 1, "REG_DWORD"		
			printf ""
			printf " INFO: Telemetria deshabilitada correctamente"
			pathLOG = oWSH.ExpandEnvironmentStrings("%ProgramData%") & "\Microsoft\Diagnosis\ETLLogs\AutoLogger\AutoLogger-Diagtrack-Listener.etl"
			printf ""
			printf " Borrando DiagTrack Log..."
			wait(1)
				If oFSO.FileExists(pathLOG) Then oFSO.DeleteFile(pathLOG)
				oWSH.Run "cmd /C echo " & chr(34) & chr(34) & " > " & pathLOG
			printf ""
			printf " INFO: DiagTrack Log borrado correctamente"
			printf ""
			printf " Deshabilitando servicios de seguimiento..."
			wait(1)
				oWSH.Run "sc stop TrkWks"
				oWSH.Run "sc stop DiagTrack"
				oWSH.Run "sc stop RetailDemo"
				oWSH.Run "sc stop WMPNetworkSvc"
				oWSH.Run "sc stop dmwappushservice"
				oWSH.Run "sc stop diagnosticshub.standardcollector.service"
				oWSH.Run "sc config TrkWks start=disabled"
				oWSH.Run "sc config DiagTrack start=disabled"
				oWSH.Run "sc config RetailDemo start=disabled"
				oWSH.Run "sc config WMPNetworkSvc start=disabled"
				oWSH.Run "sc config dmwappushservice start=disabled"
				oWSH.Run "sc config diagnosticshub.standardcollector.service start=disabled"
			printf ""
			printf " INFO: Servicios de seguimiento deshabilitados"
			printf ""			
			printf " Deshabilitando tareas programadas que envian datos a Microsoft..."	
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Application Experience\ProgramDataUpdater" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\KernelCeipTask" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\Uploader" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\AppID\SmartScreenSpecific" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\NetTrace\GatherNetworkInfo" & chr(34) & " /DISABLE"
				oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Error Reporting\QueueReporting" & chr(34) & " /DISABLE"
			printf ""
			printf " INFO: Tareas programadas de seguimiento deshabilitadas"
			'printf ""
			'printf " Descargando listado actualizado para bloquear los servidores de publicidad de Microsoft..."
			'wait(1)
			'hostsFile = oWSH.ExpandEnvironmentStrings("%WinDir%") & "\System32\drivers\etc\hosts"
			'If oFSO.FileExists(hostsFile & ".cwd") = False Then
			'	oFSO.CopyFile hostsFile, hostsFile & ".cwd"
			'	Set F = oFSO.OpenTextFile(hostsFile, 8, True)
			'		F.Write oWEB.ResponseText
			'	F.Close
			'	Set F = oFSO.OpenTextFile(hostsFile, 8, True)
			'		F.WriteLine "#Antimalware"
			'		F.WriteLine "0.0.0.0 tracking.opencandy.com.s3.amazonaws.com"
			'		F.WriteLine "0.0.0.0 media.opencandy.com"
			'		F.WriteLine "0.0.0.0 cdn.opencandy.com"
			'		F.WriteLine "0.0.0.0 tracking.opencandy.com"
			'		F.WriteLine "0.0.0.0 api.opencandy.com"
			'		F.WriteLine "0.0.0.0 api.recommendedsw.com"
			'		F.WriteLine "0.0.0.0 installer.betterinstaller.com"
			'		F.WriteLine "0.0.0.0 installer.filebulldog.com"
			'		F.WriteLine "0.0.0.0 d3oxtn1x3b8d7i.cloudfront.net"
			'		F.WriteLine "0.0.0.0 inno.bisrv.com"
			'		F.WriteLine "0.0.0.0 nsis.bisrv.com"
			'		F.WriteLine "0.0.0.0 cdn.file2desktop.com"
			'		F.WriteLine "0.0.0.0 cdn.goateastcach.us"
			'		F.WriteLine "0.0.0.0 cdn.guttastatdk.us"
			'		F.WriteLine "0.0.0.0 cdn.inskinmedia.com"
			'		F.WriteLine "0.0.0.0 cdn.insta.oibundles2.com"
			'		F.WriteLine "0.0.0.0 cdn.insta.playbryte.com"
			'		F.WriteLine "0.0.0.0 cdn.llogetfastcach.us"
			'		F.WriteLine "0.0.0.0 cdn.montiera.com"
			'		F.WriteLine "0.0.0.0 cdn.msdwnld.com"
			'		F.WriteLine "0.0.0.0 cdn.mypcbackup.com"
			'		F.WriteLine "0.0.0.0 cdn.ppdownload.com"
			'		F.WriteLine "0.0.0.0 cdn.riceateastcach.us"
			'		F.WriteLine "0.0.0.0 cdn.shyapotato.us"
			'		F.WriteLine "0.0.0.0 cdn.solimba.com"
			'		F.WriteLine "0.0.0.0 cdn.tuto4pc.com"
			'		F.WriteLine "0.0.0.0 cdn.appround.biz"
			'		F.WriteLine "0.0.0.0 cdn.bigspeedpro.com"
			'		F.WriteLine "0.0.0.0 cdn.bispd.com"
			'		F.WriteLine "0.0.0.0 cdn.bisrv.com"
			'		F.WriteLine "0.0.0.0 cdn.cdndp.com"
			'		F.WriteLine "0.0.0.0 cdn.download.sweetpacks.com"
			'		F.WriteLine "0.0.0.0 cdn.dpdownload.com"
			'		F.WriteLine "0.0.0.0 cdn.visualbee.net"
			'		F.WriteLine "#Telemetry"
			'		F.WriteLine "0.0.0.0 vortex.data.microsoft.com"
			'		F.WriteLine "0.0.0.0 vortex-win.data.microsoft.com"
			'		F.WriteLine "0.0.0.0 telecommand.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 telecommand.telemetry.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 oca.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 oca.telemetry.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 sqm.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 sqm.telemetry.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 watson.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 watson.telemetry.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 redir.metaservices.microsoft.com"
			'		F.WriteLine "0.0.0.0 choice.microsoft.com"
			'		F.WriteLine "0.0.0.0 choice.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 df.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 wes.df.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 reports.wes.df.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 services.wes.df.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 sqm.df.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 watson.ppe.telemetry.microsoft.com"
			'		F.WriteLine "0.0.0.0 telemetry.appex.bing.net"
			'		F.WriteLine "0.0.0.0 telemetry.urs.microsoft.com"
			'		F.WriteLine "0.0.0.0 telemetry.appex.bing.net:443"
			'		F.WriteLine "0.0.0.0 settings-sandbox.data.microsoft.com"
			'		F.WriteLine "0.0.0.0 vortex-sandbox.data.microsoft.com"
			'		F.WriteLine "0.0.0.0 survey.watson.microsoft.com"
			'		F.WriteLine "0.0.0.0 watson.live.com"
			'		F.WriteLine "0.0.0.0 watson.microsoft.com"
			'		F.WriteLine "0.0.0.0 statsfe2.ws.microsoft.com"
			'		F.WriteLine "0.0.0.0 corpext.msitadfs.glbdns2.microsoft.com"
			'		F.WriteLine "0.0.0.0 compatexchange.cloudapp.net"
			'		F.WriteLine "0.0.0.0 cs1.wpc.v0cdn.net"
			'		F.WriteLine "0.0.0.0 a-0001.a-msedge.net"
			'		F.WriteLine "0.0.0.0 statsfe2.update.microsoft.com.akadns.net"
			'		F.WriteLine "0.0.0.0 sls.update.microsoft.com.akadns.net"
			'		F.WriteLine "0.0.0.0 fe2.update.microsoft.com.akadns.net"
			'		F.WriteLine "0.0.0.0 diagnostics.support.microsoft.com"
			'		F.WriteLine "0.0.0.0 corp.sts.microsoft.com"
			'		F.WriteLine "0.0.0.0 statsfe1.ws.microsoft.com"
			'		F.WriteLine "0.0.0.0 pre.footprintpredict.com"
			'		F.WriteLine "0.0.0.0 i1.services.social.microsoft.com"
			'		F.WriteLine "0.0.0.0 i1.services.social.microsoft.com.nsatc.net"
			'		F.WriteLine "0.0.0.0 feedback.windows.com"
			'		F.WriteLine "0.0.0.0 feedback.microsoft-hohm.com"
			'		F.WriteLine "0.0.0.0 feedback.search.microsoft.com"
			'	F.Close
			'	printf ""
			'	printf " INFO: Fichero HOSTS escrito correctamente"
			'End If
			'wait(2)
		Case "0"
			Call showMenu(0)
		Case Else
			Call menuTelemetry()
	End Select
	Call menuTelemetry()
End Function

Function menuWindowsDefender()
	cls
	On Error Resume Next
	printf "   __  __ ____    ____        __                _           "
	printf "  |  \/  / ___|  |  _ \  ___ / _| ___ _ __   __| | ___ _ __ "
	printf "  | |\/| \___ \  | | | |/ _ \ |_ / _ \ '_ \ / _` |/ _ \ '__|"
	printf "  | |  | |___) | | |_| |  __/  _|  __/ | | | (_| |  __/ |   "
	printf "  |_|  |_|____/  |____/ \___|_|  \___|_| |_|\__,_|\___|_|   "
	printf ""
	printf " Selecciona una opcion:"
	printf ""
	printf "  1 = Deshabilitar Windows Defender"
	printf "  2 = Habilitar Windows Defender"
	printf ""
	printf "  0 = Volver al menu principal"
	printf ""
	printl "  > "
	Select Case scanf
		Case "1"
			printf ""
			printf " Deshabilitando Windows Defender..."
			wait(1)
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
			printf ""
			printf " INFO: Windows Defender deshabilitado correctamente"
			wait(3)
		Case "2"
			printf ""
			printf " Habilitando Windows Defender..."
			wait(1)
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
			printf ""
			printf " INFO: Windows Defender habilitado correctamente"
			wait(2)
		Case "0"
			Call showMenu(0)
		Case Else
			Call menuWindowsDefender()
	End Select
	Call menuWindowsDefender()
End Function

Function menuWindowsUpdate()
	cls
	On Error Resume Next
	printf "  __        ___           _                     _   _           _       _       "
	printf "  \ \      / (_)_ __   __| | _____      _____  | | | |_ __   __| | __ _| |_ ___ "
	printf "   \ \ /\ / /| | '_ \ / _` |/ _ \ \ /\ / / __| | | | | '_ \ / _` |/ _` | __/ _ \"
	printf "    \ V  V / | | | | | (_| | (_) \ V  V /\__ \ | |_| | |_) | (_| | (_| | ||  __/"
	printf "     \_/\_/  |_|_| |_|\__,_|\___/ \_/\_/ |___/  \___/| .__/ \__,_|\__,_|\__\___|"
	printf "                                                     |_|                        "
	printf ""
	printf "Para activar Windows Update, selecciona <<n>> en cada opcion"
	printf ""
	printl " # Deshabilitar 'Windows Update Service'? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\DeferUpgrade", 1, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\AUOptions", 2, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 1, "REG_DWROD"
		oWSH.Run "sc stop wuauserv"
		oWSH.Run "sc config wuauserv start=disabled"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\DeferUpgrade", 0, "REG_DWORD"
		oWSH.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\AUOptions"
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWROD"
		oWSH.Run "sc config wuauserv start=auto"
		oWSH.Run "sc start wuauserv"
	End If
	printl " # Deshabilitar 'Windows Update Sharing'? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config\DownloadMode", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config\DODownloadMode", 0, "REG_DWORD"
		oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\SystemSettingsDownloadMode", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config\DownloadMode", 3, "REG_DWORD"
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config\DODownloadMode", 3, "REG_DWORD"
		oWSH.RegDelete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\SystemSettingsDownloadMode"
	End If
	printl " # Deshabilitar 'Windows Update App'? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsStore\WindowsUpdate\AutoDownload", 2, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsStore\WindowsUpdate\AutoDownload", 4, "REG_DWORD"
	End If
	printl " # Deshabilitar 'Windows Update Driver'? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\DriverSearching\DontSearchWindowsUpdate", 1, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\DriverSearching\DontSearchWindowsUpdate", 0, "REG_DWORD"
	End If
	printf ""
	printf "Todos los tweaks de Windows Update se han aplicado correctamente"
	Call showMenu(2)
End Function

Function menuPerfomance()
	cls
	On Error Resume Next
	printf "   _____                    _              _                           _ _           _            _        "
	printf "  |_   _|_      _____  __ _| | _____    __| | ___   _ __ ___ _ __   __| (_)_ __ ___ (_) ___ _ __ | |_ ___  "
	printf "    | | \ \ /\ / / _ \/ _` | |/ / __|  / _` |/ _ \ | '__/ _ \ '_ \ / _` | | '_ ` _ \| |/ _ \ '_ \| __/ _ \ "
	printf "    | |  \ V  V /  __/ (_| |   <\__ \ | (_| |  __/ | | |  __/ | | | (_| | | | | | | | |  __/ | | | || (_) |"
	printf "    |_|   \_/\_/ \___|\__,_|_|\_\___/  \__,_|\___| |_|  \___|_| |_|\__,_|_|_| |_| |_|_|\___|_| |_|\__\___/ "                                                                     
	printf ""
	printl " # Acelerar el cierre de aplicaciones y servicios? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKCU\Control Panel\Desktop\WaitToKillAppTimeout", 1000, "REG_SZ"
		oWSH.RegWrite "HKCU\Control Panel\Desktop\AutoEndTasks", 1, "REG_SZ"
		oWSH.RegWrite "HKCU\Control Panel\Desktop\HungAppTimeout", 1000, "REG_SZ"
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\WaitToKillServiceTimeout", 1000, "REG_SZ"
		oWSH.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize\StartupDelayInMSec", 0, "REG_DWORD"
	End If
	printl " # Deshabilitar servicios: BitLocker, Cifrado y OfflineFiles? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.Run "sc config BDESVC start=disabled"
		oWSH.Run "sc config EFS start=disabled"
		oWSH.Run "sc config CscService start=disabled"
		oWSH.Run "sc stop BDESVC"
		oWSH.Run "sc stop EFS"
		oWSH.Run "sc stop CscService"
	Else
		oWSH.Run "sc config BDESVC start=auto"
		oWSH.Run "sc config EFS start=auto"
		oWSH.Run "sc config CscService start=auto"
		oWSH.Run "sc start BDESVC"
		oWSH.Run "sc start EFS"
		oWSH.Run "sc start CscService"
	End If
	printf ""
	printf " >> No utilizar si usas un portatil o WiFi <<"
	printf ""
	printl " # Deshabilitar servicios WiFi? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.Run "sc config WlanSvc start=disabled"
		oWSH.Run "sc stop WlanSvc"
	Else
		oWSH.Run "sc config WlanSvc start=auto"
		oWSH.Run "sc start WlanSvc"
	End If
	printl " # Ejecutar limpiador de Windows. Libera espacio y borrar Windows.old (s/n) > "
	If LCase(scanf) = "s" Then	
		printf ""
		printf " Ahora se ejecutara una ventana..."
		printf " Marca las opciones deseadas de limpieza"
		printf " Acepta los cambios y reinicia el ordenador"
		wait(2)
		printf ""
		printf " > Executing cleanmgr.exe"
		oWSH.Run "cleanmgr.exe"
		printf ""
	End If
	printl " # Instalar/Desinstalar caracteristicas adicionales de Windows (s/n) > "
	If LCase(scanf) = "s" Then
		printf ""
		printf " Ahora se ejecutara una ventana..."
		printf " Marca/Desmarca las opciones deseadas"
		printf " Acepta los cambios y reinicia el ordenador"
		wait(2)
		printf ""
		printf " > Executing optionalfeatures.exe"
		oWSH.Run "optionalfeatures.exe"
		printf ""
	End If
	printl " # Cambiar la configuracion de la compresion de ficheros? (tarda un poco!) (s/n) > "
	If LCase(scanf) = "s" Then
		printl " -> Deshabilitar la compresion de ficheros en el disco duro principal? (s/n) > "
		If LCase(scanf) = "s" Then
			oWSH.Run "compact /CompactOs:never"
		Else
			oWSH.Run "compact /CompactOs:always"
		End If
		wait(3)
	End If
	printl " # Habilitar el 100% del ancho de banda para el sistema? (s/n) > "
	If LCase(scanf) = "s" Then
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Psched\Psched", 0, "REG_DWORD"
	Else
		oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\Psched\Psched", 20, "REG_DWORD"
	End If
	printf ""
	printf " Todos los tweaks de sistema se han aplicado correctamente"	
	showMenu(3)
End Function

Function menuPowerSSD()
	cls
	On Error Resume Next
	printf " _____ _____ ____  _____     _   _       _             "
	printf "|   __|   __|    \|     |___| |_|_|_____|_|___ ___ ___ "
	printf "|__   |__   |  |  |  |  | . |  _| |     | |- _| -_|  _|"
	printf "|_____|_____|____/|_____|  _|_| |_|_|_|_|_|___|___|_|  "
	printf "                        |_|                            "
	printf ""
	printf " Felicidades, has descubierto la opcion oculta: Optimizar SSD"
	printf " "
	printf " Esta opcion se elimino porque podia causar problemas serios a usuarios con HDD"
	printf "  y causar inestabilidad en ciertos SSD"
	printf " "
	printf " Asi que utiliza esta opcion bajo TU PROPIO RIESGO, te he avisado :P"
	printf ""
	printf ""
	printf " Este script va a modificar las siguientes configuraciones:"
	printf ""
	printf "  > Habilitar TRIM"
	printf "  > Deshabilitar VSS (Shadow Copy)"
	printf "  > Deshabilitar Windows Search"
	printf "  > Deshabilitar Servicios de Indexacion"
	printf "  > Deshabilitar defragmentador de discos"
	printf "  > Deshabilitar hibernacion del sistema"
	printf "  > Deshabilitar Prefetcher + Superfetch"
	printf "  > Deshabilitar ClearPageFileAtShutdown + LargeSystemCache"
	printf ""
	printl "  # Deseas continuar y aplicar los cambios? (s/n) "	
	If scanf = "s" Then
		printf ""
		oWSH.Run "fsutil behavior set disabledeletenotify 0"
		printf " # TRIM habilitado"
		wait(1)
		oWSH.Run "vssadmin Delete Shadows /All /Quiet"
		oWSH.Run "sc stop VSS"
		oWSH.Run "sc config VSS start=disabled"
		printf " # Shadow Copy eliminada y deshabilitada"
		wait(1)
		oWSH.Run "sc stop WSearch"
		oWSH.Run "sc config WSearch start=disabled"
		printf " # Windows Search + Indexing Service deshabilitados"
		wait(1)
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\OptimizeComplete", "No"
		oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\Enable", "N"
		oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Defrag\ScheduledDefrag" & chr(34) & " /DISABLE"
		printf " # Defragmentador de disco deshabilitado"
		wait(1)		
		oWSH.Run "powercfg -h off"
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power\HiberbootEnabled", 0, "REG_DWORD"
		printf " # Hibernacion deshabilitada"
		wait(1)
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnablePrefetcher", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnableSuperfetch", 0, "REG_DWORD"
		oWSH.Run "sc stop SysMain"
		oWSH.Run "sc config SysMain start=disabled"
		printf " # Prefetcher + Superfetch deshabilitados"
		wait(1)
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\ClearPageFileAtShutdown", 0, "REG_DWORD"
		oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\LargeSystemCache", 0, "REG_DWORD"
		printf " # ClearPageFileAtShutdown + LargeSystemCache deshabilitados"
		wait(1)
		printf ""
		printf " INFO: Felicidades, acabas de prolongar la vida y el rendimiento de tu SSD"
		printf "       Es recomendable que reinicies tu PC para aplicar cambios..."
	Else
		printf ""
		printf " Operacion cancelada."
	End If
	Call showMenu(3)
End Function

Function menuCleanApps()
	cls
	On Error Resume Next
	printf "      _                      _   ___        ______  "
	printf "     / \   _ __  _ __  ___  | | | \ \      / /  _ \ "
	printf "    / _ \ | '_ \| '_ \/ __| | | | |\ \ /\ / /| |_) |"
	printf "   / ___ \| |_) | |_) \__ \ | |_| | \ V  V / |  __/ "
	printf "  /_/   \_\ .__/| .__/|___/  \___/   \_/\_/  |_|    "
	printf "          |_|   |_|                                 "
	printf " "
	printf " Este script va a desinstalar el siguiente listado de Apps:"
	printf ""
	printf "  > Bing, Zune, Skype, XboxApp"
	printf "  > Getstarted, Messagin, 3D Builder"
	printf "  > Windows Maps, Phone, Camera, Alarms, People"
	printf "  > Windows Communications Apps, Sound Recorder"
	printf "  > Microsoft Office Hub, Office Sway, OneNote"
	printf "  > Solitaire Collection, CandyCrushSaga"
	printf ""
	printl " La opcion NO es reversible. Deseas continuar? (s/n) "
	
	If scanf = "s" Then
		oWSH.Run "powershell get-appxpackage -Name *Bing* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *Zune* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *XboxApp* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *OneNote* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *SkypeApp* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *3DBuilder* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *Getstarted* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *Microsoft.People* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *MicrosoftOfficeHub* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *MicrosoftSolitaireCollection* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *WindowsCamera* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *WindowsAlarms* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *WindowsMaps* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *WindowsPhone* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *WindowsSoundRecorder* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *windowscommunicationsapps* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *CandyCrushSaga* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *Messagin* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *ConnectivityStore* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *CommsPhone* | Remove-AppxPackage", 1, True
		oWSH.Run "powershell get-appxpackage -Name *Office.Sway* | Remove-AppxPackage", 1, True
		printf ""
		printf " > Las Apps se han desinstalado correctamente..."
	Else
		printf ""
		printf " > Operacion cancelada."
	End If
	wait(1)
	Call showMenu(2)
End Function

Function showKeyboardTips()
	msg = msg & "WIN+A		Abre el centro de actividades" & vbcrlf
	msg = msg & "WIN+C		Activa el reconocimiento de voz de Cortana" & vbcrlf
	msg = msg & "WIN+D		Muestra el escritorio" & vbcrlf
	msg = msg & "WIN+E		Abre el explorador de Windows" & vbcrlf
	msg = msg & "WIN+G		Activa Game DVR para grabar la pantalla" & vbcrlf
	msg = msg & "WIN+H		Compartir en las apps Modern para Windows 10" & vbcrlf
	msg = msg & "WIN+I		Abre la configuracion del sistema" & vbcrlf
	msg = msg & "WIN+K		Inicia 'Conectar' para enviar datos a dispositivos" & vbcrlf
	msg = msg & "WIN+L		Bloquea el equipo" & vbcrlf
	msg = msg & "WIN+R		Ejecutar un comando" & vbcrlf
	msg = msg & "WIN+S		Activa Cortana" & vbcrlf
	msg = msg & "WIN+X		Abre el menu de opciones avanzadas" & vbcrlf
	msg = msg & "WIN+TAB		Abre la vista de tareas" & vbcrlf
	msg = msg & "WIN+Flechas	Pega una ventana a la pantalla (Windows Snap)" & vbcrlf
	msg = msg & "WIN+CTRL+D		Crea un nuevo escritorio virtual" & vbcrlf
	msg = msg & "WIN+CTRL+F4	Cierra un escritorio virtual" & vbcrlf
	msg = msg & "WIN+CTRL+Flechas	Cambia de escritorio virtual" & vbcrlf
	msg = msg & "WIN+SHIFT+Flechas	Mueve la ventana actual de un monitor a otro" & vbcrlf
	
	MsgBox msg, vbOkOnly, "DFX Tweaker: Atajos de teclado"
	Call showMenu(0)
End Function

Function restoreMenu()
	cls
	printf "   ____           _                             _   __        "
	printf "  |  _ \ ___  ___| |_ __ _ _   _ _ __ __ _  ___(_) /_/  _ __  "
	printf "  | |_) / _ \/ __| __/ _` | | | | '__/ _` |/ __| |/ _ \| '_ \ "
	printf "  |  _ <  __/\__ \ || (_| | |_| | | | (_| | (__| | (_) | | | |"
	printf "  |_| \_\___||___/\__\__,_|\__,_|_|  \__,_|\___|_|\___/|_| |_|"
	printf ""
	printf " Selecciona una opcion:"
	printf ""
	printf "   1 = Habilitar Telemetria"
	printf "   2 = Habilitar servicios DiagTrack, RetailDemo y Dmwappush"
	printf "   3 = Habilitar tareas programadas que envian datos a Microsoft"
	printf "   4 = Restaurar hosts y acceso a servidores de publicidad de Microsoft"
	printf "   5 = Habilitar Windows Defender Antivirus"
	printf "   6 = Habilitar OneDrive"
	printf ""
	printf "   7 = Habilitar Shadow Copy (VSS) e Instantaneas de Volumen"
	printf "   8 = Habilitar Windows Search + Indexing Service"
	printf "   9 = Habilitar tarea programada del Defragmentador de discos"
	printf "  10 = Habilitar la hibernacion en el sistema"
	printf "  11 = Habilitar Prefetcher + Superfetch"
	printf "  12 = Deshabilitar el tema oscuro (Dark Theme)"
	printf ""															
	printf "  13 = Habilitar Monitorizacion para Sensores de Tablets con Windows 10"
	printf "   0 = Regresar al menu principal"
	printf ""
	printl " > "
	RP = scanf
	If Not isNumeric(RP) = True Then
		printf ""
		printf "  Solo se permiten numeros."
		Call restoreMenu()
		Exit Function
	End If
	Select Case RP
		Case 1
			printf ""
			printf " INFO: La opcion de Telemetria se ha restaurado a su valor original"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection\AllowTelemetry", 3, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\DataCollection\AllowTelemetry", 3, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection\AllowTelemetry", 3, "REG_DWORD"
			oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..riencehost.appxmain_31bf3856ad364e35_10.0.10240.16384_none_0ab8ea80e84d4093\f!telemetry.js", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!dss-winrt-telemetry.js", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry.js", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry-event_8ac43a41e5030538", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\COMPONENTS\DerivedData\Components\amd64_microsoft-windows-c..lemetry.lib.cortana_31bf3856ad364e35_10.0.10240.16384_none_40ba2ec3d03bceb0\f!proactive-telemetry-inter_58073761d33f144b", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\MRT\DontOfferThroughWUAU", 0, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat\AITEnable", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\SQMClient\Windows\CEIPEnable", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat\DisableUAR", 0, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata\PreventDeviceMetadataFromNetwork", 0, "REG_DWORD"
		Case 2
			printf ""
			printf " INFO: Se han habilitado los servicios DiagTrack, RetailDemo y Dmwappush"
			oWSH.Run "sc config DiagTrack start=auto"
			oWSH.Run "sc config RetailDemo start=auto"
			oWSH.Run "sc config dmwappushservice start=auto"
			oWSH.Run "sc config WMPNetworkSvc start=auto"
			oWSH.Run "sc config diagnosticshub.standardcollector.service start=auto"
			oWSH.Run "sc start DiagTrack"
			oWSH.Run "sc start RetailDemo"
			oWSH.Run "sc start dmwappushservice"
			oWSH.Run "sc start WMPNetworkSvc"		
			oWSH.Run "sc start diagnosticshub.standardcollector.service"
		Case 3
			printf ""
			printf " INFO: Se han habilitado las tareas programadas que envian datos a Microsoft"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Application Experience\ProgramDataUpdater" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\KernelCeipTask" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\Uploader" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\AppID\SmartScreenSpecific" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\NetTrace\GatherNetworkInfo" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Error Reporting\QueueReporting" & chr(34) & " /ENABLE"			
		Case 4
			hostsFile = oWSH.ExpandEnvironmentStrings("%WinDir%") & "\System32\drivers\etc\hosts"
			If oFSO.FileExists(hostsFile & ".cwd") = True Then
				oFSO.DeleteFile	hostsFile
				oFSO.CopyFile	hostsFile & ".cwd", hostsFile
			Else
				Set F = oFSO.CreateTextFile("C:\Windows\System32\drivers\etc\hosts", True)
					F.WriteLine "127.0.0.1	localhost"
					F.WriteLine "::1		localhost"
					F.WriteLine "127.0.0.1	local"
				F.Close
			End If
			printf ""
			printf " INFO: El fichero hosts se ha restablecido correctamente"
		Case 6
			printf ""
			printf " INFO: Se ha habilitado One Drive correctamente"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 0, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\OneDrive\DisableFileSyncNGSC", 0, "REG_DWORD"
		Case 5
			printf ""
			printf " INFO: Se ha habilitado Windows Defender Antivirus correctamente"
			oWSH.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 0, "REG_DWORD"
			oWSH.RegWrite "HKLM\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows Defender\DisableAntiSpyware", 0, "REG_DWORD"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cache Maintenance" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Cleanup" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Scheduled Scan" & chr(34) & " /ENABLE"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Windows Defender\Windows Defender Verification" & chr(34) & " /ENABLE"
			oWSH.Run "sc config WdNisSvc start=auto"
			oWSH.Run "sc config WinDefend start=auto"	
			oWSH.Run "sc start WdNisSvc"
			oWSH.Run "sc start WinDefend"
		Case 7
			printf ""
			printf " INFO: Se ha habilitado el servicio de VSS (Shadow Copy)"
			oWSH.Run "sc config VSS start=auto"
			oWSH.Run "sc start VSS"
		Case 8
			printf ""
			printf " INFO: Se ha habilitado el servicio de Windows Search + Indexing Service"
			oWSH.Run "sc config WSearch start=auto"
			oWSH.Run "sc start WSearch"
		Case 9
			printf ""
			printf " INFO: Se ha habilitado la tarea programada del defragmentador de discos de Windows"
			oWSH.Run "schtasks /change /TN " & chr(34) & "\Microsoft\Windows\Defrag\ScheduledDefrag" & chr(34) & " /ENABLE"
		Case 10
			printf ""
			printf " INFO: Hibernacion del sistema activada correctamente"
			oWSH.Run "powercfg -h on"
			oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power\HiberbootEnabled", 1, "REG_DWORD"
		Case 11
			printf ""
			printf " INFO. Se ha habilitado Prefetcher + Superfetch en el registro y en el servicio"
			oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnablePrefetcher", 1, "REG_DWORD"
			oWSH.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnableSuperfetch", 1, "REG_DWORD"
			oWSH.Run "sc config SysMain start=auto"
			oWSH.Run "sc start SysMain"
		Case 12
			printf ""
			printf " INFO: Se ha deshabilitado el tema oscuro (Dark Theme)"
			oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"
			oWSH.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"		
		Case 13 
			printf ""
			printf " Info: Se ha habilitado el Sensor preview, comprueba si ya funcionan los sensores de acelerometro y luz."
			oWSH.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}\SensorPermissionState", 1,  "REG_DWORD"
                        oWSH.Run "sc start SensorDataService"
			oWSH.Run "sc start SensrSvc"
		Case 0
			MsgBox "Si has restaurado alguna opcion/configuracion, te recomiendo que reinicies el sistema ahora", vbInformation + vbOkOnly, "DFX Tweaker"
			Call showMenu(0)
		Case Else
			printf ""
			printf "  Ese numero no esta disponible."
			Call restoreMenu()
			Exit Function
	End Select
	wait(2)
	Call restoreMenu()
End Function

Function printf(txt)
	WScript.StdOut.WriteLine txt
End Function

Function printl(txt)
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
		printf ""
	Next
End Function

Function ForceConsole()
	If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
		oWSH.Run "cscript //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
		WScript.Quit
	End If
End Function

Function checkW10orW11()
	If getNTversion < 10 Then
		printf "  ERROR: Necesitas ejecutar DFX Tweaker bajo Windows 10 o Windows 11"
		printf ""
		printf "  Pulsa <<Enter>> para salir"
		scanf
		WScript.Quit
	End If
End Function

Function runElevated()
	If isUACRequired Then
		If Not isElevated Then RunAsUAC
	Else
		If Not isAdmin Then
			printf "  ERROR: Necesitas ejecutar DFX Tweaker como Administrador!"
			printf ""
			printf "  Pulsa <<Enter>> para salir"
			scanf
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
		printf ""
		printf "  DFX Tweaker necesita ejecutarse como Administrador..."
		printf "  Espera..."
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

Function getNTversion()
	Set cWin = oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	For Each OS In cWin
		getNTversion = Split(OS.Version,".")(0)
	Next
End Function
