Set oWSH = CreateObject("WScript.Shell")

Call ForceConsole()
Call TestMenu()
Function TestMenu()
textf " "
textf " "
textf " "
textf " "
textf " "
textf " "
textf " "
textf "  THIS SCRIPT HAS BEEN CREATED TO TEST THE UPDATE FUNCTION ON CERTAIN DEVELOPMENT BUILDS OF DFX WINTWEAKS"
textf " "
textf " "
textf "	"
textf "	"
textf "	"
textf "  Update completed. Test finished."
textf " "
textf " "
textf "  ivandfx"
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
	RP = scanf
	If isNumeric(RP) = False Then
		Exit Function
	End If
			Exit Function
End Function

Function textf(txt)
	WScript.StdOut.WriteLine txt
End Function

Function scanf()
	scanf = LCase(WScript.StdIn.ReadLine)
End Function

Function ForceConsole()
	If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
		oWSH.Run "cscript //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
		WScript.Quit
	End If
End Function