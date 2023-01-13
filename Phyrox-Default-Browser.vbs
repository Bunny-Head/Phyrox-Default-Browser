'Registers Phyrox with Default Programs or Default Apps in Windows
'FirefoxPortable.vbs - created by Ramesh Srinivasan for Winhelponline.com
'Phyrox-Default-Browser.vbs - modified by Bunny-Head
'v1.0 17-July-2022 - Initial release. Tested on Mozilla Firefox 102.0.1.0.
'v1.1 23-July-2022 - Minor bug fixes.
'v1.2 27-July-2022 - Minor revision. Cleaned up the code.
'Phyrox-Default-Browser fork created 13-January-2023
'Suitable for all Windows versions, including Windows 10/11.
'Tutorial: https://www.winhelponline.com/blog/register-firefox-portable-with-default-apps/

Option Explicit
Dim sAction, sAppPath, sExecPath, sIconPath, objFile, sbaseKey, sbaseKey2, sAppDesc
Dim sClsKey, ArrKeys, regkey
Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = oFSO.GetFile(WScript.ScriptFullName)
sAppPath = oFSO.GetParentFolderName(objFile)
sExecPath = sAppPath & "\phyrox-portable.exe"
sIconPath = sAppPath & "\phyrox-portable.exe"
sAppDesc = "Firefox delivers safe, easy web browsing. " & _
"A familiar user interface, enhanced security features including " & _
"protection from online identity theft, and integrated search let " & _
"you get the most out of the web."

'Quit if phyrox-portable.exe is missing in the current folder!
If Not oFSO.FileExists (sExecPath) Then
   MsgBox "Please run this script from Phyrox folder. The script will now quit.", _
   vbOKOnly + vbInformation, "Register Phyrox with Default Apps"
   WScript.Quit
End If

If InStr(sExecPath, " ") > 0 Then
   sExecPath = """" & sExecPath & """"
   sIconPath = """" & sIconPath & """"
End If

sbaseKey = "HKCU\Software\"
sbaseKey2 = sbaseKey & "Clients\StartmenuInternet\Phyrox\"
sClsKey = sbaseKey & "Classes\"

If WScript.Arguments.Count > 0 Then
   If UCase(Trim(WScript.Arguments(0))) = "-REG" Then Call RegisterPhyrox
   If UCase(Trim(WScript.Arguments(0))) = "-UNREG" Then Call UnRegisterPhyrox
Else
   sAction = InputBox ("Type REGISTER to add Phyrox to Default Apps. " & _
   "Type UNREGISTER To remove.", "Phyrox Registration", "REGISTER")
   If UCase(Trim(sAction)) = "REGISTER" Then Call RegisterPhyrox
   If UCase(Trim(sAction)) = "UNREGISTER" Then Call UnRegisterPhyrox
End If

Sub RegisterPhyrox   
   WshShell.RegWrite sbaseKey & "RegisteredApplications\Phyrox", _
   "Software\Clients\StartMenuInternet\Phyrox\Capabilities", "REG_SZ"
   
   'FirefoxHTML registration
   WshShell.RegWrite sClsKey & "FirefoxHTML2\", "Firefox HTML Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxHTML2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sClsKey & "FirefoxHTML2\FriendlyTypeName", "Firefox HTML Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxHTML2\DefaultIcon\", sIconPath & ",1", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxHTML2\shell\", "open", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxHTML2\shell\open\command\", sExecPath & _
   " -url " & """" & "%1" & """", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxHTML2\shell\open\ddeexec\", "", "REG_SZ"
   
   'FirefoxPDF registration
   WshShell.RegWrite sClsKey & "FirefoxPDF2\", "Firefox PDF Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxPDF2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sClsKey & "FirefoxPDF2\FriendlyTypeName", "Firefox PDF Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxPDF2\DefaultIcon\", sIconPath & ",5", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxPDF2\shell\open\", "open", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxPDF2\shell\open\command\", sExecPath & _
   " -url " & """" & "%1" & """", "REG_SZ"
   
   'FirefoxURL registration
   WshShell.RegWrite sClsKey & "FirefoxURL2\", "Firefox URL", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxURL2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sClsKey & "FirefoxURL2\FriendlyTypeName", "Firefox URL", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxURL2\URL Protocol", "", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxURL2\DefaultIcon\", sIconPath & ",1", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxURL2\shell\open\", "open", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxURL2\shell\open\command\", sExecPath & _
   " -url " & """" & "%1" & """", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxURL2\shell\open\ddeexec\", "", "REG_SZ"   
   
   'Default Apps Registration/Capabilities
   WshShell.RegWrite sbaseKey2, "Phyrox", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\ApplicationDescription", sAppDesc, "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\ApplicationIcon", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\ApplicationName", "Phyrox", "REG_SZ" 
   WshShell.RegWrite sbaseKey2 & "Capabilities\FileAssociations\.pdf", "FirefoxPDF2", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\StartMenu", "Phyrox", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "shell\open\command\", sExecPath, "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "shell\properties\", "Firefox &Options", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "shell\properties\command\", sExecPath & " -preferences", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "shell\safemode\", "Firefox &Safe Mode", "REG_SZ"   
   WshShell.RegWrite sbaseKey2 & "shell\safemode\command\", sExecPath & " -safe-mode", "REG_SZ"
   
   ArrKeys = Array ( _
   "FileAssociations\.avif", _
   "FileAssociations\.htm", _
   "FileAssociations\.html", _
   "FileAssociations\.shtml", _
   "FileAssociations\.svg", _
   "FileAssociations\.webp", _
   "FileAssociations\.xht", _
   "FileAssociations\.xhtml", _
   "URLAssociations\http", _
   "URLAssociations\https", _
   "URLAssociations\mailto" _
   )
   
   For Each regkey In ArrKeys
      WshShell.RegWrite sbaseKey2 & "Capabilities\" & regkey, "FirefoxHTML2", "REG_SZ"
   Next      
   
   'Override the default app name by which the program appears in Default Apps  (*Optional*)
   '(i.e., -- "Mozilla Firefox, Portable Edition" Vs. "Phyrox")
   'The official Mozilla Firefox setup doesn't add this registry key.
   WshShell.RegWrite sClsKey & "FirefoxHTML2\Application\ApplicationIcon", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sClsKey & "FirefoxHTML2\Application\ApplicationName", "Phyrox", "REG_SZ"
   
   'Launch Default Programs or Default Apps after registering Phyrox   
   WshShell.Run "control /name Microsoft.DefaultPrograms /page pageDefaultProgram"
End Sub


Sub UnRegisterPhyrox
   sbaseKey = "HKCU\Software\"
   sbaseKey2 = "HKCU\Software\Clients\StartmenuInternet\Phyrox"   
   
   On Error Resume Next
   WshShell.RegDelete sbaseKey & "RegisteredApplications\Phyrox"
   On Error GoTo 0
   
   WshShell.Run "reg.exe delete " & sClsKey & "FirefoxHTML2" & " /f", 0
   WshShell.Run "reg.exe delete " & sClsKey & "FirefoxPDF2" & " /f", 0
   WshShell.Run "reg.exe delete " & sClsKey & "FirefoxURL2" & " /f", 0
   WshShell.Run "reg.exe delete " & chr(34) & sbaseKey2 & chr(34) & " /f", 0
   
   'Launch Default Apps after unregistering Phyrox   
   WshShell.Run "control /name Microsoft.DefaultPrograms /page pageDefaultProgram"   
End Sub