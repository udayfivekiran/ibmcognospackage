On Error Resume Next 
Dim strComputer, WshShell,  strExpandedString, oReg, strKeyPath, strKeyPath1, strValueName, strValue,strValueName1,strValue1
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."


Set WshShell = CreateObject("WScript.Shell")

strExpandedString = WshShell.ExpandEnvironmentStrings("%WinDir%")


Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Active Setup\Installed Components\IBM_Planning_Analytics_Integration_x64-UNINSTALL"
strKeyPath1 = "SOFTWARE\Microsoft\Active Setup\Installed Components\{52A8A7EE-7314-49FD-BB47-8892289B3DA5}"
strValueName = "StubPath"
strValueName1 = "Version"
strValue = strExpandedString +"\SysWOW64\IBM_Planning_Analytics_Integration_x64\Uninstall_deleting_Open_Keys.vbs"

oReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName1,strValue1
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath1,strValueName1,strValue2

v1 = "1,0"

If IsNull(strValue1) Then
	
    	oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName1,v1
	
Else
	VValue=WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\IBM_Planning_Analytics_Integration_x64-UNINSTALL\Version")
	Vval = Split(Vvalue,",")
	valn = Vval(0) + 1
	oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName1,valn & ",0"
	


End If

