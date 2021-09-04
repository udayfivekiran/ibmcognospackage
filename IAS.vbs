On Error Resume Next

Dim Oshell, strComputer, objRegistry, strKeyPath, strValueName, strValue, strKeyPath1, strValueName1, strValue1

Const HKEY_LOCAL_MACHINE = &H80000002

Set Oshell = CreateObject("Wscript.Shell")

strKeyPath = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
strValueName = "Identifier"

strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & _ 
   strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Active Setup\Installed Components\IBM_Planning_Analytics_Integration_x64-UNINSTALL"
strValueName = "StubPath"

objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

If IsNull(strValue) Then
    'Dont do anything
Else
    objRegistry.DeleteValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName
End If


strValueName1 = "Version"
strKeyPath1 = "SOFTWARE\Microsoft\Active Setup\Installed Components\{52A8A7EE-7314-49FD-BB47-8892289B3DA5}"
objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath1,strValueName1,strValue1
v1 = "1,0"

If IsNull(strValue1) Then

	objRegistry.CreateKey HKEY_LOCAL_MACHINE,strKeyPath1
    	objRegistry.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath1,strValueName1,v1
Else
	VValue=	Oshell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{52A8A7EE-7314-49FD-BB47-8892289B3DA5}\Version")

	Vval = Split(Vvalue,",")
	valn = Vval(0) + 1
	objRegistry.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath1,strValueName1,valn & ",0"

End If


