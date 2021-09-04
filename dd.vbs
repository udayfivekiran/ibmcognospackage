Set objFS = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
dest = objShell.expandenvironmentstrings("%localappdata%") & "\Cognos"
objFS.CreateFolder dest
dest = objShell.expandenvironmentstrings("%localappdata%") & "\Cognos\Office Connection"
objFS.CreateFolder dest
sd = objFS.GetParentFolderName(WScript.ScriptFullName)

objFS.CopyFile sd&"\CognosOfficeReportingSettings.xml", dest&"\CognosOfficeReportingSettings.xml"
objFS.CopyFile sd&"\CognosOfficeXLLSettings.xml", dest&"\CognosOfficeXLLSettings.xml"