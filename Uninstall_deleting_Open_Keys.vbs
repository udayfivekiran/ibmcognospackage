on error resume next

Const HKEY_CURRENT_USER =  &H80000001

  Set objShell = CreateObject("WScript.Shell")
  Set objFS = CreateObject("Scripting.FileSystemObject")

  strComputer = "."

  Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
  strComputer & "\root\default:StdRegProv")

  strKeyPath = "Software\Microsoft\Office\15.0\Excel\Options"

  addin_val1="/R " & chr(34) & objShell.expandenvironmentstrings("%programfiles(x86)%") & "\IBM_Planning_Analytics\IBM_PAfE_x64_2.0.65.11.xll" & chr(34)
addin_val2= chr(34) & objShell.expandenvironmentstrings("%programfiles(x86)%") & "\IBM_Planning_Analytics\IBM_PAfE_x64_2.0.65.11.xll" & chr(34)


  objReg.EnumValues HKEY_CURRENT_USER, strKeyPath, arrValueNames, arrValueTypes

strvalue1="HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Excel\Options"
  
  For intLoop = 0 To UBound(arrValueNames) 

 my_val=arrvaluenames(intLoop)
  
  
  If UCase(Left(my_val,4))="OPEN" Then 


strvalue=strvalue1 & "\" & arrvaluenames(intLoop)


if objShell.regread(strvalue)=addin_val1  then 

objShell.regdelete(strvalue)


elseif objShell.regread(strvalue)=addin_val2  then 

objShell.regdelete(strvalue)




End If

 End If 

Next

  
dest = objShell.expandenvironmentstrings("%localappdata%") & "\Cognos\Office Connection"
objFS.DeleteFolder dest,True		


 
 


