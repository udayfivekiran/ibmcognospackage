On Error Resume Next 
Const HKEY_CURRENT_USER =  &H80000001

  Set objShell = CreateObject("WScript.Shell")
  Set objFS = CreateObject("Scripting.FileSystemObject")

  strComputer = "."

  Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
  strComputer & "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Office\15.0\Excel\Options"

  addin_val="/R " & chr(34) & objShell.expandenvironmentstrings("%programfiles(x86)%") & "\ibm\cognos\IBM for Microsoft Office\PAforExcel.xll" & chr(34)



  objReg.EnumValues HKEY_CURRENT_USER, strKeyPath, arrValueNames, arrValueTypes
strvalue1="HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Excel\Options"
  
  For intLoop = 0 To UBound(arrValueNames) 

 my_val=arrvaluenames(intLoop)
  
  
  If UCase(Left(my_val,4))="OPEN" Then 


strvalue=strvalue1 & "\" & arrvaluenames(intLoop)


if objShell.regread(strvalue)=addin_val  then 

objShell.regdelete(strvalue)

End If
 End If 

  Next


  strKeyPath = "Software\Microsoft\Office\15.0\Excel\Options"


  loop_count=0
  crt_counter=0
  objReg.EnumValues HKEY_CURRENT_USER, strKeyPath, arrValueNames, arrValueTypes
  
  For intLoop = 0 To UBound(arrValueNames) 
  my_val=arrValueNames(intLoop) 
 
  
  If UCase(Left(my_val,4))="OPEN" Then 

  loop_count=1
     counter=Right(my_val,1)    
     If IsNumeric(counter) And crt_counter < counter Then
          crt_counter=counter
 
  End If  
  End If 

  Next
 
 
 
 If loop_count=0 Then
 
 reg_var="HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Excel\Options\OPEN"
 reg_val="/R " & chr(34) & objShell.expandenvironmentstrings("%programfiles(x86)%") & "\ibm\cognos\IBM for Microsoft Office\PAforExcel.xll" & chr(34)

 objShell.RegWrite reg_var,reg_val
 Else 
 
 reg_var="HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Excel\Options\OPEN"&crt_counter+1
 reg_val="/R " & chr(34) & objShell.expandenvironmentstrings("%programfiles(x86)%") & "\ibm\cognos\IBM for Microsoft Office\PAforExcel.xll" & chr(34)
 objShell.RegWrite reg_var,reg_val
 
 End if

dest = objShell.expandenvironmentstrings("%localappdata%") & "\Cognos"
objFS.CreateFolder dest
dest = objShell.expandenvironmentstrings("%localappdata%") & "\Cognos\Office Connection"
objFS.CreateFolder dest
sd = objFS.GetParentFolderName(WScript.ScriptFullName)

objFS.CopyFile sd&"\CognosOfficeReportingSettings.xml", dest&"\CognosOfficeReportingSettings.xml"
objFS.CopyFile sd&"\CognosOfficeXLLSettings.xml", dest&"\CognosOfficeXLLSettings.xml"