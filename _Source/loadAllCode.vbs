Set objExcel = CreateObject("Excel.Application")
set objShell = CreateObject("WScript.Shell")
currentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
Set objWorkbook = objExcel.Workbooks.Open(currentDirectory & "deployer.xlam")

objExcel.Application.Visible = True
objExcel.Application.Run "importModules"

saveaddress = currentDirectory & "_CPNM.xlam"
objWorkbook.Saveas saveaddress, 55

objExcel.Application.Quit



' Adicionando os ignores corretos no git.
on error resume next
objShell.Run "git update-index --assume-unchanged " & currentDirectory & "\_CPNM.xlam"
objShell.Run "git update-index --assume-unchanged " & currentDirectory & "\_CPNM.dotm"
objShell.Run "git update-index --assume-unchanged " & currentDirectory & "\_CPNM.dvb"
objShell.Run "git update-index --assume-unchanged " & currentDirectory & "\deployer.xlam"
on error goto 0