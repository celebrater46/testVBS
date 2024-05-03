' Get the current directory
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
' Msgbox fso.getParentFolderName(WScript.ScriptFullName)

Dim strPath, objFS, objFile
strPath = fso.getParentFolderName(WScript.ScriptFullName) & "\test\testCTF.txt"

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.CreateTextFile(strPath)

objFile.WriteLine("Hello world!!!!!!!!")
objFile.Close