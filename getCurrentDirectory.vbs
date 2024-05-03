' Get the current directory
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Msgbox fso.getParentFolderName(WScript.ScriptFullName)