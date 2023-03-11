option explicit 

' USE CONSTANTS INSTEAD OF NUMBERS
const vbHide = 0
const vbNormalFocus = 1
const vbMinimizedFocus = 2
const vbMaximizedFocus = 3
const vbNormalNoFocus = 4
const vbMinimizedNoFocus = 6

' CREATE WSH-SHELL OBJECT
dim objWsh 
set objWsh = createObject("wscript.shell")

' OPEN THE APP
' objWsh.run "C:\Windows\system32\notepad.exe", vbNormalFocus, False
objWsh.run "C:\Windows\system32\notepad.exe", 1, False

' DELETE WSH OBJECT
set objWsh = nothing