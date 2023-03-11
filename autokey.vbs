' https://learn.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364423(v=msdn.10)?redirectedfrom=MSDN
' BACKSPACE	{BACKSPACE}, {BS}, or {BKSP}
' BREAK	{BREAK}
' CAPS LOCK	{CAPSLOCK}
' DEL or DELETE	{DELETE} or {DEL}
' DOWN ARROW	{DOWN}
' END	{END}
' ENTER	{ENTER} or ~
' ESC	{ESC}
' HELP	{HELP}
' HOME	{HOME}
' INS or INSERT	{INSERT} or {INS}
' LEFT ARROW	{LEFT}
' NUM LOCK	{NUMLOCK}
' PAGE DOWN	{PGDN}
' PAGE UP	{PGUP}
' PRINT SCREEN	{PRTSC}
' RIGHT ARROW	{RIGHT}
' SCROLL LOCK	{SCROLLLOCK}
' TAB	{TAB}
' UP ARROW	{UP}
' F1	{F1}
' F2	{F2}
' F3	{F3}
' F4	{F4}
' F5	{F5}
' F6	{F6}
' F7	{F7}
' F8	{F8}
' F9	{F9}
' F10	{F10}
' F11	{F11}
' F12	{F12}
' F13	{F13}
' F14	{F14}
' F15	{F15}
' F16	{F16}
' SHIFT	+
' CTRL	^
' ALT	%

option explicit 

' CREATE WSH-SHELL OBJECT
dim objWsh 
set objWsh = createObject("wscript.shell")

' OPEN THE APP
objWsh.run "C:\Windows\system32\notepad.exe", 1, False

wscript.sleep(500) ' msec
objWsh.appActivate("notepad")
wscript.sleep(500)

dim msg, interval, i
msg = "hello world!!"
interval = 300

for i = 0 to len(msg)
    objWsh.sendkeys(mid(msg, i+1, 1))
    wscript.sleep(interval)
next 

' DELETE WSH OBJECT
set objWsh = nothing