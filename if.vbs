' a = b	  a == b
' a <> b   a != b
' a > b
' a >= b
' a < b
' a <= b
' and   &&
' or ||

dim score
score = inputbox("Tell me your score.")
if score > 80 Then
    wscript.echo "Are you a genius??"
elseif score > 60 Then
    wscript.echo "You pass."
else
    wscript.echo "Are you an idiot??"
end if