arr = array(1, 2, 3, 8)

for each b in arr
    wscript.echo b ' 1 2 3
next

' ubound means maxnum in the array (.length-1)
' lbound means minnum in the array (0)
for i = 0 to ubound(arr)
    wscript.echo arr(i)
next
wscript.echo ubound(arr) '3
wscript.echo lbound(arr) '0
