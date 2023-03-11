' VBS IS CASE INSENSITIVE

' String
' Integer
' Long
' Single
' Double
' Byte
' Currency
' Boolean
' Date
' Empty (undefined)
' Null
' Object
' Nothing (not referencing an object)
' Error
' Unknown

wscript.echo typename("hello") ' String
wscript.echo typename(123) ' Integer
wscript.echo typename(123456) ' Long
wscript.echo typename(3.14) ' Double
' wscript.echo typename(3.141f) ' ERROR
wscript.echo typename(true) ' Boolean

Dim num
' num = "100"
' Msgbox num + num ' 100100
' num = 100
' Msgbox num + num ' 200
num = "100"
Msgbox num + num ' 100100