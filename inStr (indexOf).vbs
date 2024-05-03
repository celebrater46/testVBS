Dim message, strSearch
message = inputbox("type something.", "TEST INPUT BOX")

strSearch = "unko" ' includes this word?

If InStr(message, strSearch) > 0 Then
	WScript.Echo strSearch & " is included."
Else
	WScript.Echo strSearch & " is not included"
End If
