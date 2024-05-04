Dim message, strSearch
message = inputbox("type something.", "TEST INPUT BOX")

strSearch = "unko" ' includes this word?

' If InStr(message, strSearch) > 0 Then
If InStr(1, message, strSearch, vbTextCompare) > 0 Then ' if using "vbTextCompare", the number of augument is 4
	WScript.Echo strSearch & " is included."
Else
	WScript.Echo strSearch & " is not included"
End If

' 「比較モード」の値は次の通りで、省略した場合はバイナリーモードでの 検索となります。
' vbBinaryCompare	0	大文字小文字を区別する（バイナリーモード）
' vbTextCompare	1	大文字小文字を区別しない（テキストモード）