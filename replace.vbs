' https://step-learn.com/article/vbscript/025-string-replace.html

Dim str
str = "abc-def-abc-def"

' すべて置換
WScript.Echo Replace(str, "abc", "xyz")  ' xyz-def-xyz-def

' 開始位置を指定
WScript.Echo Replace(str, "abc", "xyz", 5)  ' def-xyz-def

' 1回のみ置換
WScript.Echo Replace(str, "abc", "xyz", 1, 1) ' xyz-def-abc-def

' vbBinaryCompare
WScript.Echo Replace(str, "ABC", "xyz", 1, -1, 0) ' abc-def-abc-def

' vbTextCompare
WScript.Echo Replace(str, "ABC", "xyz", 1, -1, 1) ' xyz-def-xyz-def