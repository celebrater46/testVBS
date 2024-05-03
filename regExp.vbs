' https://www.sejuku.net/blog/33541

'RegExpオブジェクトの作成
Dim reg As Object
Set reg = CreateObject("VBScript.RegExp")

'正規表現の指定
With reg
    .Pattern = "[0-9]"       'パターンを指定
    .IgnoreCase = False '大文字と小文字を区別するか(False)、しないか(True)
    .Global = True           '文字列全体を検索するか(True)、しないか(False)
End With