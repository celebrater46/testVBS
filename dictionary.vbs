dim prices
set prices = createobject("scripting.dictionary")

' hoge.add key, value
prices.add "AK-47", 300
prices.add "M16A1", 500

wscript.echo prices("AK-47")
wscript.echo prices("M16A1")

' whether a key with a specific name exists
dim str : str = "M16A2"
if prices.exists(str) Then
    wscript.echo "A rifle " + str + " exists!"
else
    wscript.echo "A rifle " + str + " does not exist!"
end if

' cannot confirm whether a value a specific value exists