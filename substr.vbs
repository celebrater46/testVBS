dim str : str = "ABCDEFG"

wscript.echo left(str, 3)

wscript.echo mid(str, 3) ' CDEFG
wscript.echo mid(str, 3, 2) ' CD
wscript.echo mid(str, len(str) - 3, 2) ' DE

wscript.echo right(str, 3) ' EFG