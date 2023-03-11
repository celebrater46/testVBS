dim answer
answer = inputbox("Janken...(Type 'goo', 'choki' or 'par')")

' com is par
select case answer
    case "goo"
        wscript.echo "You Lose."
    case "choki"
        wscript.echo "You Win."
    case "par"
        wscript.echo "Draw."
    case else 
        wscript.echo "Type 'goo', 'choki' or 'par'"
end select 
