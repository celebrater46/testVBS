option Explicit
on error resume next

class SoldierClass
    public id
    public name
    public skills

    ' private function class_initialize(num, str, array)
    ' private sub class_initialize(num, str, array)
    public function setInfo(num, str, array)
        id = num
        name = str
        skills = array
    end function

    public function echoInfo()
        wscript.echo id
        wscript.echo name
        wscript.echo skills
    end function
end class

' class_initialize() could not receive argument
' set soldier = new Soldier(1, "hideru", array("josou", "spy"))
dim soldier
set soldier = new SoldierClass
' arr = array("josou", "spy")
' soldier.setInfo 1 hideru josou ' ERROR

' soldier.setInfo(18, "miyon", "charm") ' ERROR
soldier.id = 18
soldier.name = "miyon"
soldier.skills = "charm"

' wscript.echo soldier.id
' wscript.echo soldier.name
' wscript.echo soldier.skills
wscript.echo "hello"

soldier.echoInfo()

' do not use the same name for a variable and a class
' VBS is CASE INSENSITIVE
