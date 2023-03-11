option Explicit
on error resume next

class SoldierClass
    private id, name

    private sub class_initialize()
        ' do stuff
    end sub

    public default property get constructor(num, str)
        id = num
        name = str
        set constructor = me
    end property

    public function echoInfo()
        wscript.echo id
        wscript.echo name
    end function
end class

' dim num
' dim name
dim num : num = 14
dim name : name = "seiko"

dim obj
set obj = new SoldierClass
call obj.constructor(num, name)
' set obj = (new SoldierClass)(num, name)
obj.echoInfo()