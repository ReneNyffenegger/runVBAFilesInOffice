'
'  ..\..\..\runVBAFilesInOffice.vbs -word add -c add
'

sub add()

    activeDocument.saved = true

    dim doc(3) as document

    for i = 1 to 3 
        set doc(i) = documents.add()
    next i

    for i = 1 to 3 

        call doc(i).select

        call selection.moveLeft(unit:=wdCharacter, count:=1, extend:=wdExtend)

        call selection.typeText(i)
        call selection.moveLeft(unit:=wdCharacter, count:=1, extend:=wdExtend)

        selection.font.name = "Courier New"
        selection.font.size =  100


    next i

    for i = 1 to 3 
        doc(i).saved = true
    next i


end sub
