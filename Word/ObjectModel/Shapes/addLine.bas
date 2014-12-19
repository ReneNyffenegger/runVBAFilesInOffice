'
'  ..\..\..\runVBAFilesInOffice.vbs -word addLine -c main
'

sub main()

    dim connector as shape

    set shape = activeDocument.shapes.addLine(20, 50, 400, 200)

    shape.line.weight        = 5#
    shape.line.foreColor.rgb = rgb(255, 50, 20)

    activeDocument.saved = true

end sub
