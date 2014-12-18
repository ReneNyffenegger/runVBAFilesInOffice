'
'  ..\..\..\runVBAFilesInOffice.vbs -word addConnector -c main
'

sub main()

    dim connector as shape

    set shape = activeDocument.shapes.addConnector(msoConnectorStraight, 20, 50, 400, 200)

    shape.line.weight        = 5#
    shape.line.foreColor.rgb = rgb(255, 50, 20)

    activeDocument.saved = true

end sub
