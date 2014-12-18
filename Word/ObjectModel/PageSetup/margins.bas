'
'  ..\..\..\runVBAFilesInOffice.vbs -word margins -c main
'

sub main()

    dim ps as pageSetup

    set ps = activeDocument.pageSetup

    ps.leftMargin   = centimetersToPoints (1  )
    ps.rightMargin  = centimetersToPoints (2  )
    ps.topMargin    = centimetersToPoints (1.5)
    ps.bottomMargin = centimetersToPoints (3.0)

    activeDocument.saved = true

end sub

