'
'  ..\..\..\runVBAFilesInOffice.vbs -excel margins -c main
'

public sub main()

    dim ps as pageSetup

    set ps = activeSheet.pageSetup

    ps.leftMargin    = application.centimetersToPoints(0.5)
    ps.rightMargin   = application.centimetersToPoints(0.5)
    ps.topMargin     = application.centimetersToPoints(0.5)
    ps.bottomMargin  = application.centimetersToPoints(0.5)

    ps.footerMargin  = application.centimetersToPoints(0  )
    ps.headerMargin  = application.centimetersToPoints(0  )

    activeWorkbook.saved = true

end sub
