'
'  ..\..\..\runVBAFilesInOffice.vbs -excel page -c main
'

public sub main()

    dim ps as pageSetup

    set ps = activeSheet.pageSetup

    ps.paperSize   = xlPaperA4
    ps.orientation = xlLandscape

    activeWorkbook.saved = true

end sub
