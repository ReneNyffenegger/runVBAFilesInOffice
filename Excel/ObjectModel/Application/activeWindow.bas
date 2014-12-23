
'
'   ..\..\..\runVBAFilesInOffice.vbs -excel activeWindow -c main
'

public sub main()

    dim w as window

    set w = application.activeWindow

    application.cells(1,1).value = w.caption

    activeWorkbook.saved = true

end sub
