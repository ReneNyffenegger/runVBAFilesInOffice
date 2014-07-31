'
'   ..\..\..\runVBAFilesInOffice.vbs -excel cells -c Go
'

public sub Go()

    cells(5, 3).interior.color = rgb(255, 127, 30)

    activeWorkbook.saved = true

end sub
