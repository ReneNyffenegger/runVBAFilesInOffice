'
'   ..\..\..\runVBAFilesInOffice.vbs -excel b3-e9 -c Go
'

public sub Go() ' {

    col_start = 2  ' B
    row_start = 3

    col_end   = 5  ' E
    row_end   = 9

    range(cells(row_start, col_start), cells(row_end, col_end)).formula = "=rand()"

    activeWorkbook.saved = true

end sub ' }
