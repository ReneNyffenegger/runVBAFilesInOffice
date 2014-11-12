'   ..\..\..\runVBAFilesInOffice.vbs -excel cells -c Main
'

public sub Main() ' {

    dim cur_worksheet as worksheet

    dim range_start   as range
    dim range_end     as range
    dim range_all     as range

    set cur_worksheet = activeSheet

 '  Note, instead of the explicit cur_worksheet.cells(...)
 '  the line can also be written just as
 '    cells(...) 
 '  without cur_worksheet.
    set range_start = cur_worksheet.cells( 4, 2) ' 4th row, 2nd column
    set range_end   = cur_worksheet.cells( 7, 5) ' 7th row, 5th column

    set range_all = range(range_start, range_end) 

    range_all.value = "X"

    activeWorkbook.saved = true

end sub ' }

