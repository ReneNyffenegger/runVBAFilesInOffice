'
'   ..\..\..\runVBAFilesInOffice.vbs -excel intersect -c Go
'
'   See also -> union.bas
'

public sub Go()

    dim range_1      as range
    dim range_2      as range
    dim range_result as range


    set range_1      = activeSheet.range("d3:f9")
    set range_2      = activeSheet.range("c5:k7")
    set range_result = intersect (range_1, range_2)

    range_1.value="Range 1"
    range_2.value="Range 2"
    range_result.value="Interection"

    activeWorkbook.saved = true

end sub
