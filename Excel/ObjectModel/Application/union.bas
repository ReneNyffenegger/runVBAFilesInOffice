'
'   ..\..\..\runVBAFilesInOffice.vbs -excel union -c Go
'
'   See also -> intersect.bas
'

public sub Go()

    dim range_1      as range
    dim range_2      as range
    dim range_result as range


    set range_1 = activeSheet.range("f3:f9")
    set range_2 = activeSheet.range("c6:k6")

    set range_result = union (range_1, range_2)

    range_result.interior.color = RGB(255, 127 , 30)

    activeWorkbook.saved = true

end sub
