'
'   ..\..\..\runVBAFilesInOffice.vbs -excel offset -c Go
'

public sub Go()

    dim range_orig   as range
    dim range_offset as range

    set range_orig = range("d3:g6")

  ' New range: 1 downwards, 3 leftwards
    set range_offset = range_orig.offset(1, 3)

    range_orig.value = "Orig"
    range_offset.interior.color = rgb(255, 127, 30)

    activeWorkbook.saved = true

end sub
