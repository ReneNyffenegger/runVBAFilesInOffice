'
'   ..\..\..\runVBAFilesInOffice.vbs -excel name -c Go
'

public sub Go() ' {

    range("b3:d6").name = "range_one"
    range("e2:f8").name = "range_two"

    range("range_one").formula = "=rand()"
    range("range_one").interior.color = rgb(255, 140,  30)

    range("range_two").value   = "two"
    range("range_two").interior.color = rgb(215, 215, 215)

    activeWorkbook.saved = true

end sub ' }

