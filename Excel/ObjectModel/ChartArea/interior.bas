'
'  ..\..\..\runVBAFilesInOffice.vbs -excel interior -c interior
'

public sub interior()

    dim row_        as integer
    dim shape_      as shape
    dim chart_      as chart
    dim chart_area_ as chartArea

    row_ = 1

    cells(1,1).value = "x"
    cells(1,2).value = "sin(x) * x/3 + x"

    for x = 0 to 10 step 0.1

        row_ = row_ + 1

        cells(row_, 1).value = x
        cells(row_, 2).value = sin(x) * x / 3 + x

    next x

    set shape_ = activeSheet.shapes.addChart
    set chart_ = shape_.chart

    chart_.chartType = xlXYScatterSmoothNoMarkers
    chart_.setSourceData source := range(cells(1,1), cells(row_, 2))

    set chart_area_ = chart_.chartArea

 '  TODO: Where is ChartArea.Interior documented?
    chart_area_.interior.colorIndex = 5

    activeWorkbook.saved = true

end sub
