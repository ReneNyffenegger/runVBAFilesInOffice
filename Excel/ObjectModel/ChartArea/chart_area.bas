'
'  ..\..\..\runVBAFilesInOffice.vbs -excel chart_area -c chart_area
'

public sub chart_area()

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

    chart_area_.format.fill.foreColor.rgb = rgb(255, 200,  50)
    chart_area_.format.line.foreColor.rgb = rgb(200,  50, 255)
    chart_area_.format.line.weight        = 3

    activeWorkbook.saved = true

end sub
