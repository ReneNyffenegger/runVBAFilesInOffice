'
'  ..\..\..\runVBAFilesInOffice.vbs -excel export -c export %cd%\..\..\..\Word\ObjectModel\InlineShapes\exported_chart.gif
'

public sub export(file_name as string)

    dim row_   as integer
    dim shape_ as shape
    dim chart_ as chart

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

    call chart_.export(fileName := file_name, filterName := "gif")

    activeWorkbook.saved = true

end sub
