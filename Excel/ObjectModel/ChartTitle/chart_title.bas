'
'  ..\..\..\runVBAFilesInOffice.vbs -excel chart_title -c main
'

public sub main()

    dim row_         as integer
    dim shape_       as shape
    dim chart_       as chart
    dim chart_title_ as chartTitle

    row_ = 1

    cells(1,1).value = "x"
    cells(1,2).value = "sin(x) * x/3 + x" ' Somehow, this text becomes the chartTitle.text

    for x = 0 to 10 step 0.1

        row_ = row_ + 1

        cells(row_, 1).value = x
        cells(row_, 2).value = sin(x) * x / 3 + x 

    next x

    set shape_ = activeSheet.shapes.addChart
    set chart_ = shape_.chart

    chart_.chartType = xlXYScatterSmoothNoMarkers
    chart_.setSourceData source := range(cells(1,1), cells(row_, 2))

    set chart_title_ = chart_.chartTitle

    chart_title_.text = chart_title_.text & " (Title of the Chart)"
    chart_title_.format.fill.foreColor.rgb = rgb(255, 200,  50)
    chart_title_.format.line.foreColor.rgb = rgb(200,  50, 255)
    chart_title_.format.line.weight        = 3

    activeWorkbook.saved = true

end sub
