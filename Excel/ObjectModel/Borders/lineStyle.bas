'
'   ..\..\..\runVBAFilesInOffice.vbs -excel lineStyle -c main
'

dim r as long
public sub main() ' {

    r = 1

    border(xlContinuous   )
    border(xlDash         )
    border(xlDashDot      )
    border(xlDashDotDot   )
    border(xlDot          )
    border(xlDouble       )
    border(xlLineStyleNone)
    border(xlSlantDashDot )


    range( cells(1,1), cells(8,1) ).rowHeight = application.centimetersToPoints(2)

    activeWorkbook.saved = true

end sub ' }

private sub border(style as xlLineStyle) ' {

    dim b as borders

    set b = cells(r, 2).borders
    
    b.lineStyle = style

    r = r+1

end sub ' }
