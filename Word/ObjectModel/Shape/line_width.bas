'
'  ..\..\..\runVBAFilesInOffice.vbs -word line_width -c main
'

option explicit

dim ad_sh     as shapes
dim pageWidth as double

dim curLineTop as double

sub main()

    set ad_sh  = activeDocument.shapes
    pageWidth  = activeDocument.pageSetup.pageWidth
    curLineTop = 3

    call nextLine(0.25)
    call nextLine(0.50)
    call nextLine(1   )
    call nextLine(2   )
    call nextLine(5   )
    call nextLine(10  )

    activeDocument.saved = true

end sub

private sub nextLine(w as double)

    dim line_ as shape

    set line_ = ad_sh.addLine(centimetersToPoints( 3), _
                              centimetersToPoints( curLineTop), _
                              pageWidth - centimetersToPoints(3), _
                              centimetersToPoints( curLineTop))

    line_.line.weight = w


    curLineTop = curLineTop + 1

end sub
