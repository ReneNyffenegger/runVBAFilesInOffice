'
'  ..\..\..\runVBAFilesInOffice.vbs -word addPicture -c addPicture %CD%
'

sub addPicture(cur_dir as string)

  ' exported_chart.gif was created with https://github.com/ReneNyffenegger/runVBAFilesInOffice/blob/master/Excel/ObjectModel/Chart/export.bas
    call activeDocument.inlineShapes.addPicture(fileName := cur_dir & "\exported_chart.gif")

    activeDocument.saved = true

end sub
