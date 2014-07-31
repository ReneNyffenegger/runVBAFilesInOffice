'
'   ..\..\..\runVBAFilesInOffice.vbs -excel style -c Go
'

public sub Go()

  dim f as lineFormat

  call create_line("b2", "e2", msoLineSingle          )
  call create_line("b3", "e3", msoLineThickBetweenThin)
  call create_line("b4", "e4", msoLineThickThin       )
  call create_line("b5", "e5", msoLineThinThick       )
  call create_line("b6", "e6", msoLineThinThin        )

  activeWorkbook.saved = true

end sub

private sub create_line(fromCell as string, toCell as string, style_ as msoLineStyle) 

  dim line_   as shape
  dim format_ as lineFormat

  set line_ = activeSheet.shapes.addline(  _
                beginX :=  range(fromCell).left + range(fromCell).width   / 2, _
                beginY :=  range(fromCell).top  + range(fromCell).height  / 2, _
                endX   :=  range(toCell  ).left + range(toCell  ).width   / 2, _
                endY   :=  range(toCell  ).top  + range(toCell  ).height  / 2  )

  set format_ = line_.line

  format_.weight =     10
  format_.style  = style_

end sub
