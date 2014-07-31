'
'   ..\..\..\runVBAFilesInOffice.vbs -excel addLine -c Go
'

public sub Go()

  dim line as shape

  set line = create_line("b2", "e2")

  set line = create_line("c3", "f9")



  activeWorkbook.saved = true

end sub

private function create_line(fromCell as string, toCell as string) as shape

  set line = activeSheet.shapes.addline(  _
                beginX :=  range(fromCell).left + range(fromCell).width   / 2, _
                beginY :=  range(fromCell).top  + range(fromCell).height  / 2, _
                endX   :=  range(toCell  ).left + range(toCell  ).width   / 2, _
                endY   :=  range(toCell  ).top  + range(toCell  ).height  / 2  )

end function
