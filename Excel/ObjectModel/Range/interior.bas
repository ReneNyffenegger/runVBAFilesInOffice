'
'   ..\..\..\runVBAFilesInOffice.vbs -excel interior -c Go
'

public sub Go()

  dim x as long, y as long

  for x = 1 to 10
  for y = 1 to 10

      cells(x, y).interior.color = rgb(25*x, 25*y, 30)

  next y
  next x
  
  activeWorkbook.saved = true

end sub
