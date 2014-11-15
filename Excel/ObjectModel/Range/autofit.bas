'
'   ..\..\..\runVBAFilesInOffice.vbs -excel autofit -c main
'

public sub main()

  for c = 1 to 20

    cells(1, c).value = string(c, "*")

    columns(c).autoFit

  next c

  activeWorkbook.saved = true

end sub

