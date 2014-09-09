'
'   ..\..\runVBAFilesInOffice.vbs -excel sumif -c main
'
sub main()

    cells(1,1) = 13: cells(1,2) = 1
    cells(2,1) = 21: cells(2,2) = 0
    cells(3,1) = 34: cells(3,2) = 0
    cells(4,1) = 47: cells(4,2) = 1
    cells(5,1) = 56: cells(5,2) = 0

  ' Calculate the sum of numbers found in a1:a5 whose corresponding
  ' value in b1:b5 equals 1.
  '
  ' The sum will be 60 = 13 + 47
    cells(7,1).formula = "=sumif(b1:b5, 1, a1:a5)"

    activeWorkbook.saved = true

end sub
