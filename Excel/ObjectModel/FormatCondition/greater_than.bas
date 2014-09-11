'  ..\..\..\runVBAFilesInOffice.vbs -excel greater_than -c main

public sub main()

   dim cond as formatCondition

   range("a1:a10").formula = "=rand()"

   set cond = range("a1:a10").formatConditions.add(type := xlCellValue, operator := xlGreater, formula1 := "=0.5")

   cond.font.color        = rgb(255,   0,   0)
   cond.interior.color    = rgb(200, 200, 200)

   activeWorkbook.saved = true

end sub
