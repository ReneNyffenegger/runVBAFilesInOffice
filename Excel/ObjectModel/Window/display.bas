
'
'   ..\..\..\runVBAFilesInOffice.vbs -excel display -c main
'

public sub main()

    dim w as window

    set w = application.activeWindow

    cells(1,1).value = "=3+4"
  ' Make formula visible
    w.displayFormulas            = true

    w.displayGridlines           = false

  ' let the column names (A ... ) and row numbers (1 ...)
  ' disappear
    w.displayHeadings            = false

  ' no scrollbars
    w.displayHorizontalScrollbar = false
    w.displayVerticalScrollbar   = false

    range( cells(2,1), cells(5,1) ).rows.group
  ' dont show grouping symbols (aka «outline»).
    w.displayOutline             = false

    w.displayRightToLeft         = true

    w.displayRuler               = true

    w.displayWorkbookTabs        = false

    cells(2,1).value = 0
    cells(3,1).value = 1
    w.displayZeros               = false

    activeWorkbook.saved         = true

end sub
