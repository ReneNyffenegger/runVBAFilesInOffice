' ..\..\..\runVBAFilesInOffice.vbs -word %CD%\typeText -c main

sub main()

  selection.typeText "Hello World"

  selection.typeParagraph

  selection.typeText "The number is, of course, forty-two"

  selection.typeText chr(13)

  selection.typeText "Another line"

  activeDocument.saved = true

end sub
