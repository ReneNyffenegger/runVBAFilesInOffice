'
'   ..\runVBAFilesInOffice.vbs -word for -c main
'
sub main()

    for x = 1 to 10

        selection.typeText text := x
        selection.typeParagraph

    next x

    activeDocument.saved = true

end sub
