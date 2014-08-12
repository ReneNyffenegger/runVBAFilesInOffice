'
'   ..\runVBAFilesInOffice.vbs -word for -c main
'
sub main()

    for x = 1 to 10

        selection.typeText text := x
        selection.typeParagraph

    next x


    for y = -1 to 1 step 0.1

        selection.typeText text := y
        selection.typeParagraph

    next y

    activeDocument.saved = true

end sub
