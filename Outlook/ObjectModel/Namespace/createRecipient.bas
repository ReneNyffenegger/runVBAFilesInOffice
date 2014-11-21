'
'    ..\..\..\runVBAFilesInOffice.vbs -excel -ol createRecipient -c main
'
sub main()

    dim outl as outlook.application
    dim nmsp as outlook.namespace
    dim cald as outlook.folder
    dim rcpt as outlook.recipient


    set outl = new outlook.application
    set nmsp = outl.getNamespace("MAPI")

    set rcpt = nmsp.createRecipient("Nyffenegger Rene")

    rcpt.resolve

    if rcpt.resolved then

       set fold = nmsp.getSharedDefaultFolder(rcpt, olFolderCalendar)

       fold.display

    else

       msGBox("Could not resolve name")

    end if

    activeWorkbook.saved = true

end sub
