' ..\..\..\runVBAFilesInOffice.vbs -excel -ol getGlobalAddressList -c main

sub main()

    dim outl as outlook.application
    dim nmsp as outlook.namespace
    dim glal as outlook.addressList
    dim ents as outlook.addressEntries
    dim entr as outlook.addressEntry

    set outl = new outlook.application
    set nmsp = outl.getNamespace("MAPI")
    set glal = nmsp.getGlobalAddressList()
    set ents = glal.addressEntries

    dim c as integer: c=0

    for each entr in ents

        c = c+1

        cells(c, 1).value = entr.name
        cells(c, 2).value = entr.address

    next entr

    activeWorkbook.saved = true

end sub
