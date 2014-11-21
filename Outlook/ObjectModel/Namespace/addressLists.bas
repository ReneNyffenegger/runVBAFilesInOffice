'
'   ..\..\..runVBAFilesInOffice.vbs -excel -ol addressLists -c main
'
sub main

    dim outl       as outlook.application
    dim nmsp       as outlook.namespace

    dim addrLists  as outlook.addressLists
    dim addrList   as outlook.addressList


    set outl = new outlook.application
    set nmsp = outl.getNamespace("MAPI")

    set addrLists = nmsp.addressLists

    dim  c as integer: c=0

    for  each addrList in addrLists

         c=c+1

         cells(c, 1).value = addrList.name

    next addrList

    activeWorkbook.saved = true

end sub
