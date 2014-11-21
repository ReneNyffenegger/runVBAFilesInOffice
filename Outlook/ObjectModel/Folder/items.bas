'
'   ..\..\..\runVBAFilesInOffice -excel -ol appointments -c main
'
sub main

    dim outl as outlook.application
    dim nmsp as outlook.namespace
    dim cald as outlook.folder
    dim aptm as outlook.appointmentItem
    dim atta as outlook.attachment

    set outl = new outlook.application
    set nmsp = outl.getNamespace("MAPI")

    set cald = nmsp.getDefaultFolder(olFolderCalendar)

    dim c as integer: c = 0
    for each aptm in cald.items

        c = c+1

        cells(c, 1).value = aptm.start
        cells(c, 2).value = aptm.end
        cells(c, 3).value = aptm.body

        dim r as integer: r = 3

        for each atta in aptm.attachments

            r = r+1

            if     atta.type = olByValue then
                   cells(c, r).value = atta.fileName
                   cells(c, r).value = atta.DisplayName
                   
            elseif atta.type = olEmbeddedItem then
                   cells(c, r).value = "embedded .msg"

            elseif atta.type = olOLE then
                   cells(c, r).value = "OLE document" & atta.DisplayName

            end if
                   

        next
        

    next aptm

    activeWorkbook.saved =true


end sub
