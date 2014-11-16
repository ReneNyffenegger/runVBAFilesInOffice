'
'    ..\..\..\runVBAFilesInOffice.vbs -excel -vbe VBProjects -c main
'

sub main()

  ' VBE needs reference {0002E157-0000-0000-C000-000000000046} «Microsoft Visual Basic for Applications Extensibility» 
  '(the -vbe flag in runVBAFilesInOffice)
    dim vbe_ as VBE      

    dim r    as integer

    set vbe_ = application.VBE

   
    r = 1
    for each VBProject_ in vbe_.VBProjects 

        cells(r, 1).value = VBProject_.name

        r = r+1

    next VBProject_

    activeWorkbook.saved = true

end sub
