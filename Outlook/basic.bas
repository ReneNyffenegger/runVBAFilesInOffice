'
'   ..\runVBAFilesInOjfice.vbs -excel -ol basic -c main
'
'   Starts and displays outlook.
'
sub main()

    dim outl as outlook.application
    dim expl as outlook.explorer
    dim nmsp as outlook.namespace
    dim fold as outlook.folder
        

    set outl = new outlook.application

  ' Currently, only parameter for getNamespace is «MAPI».
  '
  ' Alternatively, «set nmsp = outl.session» would also
  ' work.
    set nmsp = outl.getNamespace("MAPI")


    set fold = nmsp.getDefaultFolder(olFolderInbox)
    set expl = outl.explorers.add(fold, olFolderDisplayNormal)
        
    expl.Display
  ' expl.Activate

  ' outl.Quit

    activeWorkbook.saved = true


end sub

