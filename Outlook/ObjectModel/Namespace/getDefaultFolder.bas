'
'    ..\..\..\runVBAFilesInOffice.vbs -excel -ol getDefaultFolder -c main
'
private sub e(                                     _
       outl        as outlook.application,         _
       nmsp        as outlook.namespace,           _
       folder_type as outlook.olDefaultFolders     _
    )

    dim fold as outlook.folder
    dim expl as outlook.explorer

    set fold = nmsp.getDefaultFolder(folder_type)
    set expl = outl.explorers.add(fold, olFolderDisplayNormal)

    expl.activate 

end sub

sub main()

    dim outl as outlook.application
    dim nmsp as outlook.namespace

    set outl = new outlook.application
    set nmsp = outl.getNamespace("MAPI")

    call e(outl, nmsp, olFolderCalendar               )
    call e(outl, nmsp, olFolderConflicts              )
    call e(outl, nmsp, olFolderContacts               )
    call e(outl, nmsp, olFolderDeletedItems           )
    call e(outl, nmsp, olFolderDrafts                 )
    call e(outl, nmsp, olFolderInbox                  )
    call e(outl, nmsp, olFolderJournal                )
    call e(outl, nmsp, olFolderJunk                   )
    call e(outl, nmsp, olFolderLocalFailures          )
    call e(outl, nmsp, olFolderManagedEmail           )
    call e(outl, nmsp, olFolderNotes                  )
    call e(outl, nmsp, olFolderOutbox                 )
    call e(outl, nmsp, olFolderSentMail               )
    call e(outl, nmsp, olFolderServerFailures         )
    call e(outl, nmsp, olFolderSuggestedContacts      )
    call e(outl, nmsp, olFolderSyncIssues             )
    call e(outl, nmsp, olFolderTasks                  )
    call e(outl, nmsp, olFolderToDo                   )
    call e(outl, nmsp, olPublicFoldersAllPublicFolders)
    call e(outl, nmsp, olFolderRssFeeds               )

    activeWorkbook.saved = true 

end sub
