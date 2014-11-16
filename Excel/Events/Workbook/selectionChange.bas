'
'  ..\..\..\runVBAFilesInOffice -excel -vbe selectionChange -c main

sub main()

  ' VBIDE needs reference {0002E157-0000-0000-C000-000000000046} «Microsoft Visual Basic for Applications Extensibility» 
  '(the -vbe flag in runVBAFilesInOffice)
'   dim vbe_ as VBIDE.VBE      

    dim codeMod  as VBIDE.codeModule
    dim codeLine as long


  ' Get code module for «active worksheet»
    set codeMod = activeWorkbook.VBProject.VBComponents(activeSheet.name).codeModule

    codeLine = codeMod.countOflines

    codeLine = codeLine + 1
    codeMod.insertLines codeLine, "private sub worksheet_selectionChange(byVal target as range)"

    codeLine = codeLine + 1
'   codeMod.insertLines codeLine, "  msgBox(""x"")"
    codeMod.insertLines codeLine, "  cells(2,1).value=""You clicked "" & target.address"

    codeLine = codeLine + 1
    codeMod.insertLines codeLine, "end sub"

    cells(1,1).value = "Try clicking into cells"

    activeWorkbook.saved = true

end sub
