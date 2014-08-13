'
'   ..\runVBAFilesInOffice.vbs -word dim -c d
'
sub d()

    dim q(10) as integer
    dim s     as integer

    for i = 1 to 10
        q(i) = i
    next i

    s = 0
    for i = 1 to 10
        s = s + q(i)
    next i

    call selection.typeText("Sum is: " & s)
      
    activeDocument.saved = 1

end sub
