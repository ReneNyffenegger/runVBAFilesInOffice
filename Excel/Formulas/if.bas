'
'   ..\..\runVBAFilesInOffice.vbs -excel if -c Go
'
sub Go()

    range("a1:b10").formula = "=rand()"

  ' Note:  «a1<b1» dynamically changes with the rows in which
  '        they appear.
    range("c1:c10").formula = "=if( a1<b1 , ""less"" , ""greater or equal"")"

end sub
