'
'   ..\..\runVBAFilesInOffice.vbs -excel rand -c Go
'
'   Compare with -> randbetween()

sub Go()

  range("a1").formula = "=rand()"

end sub

