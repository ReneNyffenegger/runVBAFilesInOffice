'
'   ..\..\runVBAFilesInOffice.vbs -excel rand -c Go
'

sub Go()

  range("a1").formula = "=rand()"

end sub

