'
'   ..\..\runVBAFilesInOffice.vbs -excel randBetween -c Go
'
'   Note, the function returns an integer.
'
'   Compare with -> rand()

sub Go()

  range("a1").formula = "=randBetween(100, 110)"
  range("a2").formula = "=randBetween(100, 110)"
  range("a3").formula = "=randBetween(100, 110)"

end sub

