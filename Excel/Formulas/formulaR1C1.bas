'
'   ..\..\runVBAFilesInOffice.vbs -excel formulaR1C1 -c Go
'
sub Go()

    range("b2").formula = "=rand()"
    range("c2").formula = "=rand()"

  ' Note, the formula turns into
  '    «  =IF( B2 < C2; "less"; "greater or equal" )  »
  ' in the produced formula
    range("d2").formulaR1C1 = "=if( rc[-2] < rc[-1] , ""less"" , ""greater or equal"" )"

end sub
