'
' ..\runVBAFilesInOffice.vbs -excel SetValueInRowAndColumn -c Run
'
public sub Run() ' {

  dim row as long
  dim col as long

  for row = 1 to 10
  for col = 1 to row

      cells(row, col) = row * col

  next col
  next row

end sub ' }
