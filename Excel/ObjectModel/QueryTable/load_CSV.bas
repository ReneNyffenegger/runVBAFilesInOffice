'
'   ..\..\..\runVBAFilesInOffice.vbs -excel load_CSV -c Go %CD%\data.csv
'

public sub Go(csv_file_name as string) ' {

  call ImportCSV(csv_file_name := csv_file_name              , _
                 sheet_        := activeSheet                , _
                 range_        := activeSheet.Range("$A$1")  , _
                 name_         :="CSVData" )

  activeWorkbook.saved = true

end sub ' }

private sub ImportCSV(csv_file_name as string, sheet_ as workSheet, range_ as range, name_ as string) ' {

    With ActiveSheet.QueryTables.Add(                 _ 
               Connection:= "TEXT;" & csv_file_name , _
               Destination:=range_)

        .name                 = name_
        .fieldNames = True
        .rowNumbers = False
'       .fillAdjacentFormulas = False
        .preserveFormatting = True
'       .refreshOnFileOpen = False
'       .refreshStyle = xlInsertDeleteCells
'       .savePassword = False
'       .saveData = True
'       .adjustColumnWidth = True
'       .refreshPeriod = 0
'       .textFilePromptOnRefresh = False
        .textFilePlatform = 437
        .textFileStartRow = 1
        .textFileParseType = xlDelimited
        .textFileTextQualifier = xlTextQualifierDoubleQuote
        .textFileConsecutiveDelimiter = False
'       .textFileTabDelimiter = True
'       .textFileSemicolonDelimiter = False
        .textFileCommaDelimiter = True
'       .textFileSpaceDelimiter = False
'       .textFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .textFileTrailingMinusNumbers = True
        .refresh BackgroundQuery:=False

    end with


end sub

