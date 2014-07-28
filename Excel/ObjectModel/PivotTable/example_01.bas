'
'   ..\..\..\runVBAFilesInOffice.vbs -excel example_01 -c Go %CD%
'

Sub Go(cur_working_dir as string)

    dim pivot_sheet            as workSheet
    dim pivot_cache            as pivotCache
    dim pivot_table            as pivotTable
    
    dim pivot_table_upper_left as range
    dim pf_col_1               as pivotField
    dim pf_col_2               as pivotField
'

    call importCSV(cur_working_dir & "\pivot.csv", activeSheet, range("$a$1"), "csv_data")
'
    set pivot_sheet = sheets.add

    set pivot_cache = activeWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= "csv_data", Version:=xlPivotTableVersion14)

    set pivot_table_upper_left = pivot_sheet.range("C3")
        
    set pivot_table = pivot_cache.CreatePivotTable ( TableDestination:= pivot_table_upper_left )


    set pf_col_1 = pivot_table.pivotFields("col_1")
    set pf_col_2 = pivot_table.pivotFields("col_2")

    pf_col_1.orientation = xlRowField
    pf_col_2.orientation = xlColumnField

    call pivot_table.addDataField (pf_col_2, "Count of col_2", xlCount)

    activeWorkbook.saved = true

End Sub

private sub importCSV(csv_file_name as string, sheet_ as workSheet, range_ as range, name_ as string) ' { 
  '
  ' -> https://github.com/ReneNyffenegger/runVBAFilesInOffice/blob/master/Excel/ObjectModel/QueryTable/load_CSV.bas
  '
  

    With ActiveSheet.QueryTables.Add(                 _ 
               Connection:= "TEXT;" & csv_file_name , _
               Destination:=range_)

        .name                 = name_
        .fieldNames = True
        .rowNumbers = False
        .preserveFormatting = True
        .textFilePlatform = 437
        .textFileStartRow = 1
        .textFileParseType = xlDelimited
        .textFileTextQualifier = xlTextQualifierDoubleQuote
        .textFileConsecutiveDelimiter = False
        .textFileCommaDelimiter = True
        .textFileTrailingMinusNumbers = True
        .refresh BackgroundQuery:=False

    end with


end sub

