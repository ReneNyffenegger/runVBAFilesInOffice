'
'   ..\..\..\runVBAFilesInOffice.vbs -excel fileFormat -c main
'

public sub main()

    cells(1, 1).value = "File Format:"

    select case activeWorkbook.fileFormat

      case xlAddIn                       : cells(1,2).value = "xlAddIn"
      case xlAddIn8                      : cells(1,2).value = "xlAddIn8"
      case xlCSV                         : cells(1,2).value = "xlCSV"
      case xlCSVMac                      : cells(1,2).value = "xlCSVMac"
      case xlCSVMSDOS                    : cells(1,2).value = "xlCSVMSDOS"
      case xlCSVWindows                  : cells(1,2).value = "xlCSVWindows"
      case xlCurrentPlatformText         : cells(1,2).value = "xlCurrentPlatformText"
      case xlDBF2                        : cells(1,2).value = "xlDBF2"
      case xlDBF3                        : cells(1,2).value = "xlDBF3"
      case xlDBF4                        : cells(1,2).value = "xlDBF4"
      case xlDIF                         : cells(1,2).value = "xlDIF"
      case xlExcel12                     : cells(1,2).value = "xlExcel12"
      case xlExcel2                      : cells(1,2).value = "xlExcel2"
      case xlExcel2FarEast               : cells(1,2).value = "xlExcel2FarEast"
      case xlExcel3                      : cells(1,2).value = "xlExcel3"
      case xlExcel4                      : cells(1,2).value = "xlExcel4"
      case xlExcel4Workbook              : cells(1,2).value = "xlExcel4Workbook"
      case xlExcel5                      : cells(1,2).value = "xlExcel5"
      case xlExcel7                      : cells(1,2).value = "xlExcel7"
      case xlExcel8                      : cells(1,2).value = "xlExcel8"
      case xlExcel9795                   : cells(1,2).value = "xlExcel9795"
      case xlHtml                        : cells(1,2).value = "xlHtml"
      case xlIntlAddIn                   : cells(1,2).value = "xlIntlAddIn"
      case xlIntlMacro                   : cells(1,2).value = "xlIntlMacro"
      case xlOpenDocumentSpreadsheet     : cells(1,2).value = "xlOpenDocumentSpreadsheet"
      case xlOpenXMLAddIn                : cells(1,2).value = "xlOpenXMLAddIn"
      case xlOpenXMLTemplate             : cells(1,2).value = "xlOpenXMLTemplate"
      case xlOpenXMLTemplateMacroEnabled : cells(1,2).value = "xlOpenXMLTemplateMacroEnabled"
      case xlOpenXMLWorkbook             : cells(1,2).value = "xlOpenXMLWorkbook"
      case xlOpenXMLWorkbookMacroEnabled : cells(1,2).value = "xlOpenXMLWorkbookMacroEnabled"
      case xlSYLK                        : cells(1,2).value = "xlSYLK"
      case xlTemplate                    : cells(1,2).value = "xlTemplate"
      case xlTemplate8                   : cells(1,2).value = "xlTemplate8"
      case xlTextMac                     : cells(1,2).value = "xlTextMac"
      case xlTextMSDOS                   : cells(1,2).value = "xlTextMSDOS"
      case xlTextPrinter                 : cells(1,2).value = "xlTextPrinter"
      case xlTextWindows                 : cells(1,2).value = "xlTextWindows"
      case xlUnicodeText                 : cells(1,2).value = "xlUnicodeText"
      case xlWebArchive                  : cells(1,2).value = "xlWebArchive"
      case xlWJ2WD1                      : cells(1,2).value = "xlWJ2WD1"
      case xlWJ3                         : cells(1,2).value = "xlWJ3"
      case xlWJ3FJ3                      : cells(1,2).value = "xlWJ3FJ3"
      case xlWK1                         : cells(1,2).value = "xlWK1"
      case xlWK1ALL                      : cells(1,2).value = "xlWK1ALL"
      case xlWK1FMT                      : cells(1,2).value = "xlWK1FMT"
      case xlWK3                         : cells(1,2).value = "xlWK3"
      case xlWK3FM3                      : cells(1,2).value = "xlWK3FM3"
      case xlWK4                         : cells(1,2).value = "xlWK4"
      case xlWKS                         : cells(1,2).value = "xlWKS"
      case xlWorkbookDefault             : cells(1,2).value = "xlWorkbookDefault"
      case xlWorkbookNormal              : cells(1,2).value = "xlWorkbookNormal"
      case xlWorks2FarEast               : cells(1,2).value = "xlWorks2FarEast"
      case xlWQ1                         : cells(1,2).value = "xlWQ1"
      case xlXMLSpreadsheet              : cells(1,2).value = "xlXMLSpreadsheet"
    end select

    columns(1).autoFit

    activeWorkbook.saved = true

end sub
