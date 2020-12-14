Option Explicit

Dim arrSheet, intCount

' Read and display columns A,B, rows 2..6 of "ReadExcelTest.xlsx"
arrSheet = ReadExcel( "ReadExcelTest.xlsx", "Sheet1", "A1", "B6", True )
For intCount = 0 To UBound( arrSheet, 2 )
    WScript.Echo arrSheet( 0, intCount ) & vbTab & arrSheet( 1, intCount )
Next

WScript.Echo "==============="

' An alternative way to get the same results
arrSheet = ReadExcel( "ReadExcelTest.xlsx", "Sheet1", "A2", "B6", False )
For intCount = 0 To UBound( arrSheet, 2 )
    WScript.Echo arrSheet( 0, intCount ) & vbTab & arrSheet( 1, intCount )
Next