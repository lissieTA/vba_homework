Attribute VB_Name = "Module1"
Sub vbaHomework()
Dim WS_Count As Integer
Dim I As Integer
' Set WS_Count equal to the number of worksheets in the active
' workbook.
WS_Count = ActiveWorkbook.Worksheets.Count
Dim lRow As Long
Dim lCol As Long

Dim ticker As String
Dim pointer As Double
Dim totalVolume As Double
Dim resultsPointer As Double
Dim startingPrice, closingPrice As Double

Dim Sheet As Worksheet

' Begin the loop.
' testing

For I = 1 To WS_Count
Set Sheet = ActiveWorkbook.Worksheets(I)
'Finds the last non-blank cell in a single row or column
'Find the last non-blank cell in column A(1)
lRow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
      
pointer = 2
ticker = Sheet.Cells(pointer, 1).Value()
totalVolume = CDbl(Sheet.Cells(pointer, 7).Value())

resultsPointer = 2

Sheet.Cells(1, 9).Value() = "Ticker"
Sheet.Cells(1, 10).Value() = "Total Volume"

Do While pointer <= lRow

    Do While ticker = Sheet.Cells(pointer + 1, 1).Value()
        totalVolume = totalVolume + Sheet.Cells(pointer + 1, 7).Value()
        ticker = Sheet.Cells(pointer + 1, 1).Value()
        pointer = pointer + 1
    Loop


    Sheet.Cells(resultsPointer, 9).Value() = ticker
    Sheet.Cells(resultsPointer, 10).Value() = totalVolume
    ticker = Sheet.Cells(pointer + 1, 1).Value()
    pointer = pointer + 1
    totalVolume = 0
    resultsPointer = resultsPointer + 1
    
Loop

Next I

End Sub
