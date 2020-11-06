Attribute VB_Name = "VBA_Challenge"

Sub VBA_Challenge()
    
'Set Variables
'create loop to sort through each worksheet in workbook
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
Dim TickerSymbol As String
Dim YearlyChange As Double
Dim Table_Row As Integer
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim Counter As Long
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim LastRow As Long
Dim i As Long

'Set headers, last row, tablerow, intital openprice
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "TotalStockVolume"
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
TableRow = 2
TotalStockVolume = 0
OpenPrice = ws.Cells(TableRow, 3).Value

'set i to loop through rows
For i = 2 To LastRow

'use conditional to find the first and last iteration of each ticker & pull different tickers from column A to new location
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    TickerSymbol = ws.Cells(i, 1).Value
    ws.Range("I" & TableRow).Value = TickerSymbol
    
'Subtract opening value(first)-closing value (last) to calculate yearlychange
'Record yearlychange in new location
    ClosePrice = ws.Cells(i, 6).Value
    YearlyChange = ClosePrice - OpenPrice
    ws.Range("J" & TableRow).Value = YearlyChange
    
'Use yearlychange to calculate percent change
'Use conditional to eliminate potential errors in calculation
'Set correct formatting for percentchange values
'Record percentchange in new location
         If OpenPrice = 0 Then
                PercentChange = 0
                 ws.Range("K" & TableRow).NumberFormat = "0.00%"
                 ws.Range("K" & TableRow).Value = PercentChange
            Else
                PercentChange = (YearlyChange / OpenPrice)
                ws.Range("K" & TableRow).Value = PercentChange
                ws.Range("K" & TableRow).NumberFormat = "0.00%"
     End If
    
    
  
'Calculate totalstockvolume
'Record totalstockvolume in new location
    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    ws.Range("L" & TableRow).Value = TotalStockVolume
    
'Reset values
    TableRow = TableRow + 1
    OpenPrice = ws.Cells(i + 1, 3).Value
    TotalStockVolume = 0
    
       Else
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
                
            
            End If
            
    Next i
    
'Create loop through yearly change for color formatting
'Use conditional to reference yearly change value for color coding; if >0 then green, if <0 red, if = 0 no fill
YearlyChangeLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
For j = 2 To YearlyChangeLastRow
    If ws.Cells(j, 10).Value > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
                 Else
                    ws.Cells(j, 10).Interior.ColorIndex = 0
                
                         End If
            Next j
    Next ws
    
End Sub

    
    
 
 
  






