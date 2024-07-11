VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_data()

For Each ws In Worksheets
    
    Dim WorksheetName As String
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    WorksheetName = ws.Name
    
    Dim Ticker_symbol As String
    
    Dim Stock_volume As Double
    
    Dim quarterly_change As Double
    
    Dim Date_market As Date
    
    Dim Percent_change As Double
    
    quarterly_change = 0
    
    Stock_volume = 0
    
    Dim Summary_Table_Row As Integer
    
    Summary_Table_Row = 2
    
    ws.Cells(1, 9).Value = "Ticker"
    
    ws.Cells(1, 10).Value = "Quarterly Change"
    
    ws.Cells(1, 11).Value = "Percent Change"
    
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    
    ws.Cells(1, 17).Value = "Value"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    
    ws.Cells(4, 15).Value = "Greatest total Volume"
    
    
    Open_price = ws.Cells(2, 3).Value
    

    For i = 2 To lastrow
    
    


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Ticker_symbol = ws.Cells(i, 1).Value
    
    Stock_volume = Stock_volume + ws.Cells(i, 7)
    
    quarterly_change = (ws.Cells(i, 6).Value - Open_price)
    
    
    
    ws.Range("I" & Summary_Table_Row).Value = Ticker_symbol
    
    ws.Range("J" & Summary_Table_Row).Value = quarterly_change
    
    ws.Range("L" & Summary_Table_Row).Value = Stock_volume
    
    Percent_change = (quarterly_change / Open_price)

    ws.Range("K" & Summary_Table_Row).Value = Percent_change

    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    Stock_volume = 0
    
    Open_price = ws.Cells(i + 1, 3).Value
    
Else

    Stock_volume = Stock_volume + ws.Cells(i, 7)

    End If

If ws.Cells(i, 10).Value > 0 Then
    
        ws.Cells(i, 10).Interior.ColorIndex = 4 ' Green
        
    ElseIf ws.Cells(i, 10).Value < 0 Then
    
        ws.Cells(i, 10).Interior.ColorIndex = 3 ' Red
        
    Else
        
        ws.Cells(i, 10).Interior.ColorIndex = 0 ' No color (default)
        
        End If
    
    Next i
    
    ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
    
    ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
    
    ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
    
    ws.Range("Q2").NumberFormat = "0.00%"
    
    ws.Range("Q3").NumberFormat = "0.00%"
    
    max_index = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & lastrow), 0)
    
    min_index = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & lastrow), 0)
    
    max_volume_index = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & lastrow), 0)
    
    ws.Cells(2, 16).Value = ws.Cells(max_index + 1, 9).Value
    
    ws.Cells(3, 16).Value = ws.Cells(min_index + 1, 9).Value
    
    ws.Cells(4, 16).Value = ws.Cells(max_volume_index + 1, 9).Value
    
    Next ws

End Sub



