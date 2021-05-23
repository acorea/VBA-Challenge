Attribute VB_Name = "Module1"
Sub StockData()

For Each ws In Worksheets

'Define Variables
Dim Ticker As String
Dim TickerCount As Integer
    TickerCount = 0
Dim TotalStockVolume As LongLong
    Volume = 0
Dim SummaryRow As Integer
    SummaryRow = 2
    

'Add summary headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Determine the quantitiy of rows
Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row


    'Loop through rows
    For i = 2 To Last_Row
    
        'Look for ticker change in proceeding row if not keep going
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'set ticker
            Ticker = ws.Cells(i, 1).Value
            
            'Add total stock volume for year
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
            'Add ticker symbol to summary section
            ws.Cells(SummaryRow, 9).Value = Ticker
        
            'add total ticker volume to summary section
            ws.Cells(SummaryRow, 12).Value = TotalStockVolume
        
            'calculate yearly change add to summary section
            ws.Cells(SummaryRow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i - TickerCount, 3).Value

    
                'add color fill to yearly change
                If ws.Cells(SummaryRow, 10).Value > 0 Then
                    
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                
                ElseIf ws.Cells(SummaryRow, 10).Value < 0 Then
        
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
            
                'no color for no yearly change
                Else
            
                End If
            
                'Print percentage change
                If ws.Cells(i - TickerCount, 3).Value = 0 Then
            
                ws.Cells(SummaryRow, 11).Value = 0
                Else
                ws.Cells(SummaryRow, 11).Value = Format((ws.Cells(i, 6).Value - ws.Cells(i - TickerCount, 3).Value) / ws.Cells(i - TickerCount, 3).Value, "Percent")
        
                End If
        
        'move on to next row in summary section
        SummaryRow = SummaryRow + 1
        
        'reset Total ticker volume to 0
        TotalStockVolume = 0
        
        'clear ticker count and move to next
        TickerCount = 0
        
    'Continue adding if ticker symbol stays the same
    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    
    TickerCount = TickerCount + 1
    
    End If
    
    Next i
    
    Next ws
  
End Sub
