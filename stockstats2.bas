Attribute VB_Name = "Module2"
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:
'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

Sub StockStats2()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets

    'Labelling the Table
    ws.Cells(2, 14).Value = "Greatest Percent Increase"
    ws.Cells(3, 14).Value = "Greatest Percent Decrease"
    ws.Cells(4, 14).Value = "Greatest Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"

    'put the values in the ws.Cells
    ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(ws.Range("k:k"))
    ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(ws.Range("k:k"))
    ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(ws.Range("l:l"))
    
            
    'grabbing the ticker names w a for loop
    For k = 2 To ws.UsedRange.Rows.Count
    
        If ws.Cells(k, 11).Value = ws.Cells(2, 16).Value Then
            ws.Cells(2, 15).Value = ws.Cells(k, 9).Value
        End If
        
        If ws.Cells(k, 11).Value = ws.Cells(3, 16).Value Then
            ws.Cells(3, 15).Value = ws.Cells(k, 9).Value
        End If
        
        If ws.Cells(k, 12).Value = ws.Cells(4, 16).Value Then
            ws.Cells(4, 15).Value = ws.Cells(k, 9).Value
        End If
        
    Next k
            
    Next ws

End Sub


