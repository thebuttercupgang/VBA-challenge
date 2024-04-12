Attribute VB_Name = "Module1"
Sub StockStats()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
    'create column headers
    ws.Cells(1, 9).Value = "Ticker Name"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'create the variables
    Dim Ticker_Name As String
    Dim Yearly_Change, Percent_Change, Total_Vol As Double 'code won't run unless this is a double idk
    Total_Vol = 0
    
    'name the row that the table will begin
    Dim summary_table As Integer
    summary_table = 2

    'begin the for loop
    Dim open_value As Double
    open_value = ws.Cells(2, 3).Value
    
    For i = 2 To ws.UsedRange.Rows.Count
    
        'if the previous ticker does not match the next ticker, then
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'putting unique tickers in a column
            Ticker_Name = ws.Cells(i, 1).Value
            ws.Range("I" & summary_table).Value = Ticker_Name
            
            'putting yearly change in a column
            'unsure if this is the "proper" way to do it, but there are only 252 trading days in the stock market sooo
            Yearly_Change = ws.Cells(i, 6).Value - open_value
            
            ws.Range("J" & summary_table).Value = Yearly_Change
           
            Percent_Change = Yearly_Change / open_value
            ws.Range("K" & summary_table).Value = Percent_Change
            open_value = ws.Cells(i + 1, 3).Value
            
            'put total_change in a column
            Total_Vol = Total_Vol + ws.Cells(i, 7).Value
            ws.Range("L" & summary_table).Value = Total_Vol
            Total_Vol = 0
            
            summary_table = summary_table + 1
            
        Else
            Total_Vol = Total_Vol + ws.Cells(i, 7).Value
            
        End If
        
        'set interior colors based on increase or decrease
        If ws.Cells(i, 11).Value > 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 11).Interior.ColorIndex = 3
        End If
        
    Next i
    
    ws.Columns("K").NumberFormat = "0.00%"

Next ws

End Sub
