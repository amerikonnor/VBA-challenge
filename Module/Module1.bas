Attribute VB_Name = "Module1"
Sub stock_data()

Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim change_price As Double
Dim change As Double
Dim total_volume As Double

Dim lastRow As Long

Dim summary_row As Integer

Dim max_percent_ticker As Integer
Dim min_percent_ticker As Integer
Dim max_volume_ticker As Integer

max_percent_ticker = 2
min_percent_ticker = 2
max_volume_ticker = 2

For Each ws In Worksheets

ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"

summary_row = 2

total_volume = 0
ticker = ws.Cells(2, 1).Value
open_price = ws.Cells(2, 3).Value

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow
    
    If ws.Cells(i + 1, 1).Value <> ticker Then
        
        ws.Cells(summary_row, 10).Value = ticker
        
        close_price = ws.Cells(i, 6).Value
        change_price = close_price - open_price
        ws.Cells(summary_row, 11).Value = change_price
        
        If change_price < 0 Then
            ws.Cells(summary_row, 11).Interior.ColorIndex = 3
            
        ElseIf change_price > 0 Then
            ws.Cells(summary_row, 11).Interior.ColorIndex = 4
        
        End If
        
        If open_price = 0 Then
            ws.Cells(summary_row, 12).Value = 0
        Else
            ws.Cells(summary_row, 12).Value = change_price / open_price
        End If
        
        ws.Cells(summary_row, 12).NumberFormat = "0.00%"
        
        If ws.Cells(summary_row, 12).Value > ws.Cells(max_percent_ticker, 12) Then
            max_percent_ticker = summary_row
        End If
        
        If ws.Cells(summary_row, 12).Value < ws.Cells(min_percent_ticker, 12) Then
            min_percent_ticker = summary_row
        End If
        
        total_volume = total_volume + ws.Cells(i, 7).Value
        ws.Cells(summary_row, 13).Value = total_volume
        
        
        If total_volume > ws.Cells(max_volume_ticker, 13) Then
            max_volume_ticker = summary_row
        End If
        
        
        
        
        ticker = ws.Cells(i + 1, 1).Value
        open_price = ws.Cells(i + 1, 3).Value
        total_volume = 0
        summary_row = summary_row + 1
        
        
    
    
    Else
    
        total_volume = total_volume + ws.Cells(i, 7).Value
    
    End If

Next i

    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(2, 17).Value = ws.Cells(max_percent_ticker, 10).Value
    ws.Cells(2, 18).Value = ws.Cells(max_percent_ticker, 12).Value
    ws.Cells(2, 18).NumberFormat = "0.00%"
    
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(3, 17).Value = ws.Cells(min_percent_ticker, 10).Value
    ws.Cells(3, 18).Value = ws.Cells(min_percent_ticker, 12).Value
    ws.Cells(3, 18).NumberFormat = "0.00%"
    
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    ws.Cells(4, 17).Value = ws.Cells(max_volume_ticker, 10).Value
    ws.Cells(4, 18).Value = ws.Cells(max_volume_ticker, 13).Value

Next ws

End Sub
