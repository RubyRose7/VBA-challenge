Attribute VB_Name = "Module1"

Sub VBAChallenge():

Dim ticker As String
Dim number_tickers As Integer
Dim LastRow As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double

' loop for worksheet
For Each ws In Worksheets
Dim Worsheet As String
 
 ' Get the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
' Column Titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Variables
    number_tickers = 0
    ticker = " "
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    ' loop through the list of tickers.
    For i = 2 To LastRow

        ' Value for ticker symbol
        ticker = ws.Cells(i, 1).Value
        
        ' Opening price
        If opening_price = 0 Then
            opening_price = ws.Cells(i, 3).Value
        End If
        
        ' Total stock volume
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ticker Then
           number_tickers = number_tickers + 1
            ws.Cells(number_tickers + 1, 9) = ticker
            
            ' Closing price
            closing_price = ws.Cells(i, 6).Value
            
            'Yearly Change
            yearly_change = closing_price - opening_price
            
            ws.Cells(number_tickers + 1, 10).Value = yearly_change
            
            ' Colors
            If yearly_change > 0 Then
                ws.Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
           
            ElseIf yearly_change < 0 Then
                ws.Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
        
           
            End If
            
            
            'Percent change
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            
            ' Percent_change as a percentage
            ws.Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
             ' Set opening price back to 0
            opening_price = 0
            
            ' Add total stock volume value
            ws.Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            ' Set total stock volume back to 0
            total_stock_volume = 0
        End If
        
         
    
    
Next i
    
Next ws

    
End Sub


