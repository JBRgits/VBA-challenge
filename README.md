# VBA-challenge code for VBA Challenge Homework using Multi Year Stock Data and a Challenge Summary Table, across Worksheets

Sub Ticker()

'Summarize the ticker symbols and put them in column I
'Calculate the cumulative volume of stock for each ticker symbol and put the value in column L
'Calculate the difference between the Year End Close and Year Beginning Open and put the value in column J
'Conditional format cells in column J to be red if less than 0 and green if greater than zero
'Calculate percentage change of column J divided by Year Beginning Open and put the value in column K
'Add .End(xlUp).Row
'Add coding to allow subroutine to run across all worksheets

Dim ws As Worksheet

For Each ws In Worksheets

Dim TickerRow As Integer
TickerRow = 2

Dim StockVolTotal As Double
StockVolTotal = 0

Dim YearlyChange As Double
Dim OpeningPrice As Double
Dim ClosingPrice As Double

Dim PercentChange As Double

Dim GrtInc As Double
GrtInc = 0
GrtIncTicker = " "

Dim GrtDec As Double
GrtDec = 99999999
GrtDecTicker = " "

Dim StkMaxVol As Double
StkMaxVol = 0
MaxVolTicker = " "

Dim LastRow As Double



LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

    StockVolTotal = StockVolTotal + ws.Cells(i, 7).Value
    OpeningPrice = ws.Cells(TickerRow, 3).Value


    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
        ws.Cells(TickerRow, 9).Value = ws.Cells(i, 1).Value
        
        ws.Cells(TickerRow, 12).Value = StockVolTotal
        
        ClosingPrice = ws.Cells(i, 6).Value
        YearlyChange = ClosingPrice - OpeningPrice
        ws.Cells(TickerRow, 10).Value = YearlyChange
        
        If ws.Cells(TickerRow, 10).Value < 0 Then
            
            ws.Cells(TickerRow, 10).Interior.ColorIndex = 3
            
        Else
        
            ws.Cells(TickerRow, 10).Interior.ColorIndex = 4
            
        End If
            
        PercentChange = YearlyChange / OpeningPrice
        ws.Cells(TickerRow, 11).Value = PercentChange
                
        StockVolTotal = 0
        TickerRow = TickerRow + 1
        
              
    End If
    
    If PercentChange > GrtInc Then
    
        GrtInc = PercentChange
        ws.Cells(2, 15).Value = PercentChange
        GrtIncTicker = ws.Cells(i, 1).Value
        
    End If
    
    If PercentChange < GrtDec Then
    
        GrtDec = PercentChange
        ws.Cells(3, 15).Value = PercentChange
        GrtDecTicker = ws.Cells(i, 1).Value
        
    End If
           
    If StockVolTotal > StkMaxVol Then
    
        StkMaxVol = StockVolTotal
        ws.Cells(4, 15).Value = StkMaxVol
        MaxVolTicker = ws.Cells(i, 1).Value
        
    End If
                  
    
Next i
    
    
    ws.Cells(2, 14).Value = GrtIncTicker
    ws.Cells(3, 14).Value = GrtDecTicker
    ws.Cells(4, 14).Value = MaxVolTicker

Next ws

End Sub

