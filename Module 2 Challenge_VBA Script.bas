Attribute VB_Name = "Module1"
Sub StocksLoop()

For Each ws In Worksheets

'Headers

ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"
ws.Cells(2, 14) = "Ticker"
ws.Cells(2, 15) = "Value"
ws.Cells(3, 13) = "Greatest Percent Increase"
ws.Cells(4, 13) = "Greatest Percent Decrease"
ws.Cells(5, 13) = "Greatest Total Volume"

'Define Variables

Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Percent_Change = 0
Dim Total_Volume As Double
Total_Volume = 0
Dim Open_Price As Double
Open_Price = 0
Dim Close_Price As Double
Close_Price = 0
Dim GrPercentIn_TickerName As String
Dim GrPercentIn_TickerValue As Double
GrPercentIn_TickerValue = 0
Dim GrPercentDe_TickerName As String
Dim GrPercentDe_TickerValue As Double
GrPercentDe_TickerValue = 0
Dim Greatest_Total_Volume_Name As String
Dim Greatest_Total_Volume As Double
Greatest_Total_Volume = 0

'summary row

summary_row = 2


'Last Row


lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'set open price

Open_Price = ws.Cells(2, 3)

For i = 2 To lastRow

    If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
    
    Ticker = ws.Cells(i, 1)
    
    Close_Price = ws.Cells(i, 6)
    Yearly_Change = Close_Price - Open_Price
    
        If Open_Price <> 0 Then
       Percent_Change = (Yearly_Change / Open_Price) * 100
    
        End If
        
     Total_Volume = Total_Volume + ws.Cells(i, 7)
    
    
    ws.Cells(summary_row, 9) = Ticker
    ws.Cells(summary_row, 10) = Yearly_Change
  
    
   If (Yearly_Change > 0) Then
   ws.Cells(summary_row, 10).Interior.ColorIndex = 4
   ElseIf Yearly_Change <= 0 Then
   ws.Cells(summary_row, 10).Interior.ColorIndex = 3
   
   End If
   
   ws.Cells(summary_row, 11) = (CStr(Percent_Change) & "%")
   
   ws.Cells(summary_row, 12) = Total_Volume
   
   
summary_row = (summary_row + 1)

Open_Price = ws.Cells(i + 1, 3)

If (Percent_Change > GrPercentIn_TickerValue) Then
GrPercentIn_TickerValue = Percent_Change
GrPercentIn_TickerName = Ticker

ElseIf (Percent_Change < GrPercentDe_TickerValue) Then
GrPercentDe_TickerValue = Percent_Change
GrPercentDe_TickerName = Ticker

End If

If (Total_Volume > Greatest_Total_Volume) Then
Greatest_Total_Volume = Total_Volume
Greatest_Total_Volume_Name = Ticker

End If


Percent_Change = 0
Total_Volume = 0


Else

Total_Volume = Total_Volume + ws.Cells(i, 7)

End If


Next i

ws.Cells(3, 15) = (CStr(GrPercentIn_TickerValue) & "%")
ws.Cells(4, 15) = (CStr(GrPercentDe_TickerValue) & "%")
ws.Cells(3, 14) = GrPercentIn_TickerName
ws.Cells(4, 14) = GrPercentDe_TickerName
ws.Cells(5, 14) = Greatest_Total_Volume_Name
ws.Cells(5, 15) = Greatest_Total_Volume



Next ws


End Sub
