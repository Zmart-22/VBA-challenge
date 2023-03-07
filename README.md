# VBA-challenge
Module 2
Sub Module2()

Dim ws As Worksheet

For Each ws In Worksheets


ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"



Dim Ticker As String
Ticker = " "
Dim StockVolume As Double
StockVolume = 0


Dim Lastrow As Long
Dim i As Long
Dim j As Integer


Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


Dim stock_open As Double
stock_open = 0
Dim stock_close As Double
stock_close = 0
Dim price_change As Double
price_change = 0
Dim percent_change As Double
percent_change = 0
Dim Increasename As String
Dim decreasename As String
Dim volumename As String
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestvolume As Double



Dim endrows As Long
endrows = 2

stock_open = ws.Cells(2, 3).Value

For i = 2 To Lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker = ws.Cells(i, 1).Value
endrows = endrows + 1
ws.Cells(endrows, "I").Value = Ticker

stock_close = ws.Cells(i, 6).Value
price_change = stock_close - stock_open


ElseIf stock_open <> 0 Then
percent_change = (price_change / stock_open) * 100

End If

StockVolume = StockVolume + ws.Cells(i, 7).Value

ws.Range("I" & endrows).Value = Ticker
ws.Range("J" & endrows).Value = price_change

If (price_change > 0) Then
ws.Range("J" & endrows).Interior.ColorIndex = 4

ElseIf (prince_change <= 0) Then
ws.Range("J" & endrows).Interior.ColorIndex = 3
End If


ws.Range("K" & endrows).Value = (CStr(percent_change) & "%")

ws.Range("L" & endrows).Value = StockVolume

stock_open = ws.Cells(i + 1, 3).Value


If (percent_change > greatestincrease) Then
greatestincrease = percent_change
Increasename = Ticker

ElseIf (percent_change < greatestdecrease) Then
greatestdecrease = percent_change
decreasename = Ticker

End If

If (StockVolume > greatestvolume) Then
greatestvolume = StockVolume
volumename = Ticker

End If

percent_change = 0
StockVolume = 0


Next i


ws.Range("P2").Value = Increasename
ws.Range("P3").Value = decreasename
ws.Range("P4").Value = volumename
ws.Range("Q2").Value = greatestincrease
ws.Range("Q3").Value = greatestdecrease
ws.Range("Q4").Value = greatestvolume



Next ws

End Sub

