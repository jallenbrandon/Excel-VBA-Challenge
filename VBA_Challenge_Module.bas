Sub Stocks()
' Declar and set worksheet
Dim ws As Worksheet

'Loop for all worksheets
For Each ws In Worksheets

' column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Value"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker as variable
Dim TickerName As String
TickerName = " "
Dim Ticker_volume As Double
Ticker_volume = 0

'Set new variable for price and percent changes
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0
Dim price_change_percent As Double
price_change_percent = 0

' Set variables for min/max tables
Dim MaxTicker As String
MaxTicker = " "
Dim MinTicker As String
MinTicker = " "
Dim MaxPercent As Double
MaxPercent = 0
Dim MinPercent As Double
MinPercent = 0
Dim MaxTickerVolume As String
MaxTickerVolume = " "
Dim MaxValueVolume As Double
MaxValueVolume = 0

'Keeping track of Ticker row
Dim Summary_Table As Long
Summary_Table = 2

'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long

'Define Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'openprice iniatialization
open_price = ws.Cells(2, 3).Value

'Do loop of current worksheet to Lastrow
For i = 2 To Lastrow

'Ticker symbol output
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

TickerName = ws.Cells(i, 1).Value

'Calculate change in price
close_price = ws.Cells(i, 6).Value
price_change = close_price - open_price

'Total ticker volume calculations
Ticker_volume = Ticker_volume + ws.Cells(i, 7).Value



price_change_percent = (price_change / open_price) * 100

'Print Ticker name in worksheet
ws.Range("I" & Summary_Table).Value = TickerName

'Print Ticker change in worksheet
ws.Range("J" & Summary_Table).Value = price_change

'Fill color for price_change in worksheet
If (price_change > 0) Then
ws.Range("J" & Summary_Table).Interior.ColorIndex = 4

ElseIf (price_change < 0) Then
ws.Range("J" & Summary_Table).Interior.ColorIndex = 3

End If

'Print percentage change in worksheet
ws.Range("K" & Summary_Table).Value = (CStr(price_change_percent) & "%")

'Print tciker total in worksheet
ws.Range("L" & Summary_Table).Value = Ticker_volume

'Summary table row count
Summary_Table = Summary_Table + 1

'reset price change and price percent
price_change = 0

close_price = 0
open_price = ws.Cells(i + 1, 3).Value


' find max ticker and min percentage of ticker and the value
If (price_change_percent > MaxPercent) Then
MaxPercent = price_change_percent
MaxTicker = TickerName

ElseIf (price_change_percent < MinPercent) Then
MinPercent = price_change_percent
MinTicker = TickerName

End If

If (Ticker_volume > MaxValueVolume) Then
MaxValueVolume = Ticker_volume
MaxTickerVolume = TickerName

End If

price_change_percent = 0
Ticker_volume = 0

Else
Ticker_volume = Ticker_volume + ws.Cells(i, 7).Value

End If

Next i

ws.Range("Q2").Value = (CStr(MaxPercent) & "%")
ws.Range("Q3").Value = (CStr(MinPercent) & "%")
ws.Range("P2").Value = MaxTicker
ws.Range("P3").Value = MinTicker
ws.Range("P4").Value = MaxTickerVolume
ws.Range("Q4").Value = MaxValueVolume

Next ws

End Sub
