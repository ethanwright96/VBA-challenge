Sub StockAnalytics():

'Create for loop to run script on all sheets
For Each ws In Worksheets

'Declare variable data types
Dim Ticker As String
Dim Start_Price As Double
Dim End_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double

'Set summary table column labels
ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Start Price"
ws.Cells(1, "K").Value = "End Price"
ws.Cells(1, "L").Value = "Yearly Change"
ws.Cells(1, "M").Value = "Percent Change"
ws.Cells(1, "N").Value = "Total Stock Volume"

'Declare and initialize Total_Stock_Volume to zero
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

'Declare and set initial Summary_Table_Row to 2
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Set initial Start_Price to the first <open> obs
Start_Price = ws.Cells(2, "C").Value

'Find length of column "A" and store As Long
Dim nRows As Long
nRows = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Create for loop to detect change in Ticker
For i = 2 To nRows
    
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
'Store <ticker> for cells that satisfy condition above as Ticker
Ticker = ws.Cells(i, 1).Value

'Store last <close> obs as End_Price
End_Price = ws.Cells(i, 6).Value
    
'Calculate and store Yearly_Change
Yearly_Change = End_Price - Start_Price

'Use ifelse statement to avoid incidence of division by zero
If Start_Price = 0 Then
    Percent_Change = NA
'Calculate and store Percent_Change for all Start_Price not equal to 0
Else
Percent_Change = (End_Price - Start_Price) / Start_Price
End If

'Calculate and store Total_Stock_Volume
Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, "G").Value
         
'Set corresponding summary table cells to their respective values
ws.Cells(Summary_Table_Row, "I").Value = Ticker
ws.Cells(Summary_Table_Row, "J").Value = Start_Price
    ws.Cells(Summary_Table_Row, "J").Style = "Currency"
ws.Cells(Summary_Table_Row, "K").Value = End_Price
    ws.Cells(Summary_Table_Row, "K").Style = "Currency"
ws.Cells(Summary_Table_Row, "L").Value = Yearly_Change
    ws.Cells(Summary_Table_Row, "L").Style = "Currency"
ws.Cells(Summary_Table_Row, "M").Value = Percent_Change
    ws.Cells(Summary_Table_Row, "M").Style = "Percent"
ws.Cells(Summary_Table_Row, "N").Value = Total_Stock_Volume

'Apply conditional color formatting to Yearly Change column
If ws.Cells(Summary_Table_Row, "L").Value > 0 Then
    ws.Cells(Summary_Table_Row, "L").Interior.ColorIndex = 4
ElseIf ws.Cells(Summary_Table_Row, "L").Value < 0 Then
    ws.Cells(Summary_Table_Row, "L").Interior.ColorIndex = 3
Else: ws.Cells(Summary_Table_Row, "L").Interior.ColorIndex = 2
End If

'Apply conditional color formatting to Percent Change column
If ws.Cells(Summary_Table_Row, "M").Value > 0 Then
    ws.Cells(Summary_Table_Row, "M").Interior.ColorIndex = 4
ElseIf ws.Cells(Summary_Table_Row, "M").Value < 0 Then
    ws.Cells(Summary_Table_Row, "M").Interior.ColorIndex = 3
Else: ws.Cells(Summary_Table_Row, "M").Interior.ColorIndex = 2
End If

'Reset/adjust necessary variables for next i
Summary_Table_Row = Summary_Table_Row + 1

Total_Stock_Volume = 0

Start_Price = ws.Cells(i + 1, 3)

Else

'Complete ifelse statement for Total_Stock_Volume
Total_Stock_Volume = Total_Stock_Volume + Cells(i, "G").Value

End If

Next i

Next ws

End Sub


'Bonus Section

Sub FindExtremes():

'Create for loop to run script on all sheets
For Each ws In Worksheets

'Declare variable data types
Dim TickerRange As Range
Dim PercentChangeRange As Range
Dim TotalVolumeRange As Range
Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim MaxVolume As Double

'Declare and find LastRow
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row

'Set Variable Ranges (I,M,N)
Set TickerRange = ws.Range("I2:I" & LastRow)
Set PercentChangeRange = ws.Range("M2:M" & LastRow)
Set TotalVolumeRange = ws.Range("N2:N" & LastRow)

'Set Data Table Headers ("Q:S")
ws.Cells(2, "Q").Value = "Greatest % Increase"
ws.Cells(3, "Q").Value = "Greatest % Decrease"
ws.Cells(4, "Q").Value = "Greatest Total Volume"
ws.Cells(1, "R").Value = "Ticker"
ws.Cells(1, "S").Value = "Value"

'Find MaxIncrease using Max()
MaxIncrease = Application.WorksheetFunction.Max(PercentChangeRange)
'Assign value to target cell in new summary table and format as percent
ws.Cells(2, "S") = MaxIncrease
ws.Cells(2, "S").Style = "Percent"

'Use Index(Match()) to find the Ticker that corresponds with MaxIncrease
ws.Cells(2, "R").Value = Application.WorksheetFunction.Index(TickerRange, Application.WorksheetFunction.XMatch(MaxIncrease, PercentChangeRange))

'Find MaxDecrease using Min()
MaxDecrease = Application.WorksheetFunction.Min(PercentChangeRange)
'Assign value to target cell in new summary table and format as percent
ws.Cells(3, "S") = MaxDecrease
ws.Cells(3, "S").Style = "Percent"

'Use Index(Match()) to find the Ticker that corresponds with MaxDecrease
ws.Cells(3, "R").Value = Application.WorksheetFunction.Index(TickerRange, Application.WorksheetFunction.XMatch(MaxDecrease, PercentChangeRange))

'Find MaxVolume using Max()
MaxVolume = Application.WorksheetFunction.Max(TotalVolumeRange)
'Assign value to target cell in new summary table
ws.Cells(4, "S") = MaxVolume

'Use Index(Match()) to find the Ticker that corresponds with MaxVolume
ws.Cells(4, "R").Value = Application.WorksheetFunction.Index(TickerRange, Application.WorksheetFunction.XMatch(MaxVolume, TotalVolumeRange))

Next ws

End Sub
