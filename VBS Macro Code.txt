'-------INSTRUCTIONS--------
' Create a script that loops through all the stocks for one year and outputs the following information:
' 1) The ticker symbol.
' 2) Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
' 3) The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
' 4) The total stock volume of the stock.

Sub StockData():

' Loop through all sheets
For Each ws In Worksheets


' Define and determine the Last Row
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Add Headers for columns I,J,K, and L
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


' Define variable for holding the ticker type
Dim Ticker_Type As String

' Define variables for holding the total yearly change and percent change, and set initial yearly change value
Dim Yearly_Change As Double
Dim Percent_Change As Double
Yearly_Change = 0

' Define and set an initial variable for holding the total stock volume
Dim Total_Volume As Double
Total_Volume = 0

'Define variables for holding stock open values and stock close values
Dim Stock_Open As Double
Dim Stock_Close As Double


' Define a new variable to keep track of the location for each ticker type in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Define i to reflect number of rows to loop through all transactions
Dim i As Long
For i = 2 To LastRow

    ' Check if we are still within the same ticker type, if it is not then do the following
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    ' Set the ticker type
    Ticker_Type = ws.Cells(i, 1).Value


    ' Calculate total stock volume
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value


    ' Print the ticker type in the Summary Table
    ws.Range("I" & Summary_Table_Row).Value = Ticker_Type


    ' Print total stock volume to the Summary Table
    ws.Range("L" & Summary_Table_Row).Value = Total_Volume

    ' Reset the total stock volume for the next ticker type
    Total_Volume = 0

    ' Set stock close value
    Stock_Close = ws.Cells(i, 6).Value

' CALCULATE YEARLY CHANGE
    'Make sure we are not dividing by 0 to avoid getting errors
    If Stock_Open = 0 Then

        Yearly_Change = 0
        Percent_Change = 0

    Else

        Yearly_Change = Stock_Close - Stock_Open
        Percent_Change = (Stock_Close - Stock_Open) / Stock_Open

    End If

    ' Pring yearly change and percent change values to table

    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

    ' Add one to the summary table row to go down a row
    Summary_Table_Row = Summary_Table_Row + 1
      


' Finding the open stock value
ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    Stock_Open = ws.Cells(i, 3).Value


Else
    ' If the cell immediately following a row is the same ticker type
    ' Add to the total stock volume
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value

End If

Next i


' Loop through the rows to change color for column J
For i = 2 To LastRow



' Change Color to green if value is above 0
If ws.Range("J" & i).Value > 0 Then
    ws.Range("J" & i).Interior.ColorIndex = 4


'Change color to red if value is below 0
ElseIf ws.Range("J" & i).Value < 0 Then
    ws.Range("J" & i).Interior.ColorIndex = 3

End If

Next i



' -----BONUS SECTION-----

' Insert new cell headers to show the following

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

' Find Greatest Value Increase


'Defind new variables
Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim MaxVolume As Double


' Set new variables to intitil values of 0
MaxIncrease = 0
MaxDecrease = 0
MaxVolume = 0


' Find maximum Percent increase

For i = 2 To LastRow
If ws.Cells(i, 11).Value > MaxIncrease Then
MaxIncrease = ws.Cells(i, 11).Value


' Write values next to Greatest % Increase cells
ws.Range("P2").Value = ws.Cells(i, 9).Value
ws.Range("Q2").Value = MaxIncrease
ws.Range("Q2").NumberFormat = "0.00%"


End If
Next i



' Find maximum Percent decrease
For i = 2 To LastRow
If ws.Cells(i, 11).Value < MaxDecrease Then
MaxDecrease = ws.Cells(i, 11).Value


' Write values next to Greatest % Decrease cells
ws.Range("P3").Value = ws.Cells(i, 9).Value
ws.Range("Q3").Value = MaxDecrease
ws.Range("Q3").NumberFormat = "0.00%"

End If
Next i


' Find greatest total volume
For i = 2 To LastRow
If ws.Cells(i, 12).Value > MaxVolume Then
MaxVolume = ws.Cells(i, 12).Value


' Write values next to Greatest % Decrease cells
ws.Range("P4").Value = ws.Cells(i, 9).Value
ws.Range("Q4").Value = MaxVolume


End If
Next i




Next ws
        
        
End Sub

