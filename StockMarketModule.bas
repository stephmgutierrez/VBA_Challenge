Attribute VB_Name = "Module1"

Sub stockdatachallenge():

For Each ws In Worksheets
Dim WorksheetName As String
Dim i As Long
Dim j As Long
Dim TickCount As Long
Dim LastRowA As Long
Dim LastRowI As Long
Dim PerChange As Double
Dim GreatIncr As Double
Dim GreatDecr As Double
Dim GreatVol As Double

'Get the worksheet name

WorksheetName = ws.Name

'Create the headers

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'Set ticker counter to first row
TickCount = 2

'Set start row to 2
j = 2

'Find the last row with data in the first column
LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Create loop to move through all rows
For i = 2 To LastRowA

'Check for unique ticker name
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Write ticker in column I/9
ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value

'Calculate and write Yearly Change in column J/10
ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

'Conditional Formatting
If ws.Cells(TickCount, 10).Value < 0 Then

'Set background color to Red for Cell
ws.Cells(TickCount, 10).Interior.ColorIndex = 3

'Set background color to Green for Cell
Else
ws.Cells(TickCount, 10).Interior.ColorIndex = 4

End If

'Calculate and input % change in column K/11
If ws.Cells(j, 3).Value <> 0 Then
PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)

'Format to Percentage
ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")

Else

ws.Cells(TickCount, 11).Value = Format(0, "Percent")

End If

'Calculate and write Total Volume in column L/12
ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

'Increase TickCount by 1
TickCount = TickCount + 1

'Set new start row of the ticker block
j = i + 1

End If

Next i

'Find last non-blank cell in column I
LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Prepare for summary
GreatVol = ws.Cells(2, 12).Value
GreatIncr = ws.Cells(2, 11).Value
GreatDecr = ws.Cells(2, 11).Value

'Loop for summary
For i = 2 To LastRowI

'Check if next value is larger for greatest total volume and take over with new value if true
If ws.Cells(i, 12).Value > GreatVol Then
GreatVol = ws.Cells(i, 12).Value
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value

Else

GreatVol = GreatVol

End If

'Check for greatest increase, check for larger value and replace if true
If ws.Cells(i, 11).Value > GreatIncr Then
GreatIncr = ws.Cells(i, 11).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value

Else

GreatIncr = GreatIncr

End If

'Check for greatest decrease and if value is smaller, replace with value
If ws.Cells(i, 11).Value < GreatDecr Then
GreatDecr = ws.Cells(i, 11).Value
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

Else

GreatDecr = GreatDecr

End If

'Write Summary results in ws.Cells
ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")

Next i
'Auto column adjust
Worksheets(WorksheetName).Columns("A:Z").AutoFit

Next ws


End Sub
