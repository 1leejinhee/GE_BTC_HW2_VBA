Attribute VB_Name = "Module1"
Sub Main()
Dim xSh As Worksheet
Application.ScreenUpdating = False
For Each xSh In Worksheets
xSh.Select
Call SheetCalc
Next
Application.ScreenUpdating = True
End Sub

Sub SheetCalc()
Dim maxvolumeticker As String
Dim maxincreseticker As String
Dim maxdecreaseticker As String
Dim ticker As String
ticker = Cells(2, 1)
Dim r As Long
Dim tickerrowtart As Long
Dim maxVolume As Double
Dim maxIncPercentage As Double
Dim maxDecPercentage As Double
Dim stockVolume As Double
tickerrowtart = 2
targetRow = 2
For r = 2 To Rows.Count

If StrComp(ticker, Cells(r, 1).Value) <> 0 Then
Range("I" & targetRow).Value = ticker
Dim closingPrice As Double
Dim openingPrice As Double
openingPrice = Cells(tickerrowtart, 6).Value
closingPrice = Cells(r - 1, 6).Value
Cells(targetRow, 10).Value = closingPrice - openingPrice
Dim yearlyPercentage As Double

If openingPrice = 0 Then
yearlyPercentage = closingPrice - openingPrice
Else
yearlyPercentage = (closingPrice - openingPrice) / openingPrice
End If

Cells(targetRow, 11).Value = yearlyPercentage
Cells(targetRow, 12).Value = stockVolume
If yearlyPercentage > maxIncPercentage Then
maxIncPercentage = yearlyPercentage
maxincreseticker = ticker
End If

If yearlyPercentage < maxDecPercentage Then
maxDecPercentage = yearlyPercentage
maxdecreaseticker = ticker
End If

If stockVolume > maxVolume Then
maxVolume = stockVolume
maxvolumeticker = ticker
End If

targetRow = targetRow + 1
tickerrowtart = r
ticker = Cells(r, 1).Value
stockVolume = Cells(r, 7).Value
Else
Dim valToAdd As Double
valToAdd = Cells(r, 7).Value
stockVolume = stockVolume + valToAdd
End If

Next r

Cells(2, 15).Value = maxincreseticker
Cells(2, 16).Value = maxIncPercentage
Cells(3, 15).Value = maxdecreaseticker
Cells(3, 16).Value = maxDecPercentage
Cells(4, 15).Value = maxvolumeticker
Cells(4, 16).Value = maxVolume

End Sub
