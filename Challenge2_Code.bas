Attribute VB_Name = "Module2"
Option Explicit

Sub AlphabeticalTest()

Dim ws As Object
For Each ws In Worksheets

Dim LastRow, LastRowM, LastRowN, i, t, BestTic, WorstTic, BestTot As Integer
Dim StartPrice, EndPrice As Double
Dim Vol As LongLong


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(1, 11).Value = "Ticker"
ws.Cells(1, 12).Value = "Yearly Change"
ws.Cells(1, 13).Value = "Percent Change"
ws.Cells(1, 14).Value = "Total Stock Volume"
ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"
t = 2
Vol = 0

StartPrice = Cells(2, 3).Value

For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(t, 11).Value = ws.Cells(i, 1).Value
        
        Vol = Vol + ws.Cells(i, 7).Value
        EndPrice = ws.Cells(i, 6).Value
        
        ws.Cells(t, 12) = EndPrice - StartPrice
            If ws.Cells(t, 12) < 0 Then
                ws.Cells(t, 12).Interior.ColorIndex = 3
            Else
                ws.Cells(t, 12).Interior.ColorIndex = 4
            End If
            
        ws.Cells(t, 13) = FormatPercent(ws.Cells(t, 12).Value / StartPrice)
        ws.Cells(t, 14).Value = Vol
        StartPrice = ws.Cells(i + 1, 3).Value
        Vol = 0
        t = t + 1
    Else
        Vol = Vol + ws.Cells(i, 7).Value
    
    End If

Next i

LastRowM = ws.Cells(Rows.Count, 13).End(xlUp).Row
LastRowN = ws.Cells(Rows.Count, 14).End(xlUp).Row

ws.Cells(2, 18) = FormatPercent(WorksheetFunction.Max(ws.Range("M2:M" & LastRowM).Value))
ws.Cells(3, 18) = FormatPercent(WorksheetFunction.Min(ws.Range("M2:M" & LastRowM).Value))
ws.Cells(4, 18) = WorksheetFunction.Max(ws.Range("N2:N" & LastRowN).Value)

BestTic = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("M2:M" & LastRowM).Value), ws.Range("M2:M" & LastRowM).Value, 0)
WorstTic = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("M2:M" & LastRowM).Value), ws.Range("M2:M" & LastRowM).Value, 0)
BestTot = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("N2:N" & LastRowN).Value), ws.Range("N2:N" & LastRowN).Value, 0)

ws.Cells(2, 17) = ws.Cells(BestTic + 1, 11).Value
ws.Cells(3, 17) = ws.Cells(WorstTic + 1, 11).Value
ws.Cells(4, 17) = ws.Cells(BestTot + 1, 11).Value

Next ws

End Sub



