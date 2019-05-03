sub tickercounter ()
Dim ws As Worksheets

For Each ws In Worksheets

'Define Variables
Dim ticker As String
Dim ticker_count As Double
Dim lastrow As Double
Dim summary_row As Double
Dim vol As Double

'Declare initial variables
lastrow = ws.Cells(Rows.Count, "A").End(xlup).Row
summary_row = 2
vol = 0


'Set headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Stock Volume"


For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'Set ticker
        ticker = Cells(i, 1).Value
        Cells(summary_row, 9).value = ticker

        'Set volume
        vol = vol + Cells(i, 7).Value
        Cells(summary_row, 10).Value = vol

        summary_row = summary_row + 1

        vol = 0

    Else vol = vol + Cells(i, 7).Value  
            

    End If

    
Next i

Next ws

End Sub