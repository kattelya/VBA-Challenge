Attribute VB_Name = "Module1"
Sub ticker()

    Dim ticker As String
    Dim vol_total As Double
    vol_total = 0

    Dim summary_table_row As Integer
    summary_table_row = 2

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            vol_total = vol_total + Cells(i, 7).Value
            Range("I" & summary_table_row).Value = ticker
            Range("J" & summary_table_row).Value = vol_total
            summary_table_row = summary_table_row + 1
            vol_total = 0
        Else
            vol_total = vol_total + Cells(i, 7).Value
        End If
    Next i
End Sub
