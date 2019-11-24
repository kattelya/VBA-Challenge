Sub Ticker()

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim Op_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
    
        For i = 2 To LastRow
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                Ticker = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker
                Close_Price = Cells(i, Column + 5).Value
                Yearly_Change = Close_Price - Op_Price
                Cells(Row, Column + 9).Value = Yearly_Change
                If (Op_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Op_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Op_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                Row = Row + 1
                Op_Price = Cells(i + 1, Column + 2)
                Volume = 0
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
     
        YearlyChangeLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        For j = 2 To YearlyChangeLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
    Next WS
End Sub
