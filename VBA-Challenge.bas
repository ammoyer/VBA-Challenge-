Attribute VB_Name = "Module1"
Sub ticker()
Dim ticker As String
    Dim j As Integer
    Dim total As Double
    Dim ws As Worksheet
    Dim last_row As Variant
    For Each ws In Worksheets
    j = 0
    Start = 2
        last_row = Cells(Rows.Count, "A").End(xlUp).Row
        ws.Range("I1").Value = "ticker"
        ws.Range("J1").Value = "change"
        ws.Range("K1").Value = "percentchange"
        ws.Range("L1").Value = "totalvolume"
        For Row = 2 To last_row
        If (ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value) Then
            ticker = ws.Cells(Row, 1).Value
            ws.Range("I" & 2 + j).Value = ticker
            total = total + ws.Cells(Row, 7).Value
        If total = 0 Then
                        ' print the results
                        ws.Range("I" & 2 + j).Value = ws.Cells(Row, 1).Value
                        ws.Range("J" & 2 + j).Value = 0
                        ws.Range("K" & 2 + j).Value = "%" & 0
                        ws.Range("L" & 2 + j).Value = 0
            Else
            ws.Range("L" & 2 + j).Value = total
        If ws.Cells(Start, 3) = 0 Then
            For find_value = Start To Row
                If ws.Cells(find_value, 3) <> 0 Then
                    Start = find_value
                    Exit For
                End If
            Next find_value
        End If
             Change = ws.Cells(Row, 6) - ws.Cells(Start, 3)
             percentchange = Change / ws.Cells(Start, 3) * 100
             Start = Row + 1
             ws.Range("J" & 2 + j).Value = Change
             ws.Range("K" & 2 + j).Value = percentchange
             Select Case Change
                Case Is > 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
             End Select
             End If
            j = j + 1
            total = 0
            Change = 0
        Else
            total = total + ws.Cells(Row, 7).Value
        End If
        Next Row
        Next ws

End Sub


