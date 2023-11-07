Sub TonghopDailySellReport()
    Dim ws As Worksheet
    Dim durCell As Range
    Dim dataRange As Range
    Dim cell As Range

    ' Duy?t qua các sheet du?c ch?n
    For Each ws In ActiveWindow.SelectedSheets
        ws.Activate ' Chuy?n d?n sheet hi?n t?i

        ' Replace "° " with ":" in the entire worksheet
        Cells.Replace What:="° ", Replacement:=":", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        ' Replace "'" with ":" in the entire worksheet
        Cells.Replace What:="'", Replacement:=":", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        ' Replace """ with "" in the entire worksheet
        Cells.Replace What:="""", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        ' Tìm ô ch?a "Dur"
        Set durCell = Cells.Find(What:="Dur", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
If Not durCell Is Nothing Then
        ' L?y c?t sau giá tr? "Dur" cách 4 c?t
        Dim nextCol As Long
        nextCol = durCell.Column + 6

        ' Ð?t chi?u r?ng m?i cho c?t sau giá tr? "Dur"
        Columns(nextCol).ColumnWidth = 18 ' Thay 15 b?ng chi?u r?ng m?i mà b?n mu?n

    End If
        If Not durCell Is Nothing Then
            ' Tô màu xanh lá cây, d?i màu ch? thành màu tr?ng và in d?m cho 3 ô tru?c giá tr? "Dur"
            Dim prevThreeCells As Range
            Set prevThreeCells = durCell.Offset(0, 6).Resize(1, 3)
            prevThreeCells.Interior.Color = RGB(0, 176, 80) ' Màu xanh lá cây
            With prevThreeCells.Font
                .Bold = True ' In d?m
                .Color = RGB(255, 255, 255) ' Màu tr?ng
            End With
            Set prevThreeCells = durCell.Offset(5, 6).Resize(1, 3)
            prevThreeCells.Interior.Color = RGB(0, 176, 80) ' Màu xanh lá cây
            With prevThreeCells.Font
                .Bold = True ' In d?m
                .Color = RGB(255, 255, 255) ' Màu tr?ng
            End With

            ' T?o các công th?c tính toán
             durCell.Offset(0, 4).Value = "Don hang"
            durCell.Offset(0, 3).Value = "So phut"
            Dim i As Long
            For i = 1 To 50
                durCell.Offset(i, 4).FormulaR1C1 = "=IF(LEFT(RIGHT(RC11,4),1)="","",""1"",""0"")"
                durCell.Offset(i, 3).FormulaR1C1 = "=IF(RC[-3]<>"""",HOUR(RC[-3])*60+MINUTE(RC[-3]),"""")"
            Next i
            durCell.Offset(0, 6).Value = "Th" & ChrW(7889) & "ng k" & ChrW(234) & " bao tr" & ChrW(249) & "m"
            durCell.Offset(0, 7).Value = "S" & ChrW(7889) & " Call"
            durCell.Offset(0, 8).Value = "Call c" & ChrW(243) & " " & ChrW(273) & ChrW(417) & "n"
            durCell.Offset(1, 6).Value = "Call d" & ChrW(432) & ChrW(7899) & "i 5 ph" & ChrW(250) & "t:"
            durCell.Offset(2, 6).Value = "Call t" & ChrW(7915) & " 5 - 12 ph" & ChrW(250) & "t:"
            durCell.Offset(3, 6).Value = "Call t" & ChrW(7915) & " 12 - 30 ph" & ChrW(250) & "t:"
            durCell.Offset(4, 6).Value = "Call tr" & ChrW(234) & "n 30 ph" & ChrW(250) & "t:"
            durCell.Offset(5, 6).Value = "T" & ChrW(7893) & "ng"
            Set durCell = durCell.Offset(1, 7)
            durCell.FormulaR1C1 = "=COUNTIFS(C[-4],""<=5"")"
            Set durCell = durCell.Offset(0, 1)
            durCell.FormulaR1C1 = "=COUNTIFS(C[-5],""<=5"",C[-4],""1"")"
            Set durCell = durCell.Offset(1, -1)
            durCell.FormulaR1C1 = "=COUNTIFS(C[-4],"">5"",C[-4],""<=12"")"
            Set durCell = durCell.Offset(0, 1)
            durCell.FormulaR1C1 = "=COUNTIFS(C[-5],"">5"",C[-5],""<=12"",C[-4],""1"")"
            Set durCell = durCell.Offset(1, -1)
            durCell.FormulaR1C1 = "=COUNTIFS(C[-4],"">12"",C[-4],""<=30"")"
            Set durCell = durCell.Offset(0, 1)
            durCell.FormulaR1C1 = "=COUNTIFS(C[-5],"">12"",C[-5],""<=30"",C[-4],""1"")"
            Set durCell = durCell.Offset(1, -1)
            durCell.FormulaR1C1 = "=COUNTIFS(C[-4],"">30"")-1"
            Set durCell = durCell.Offset(0, 1)
            durCell.FormulaR1C1 = "=COUNTIFS(C[-5],"">30"",C[-4],""1"")-1"
            Set durCell = durCell.Offset(1, -1)
            durCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
            Set durCell = durCell.Offset(0, 1)
            durCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
        Else
            MsgBox "Không tìm th?y giá tr? 'Dur' trong tài li?u này."
        End If

        ' Tìm ô ch?a "So phut"
        Set durCell = Cells.Find(What:="So phut", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

        If Not durCell Is Nothing Then
            ' Xác d?nh vùng có ch?a d? li?u
            Dim lastRow As Long
            Dim lastColumn As Long
            lastRow = ws.Cells(ws.Rows.Count, durCell.Column).End(xlUp).Row
            lastColumn = ws.Cells(durCell.Row, ws.Columns.Count).End(xlToLeft).Column
            Set dataRange = ws.Range(ws.Cells(durCell.Row, durCell.Column), ws.Cells(lastRow, lastColumn))

            ' K? ô và can gi?a các ô có ch?a d? li?u
            For Each cell In dataRange
                If Not IsEmpty(cell.Value) Then
                    cell.Borders.LineStyle = xlContinuous
                    cell.Borders.Color = RGB(0, 0, 0)
                    cell.HorizontalAlignment = xlCenter
                    cell.VerticalAlignment = xlCenter
                End If
            Next cell
        End If
    Next ws
End Sub


