Attribute VB_Name = "modTable"
'модуль для форматирования рамки на листе с запчастями
Dim r As Integer


Function Table() 'Вставляем рамку с  описанием
'Проверяем активен ли лист "Список запчастей"
a = Timer
If Not ActiveSheet.Name = shParts.Name Then shParts.Activate: MsgBox ("Указать место вставки таблицы на листе " & shParts.Name): Exit Function
'отключаем обновление экрана
Application.ScreenUpdating = False
r = ActiveCell.Row
Dim sA As String
sA = r & ":" & (r + 5)
Rows(sA).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow 'добавляем строки
'Первая строка
Call CommentTable(r)
With Cells(r, 2) 'добавляем имя диапазону
Navigator.Navi (ActiveCell.Row)
.Name = "Punkt" & Nrazdel & "." & PunktCount + 1
Navigator.Navi (ActiveCell.Row)
End With

'Третья строка
With Range(Cells(r + 2, 1), Cells(r + 2, 4))
    .HorizontalAlignment = xlLeft
    .Merge
    .Value = "Причина"
    .Font.Bold = True
End With
With Range(Cells(r + 2, 5), Cells(r + 2, 34))
    .HorizontalAlignment = xlLeft
    .Merge
End With
'Четвертая строка
With Range(Cells(r + 3, 1), Cells(r + 3, 4))
    .HorizontalAlignment = xlLeft
    .Merge
    .Value = "Действие"
    .Font.Bold = True
End With
With Range(Cells(r + 3, 5), Cells(r + 3, 34))
    .HorizontalAlignment = xlLeft
    .Merge
End With
'Пятая строка
r = shService.Cells(Npunkt + 1, 4).Value
With Range(Cells(r + 4, 1), Cells(r + 4, 10))
    .HorizontalAlignment = xlLeft
    .Merge
    .Value = " Запчасти не требуются"
    .Font.Bold = False
End With

'Рамки
Call Bord(Range(Cells(r, 1), Cells(r + 3, 34)))
                    'это для активации экселя
                        Dim WBN As Workbook
                        Set WBN = Workbooks.Add
                        WBN.Close
                    'это для активации экселя
Range(MeWB.Names("punkt" & Nrazdel & "." & Npunkt)).Offset(0, 3).MergeArea.Select '.Activate
Application.ScreenUpdating = True
'Range(MeWB.Names("punkt" & Nrazdel & "." & Npunkt)).Offset(0, 3)
'ActiveWindow.Activate
'AppActivate "Microsoft Excel"
'MsgBox (Timer - a & " секунд на выполнение")
End Function

Function CommentTable(r As Integer)
Cells(r, 1).Value = "п."
'Call SortirovatPunkty
With Range(Cells(r, 2), Cells(r, 4))
' добавляем название подраздела=============================================================
    .HorizontalAlignment = xlLeft
    .Merge
    .Font.Bold = False
End With
With Range(Cells(r, 5), Cells(r + 1, 34))
    .Merge
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlTop
    .WrapText = True
    .Font.Bold = True
End With
'Вторая строка
With Range(Cells(r + 1, 1), Cells(r + 1, 4))
    .Merge
    .HorizontalAlignment = xlLeft
    .Value = "Описание"
    .Font.Bold = True
End With

End Function

Function Bord(BorderRange As Range)
With BorderRange
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End With
End Function

