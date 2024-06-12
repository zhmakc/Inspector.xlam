Attribute VB_Name = "FotoComment"
Function AddPunktFoto()
a = Timer
Application.ScreenUpdating = Editmode
Dim TableRowQty As Integer
Dim r As Integer
TableRowQty = 4 'количество строк в таблице+1
If Not ActiveSheet.Name = shParts.Name Then shParts.Activate: MsgBox ("Нужно выбрать пункт на листе " & shParts.Name): Exit Function
Navigator.Navi (ActiveCell.Row)
'Debug.Print Workbooks(MeWB.Name)..RefersTo
'Debug.Print Range(MeWB.Names("mayak" & Nrazdel)).Row 'Маяк
'Range(MeWB.Names("mayak" & Nrazdel)).Value = Range(MeWB.Names("razdel" & Nrazdel)).Value & ". Фотографии"
'Debug.Print MeWB.Names("mayak" & Nrazdel)
'Debug.Print Range(MeWB.Names("mayak" & Nrazdel)).Worksheet.Name ' название листа
'вставляем текс в ячейку
'MeWB.Worksheets(Range(MeWB.Names("mayak" & Nrazdel)).Worksheet.Name).Activate
'Range(MeWB.Names("mayak" & Nrazdel)).Select
'Application.Goto Cells(Rows.Count, Columns.Count)
Call ReadNames("Foto" & Nrazdel) 'читаем все имена Foto+razdel на сервисный лист.
PunktCount = shService.Cells(101, 3).End(xlUp).Row - 1 'определяем количество пунктов с фото в текущем разделе.
'ставим курсор на маяк
Application.Goto Range(MeWB.Names("Mayak" & Nrazdel))
'Проверяем есть ли пункты вообще?
'Если да, то пробуем перейти на нужный пункт
On Error GoTo L1 'при ошибке переходим на проверку наличия пунктов.
Application.Goto Range(MeWB.Names("Foto" & Nrazdel & "." & Npunkt))
ActiveWindow.ScrollRow = Range(MeWB.Names("Foto" & Nrazdel & "." & Npunkt)).Row
Exit Function 'пункт найден, выходим из процедуры
L1: 'MsgBox ("Такой пункт с фото отсутсвует! проверяем на наличие других пунктов в разделе")
On Error GoTo 0
    If PunktCount = 0 Then 'MsgBox ("пункты отсутсвуют совсем, используем координаты маяка")
        r = Range(MeWB.Names("mayak" & Nrazdel)).Row + 3 - TableRowQty
        Else: 'MsgBox ("сейчас пунктов " & PunktCount & "переходим к поиску координаты для вставки")
        r = 0 'обнуляем координату
        For n = 2 To PunktCount + 1 'сравниваем текущий пункт с уже расположенными на листе
        If shService.Cells(n, 5).Value < Npunkt Then r = shService.Cells(n, 4).Value
        Next n
        'если условие не выполнено ни разу, используем координаты маяка
         If r = 0 Then r = Range(MeWB.Names("mayak" & Nrazdel)).Row + 3 - TableRowQty
        End If
r = TableRowQty + r
'MsgBox ("вставляем в эту ячейку " & r)

'Application.Goto Range(MeWB.Names("mayak" & Nrazdel))
'ActiveWindow.ScrollRow = Range(MeWB.Names("mayak" & Nrazdel)).Row

Dim sA As String
sA = (r) & ":" & (r + TableRowQty - 1)
Rows(sA).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow 'добавляем строки
'Первая строка
Call modTable.CommentTable(r)
With Cells(r, 2) 'добавляем имя диапазону
.Name = "Foto" & Nrazdel & "." & Npunkt
.Value = "'" & Nrazdel & "." & Npunkt
Call Bord(Range(Cells(r, 1), Cells(r + 1, 34)))
End With
With Cells(r, 5) ' переносим текст описания пункта в зависимости от настроек
    If GetSetting(MyAppName, "Settings", "LinkInDescription", False) = True Then
        Dim FormulaText As String
        FormulaText = "=OFFSET(punkt" & Nrazdel & "." & Npunkt & ",0,3)"
        '.FormulaR1C1 = "=OFFSET(punkt1.3,0,3)"
        .FormulaR1C1 = FormulaText
    ElseIf GetSetting(MyAppName, "Settings", "ActionToComment", False) = True Then
        Dim txtReason As String
        txtReason = Range("punkt" & Nrazdel & "." & Npunkt).Offset(0, 1).Value & ". " & _
                    Range("punkt" & Nrazdel & "." & Npunkt).Offset(3, 1).Value
        .Value = txtReason
    Else
        Dim txt As String
        txt = Range("punkt" & Nrazdel & "." & Npunkt).Offset(0, 1).Value
        'txt = txt & ". " & Range("punkt" & Nrazdel & "." & Npunkt).Offset(3, 1).Value
        '.Value = Range("punkt" & Nrazdel & "." & Npunkt).Offset(0, 1).Value 'Просто перенос текста вариант
        .Value = txt 'Просто перенос текста вариант
    End If
End With
With Range(Cells(r + 2, 1), Cells(r + 2, 17)) 'поле для фото 1
.Merge
.EntireRow.RowHeight = 190
End With
With Range(Cells(r + 2, 18), Cells(r + 2, 34)) 'поле для фото 2
.Merge
End With
With Range(Cells(r + 3, 1), Cells(r + 3, 17)) 'поле для фото 3
.Merge
End With
With Range(Cells(r + 3, 18), Cells(r + 3, 34)) 'поле для фото 4
.Merge
End With
'Rows(r + 2 & ":" & r + 2).RowHeight = 185
'переключаем вид на созданную таблицу
ActiveWindow.ScrollRow = Range(MeWB.Names("Foto" & Nrazdel & "." & Npunkt)).Row
Range(MeWB.Names("Foto" & Nrazdel & "." & Npunkt)).Offset(2, -1).Activate
Application.ScreenUpdating = True
'Запускаем вставку фото
'FotoTrigger = True
FotoInsert.InsertFoto
If StopRepit = True Then GoTo LAT
Range(MeWB.Names("Foto" & Nrazdel & "." & Npunkt)).Offset(2, 16).Activate
FotoInsert.InsertFoto
If StopRepit = True Then GoTo LAT
Range(MeWB.Names("Foto" & Nrazdel & "." & Npunkt)).Offset(3, -1).Activate
'Автопереход на раздел с запчастями после вставки фото
LAT:
If GetSetting(MyAppName, "Settings", "AutoTransition", False) Then shParts.Activate

'L2:
'MsgBox (Timer - a & " секунд на выполнение")
'Exit Function
End Function
