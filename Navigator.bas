Attribute VB_Name = "Navigator"
Public Function Navi(r As Integer)
Call SetGlobDim
Application.ScreenUpdating = Editmode
'выгружаем все имена на сервисную страницу
Dim NN As Name 'задаем переменную имен
'очищаем сервисную таблицу
'номер первой строки на листе Serv в которую идет запись первого значения, остальные пишутся ниже
iRo = 2 'смещение по другому
Call DelRef 'Удаляем имена с битыми ссылками (ранее удаленные ячейки)
Call ReadNames("razdel")
'поиск раздела по номеру текущей строки:
For b = 2 To shService.Cells(101, 4).End(xlUp).Row - 1
    If shService.Cells(b, 4).Value <= r And shService.Cells(b + 1, 4).Value > r Then
    Nrazdel = b - 1
        Exit For
    End If
    Nrazdel = 0
Next b
'MsgBox ("раздел №" & Nrazdel)
Call ReadNames("punkt" & Nrazdel)
    PunktCount = shService.Cells(101, 3).End(xlUp).Row - 1
    If PunktCount = 0 Then Npunkt = 0
    For b = 2 To shService.Cells(101, 3).End(xlUp).Row
        If shService.Cells(b, 4).Value <= r And shService.Cells(b + 1, 4).Value > r Then
        Npunkt = b - 1
            Exit For
        End If
    Npunkt = 0
    Next b
'MsgBox ("Раздел №" & Nrazdel & " Пункт №" & Npunkt & " из " & PunktCount)
End Function
Public Function ReadNames(n As String)
Range("serv").Clear
'Читаем Имена содержащие заданный текст в имени и записываем в сервисный лист
'Сортируем в порядке расположения по строкам
'Переименовываем
iRo = 2
For Each NN In Workbooks(MeWB.Name).Names
    If InStr(1, NN.Name, n, vbTextCompare) Then
        'Debug.Print CInt(Right(Left(NN.Name, 8), 2))
        shService.Cells(iRo, 1).Value = NN.Value
        shService.Cells(iRo, 2).Value = NN.Name
        shService.Cells(iRo, 3).Value = Range(NN.Name).Worksheet.Name
        shService.Cells(iRo, 4).Value = Range(NN.Name).Row
        shService.Cells(iRo, 5).Value = Right(NN.Name, 1)
        iRo = iRo + 1
    End If
Next NN
shService.Cells(iRo, 4).Value = 10000
'Сортировка по возрастанию номера строки.
    shService.Sort.SortFields.Clear
    shService.Sort.SortFields.Add Key:=Range("d2:d2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With shService.Sort
        .SetRange Range("A1:D97")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'Добавляем Имена по порядку как на сервисном листе:
If InStr(1, n, "razdel", vbTextCompare) = 1 Then
     For Each NN In Workbooks(MeWB.Name).Names 'Удаляем имена содержащие Razdel
        If InStr(1, NN.Name, n, vbTextCompare) Then
            NN.Delete
        End If
    Next NN
    For i = 1 To shService.Cells(101, 3).End(xlUp).Row - 1
        MeWB.Names.Add Name:=n & i, RefersToR1C1:="='" & shParts.Name & "'!R" & shService.Cells(i + 1, 4) & "C1"
    Next i
    Exit Function
End If
If InStr(1, n, "punkt", vbTextCompare) = 1 Then
    For Each NN In Workbooks(MeWB.Name).Names 'Удаляем имена содержащие punkt
        If InStr(1, NN.Name, n, vbTextCompare) Then
            NN.Delete
        End If
    Next NN
    For i = 1 To shService.Cells(101, 3).End(xlUp).Row - 1
        MeWB.Names.Add Name:=n & "." & i, RefersToR1C1:="='" & shParts.Name & "'!R" & shService.Cells(i + 1, 4) & "C2"
'Исправляем имена на листе с фотографиями.
'        If Not Range(n & "." & i).Value = Nrazdel & "." & i And Not IsEmpty(Range(n & "." & i).Value) Then
'        MsgBox (Range(n & "." & i).Value & "," & Nrazdel & "." & i)
'
'        End If
'
'
'
        Range(n & "." & i).Value = "'" & Nrazdel & "." & i
    Next i
    Exit Function
End If
'MsgBox ("условие не сработало")
End Function

