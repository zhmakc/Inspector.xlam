Attribute VB_Name = "dv"
Option Base 1
Dim ID(), Kv(), nm(), DVrcount, n
Dim stopkod As Boolean
Dim NewArrRowCount As Integer
Dim arrZap() As Variant

Sub Deved() 'Дефектовочная ведомость
a = Timer
Application.Goto Range(MeWB.Names("razdel1"))

'Debug.Print Cells(8, 5).MergeArea.Columns.Count
'Exit Sub
t = 0 'считаем количество запчастей
    For n = 1 To shParts.UsedRange.Rows.Count
        If Cells(n, 5).MergeArea.Columns.Count = 2 And Not Cells(n, 5).Value = "Кол." Then
            t = t + 1
        End If
    Next n
ReDim arrZap(t, 4)
i = 1
    For n = 1 To shParts.UsedRange.Rows.Count
        If Cells(n, 5).MergeArea.Columns.Count = 2 And Not Cells(n, 5).Value = "Кол." Then
            arrZap(i, 1) = i 'номер порядковый
            arrZap(i, 2) = Cells(n, 5).Offset(0, -4).Value 'id
            arrZap(i, 3) = Cells(n, 5).Offset(0, 1).Value 'наименование
            arrZap(i, 4) = Cells(n, 5).Value 'количество
            i = i + 1
        End If
    Next n
'сортируем
'    arrZap = CoolSort(arrZap, 2) 'сортируем по возрастанию номер детали
'    arrZap = TrimArr(arrZap, 2, 3) 'объединяем
'    Dim newarr()
'    ReDim newarr(NewArrRowCount, UBound(arrZap, 2))
'    Call DelEmptyRow(arrZap)
    'newarr =  'удаляем пустые строки
 '   arrZap = CoolSort(newarr, 1) 'сортируем по возрастанию порядкогового номера
'Суммируем дубликаты
'Dim arrZapTemp(t, 4)
'    For h = 1 To t - 1
'        arrZapTemp(t, 4) =
'        If arrZap(h, 2) = arrZap(h + 1, 2) Then
'            arrZap(h, 3) = arrZap(h, 3) + arrZap(h + 1, 3)
'        End If
'    Next h
'Делаем новую книгу с листом ДВ
Dim wbDV As Workbook: Set wbDV = Workbooks.Add
    'Dim wbDV As Workbook
    shDVShablon.Copy after:=wbDV.Sheets(1)
    Application.DisplayAlerts = False
    wbDV.Sheets(1).Delete
    Application.DisplayAlerts = True
'Сохраняем сгенерированную ДВ
Dim nDV As Byte: nDV = 1
L1:
If Len(Dir$(MeWB.Path & "\ДВ №" & nDV & " " & MeWB.Name)) > 0 Then
    nDV = nDV + 1
    GoTo L1
Else
    wbDV.SaveAs FileName:=MeWB.Path & "\ДВ №" & nDV & " " & MeWB.Name, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End If
'Debug.Print wbDV.Sheets(1).Range("C24").Value
wbDV.Sheets(1).Range("C21").Value = shMain.Range("J7").Value 'фамилия
wbDV.Sheets(1).Range("F1").Value = shMain.Range("A11").Value 'Машина
wbDV.Sheets(1).Range("F2").Value = shMain.Range("G11").Value 'SN:
wbDV.Sheets(1).Range("F4").Value = shMain.Range("L11").Value 'Наработка
wbDV.Sheets(1).Range("I4").Value = shMain.Range("S11").Value 'Наработка ходом
wbDV.Sheets(1).Range("F5").Value = Date
wbDV.Sheets(1).Range("H21").Value = Date
wbDV.Sheets(1).Range("F6").Value = shMain.Range("K14").Value 'Клиент
wbDV.Sheets(1).Range("F7").Value = shMain.Range("K15").Value 'Регион
wbDV.Sheets(1).Range("F12").Value = shMain.Range("AB6").Value 'Номер ДВ
'Добавляем строки для запчастей
wbDV.Sheets(1).Range(Cells(18, 1), Cells(18 + UBound(arrZap, 1), 1)).EntireRow.Insert Shift:=xlUp, CopyOrigin:=xlFormatFromRightOrBelow
'Переносим массив на лист
wbDV.Sheets(1).Range("A17").Resize(UBound(arrZap, 1), UBound(arrZap, 2)).Value = arrZap
'Делаем красиво
    With wbDV.Sheets(1).Range("A17").Resize(UBound(arrZap, 1) + 2, UBound(arrZap, 2) + 5).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

''выносим массив на лист
''Dim sh As Worksheet: Set sh = Workbooks.Add(-4167).Worksheets(1): sh.Name = "Массив"
''    ' заносим данные из массива MyArr на лист sh
''    Call Array2worksheet(sh, arrZap, Array("№пп", "ID", "Кол-во", "Наименование"))
''MsgBox (Timer - a & " секунд на выполнение")
End Sub
Sub Array2worksheet(ByRef sh As Worksheet, ByVal Arr, ByVal ColumnsNames)
    ' https://excelvba.ru/ Получает двумерный массив Arr с данными,
    ' и массив заголовков столбцов ColumnsNames.
    ' Заносит данные из массива на лист sh
    If UBound(Arr, 1) > sh.Rows.Count - 1 Or UBound(Arr, 2) > sh.Columns.Count Then
        MsgBox "Массив не влезет на лист " & sh.Name, vbCritical, _
               "Размеры массива: " & UBound(Arr, 1) & "*" & UBound(Arr, 2): End
    End If
    With sh
        .UsedRange.Clear
        ColumnsNamesCount = UBound(ColumnsNames) - LBound(ColumnsNames) + 1
        .Range("a1").Resize(, ColumnsNamesCount).Value = ColumnsNames
        .Range("a1").Resize(, ColumnsNamesCount).Interior.ColorIndex = 15
        .Range("a2").Resize(UBound(Arr, 1), UBound(Arr, 2)).Value = Arr
        .UsedRange.EntireColumn.AutoFit
    End With
End Sub
Function CoolSort(SourceArr As Variant, ByVal n As Integer) As Variant
    ' https://excelvba.ru/ сортировка двумерного массива по столбцу N
    If n > UBound(SourceArr, 2) Or n < LBound(SourceArr, 2) Then _
       MsgBox "Нет такого столбца в массиве!", vbCritical: Exit Function
    Dim Check As Boolean, iCount As Integer, jCount As Integer, nCount As Integer
    ReDim tmparr(UBound(SourceArr, 2)) As Variant
    Do Until Check
        Check = True
        For iCount = LBound(SourceArr, 1) To UBound(SourceArr, 1) - 1
            If Val(SourceArr(iCount, n)) > Val(SourceArr(iCount + 1, n)) Then
                For jCount = LBound(SourceArr, 2) To UBound(SourceArr, 2)
                    tmparr(jCount) = SourceArr(iCount, jCount)
                    SourceArr(iCount, jCount) = SourceArr(iCount + 1, jCount)
                    SourceArr(iCount + 1, jCount) = tmparr(jCount)
                    Check = False
                Next
            End If
        Next
    Loop
    CoolSort = SourceArr
End Function
Function TrimArr(SourceArr As Variant, ByVal n As Integer, ByVal N2 As Integer) As Variant
    ' суммируем дубли
    'N - номер столбика с дубликатами
    'N2 - номер столбика для сложения
    Dim Check As Boolean 'проверка окончания цикла
    Dim iCount As Integer 'количество строк массива
    Dim jCount As Integer 'количество столбцов массива
    Dim nCount As Integer 'номер строки нового массива
    Dim kCount As Integer 'номер строки массива который сравнивается сейчас
    ReDim tmparr(UBound(SourceArr, 2)) As Variant 'временный двухмерный массив длиной в количество столбцов массива
    'цыкл поиска одинаковых значений
    ReDim newarr(UBound(SourceArr, 1), UBound(SourceArr, 2)) As Variant
    nCount = 0
    iCount = 2
'Debug.Print SourceArr(1, 1), SourceArr(1, 2), SourceArr(1, 3), SourceArr(1, 4)
Do Until iCount = UBound(SourceArr, 1) 'в цикле перебираем все строки
        Check = False
        nCount = nCount + 1
        iCount = iCount - 1
        'первую строку записываем в новый массив
            For jCount = LBound(SourceArr, 2) To UBound(SourceArr, 2)
                newarr(nCount, jCount) = SourceArr(iCount, jCount)
            Next
        'сравниваем запись текущую в новом массиве с следующей в исходом
        'для этого плюсуем счетчик iCount и запускаем условие в цикле
    iCount = iCount + 1
    Do Until Check 'проверяем до первого не совпадения
        Check = True
        If Val(newarr(nCount, n)) = Val(SourceArr(iCount, n)) Then
            'Если условие работает, плюсуем количество в новый массив
            newarr(nCount, N2) = newarr(nCount, N2) + SourceArr(iCount, N2)
            Check = False 'запускаем цикл заново
        End If
        iCount = iCount + 1 'берем следующий пункт
    Loop
Loop
    NewArrRowCount = nCount 'количество уникальных значений
    TrimArr = newarr
Exit Function
End Function

Function DelEmptyRow(SourceArr As Variant) '

ReDim tempArr(UBound(SourceArr, 1), UBound(SourceArr, 2)) As Variant
tempArr = SourceArr
ReDim arrZap(NewArrRowCount, UBound(SourceArr, 2))

    For rw = 1 To NewArrRowCount
            For jCount = LBound(SourceArr, 2) To UBound(SourceArr, 2)
                arrZap(rw, jCount) = tempArr(rw, jCount)
            Next
    Next
End Function

Sub Korzina() 'Запись в массив данных из корзины
'выбор версии Lidos
On Error GoTo 0
stopkod = False
If MainForm.PDFMode = True Then
'        strFile = LastFile$(MainForm.TextBoxLidos20Path, ".xml", 1)
'        'MsgBox (strFile)
'        Set WB = Workbooks.OpenXML(FileName:=strFile, LoadOption:=xlXmlLoadImportToList)
    Dim TextFromClipBoard As String
    Dim myData As New DataObject
    myData.GetFromClipboard
    On Error GoTo 1
    
 '   Debug.Print IsError(myData.GetText)
    
    
    TextFromClipBoard = myData.GetText
    'TextFromClipBoard = StrConv(myData.GetText, vbFromUnicode)
    'для правильного отображения шрифта меняем кодовую страницу
    
    If MainForm.CodePage.Value Then
        TextFromClipBoard = ChangeTextCharset(TextFromClipBoard, "Windows-1251", "Windows-1252")
    End If
    
    ListPDFZap.Range("A2") = TextFromClipBoard
'    Application.Calculate
    ListPQZap.Range("A1").ListObject.QueryTable.Refresh BackgroundQuery:=False
'    MsgBox ListPQZap.Range("A2")
'    Debug.Print ListPQZap.Range("A1").ListObject.QueryTable.RowNumbers
'    Debug.Print ListPQZap.ListObjects("TablePDF").Range.Rows.Count
    
    Set QT = ListPQZap.ListObjects.Item("TablePDF")
   
    DVrcount = QT.Range.Rows.Count - 1
    ReDim ID(DVrcount)
    ReDim nm(DVrcount)
    ReDim Kv(DVrcount)
    For n = 1 To DVrcount
    ID(n) = QT.Range(1 + n, 1) 'номер детали
    nm(n) = QT.Range(1 + n, 3) 'название детали
    Kv(n) = QT.Range(1 + n, 2) 'количество
    Next n
Exit Sub
1:     MsgBox ("Ошибка. Не удалось преобразовать данные буфера обмена")
    stopkod = True
Else

        Application.ScreenUpdating = False
        Call FindAndReplace("windows-1252", "windows-1251")
         Dim strTargetFile As String
         Application.DisplayAlerts = False
         strTargetFile = "c:\LIDOS\User_files\Orders\Standard.pro"
         Set WB = Workbooks.OpenXML(FileName:=strTargetFile, LoadOption:=xlXmlLoadImportToList)
        Call FindAndReplace("windows-1251", "windows-1252")
L1:
        DVrcount = WB.Sheets(1).UsedRange.Rows.Count - 1
            If DVrcount = 0 Then
            Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
            Range("A1").FormulaR1C1 = "1"
            GoTo L1
            End If
        ReDim ID(DVrcount)
        ReDim nm(DVrcount)
        ReDim Kv(DVrcount)
        For n = 1 To DVrcount
        ID(n) = WB.Sheets(1).UsedRange.Cells(1 + n, 28) 'номер детали
        nm(n) = WB.Sheets(1).UsedRange.Cells(1 + n, 30) 'название детали
        Kv(n) = WB.Sheets(1).UsedRange.Cells(1 + n, 29) 'количество
        Next n
        If IsEmpty(ID(1)) And IsEmpty(nm(1)) And IsEmpty(Kv(1)) Then stopkod = True
        WB.Close False
        'Application.DisplayAlerts = True

End If

End Sub
Function ChangeTextCharset(ByVal txt$, ByVal DestCharset$, _
                           Optional ByVal SourceCharset$) As String
    On Error Resume Next: Err.Clear
    With CreateObject("ADODB.Stream")
        .Type = 2: .Mode = 3
        If Len(SourceCharset$) Then .Charset = SourceCharset$
        .Open
        .WriteText txt$
        .Position = 0
        .Charset = DestCharset$
        ChangeTextCharset = .ReadText
        .Close
    End With
End Function

Sub ImportXML() 'старт
    ActiveWorkbook.Saved = False
    UFDV.Show
End Sub

Function str(s) As String
Dim Title As String
If UFDV.CheckBox1 = True Then Title = "" Else Title = " - " & nm(s)
str = ID(s) & Title & " - " & Kv(s) & " шт."
End Function

Sub FindAndReplace(FindValues As Variant, ReplaceValues As Variant)
    Dim FileName As String
    Dim FSO As Object
    Dim i As Long
    Dim Text As String
    Dim TextFile As Object
    Dim Wks As Worksheet
On Error GoTo 1
    FileName = "c:\LIDOS\User_files\Orders\Standard.pro"
    Set Wks = Worksheets(1)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TextFile = FSO.OpenTextFile(FileName, 1, False)
    Text = TextFile.ReadAll
    TextFile.Close
    Text = Replace(Text, FindValues, ReplaceValues)
    Set TextFile = FSO.OpenTextFile(FileName, 2, False)
    TextFile.Write Text
    TextFile.Close
Exit Sub
1: MsgBox ("Не найден файл c:\LIDOS\User_files\Orders\Standard.pro. Нужно сохранить бланк заказа с запчастями в программе Lidos.")
End
End Sub

Sub Flash()
If Not ActiveSheet.Name = shParts.Name Then shParts.Activate: MsgBox ("Указать пункт раздела для вставки запчастей"):   Exit Sub
Application.ScreenUpdating = False
Navigator.Navi (ActiveCell.Row)
a = Timer
Call Korzina
If stopkod = True Then MsgBox ("Корзина пуста"): Exit Sub
'обработка ошибки при вставке в раздел без пункта
If Not IsNumeric(shService.Cells(Npunkt + 1, 4).Value) Then MsgBox ("Указать пункт раздела для вставки запчастей"): Exit Sub
Dim r As Integer
r = shService.Cells(Npunkt + 1, 4).Value + 5  'определяем адрес текущего пункта
'MsgBox (r)
Dim sA As String
sA = r & ":" & (r + DVrcount - 1) 'добавляем строки для вставки
Rows(sA).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow 'добавляем строки

Set cur = Cells(r, 1)
For n = 1 To DVrcount
With Range(Cells(r + n - 1, 1), Cells(r + n - 1, 4))
    .Merge
    .HorizontalAlignment = xlLeft
    .Font.Bold = False
    .Value = ID(n)
End With
With Range(Cells(r + n - 1, 5), Cells(r + n - 1, 6))
    .Merge
    .HorizontalAlignment = xlCenter
    .Font.Bold = True
    .Value = Abs(Kv(n) / 1)
End With
With Range(Cells(r + n - 1, 7), Cells(r + n - 1, 34))
    .Merge
    .HorizontalAlignment = xlLeft
    .Font.Bold = False
    .Value = nm(n)
End With
Next n
'For N = 1 To DVrcount
'    cur.Offset(N - 1, 0) = ID(N)
'    cur.Offset(N - 1, 4) = Abs(Kv(N) / 1)
'    cur.Offset(N - 1, 6) = nm(N)
'Next N
'Добавляем шапку таблицы
'Пятая строка
r = shService.Cells(Npunkt + 1, 4).Value
Cells(r + 4, 1).EntireRow.UnMerge
With Range(Cells(r + 4, 1), Cells(r + 4, 4))
    .HorizontalAlignment = xlLeft
    .Merge
    .Value = " ID"
    .Font.Bold = False
End With
With Range(Cells(r + 4, 5), Cells(r + 4, 6))
    .HorizontalAlignment = xlCenter
    .Merge
    .Value = "Кол."
    .Font.Bold = False
End With
With Range(Cells(r + 4, 7), Cells(r + 4, 34))
    .HorizontalAlignment = xlLeft
    .Merge
    .Value = "     Наименование запчасти"
    .Font.Bold = False
End With

'Перемещаем курсор для вставки новой таблицы(Ищем пустую ячейку)
Cells(r + DVrcount + 4, 1).Activate
Do While Not IsEmpty(ActiveCell)
ActiveCell.Offset(1, 0).Activate
Loop
ActiveCell.Offset(1, 0).Activate
'If Not IsEmpty(Cells(r + DVrcount + 6, 1)) Then
'End If

Application.DisplayAlerts = True
'MsgBox (Timer - a & " секунд на выполнение")
End Sub

