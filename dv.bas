Attribute VB_Name = "dv"
Option Base 1
Dim ID(), Kv(), nm(), DVrcount, n
Dim stopkod As Boolean
Dim NewArrRowCount As Integer
Dim arrZap() As Variant

Sub Deved() '������������� ���������
a = Timer
Application.Goto Range(MeWB.Names("razdel1"))

'Debug.Print Cells(8, 5).MergeArea.Columns.Count
'Exit Sub
t = 0 '������� ���������� ���������
    For n = 1 To shParts.UsedRange.Rows.Count
        If Cells(n, 5).MergeArea.Columns.Count = 2 And Not Cells(n, 5).Value = "���." Then
            t = t + 1
        End If
    Next n
ReDim arrZap(t, 4)
i = 1
    For n = 1 To shParts.UsedRange.Rows.Count
        If Cells(n, 5).MergeArea.Columns.Count = 2 And Not Cells(n, 5).Value = "���." Then
            arrZap(i, 1) = i '����� ����������
            arrZap(i, 2) = Cells(n, 5).Offset(0, -4).Value 'id
            arrZap(i, 3) = Cells(n, 5).Offset(0, 1).Value '������������
            arrZap(i, 4) = Cells(n, 5).Value '����������
            i = i + 1
        End If
    Next n
'���������
'    arrZap = CoolSort(arrZap, 2) '��������� �� ����������� ����� ������
'    arrZap = TrimArr(arrZap, 2, 3) '����������
'    Dim newarr()
'    ReDim newarr(NewArrRowCount, UBound(arrZap, 2))
'    Call DelEmptyRow(arrZap)
    'newarr =  '������� ������ ������
 '   arrZap = CoolSort(newarr, 1) '��������� �� ����������� ������������� ������
'��������� ���������
'Dim arrZapTemp(t, 4)
'    For h = 1 To t - 1
'        arrZapTemp(t, 4) =
'        If arrZap(h, 2) = arrZap(h + 1, 2) Then
'            arrZap(h, 3) = arrZap(h, 3) + arrZap(h + 1, 3)
'        End If
'    Next h
'������ ����� ����� � ������ ��
Dim wbDV As Workbook: Set wbDV = Workbooks.Add
    'Dim wbDV As Workbook
    shDVShablon.Copy after:=wbDV.Sheets(1)
    Application.DisplayAlerts = False
    wbDV.Sheets(1).Delete
    Application.DisplayAlerts = True
'��������� ��������������� ��
Dim nDV As Byte: nDV = 1
L1:
If Len(Dir$(MeWB.Path & "\�� �" & nDV & " " & MeWB.Name)) > 0 Then
    nDV = nDV + 1
    GoTo L1
Else
    wbDV.SaveAs FileName:=MeWB.Path & "\�� �" & nDV & " " & MeWB.Name, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End If
'Debug.Print wbDV.Sheets(1).Range("C24").Value
wbDV.Sheets(1).Range("C21").Value = shMain.Range("J7").Value '�������
wbDV.Sheets(1).Range("F1").Value = shMain.Range("A11").Value '������
wbDV.Sheets(1).Range("F2").Value = shMain.Range("G11").Value 'SN:
wbDV.Sheets(1).Range("F4").Value = shMain.Range("L11").Value '���������
wbDV.Sheets(1).Range("I4").Value = shMain.Range("S11").Value '��������� �����
wbDV.Sheets(1).Range("F5").Value = Date
wbDV.Sheets(1).Range("H21").Value = Date
wbDV.Sheets(1).Range("F6").Value = shMain.Range("K14").Value '������
wbDV.Sheets(1).Range("F7").Value = shMain.Range("K15").Value '������
wbDV.Sheets(1).Range("F12").Value = shMain.Range("AB6").Value '����� ��
'��������� ������ ��� ���������
wbDV.Sheets(1).Range(Cells(18, 1), Cells(18 + UBound(arrZap, 1), 1)).EntireRow.Insert Shift:=xlUp, CopyOrigin:=xlFormatFromRightOrBelow
'��������� ������ �� ����
wbDV.Sheets(1).Range("A17").Resize(UBound(arrZap, 1), UBound(arrZap, 2)).Value = arrZap
'������ �������
    With wbDV.Sheets(1).Range("A17").Resize(UBound(arrZap, 1) + 2, UBound(arrZap, 2) + 5).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

''������� ������ �� ����
''Dim sh As Worksheet: Set sh = Workbooks.Add(-4167).Worksheets(1): sh.Name = "������"
''    ' ������� ������ �� ������� MyArr �� ���� sh
''    Call Array2worksheet(sh, arrZap, Array("���", "ID", "���-��", "������������"))
''MsgBox (Timer - a & " ������ �� ����������")
End Sub
Sub Array2worksheet(ByRef sh As Worksheet, ByVal Arr, ByVal ColumnsNames)
    ' https://excelvba.ru/ �������� ��������� ������ Arr � �������,
    ' � ������ ���������� �������� ColumnsNames.
    ' ������� ������ �� ������� �� ���� sh
    If UBound(Arr, 1) > sh.Rows.Count - 1 Or UBound(Arr, 2) > sh.Columns.Count Then
        MsgBox "������ �� ������ �� ���� " & sh.Name, vbCritical, _
               "������� �������: " & UBound(Arr, 1) & "*" & UBound(Arr, 2): End
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
    ' https://excelvba.ru/ ���������� ���������� ������� �� ������� N
    If n > UBound(SourceArr, 2) Or n < LBound(SourceArr, 2) Then _
       MsgBox "��� ������ ������� � �������!", vbCritical: Exit Function
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
    ' ��������� �����
    'N - ����� �������� � �����������
    'N2 - ����� �������� ��� ��������
    Dim Check As Boolean '�������� ��������� �����
    Dim iCount As Integer '���������� ����� �������
    Dim jCount As Integer '���������� �������� �������
    Dim nCount As Integer '����� ������ ������ �������
    Dim kCount As Integer '����� ������ ������� ������� ������������ ������
    ReDim tmparr(UBound(SourceArr, 2)) As Variant '��������� ���������� ������ ������ � ���������� �������� �������
    '���� ������ ���������� ��������
    ReDim newarr(UBound(SourceArr, 1), UBound(SourceArr, 2)) As Variant
    nCount = 0
    iCount = 2
'Debug.Print SourceArr(1, 1), SourceArr(1, 2), SourceArr(1, 3), SourceArr(1, 4)
Do Until iCount = UBound(SourceArr, 1) '� ����� ���������� ��� ������
        Check = False
        nCount = nCount + 1
        iCount = iCount - 1
        '������ ������ ���������� � ����� ������
            For jCount = LBound(SourceArr, 2) To UBound(SourceArr, 2)
                newarr(nCount, jCount) = SourceArr(iCount, jCount)
            Next
        '���������� ������ ������� � ����� ������� � ��������� � �������
        '��� ����� ������� ������� iCount � ��������� ������� � �����
    iCount = iCount + 1
    Do Until Check '��������� �� ������� �� ����������
        Check = True
        If Val(newarr(nCount, n)) = Val(SourceArr(iCount, n)) Then
            '���� ������� ��������, ������� ���������� � ����� ������
            newarr(nCount, N2) = newarr(nCount, N2) + SourceArr(iCount, N2)
            Check = False '��������� ���� ������
        End If
        iCount = iCount + 1 '����� ��������� �����
    Loop
Loop
    NewArrRowCount = nCount '���������� ���������� ��������
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

Sub Korzina() '������ � ������ ������ �� �������
'����� ������ Lidos
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
    '��� ����������� ����������� ������ ������ ������� ��������
    
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
    ID(n) = QT.Range(1 + n, 1) '����� ������
    nm(n) = QT.Range(1 + n, 3) '�������� ������
    Kv(n) = QT.Range(1 + n, 2) '����������
    Next n
Exit Sub
1:     MsgBox ("������. �� ������� ������������� ������ ������ ������")
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
        ID(n) = WB.Sheets(1).UsedRange.Cells(1 + n, 28) '����� ������
        nm(n) = WB.Sheets(1).UsedRange.Cells(1 + n, 30) '�������� ������
        Kv(n) = WB.Sheets(1).UsedRange.Cells(1 + n, 29) '����������
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

Sub ImportXML() '�����
    ActiveWorkbook.Saved = False
    UFDV.Show
End Sub

Function str(s) As String
Dim Title As String
If UFDV.CheckBox1 = True Then Title = "" Else Title = " - " & nm(s)
str = ID(s) & Title & " - " & Kv(s) & " ��."
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
1: MsgBox ("�� ������ ���� c:\LIDOS\User_files\Orders\Standard.pro. ����� ��������� ����� ������ � ���������� � ��������� Lidos.")
End
End Sub

Sub Flash()
If Not ActiveSheet.Name = shParts.Name Then shParts.Activate: MsgBox ("������� ����� ������� ��� ������� ���������"):   Exit Sub
Application.ScreenUpdating = False
Navigator.Navi (ActiveCell.Row)
a = Timer
Call Korzina
If stopkod = True Then MsgBox ("������� �����"): Exit Sub
'��������� ������ ��� ������� � ������ ��� ������
If Not IsNumeric(shService.Cells(Npunkt + 1, 4).Value) Then MsgBox ("������� ����� ������� ��� ������� ���������"): Exit Sub
Dim r As Integer
r = shService.Cells(Npunkt + 1, 4).Value + 5  '���������� ����� �������� ������
'MsgBox (r)
Dim sA As String
sA = r & ":" & (r + DVrcount - 1) '��������� ������ ��� �������
Rows(sA).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow '��������� ������

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
'��������� ����� �������
'����� ������
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
    .Value = "���."
    .Font.Bold = False
End With
With Range(Cells(r + 4, 7), Cells(r + 4, 34))
    .HorizontalAlignment = xlLeft
    .Merge
    .Value = "     ������������ ��������"
    .Font.Bold = False
End With

'���������� ������ ��� ������� ����� �������(���� ������ ������)
Cells(r + DVrcount + 4, 1).Activate
Do While Not IsEmpty(ActiveCell)
ActiveCell.Offset(1, 0).Activate
Loop
ActiveCell.Offset(1, 0).Activate
'If Not IsEmpty(Cells(r + DVrcount + 6, 1)) Then
'End If

Application.DisplayAlerts = True
'MsgBox (Timer - a & " ������ �� ����������")
End Sub

