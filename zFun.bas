Attribute VB_Name = "zFun"
Option Explicit
'#If VBA7 Then
'    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal uFormat As Long) As Long
'    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
'    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
'    Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal drop_handle As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
'#Else
'    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal uFormat As Long) As Long
'    Private Declare Function OpenClipboard Lib "user32" (ByVal Hwnd As Long) As Long
'    Private Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
'    Private Declare Function CloseClipboard Lib "user32" () As Long
'    Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal drop_handle As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
'#End If
' Clipboard routines.
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileW" (ByVal drop_handle As LongPtr, ByVal UINT As Long, ByVal lpStr As LongPtr, ByVal ch As Long) As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
' Global memory routines.
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Dst As Any, Src As Any, ByVal Sz As LongPtr)
Private Type POINTAPI ' DROPFILES data structure.
    X As Long
    Y As Long
End Type
Private Type DropFiles
    pFiles As Long
    pt As POINTAPI
    fNC As Long
    fWide As Long
End Type
Private Const CF_HDROP = 15 ' File list clipboard format code.
' Global Memory Flags
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
 
Private FormatValue() As Integer ' List of format values
 
' Get an array of the files listed in the clipboard.
Public Function ClipboardGetFiles() As String '()
    Dim hDrop As LongPtr, i&, CntPaths$, ln&
    
    ' Make sure there is file data.
    If IsClipboardFormatAvailable(CF_HDROP) Then                                'File data exists. Get it.
        If OpenClipboard(0) Then                                                'Open the clipboard.
            hDrop = GetClipboardData(CF_HDROP)                                  'Get the handle to the dropped list of files.
            CntPaths = DragQueryFile(hDrop, -1, 0, 0)                           'Get the number of dropped files.
            
            ReDim FilePaths$(CntPaths - 1)                                      'Get the file names.
            For i = 0 To CntPaths - 1
                ln = DragQueryFile(hDrop, i, 0, 0)
                FilePaths(i) = String(ln, vbNullChar)
                DragQueryFile hDrop, i, StrPtr(FilePaths(i)), ln + 1            'Get the file name.
            Next
            CloseClipboard                                                      'Close the clipboard.
            
            GoTo resp '???? ? ????? ???????, ????? ?? ?????????? ??????, ? ??????????? ??????
        End If
    End If
MsgBox ("����� ������ �������� �� ������ ����������")
Exit Function
resp:
    ClipboardGetFiles = FilePaths(0)  ' Assign the return value.
End Function
 
' Copy the file names into the clipboard. Return True if we succeed.
Public Function ClipboardSetFiles(FilePaths() As String) As Boolean
    Dim File_String As String
    Dim Drop_Files As DropFiles
    Dim hGlb As LongPtr
    Dim pGlbLock As LongPtr
    Dim i As Long
 
'    Clipboard.Clear ' Clear the clipboard.
    If OpenClipboard(0) Then ' Open the clipboard.
        For i = LBound(FilePaths) To UBound(FilePaths) ' Build a null-terminated list of file names.
            File_String = File_String & FilePaths(i) & vbNullChar
        Next
        Drop_Files.pFiles = LenB(Drop_Files)    ' Initialize the DROPFILES structure.
        Drop_Files.fWide = 1                    ' Unicode characters.
        Drop_Files.fNC = 0                      ' Client coordinates.
 
        ' Get global memory to hold the DROPFILES structure and the file list string.
        hGlb = GlobalAlloc(GHND, LenB(Drop_Files) + LenB(File_String))
        If hGlb Then
            pGlbLock = GlobalLock(hGlb) ' Lock the memory while we initialize it.
 
            ' Copy the DROPFILES structure and the file string into the global memory.
            CopyMem ByVal pGlbLock, Drop_Files, LenB(Drop_Files)
            CopyMem ByVal pGlbLock + LenB(Drop_Files), ByVal StrPtr(File_String), LenB(File_String)
            GlobalUnlock hGlb
 
            ' Copy the data to the clipboard.
            SetClipboardData CF_HDROP, hGlb
            ClipboardSetFiles = True
        End If
        
        CloseClipboard ' Close the clipboard.
    End If
End Function
 
Private Sub Example()
    Dim FilePaths() As String
    
    FilePaths = ClipboardGetFiles
    Stop
    
    ReDim FilePaths(2)
    FilePaths(0) = "\\User-pc\????? ?????\???????????\TextToClipboard.xlsm"
    FilePaths(1) = "\\User-pc\????? ?????\???????????\MineTextToClipboard.xlsm"
    FilePaths(2) = "\\User-pc\????? ?????\???????????\TextToClipboard2.xlsm"
    
    ClipboardSetFiles FilePaths
End Sub
 
 
Function changeA1(Optional ByVal a&)
    Dim ret
    If a Then
        ret = Application.Evaluate("=changeA1")
    Else
        ActiveSheet.Cells("A1").Interior.Color = vbRed
    End If
End Function
Sub sffa()
    Dim a
    a = [changeA1]
End Sub


'- ���� � ����� https://www.cyberforum.ru/vba/thread3168926-page2.html

Public Function SheetByCodename(ByRef WB As Workbook, ByVal Codename$) As Worksheet
'������� ������� � ���������� ���� �� ��������� �������� �����
    On Error Resume Next: Dim sh As Worksheet
    For Each sh In WB.Worksheets
        If sh.Codename = Codename$ Then Set SheetByCodename = sh: Exit Function
    Next sh
End Function
Public Function SetGlobDim() '������ ���������� ����������, ������������ � ����
'��������� �� �������� �� ��� ���������� � ���� �� �� �� ������ ��� ��������
If Not TypeName(shParts) = "Nothing" Then Exit Function
'������� ���� � ����������
Set shParts = SheetByCodename(ActiveWorkbook, "AParts")
Set shMain = SheetByCodename(ActiveWorkbook, "Main")
'Dpi = MonitorDpiSetting(1, 0)
If TypeName(shParts) = "Nothing" Then Exit Function '���� ���� �� ������, ������� �� ���������
'���������� � ���������� ����� � �������
Set MeWB = shParts.Parent
'������� ��������� ����
Set shService = SheetByCodename(ActiveWorkbook, "ZService")
If TypeName(shService) = "Nothing" Then Exit Function '���� ���� �� ������, ������� �� ���������
'��������� ����� �������� � �����������
End Function
Public Function DelRef() '������� ����� � ������ �������� (����� ��������� ������)
'Dim NN As Name
For Each NN In Workbooks(MeWB.Name).Names
 If InStr(1, NN.Value, "#REF", vbTextCompare) Then
'  MsgBox ("������� ��� � �������� �������: " & NN.Name)
  NN.Delete
 End If
Next NN
End Function


Public Function ZapolnitPunkty()
iRo = 2
For Each NN In Workbooks(MeWB.Name).Names
'Debug.Print InStr(1, NN.Name, "razdel", vbTextCompare)
    If InStr(1, NN.Name, "Punkt" & Nrazdel, vbTextCompare) Then
        'Debug.Print CInt(Right(Left(NN.Name, 8), 2))
        shService.Cells(iRo, 11).Value = NN.Value
        shService.Cells(iRo, 12).Value = NN.Name
        shService.Cells(iRo, 13).Value = Range(NN.Name).Worksheet.Name
        shService.Cells(iRo, 14).Value = Range(NN.Name).Row
        iRo = iRo + 1
    End If
Next NN
shService.Cells(iRo, 14).Value = 10000
End Function

Public Function SortirovatPunkty()
Call ZapolnitPunkty
'���������
    shService.Sort.SortFields.Clear
    shService.Sort.SortFields.Add Key:=Range("n2:n2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With shService.Sort
        .SetRange Range("K1:N30")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    '�������
For Each NN In Workbooks(MeWB.Name).Names
 If InStr(1, NN.Name, "Punkt" & Nrazdel, vbTextCompare) Then
  NN.Delete
 End If
Next NN
'��������� ����� �� ������� ��� �� ��������� �����:
s = shService.Cells(2, 3)
For i = 1 To shService.Cells(101, 11).End(xlUp).Row - 1
MeWB.Names.Add Name:="Punkt" & Nrazdel & "." & i, RefersToR1C1:="='" & s & "'!R" & shService.Cells(i + 1, 14) & "C2"
Range("Punkt" & Nrazdel & "." & i).Value = Nrazdel & "." & i
Next i
Call ZapolnitPunkty
End Function

Public Function LastFile$(ByVal FolderPath As String, Optional ByVal Mask As String = "", _
                   Optional ByVal SearchDeep As Long = 999)
    ' �������� � �������� ��������� ���� � ����� FolderPath,
    ' ����� ����� ������� ������ Mask (����� ��������� ������ ����� � ����� ������/�����������)
    ' � ������� ������ SearchDeep � ��������� (���� SearchDeep=1, �� �������� �� ���������������).
    ' ���������� ������ ���� � �����, �������� ����� ������� ���� ��������
    ' (����������� ����������� ����� ��������� GetAllFileNamesUsingFSO)

    Dim FilenamesCollection As New Collection    ' ������ ������ ���������
    Set FSO = CreateObject("Scripting.FileSystemObject")    ' ������ ��������� FileSystemObject
    GetAllFileNamesUsingFSO FolderPath, Mask, FSO, FilenamesCollection, SearchDeep    ' �����
    Set FSO = Nothing: Application.StatusBar = False    ' ������� ������ ��������� Excel
    Dim maxFileDate As Double
    For Each File In FilenamesCollection ' ���������� ��� ����� ����� ���������
        currFileDate = FileDateTime(File) ' ��������� ���� ���������� ����������
        ' ��������� ��������� ���� - �� ����� �� �� ����������
        If currFileDate > maxFileDate Then LastFile$ = File: maxFileDate = currFileDate
    Next File
End Function
 
Public Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef FSO, _
                                 ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
    ' ���������� ��� ����� � �������� � ����� FolderPath, ��������� ������ FSO
    ' ������� ����� �������������� � ��� ������, ���� SearchDeep > 1
    ' ��������� ���� ��������� ������ � ��������� FileNamesColl
    On Error Resume Next: Set curfold = FSO.GetFolder(FolderPath)
    If Not curfold Is Nothing Then    ' ���� ������� �������� ������ � �����
        Application.StatusBar = "����� � �����: " & FolderPath
 
        For Each fil In curfold.Files    ' ���������� ��� ����� � ����� FolderPath
            If fil.Name Like "*" & Mask Then FileNamesColl.Add fil.Path
        Next
        SearchDeep = SearchDeep - 1    ' ��������� ������� ������ � ���������
        If SearchDeep Then    ' ���� ���� ������ ������
            For Each sfol In curfold.SubFolders    ' ' ���������� ��� �������� � ����� FolderPath
                GetAllFileNamesUsingFSO sfol.Path, Mask, FSO, FileNamesColl, SearchDeep
            Next
        End If
        Set fil = Nothing: Set curfold = Nothing    ' ������� ����������
    End If
End Function
'Sub ��������������������������_LastFile()
'    ' ���� �� ������� ����� ��� ����� TXT, � ������� ��� ������ ������ �����.
'    ' ��������������� ����� � �������� �������� �� ����� ���.''
'
'    Dim ����������$, ������������������$
'    ' �������� ���� � ����� ������� ����
'    ���������� = "C:\Users\lrushm0\Downloads"
'    ' �������� ���� � ������ ������ ����� (����������� ���� ���������� ����������)
'    ������������������$ = LastFile$("C:\Users\lrushm0\Downloads", ".csv", 3)
'
'    If ������������������$ = "" Then MsgBox "�� ������� �� ������ �����", vbExclamation: Exit Sub
'    MsgBox ������������������$, vbInformation, "����� ������ ����"
'End Sub
