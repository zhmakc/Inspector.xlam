VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Инспектор"
   ClientHeight    =   3420
   ClientLeft      =   64
   ClientTop       =   168
   ClientWidth     =   5704
   OleObjectBlob   =   "MainForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Init As Boolean

Private Sub CBCrop_Click()
     Select Case TypeName(Selection)
        Case "Picture"
            Application.CommandBars.ExecuteMso ("PictureCropAspectRatio4To3")
        Case Else
            MsgBox "Выбор не верный"
    End Select
End Sub

Private Sub ChBoxActionToComment_Click()
If Init = True Then Exit Sub

SaveSetting MyAppName, "Settings", "ActionToComment", Me.ChBoxActionToComment.Value
If ChBoxActionToComment = False Then Exit Sub
ChBoxLink = False
End Sub


Private Sub FotoOpisanie_Click()

End Sub

Private Sub CodePage_Click()
If Init = True Then Exit Sub
SaveSetting MyAppName, "Settings", "CodePage", Me.CodePage.Value
End Sub

Private Sub FotoVyacheyku_Click()
Dim str As String
    Select Case TypeName(Selection)
    
    Case "Nothing"
        MsgBox "Выбор не сделан"
    
    Case "Range"
    
        If ActiveSheet.Name = shParts.Name Then
            Call FotoComment.AddPunktFoto
        Else
            Call FotoInsert.InsertFoto
        End If
'размещение картинки в ячейке
    Case "Picture"
        Application.ScreenUpdating = Editmode
        Application.DisplayAlerts = Editmode
            
            Call FotoInsert.PictPos
            Call FotoInsert.GetQTY
            Call FotoInsert.SaveAsPicture
            avFile = CompressFoto(avFile, QTY)
            shp.Delete
            Call FotoInsert.SetShp
'            shp.ScaleHeight 1, msoTrue
'            shp.ScaleWidth 1, msoTrue

'            With shp
'                .LockAspectRatio = msoFalse
'                Call Addshp
'            End With
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    Case Else
        MsgBox "Выбор не верный"
    End Select
'FotoInsert.InsertFoto
'MsgBox ("надо настроить. Вставка будет осуществляться в указанную объединенную ячейку")
End Sub

Private Sub CBPaint_Click()

Select Case TypeName(Selection)

    Case "Picture"
        
        Application.ScreenUpdating = Editmode
        Application.DisplayAlerts = Editmode
        
        Call FotoInsert.PictPos
        Call FotoInsert.GetQTY
        Call FotoInsert.SaveAsPicture
        
    ''    avFile = CompressFoto(avFile, GetQTY * 2)
'         MsgBox Dpi
        'Exit Sub
        'Редактируем фото в PAINT
        
        'On Error GoTo TheEnd
        MainForm.Enabled = False
            shp.Delete
            Set WshShell = CreateObject("WScript.Shell")
            WshShell.Run "c:\windows\system32\mspaint.exe " & """" & avFile & """", 3, True
    '        SendKeys "%{TAB}"
    
    
    avFile = CompressFoto(avFile, QTY)
             
            Call FotoInsert.SetShp
        
        MainForm.Enabled = True
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
     
    Case Else
            MsgBox "Выбор не верный"
    
End Select
End Sub


Private Sub AddRow_Click()
Application.ScreenUpdating = False
On Error GoTo 1
st = Selection.Row
ed = st + Selection.Rows.Count - 1
Rows(st & ":" & ed).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
'Exit Sub
'asdf dfas sdf as
With Range(Cells(st, 1), Cells(ed, 34))
'    .Borders(xlDiagonalDown).LineStyle = xlNone ' = xlContinuous
'    .Borders(xlDiagonalUp).LineStyle = xlNone
    .Borders(xlEdgeLeft).LineStyle = xlNone
'    .Borders(xlEdgeTop).LineStyle = xlNone
'    .Borders(xlEdgeBottom).LineStyle = xlNone
    .Borders(xlEdgeRight).LineStyle = xlNone
'    .Borders(xlInsideHorizontal).LineStyle = xlNone
    .Borders(xlInsideVertical).LineStyle = xlNone
End With
Exit Sub
1:
Application.ScreenUpdating = True
MsgBox ("Не верное выделение!")
End Sub

Private Sub AddTable_Click()
Call Table
End Sub

Private Sub ButAbout_Click()
About.Show 1
End Sub

Private Sub cboxRazdel_Change()
'переходим на выбранный раздел
'On Error Resume Next
If cboxRazdel.ListIndex = -1 Then Exit Sub
shParts.Activate
Range("_" & cboxRazdel.ListIndex + 1).Activate
End Sub

Private Sub cboxSheets_Change()
'Переходим на выбранный лист
On Error Resume Next
Worksheets(cboxSheets.Text).Activate
cboxSheets.Value = "Выбор раздела"
MultiPage1.SetFocus
End Sub


Private Sub cboxSheets_Enter()
cboxSheets.DropDown
End Sub

Private Sub ChBAutoHieght_Click()
If Init = True Then Exit Sub
SaveSetting MyAppName, "Settings", "AutoHieght", Me.ChBAutoHieght.Value
'Call UserForm_Initialize
End Sub

Private Sub ChBoxAutoTrans_Click()
If Init = True Then Exit Sub
SaveSetting MyAppName, "Settings", "AutoTransition", Me.ChBoxAutoTrans.Value
End Sub

Private Sub ChBoxLidos20_Click()
If Init = True Then Exit Sub
SaveSetting MyAppName, "Settings", "Lidos20", Me.ChBoxLidos20.Value
End Sub

Private Sub ChBoxLink_Click()
If Init = True Then Exit Sub
SaveSetting MyAppName, "Settings", "LinkInDescription", Me.ChBoxLink.Value
If ChBoxLink = False Then Exit Sub
ChBoxActionToComment = False
'Call UserForm_Initialize
End Sub


Private Sub CommandButton1_Click()
Application.ScreenUpdating = False
Dim Wsh As Worksheet

If ActiveWindow.View = 2 Then
    ActiveWindow.View = 3
    CommandButton1.Caption = "Режим разметки"
    
    Else:
    ActiveWindow.View = 2
    CommandButton1.Caption = "Страничный режим"
End If

'For Each WSh In MeWB.Worksheets
'    WSh.Activate
'    'ActiveWindow.View = xlPageBreakPreview
'    ActiveWindow.View = shmode
'Next

End Sub

Private Sub CommandButton2_Click()
Cells.CheckSpelling SpellLang:=1049
End Sub

Private Sub DelRow_Click() ' удалить строки
Application.ScreenUpdating = False
On Error GoTo 1
st = Selection.Row
ed = st + Selection.Rows.Count - 1
Rows(st & ":" & ed).Delete Shift:=xlUp
Exit Sub
1:
Application.ScreenUpdating = True
MsgBox ("Не верное выделение!")
End Sub

Private Sub GrupZap_Click()
dv.Deved
End Sub

Private Sub MultiPage1_Change()
Exit Sub
    Select Case Me.MultiPage1.Value
        Case 0:    ' ??????? 1
            
        Case 1:    ' ??????? 2
            
        Case 2:    ' ??????? 3
            
        Case 3:    '??????? 4 -?????????
            
    End Select
End Sub

Private Sub NewRazdel_Click()
    Dim VMode As Integer
        Application.ScreenUpdating = False
            VMode = ActiveWindow.View
                Call NewPage
            ActiveWindow.View = VMode
        Application.ScreenUpdating = True
End Sub

Private Sub PDFMode_Click()
If Init = True Then Exit Sub
SaveSetting MyAppName, "Settings", "PDFMode", Me.PDFMode.Value
End Sub

Private Sub PrintReset_Click()
Call ResetPage
End Sub

Private Sub SavePDF_Click()
dsda = DatePart("YYYY", Now) & "-" & DatePart("M", Now) & "-" & DatePart("D", Now) & " " & DatePart("H", Now) & "." & DatePart("N", Now) & " "
'MsgBox (ActiveWorkbook.Name & " " & dsda)
MsgBox ("Сохранено с именем " & MeWB.Path & "\" & dsda & Left(MeWB.Name, Len(MeWB.Name) - 5) & ".pdf")

    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        MeWB.Path & "\" & dsda & Left(MeWB.Name, Len(MeWB.Name) - 5) & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
        
'            ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
'        MeWB.Path & "\" & Left(MeWB.Name, InStr(MeWB.Name, ".") - 1) & ".pdf" _
'        , Quality:=xlQualityStandard, IncludeDocProperties:=False, _
'        IgnorePrintAreas:=False, OpenAfterPublish:=True

End Sub

Private Sub Shpartsact_Click()
shParts.Activate
End Sub

Private Sub TBoxQTY_Change()
Dim Val As Integer

On Error GoTo 1
If TBoxQTY.Value < 1 Then GoTo 1
Val = TBoxQTY.Value
SaveSetting MyAppName, "Settings", "QTY", Val
1:
End Sub

'Private Sub TBoxQTY_Enter()
'MsgBox ("В поле качество фото задается коэффициент в процентах. По умолчанию " & Q & "%" & vbNewLine & "Влияет только на качество вновь вставленных фотографий" & vbNewLine & "При увеличении коэффициента увеличивается качество фото и размер файла")
'MsgBox ("Дополнительные настройки для каждого файла отчета ОБЯЗАТЕЛЬНО:" & vbNewLine & "Файл->Параметры->Дополнительно->Размер и качество изображений" & vbNewLine & "Выбрать пункт -> Не сжимать изображения в файле")
'End Sub



Private Sub TBoxQTY_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Len(TBoxQTY.Value) > 2 Then GoTo 1
    If InStr("0123456789", ChrW(KeyAscii)) = 0 Then GoTo 1
    Exit Sub
1:
    KeyAscii = 0
End Sub

Private Sub test_Click()
Call Navigator.Navi(ActiveCell.Row)
End Sub


Private Sub TextBoxLidos20Path_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Init = True Then Exit Sub
    With Application.FileDialog(msoFileDialogFolderPicker)
        .ButtonName = "Выбрать": .Title = "Выбор папки с файлами корзины"
        If .Show <> -1 Then Exit Sub
        SaveSetting MyAppName, "Settings", "FolderLidos20", .SelectedItems(1)
    End With
'Call UserForm_Initialize
End Sub

Private Sub UserForm_Activate()
'    MsgBox MonitorDpiSetting(MonitorHandleForHwnd(MainForm.hWnd), MDT_EFFECTIVE_DPI)
'    MsgBox MonitorDpiSetting(MonitorHandleForHwnd(MainForm.hWnd), MDT_ANGULAR_DPI)
'    MsgBox MonitorDpiSetting(MonitorHandleForHwnd(MainForm.hWnd), MDT_RAW_DPI)
End Sub

Private Sub UserForm_Initialize()
Init = True
Me.Caption = "Инспектор v " & Ver
Me.ChBAutoHieght.Caption = "Автоподбор высоты ячейки"
Me.ChBAutoHieght.Value = GetSetting(MyAppName, "Settings", "AutoHieght", True)
Me.ChBoxLink.Value = GetSetting(MyAppName, "Settings", "LinkInDescription", False)
Me.ChBoxAutoTrans.Value = GetSetting(MyAppName, "Settings", "AutoTransition", False)
Me.ChBoxLidos20.Value = GetSetting(MyAppName, "Settings", "Lidos20", False)
Me.TextBoxLidos20Path.Value = GetSetting(MyAppName, "Settings", "FolderLidos20", "")
Me.TBoxQTY.Value = GetSetting(MyAppName, "Settings", "QTY", Q)
Me.ChBoxActionToComment.Value = GetSetting(MyAppName, "Settings", "ActionToComment", False)
Me.PDFMode.Value = GetSetting(MyAppName, "Settings", "PDFMode", False)
Me.CodePage.Value = GetSetting(MyAppName, "Settings", "CodePage", False)

'добавляем базовые листы
MainForm.cboxSheets.Clear
Dim Wsh As Worksheet
Dim iG As Integer

For Each Wsh In MeWB.Sheets
    If Wsh.Visible = xlSheetVisible Then
        MainForm.cboxSheets.AddItem Wsh.Name, iG
        iG = iG + 1
    End If
Next

cboxSheets.ListRows = iG

'добавляем разделы
MainForm.cboxRazdel.Clear
For n = 1 To 10
On Error GoTo 1
MainForm.cboxRazdel.AddItem Range("_" & n).Value, n - 1
Next n
1:
Me.MultiPage1.Value = 0
Init = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then Unload Me
End Sub

Private Sub VstavitZap_Click()
'MsgBox ("находится в разработке. Запчасти будут вставляться из сохраненной корзины Lidos")
dv.Flash
End Sub


