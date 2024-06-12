Attribute VB_Name = "ModUI"
Option Compare Text
  
Public Sub Update_Excel_OfficeUI_file()
    'On Error Resume Next
  
    ' ==========================      ниже можно менять      ==========================
    MacroFile$ = Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\AddIns\Inspector.xlam"
  
    Prefix$ = "<mso:customUI xmlns:x1=""http://schemas.microsoft.com/office/2009/07/customui/macro"" xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">"
    BUTTON_1$ = "<mso:button idQ=""x1:On_Off_2"" visible=""true"" label=""Инспектор"" imageMso=""Head"" onAction=""%filename%!InspectorStart""/>"
    'BUTTON_2$ = "<mso:button idQ=""x1:Fichier_Actif_1"" visible=""true"" label=""Запчасти"" imageMso=""Bullets"" onAction=""%filename%!ImportXML""/>"
    'BUTTON_3$ = "<mso:button idQ=""x1:Fichiers_Ouverts_1"" visible=""true"" label=""Bouton_A3_File_Memory_Enregistre_tous_les_Fichiers_Ouverts"" imageMso=""OutlinePromoteToHeading"" onAction=""%filename%!Bouton_A3_File_Memory_Enregistre_tous_les_Fichiers_Ouverts""/>"
    ' ========================== выше можно менять ==========================
  
  ' бэкап отключен
    Folder$ = Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\"
    FileName$ = "Excel.officeUI"
    'BackupFilename$ = Format(Now, "yyyy-mm-dd hh-nn-ss ") & FileName$
    'FileCopy Folder$ & FileName$, Environ("TMP") & "\!BACKUP " & BackupFilename$
  
On Error GoTo Error1
        With CreateObject("ADODB.Stream")    ' Загружаем данные из файла
        .Type = 2: .Charset = "utf-8": .Open
        .LoadFromFile Folder$ & FileName$
        txt = .ReadText:
return1: .Close
    End With
'ThisWorkbook.Worksheets(1).Cells(1, 1) = txt
    Const RIBBON_PREFIX$ = "<mso:ribbon>"
    If txt Like "*" & RIBBON_PREFIX$ & "*" Then
        Arr = Split(txt, RIBBON_PREFIX$, 2)
        Arr(0) = Prefix$
        txt = Join(Arr, RIBBON_PREFIX$)
    End If
  
    Const BUTTONS_SUFFIX$ = "</mso:sharedControls>"
    BUTTON_1$ = Replace(BUTTON_1$, "%filename%", MacroFile$)
    BUTTON_2$ = Replace(BUTTON_2$, "%filename%", MacroFile$)
    'BUTTON_3$ = Replace(BUTTON_3$, "%filename%", MacroFile$)
  
    ' Добавляем кнопку только один раз, повторно не добавляется
    If InStr(1, txt, BUTTON_1$) = 0 Then txt = Replace(txt, BUTTONS_SUFFIX$, BUTTON_1$ & BUTTONS_SUFFIX$)
    If InStr(1, txt, BUTTON_2$) = 0 Then txt = Replace(txt, BUTTONS_SUFFIX$, BUTTON_2$ & BUTTONS_SUFFIX$)
    'If InStr(1, txt, BUTTON_3$) = 0 Then txt = Replace(txt, BUTTONS_SUFFIX$, BUTTON_3$ & BUTTONS_SUFFIX$)
   
    ' сохраняем изменения в файле с учетом кодировки utf-8 no BOM
    With CreateObject("ADODB.Stream")
        .Type = 2: .Charset = "utf-8": .Open
        .WriteText txt
  
        Set binaryStream = CreateObject("ADODB.Stream")
        binaryStream.Type = 1: binaryStream.Mode = 3: binaryStream.Open
        .Position = 3: .CopyTo binaryStream        'Skip BOM bytes
        .flush: .Close
        binaryStream.SaveToFile Folder$ & FileName$, 2
        binaryStream.Close
    End With
Exit Sub:
Error1:     'Обработка ошибки
    txt = ThisWorkbook.Worksheets("Лист1").Cells(1, 1)
GoTo return1
End Sub

