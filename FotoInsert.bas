Attribute VB_Name = "FotoInsert"
'Option Explicit
'
'Автор: Жгулёв Максим
'Почта: Maxim.Shguljov@liebherr.com
'
Public MeWbName As String
Public MeWbPath As String
Dim TargetWbName As String
Public cel As Range
Public avFile As String 'путь к файлу с фотографией
'Dim TTL As Variant
'Dim n As Byte 'Номер фото
'Dim K As Byte 'Номер объекта Shape
'Dim c As Byte 'Количество объектов Shape на листе
Public CellHeight As Variant 'Высота ячейки
Public CellWidth As Variant 'Ширина ячейки
Public СеllRatio As Variant 'Пропорции ячейки
Public ShpRatio As Variant 'Пропорции картинки
Dim NRG(99) As String 'Массив адресов ячеек для вставки фото
Public shp As Shape
Public Rot As Integer 'Переменная положения фотографии 0,90,180,270 град
Public Const Clir = 2 'Клиренс
Public StopRepit As Boolean

Dim X
Dim Y
Dim r '1 - повернута на 90: 0 - не повернута.
Dim Illcount As String
Dim l As Integer
Dim UFKtop As Integer 'Координаты окна с комментом
Dim UFKleft As Integer 'Координаты окна с комментом
Public WshShell
Dim ch As Double
Dim Folder$
Dim SavePath As Boolean
Dim nm As String 'Номер для обращения к объекту Shape по имени
Dim ttt As Variant
Public Const Q = 350 ' качество по умолчанию
Public QTY As Integer 'Размер фото в пикселях
Dim Zum As Double
Dim BuferError As Boolean




Sub Проба()
Set cel = ActiveCell
With cel
    CellWidth = .MergeArea.Columns.Width
    CellHeight = .MergeArea.Rows.Height
End With
MsgBox CellWidth & " " & CellHeight
'Application.CommandBars.ExecuteMso ("PictureCrop")
'For Each cbar In CommandBars
'    Debug.Print cbar.Name & "," & cbar.NameLocal & "," & cbar.Visible
'Next


'Dim octl As CommandBarControl
 
'With Selection
'Set octl = Application.CommandBars.ExecuteMso("Help")
'    Set octl = Application.CommandBars.FindControl(ID:=6382)
'   Application.SendKeys "%e~"
'   Application.SendKeys "%a~"
'    octl.Execute
'End With


'MsgBox shp.PictureFormat.Application
End Sub



Sub SaveAsPicture()
Zum = ActiveWindow.Zoom
ActiveWindow.Zoom = 100


Dim oObj As Object ', wsTmpSh As Worksheet
    
'If VarType(Selection) <> vbObject Then
'        MsgBox "Выделенная область не является объектом!", vbCritical, "www.excel-vba.ru"
'    Exit Sub
'End If
    
  
Set shp = ActiveSheet.Shapes(Selection.Name)
    'Application.CommandBars.ExecuteMso ("PictureResetAndSize")
    Dim MonitorDPI As Integer
    Dim shW As Single 'Ширина картинки
    Dim shH As Single 'Высота картинки
    Dim ojW As Single 'Ширина картинки
    Dim ojH As Single 'Высота картинки
    Dim WxH As Single 'Коэффициент
    Dim CropW As Single 'Значение обрезки по ширине
    Dim CropH As Single 'Значение обрезки по высоте
    'MsgBox VarType(Selection.ShapeRange.PictureFormat.Crop.PictureWidth)
    MonitorDPI = GetDpi
'    GoTo 1
      shp.ScaleWidth 1, msoTrue
      shp.ScaleHeight 1, msoTrue
1:
'If shp.Rotation = 0 Or shp.Rotation = 180 Then




    shW = shp.Width
    shH = shp.Height
    CropW = shp.PictureFormat.Crop.PictureWidth / shW
    CropH = shp.PictureFormat.Crop.PictureHeight / shH
    WxH = shW / shH
        If shp.Rotation = 90 Or shp.Rotation = 270 Then
            shp.Height = QTY / WxH / MonitorDPI * 72 / CropW
            shp.Width = shp.Height * WxH  ' - CropH
        Else
            shp.Width = QTY / MonitorDPI * 72 / CropW
            shp.Height = shp.Width / WxH  ' - CropH
        End If
'Else
'    'shW = shp.Height
'    'shH = shp.Width
'    shW = shp.Width
'    shH = shp.Height
'    CropW = shp.PictureFormat.Crop.PictureWidth / shW
'    CropH = shp.PictureFormat.Crop.PictureHeight / shH
'    WxH = shW / shH
'    shp.Width = QTY / MonitorDPI * 72 / CropW
'    shp.Height = shp.Width / WxH  ' - CropH
'    'shp.Height = QTY / MonitorDPI * 72 / CropW
'    'shp.Width = shp.Height / WxH  ' - CropH'
'
'End If
'shp.PictureFormat.Crop.PictureWidth = QTY / MonitorDPI * 72
'shp.PictureFormat.Crop.PictureHeight = QTY / MonitorDPI * 72 / WxH

   
Set oObj = Selection: oObj.Copy
    If shp.Rotation = 0 Or shp.Rotation = 180 Then
        ojW = oObj.Width
        ojH = oObj.Height
    Else
        ojW = oObj.Height
        ojH = oObj.Width
    End If




'Делаем текущую вкладку временной
'Set wsTmpSh = ActiveSheet
'В временной папке создаем временный файл
    Set FSO = CreateObject("Scripting.FileSystemObject")
    tmpName = Environ$("Temp") & "\" & FSO.GetTempName
    pos = InStrRev(tmpName, "."): tmpName = Mid$(tmpName, 1, pos) & "jpg"

'sName = ActiveWorkbook.FullName & "_" & ActiveSheet.Name & "_" & oObj.Name
'avFile = ActiveWorkbook.FullName & "_" & oObj.Name & ".jpg"
QTY = QTY / CropW

avFile = tmpName
    With ActiveSheet.ChartObjects.Add(30, 30, ojW, ojH).Chart
        .ChartArea.Border.LineStyle = 0
        .Parent.Select
        .Paste
        .Export FileName:=avFile, FilterName:="JPG"
        .Parent.Delete
    End With
ActiveWindow.Zoom = Zum
End Sub

Sub PictPos()
Set shp = ActiveSheet.Shapes(Selection.Name)
Set cel = shp.TopLeftCell
End Sub

Function GetQTY()
Dim kQ As Single 'множитель
    kQ = GetSetting(MyAppName, "Settings", "QTY", Q) / 100
        With cel
            CellWidth = .MergeArea.Columns.Width
            CellHeight = .MergeArea.Rows.Height
        End With

    If CellWidth > CellHeight Then QTY = CellWidth * kQ Else QTY = CellHeight * kQ
GetQTY = QTY
End Function

Function CompressFoto(File As String, Rasmer As Integer)
    Dim Img As Object 'файл
    Dim IP As Object 'процесс
    Dim j As Integer
    Dim FullPath As String
    Dim Name As String
    Dim Folder As String
    Dim Name_I As String
'    Folder = ""
'   создаем объект Windows Image Acquisition (WIA)
    Set Img = CreateObject("WIA.ImageFile")
    Set IP = CreateObject("WIA.ImageProcess")
'   Загружаем фотографию
On Error GoTo 1
    Img.LoadFile File
'   Проверка размера фотки. Если меньше требуемого, размер не меняется
        If Img.Height >= Rasmer Or Img.Width >= Rasmer Then
'   Изменяем текущий размер до заданного
            IP.Filters.Add IP.FilterInfos("Scale").FilterID
            IP.Filters(1).Properties("MaximumWidth") = Rasmer
            IP.Filters(1).Properties("MaximumHeight") = Rasmer
'   Заменяем загруженную фотографию уменьшенной
            Set Img = IP.Apply(Img)
        End If
'   Сохраняем в временном файле
    Set FSO = CreateObject("Scripting.FileSystemObject")
    tmpName = Environ$("Temp") & "\" & FSO.GetTempName
    pos = InStrRev(tmpName, "."): tmpName = Mid$(tmpName, 1, pos) & "jpg"

Img.SaveFile tmpName

CompressFoto = tmpName
Exit Function
1:
MsgBox ("Не верный тип файла в буфере")
BuferError = True
End Function

Sub SetShp()
Zum = ActiveWindow.Zoom
ActiveWindow.Zoom = 100

Set shp = ActiveSheet.Shapes.AddPicture(avFile, False, True, cel.MergeArea.Left, cel.MergeArea.Top, -1, -1)

shp.Select

'MsgBox QTY
'MsgBox shp.Width
    
    With shp
        .LockAspectRatio = msoFalse
        Call Addshp
    End With
ActiveWindow.Zoom = Zum
End Sub

Sub InsertFoto()
BuferError = False
StopRepit = False
'ttt = Timer
'Application.ScreenUpdating = False
Set cel = ActiveCell

TargetWbName = ActiveWorkbook.Name
'n = 1
'c = Workbooks(TargetWbName).Sheets(2).Shapes.Count
 'Задаем переменную пути файла
 'Вставляем из буфера
 
If MainForm.CBFromBufer = True Then
   avFile = ClipboardGetFiles()
Else
   avFile = GetFilePath()
End If
'останавливаем повтор если выбора нет
If avFile = "" Then
    StopRepit = True
    Exit Sub
End If
'сжимаем фотографию
avFile = CompressFoto(avFile, GetQTY)
If BuferError = True Then Exit Sub
'ttt = Timer
'If avFile = "" Then Unload UFKomment
'If avFile = "" Then Unload UFSettings
Call SetShp

UFKomment.Show
'Проверка условия вставки фото в описание. Доп переменная boolean \\\\\\||||||||////////
'If FotoTrigger = True Then
'Range(MeWB.Names("Foto" & Nrazdel & "." & Npunkt)).Offset(2, 16).Activate
'End If
'FotoTrigger = False

'Вставляем описание над фото
'    If Not StrPtr(TTL) = 0 And Ch = 0 Then Workbooks(TargetWbName).Sheets(2).Range(NRG(n - 1)).Offset(-1, 0).Formula = TTL
 '       Workbooks(TargetWbName).Sheets(2).Range(NRG(n - 1)).Offset(-1, 0).WrapText = True ' Перенос сторок текста
'MsgBox (Timer - ttt & " секунд на выполнение")
'Application.ScreenUpdating = True
End Sub
Sub Addshp()
On Error Resume Next
If MainForm.ChBAutoHieght = True Then
    cel.RowHeight = cel.MergeArea.Columns.Width * 3 / 4
'    Cel.Rows.Height = Cel.MergeArea.Columns.Width * 3 / 4
End If
Call Shprot
'shp.Select
r = 0
End Sub

Public Sub Shprot()
On Error GoTo 44
'MsgBox (shp.Name)
'If Clir = 0 Then
'Call Dimset 'Заполняем массив адресами ячеек, в которые будет происходить вставка фотографий
'End If
With cel
    CellWidth = .MergeArea.Columns.Width
    CellHeight = .MergeArea.Rows.Height
    СеllRatio = CellWidth / CellHeight
    X = .MergeArea.Left
    Y = .MergeArea.Top
End With
'shp.Select
With shp
ShpRatio = (.Width + Clir) / (.Height + Clir)
Rot = .Rotation
'.Rotation = Rot
    
    If (Rot = 90 Or Rot = -90 Or Rot = 270) Then
    ShpRatio = 1 / ShpRatio
        If ShpRatio < СеllRatio Then
'        Debug.Print 1
            .Height = (CellHeight - Clir) * ShpRatio ' - Clir
            .Width = CellHeight - Clir
        ElseIf ShpRatio > СеllRatio Then
'        Debug.Print 2
            .Height = CellWidth - Clir
            .Width = (CellWidth - Clir) / ShpRatio '- Clir
        End If
            GoTo 33
    End If
   
    If ShpRatio >= СеllRatio Then 'Широкая, не повернутая
'        Debug.Print 3
            .Height = (CellWidth - Clir) / ShpRatio ' - Clir
            .Width = CellWidth - Clir
    ElseIf ShpRatio <= СеllRatio Then 'Узкая не повернутая
'        Debug.Print 4
            .Width = ShpRatio * (CellHeight - Clir)
            .Height = CellHeight - Clir  'Меняем размер по высоте
    End If
33:
    shp.Top = Y + CellHeight / 2 - shp.Height / 2
    shp.Left = X + CellWidth / 2 - shp.Width / 2 + 0.5
        If X + CellWidth / 2 - shp.Width / 2 + 0.5 < 0 Then
            .IncrementLeft X + CellWidth / 2 - shp.Width / 2 + 0.5
        End If
End With
44:
End Sub

Function GetFilePath(Optional ByVal Title As String = "Выберите фото для вставки", _
                     Optional ByVal InitialPath As String = "C:\", _
                     Optional ByVal FilterDescription As String = "Фотографии", _
                     Optional ByVal FilterExtention As String = "*.jpg;*.png;*.bmp;*.jpeg") As String

    On Error Resume Next
    With Application.FileDialog(msoFileDialogOpen)
        .ButtonName = "Выбрать": .Title = Title:
            If Folder$ = "" Then
                .InitialFileName = GetSetting(MyAppName, "Settings", "folder", InitialPath)
            Else
                .InitialFileName = Folder$
            End If
        .Filters.Clear: .Filters.Add FilterDescription, FilterExtention
            If .Show <> -1 Then Exit Function
        GetFilePath = .SelectedItems(1)
        Folder$ = Left(.SelectedItems(1), InStrRev(.SelectedItems(1), "\"))
            If GetSetting(MyAppName, "Settings", "SaveFolder", True) = True Then
                SaveSetting MyAppName, "Settings", "folder", Folder$
            End If
    End With
End Function

