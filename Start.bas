Attribute VB_Name = "Start"
Public Const Ver As String = "2.04"    'Версия ПО
Public shParts As Worksheet ' лист с запчастями с него всё начинается
Public shMain As Worksheet ' лист главный, отсюда заполняется шапка ДВ
Public shService As Worksheet ' лист для обслуживания работы макроса. находится на листе отчета пока что.
'Public shService As Worksheet ' лист для обслуживания приложения и сохранения информации
Public MeWB As Workbook ' Книга инспекторского отчета
Public NN As Name 'Переменная справочника имен книги
Public Nrazdel As Integer 'Номер текущего раздела
Public Npunkt As Integer 'Номер текущего пункта
Public PunktCount As Integer 'всего пунктов в разделе
Public DVrcount As Integer ' Количество строк запчастей
Public Const Editmode = False
Public Const MyAppName = "Inspector1"
Public Dpi As Integer
'Public FotoTrigger As Boolean


'Public KolRazdel As Integer 'Количество разделов в книге

Sub InspectorStart()
Call SetGlobDim

If TypeName(shParts) = "Nothing" Or TypeName(shService) = "Nothing" Then 'Если лист не найден, выходим из программы
    MsgBox ("Документ не подходит для использования. Скачать специальный бланк для инспекции")
    About.Show 1
    Exit Sub
End If
'проверяем соответствие версии бланка и программы
'Версия бланка указана на сервисном листе в shService.Range("V1")
Dim Ver1
Ver1 = Left(Ver, InStr(1, Ver, ".", vbTextCompare) - 1)
If Not Ver1 = shService.Range("V1").Text Then MsgBox ("Версия бланка (" & shService.Range("V1").Value & ") не соответствует версии программы (" & Ver1 & ")")
MainForm.Show
End Sub
Sub test()

End Sub
Function RazdelAddress() ' Определяет адреса разделов
Debug.Print ActiveWorkbook.Worksheets(1).Names.Count
Debug.Print ActiveWorkbook.Worksheets(1).Name
Debug.Print ActiveWorkbook.Worksheets(1).Name
End Function



