Attribute VB_Name = "Start"
Public Const Ver As String = "2.04"    '������ ��
Public shParts As Worksheet ' ���� � ���������� � ���� �� ����������
Public shMain As Worksheet ' ���� �������, ������ ����������� ����� ��
Public shService As Worksheet ' ���� ��� ������������ ������ �������. ��������� �� ����� ������ ���� ���.
'Public shService As Worksheet ' ���� ��� ������������ ���������� � ���������� ����������
Public MeWB As Workbook ' ����� �������������� ������
Public NN As Name '���������� ����������� ���� �����
Public Nrazdel As Integer '����� �������� �������
Public Npunkt As Integer '����� �������� ������
Public PunktCount As Integer '����� ������� � �������
Public DVrcount As Integer ' ���������� ����� ���������
Public Const Editmode = False
Public Const MyAppName = "Inspector1"
Public Dpi As Integer
'Public FotoTrigger As Boolean


'Public KolRazdel As Integer '���������� �������� � �����

Sub InspectorStart()
Call SetGlobDim

If TypeName(shParts) = "Nothing" Or TypeName(shService) = "Nothing" Then '���� ���� �� ������, ������� �� ���������
    MsgBox ("�������� �� �������� ��� �������������. ������� ����������� ����� ��� ���������")
    About.Show 1
    Exit Sub
End If
'��������� ������������ ������ ������ � ���������
'������ ������ ������� �� ��������� ����� � shService.Range("V1")
Dim Ver1
Ver1 = Left(Ver, InStr(1, Ver, ".", vbTextCompare) - 1)
If Not Ver1 = shService.Range("V1").Text Then MsgBox ("������ ������ (" & shService.Range("V1").Value & ") �� ������������� ������ ��������� (" & Ver1 & ")")
MainForm.Show
End Sub
Sub test()

End Sub
Function RazdelAddress() ' ���������� ������ ��������
Debug.Print ActiveWorkbook.Worksheets(1).Names.Count
Debug.Print ActiveWorkbook.Worksheets(1).Name
Debug.Print ActiveWorkbook.Worksheets(1).Name
End Function



