Attribute VB_Name = "modTable"
'������ ��� �������������� ����� �� ����� � ����������
Dim r As Integer


Function Table() '��������� ����� �  ���������
'��������� ������� �� ���� "������ ���������"
a = Timer
If Not ActiveSheet.Name = shParts.Name Then shParts.Activate: MsgBox ("������� ����� ������� ������� �� ����� " & shParts.Name): Exit Function
'��������� ���������� ������
Application.ScreenUpdating = False
r = ActiveCell.Row
Dim sA As String
sA = r & ":" & (r + 5)
Rows(sA).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow '��������� ������
'������ ������
Call CommentTable(r)
With Cells(r, 2) '��������� ��� ���������
Navigator.Navi (ActiveCell.Row)
.Name = "Punkt" & Nrazdel & "." & PunktCount + 1
Navigator.Navi (ActiveCell.Row)
End With

'������ ������
With Range(Cells(r + 2, 1), Cells(r + 2, 4))
    .HorizontalAlignment = xlLeft
    .Merge
    .Value = "�������"
    .Font.Bold = True
End With
With Range(Cells(r + 2, 5), Cells(r + 2, 34))
    .HorizontalAlignment = xlLeft
    .Merge
End With
'��������� ������
With Range(Cells(r + 3, 1), Cells(r + 3, 4))
    .HorizontalAlignment = xlLeft
    .Merge
    .Value = "��������"
    .Font.Bold = True
End With
With Range(Cells(r + 3, 5), Cells(r + 3, 34))
    .HorizontalAlignment = xlLeft
    .Merge
End With
'����� ������
r = shService.Cells(Npunkt + 1, 4).Value
With Range(Cells(r + 4, 1), Cells(r + 4, 10))
    .HorizontalAlignment = xlLeft
    .Merge
    .Value = " �������� �� ���������"
    .Font.Bold = False
End With

'�����
Call Bord(Range(Cells(r, 1), Cells(r + 3, 34)))
                    '��� ��� ��������� ������
                        Dim WBN As Workbook
                        Set WBN = Workbooks.Add
                        WBN.Close
                    '��� ��� ��������� ������
Range(MeWB.Names("punkt" & Nrazdel & "." & Npunkt)).Offset(0, 3).MergeArea.Select '.Activate
Application.ScreenUpdating = True
'Range(MeWB.Names("punkt" & Nrazdel & "." & Npunkt)).Offset(0, 3)
'ActiveWindow.Activate
'AppActivate "Microsoft Excel"
'MsgBox (Timer - a & " ������ �� ����������")
End Function

Function CommentTable(r As Integer)
Cells(r, 1).Value = "�."
'Call SortirovatPunkty
With Range(Cells(r, 2), Cells(r, 4))
' ��������� �������� ����������=============================================================
    .HorizontalAlignment = xlLeft
    .Merge
    .Font.Bold = False
End With
With Range(Cells(r, 5), Cells(r + 1, 34))
    .Merge
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlTop
    .WrapText = True
    .Font.Bold = True
End With
'������ ������
With Range(Cells(r + 1, 1), Cells(r + 1, 4))
    .Merge
    .HorizontalAlignment = xlLeft
    .Value = "��������"
    .Font.Bold = True
End With

End Function

Function Bord(BorderRange As Range)
With BorderRange
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End With
End Function

