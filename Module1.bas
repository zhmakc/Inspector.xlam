Attribute VB_Name = "Module1"
Sub ������2()
Attribute ������2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������2 ������
'

'
    Range("C3").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Range("A3").Select
End Sub
