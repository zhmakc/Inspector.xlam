Attribute VB_Name = "Module1"
Sub Макрос2()
Attribute Макрос2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос2 Макрос
'

'
    Range("C3").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Range("A3").Select
End Sub
