VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFKomment 
   Caption         =   "Редактирование"
   ClientHeight    =   1300
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   2744
   OleObjectBlob   =   "UFKomment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFKomment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'
'Автор: Жгулёв Максим
'Почта: Maxim.Shguljov@liebherr.com
'

Private Sub cbChancel_Click()
End Sub

Private Sub cbChange_Click()
shp.Select
Selection.Delete
UFKomment.Hide
Call InsertFoto
End Sub

''''--------------------------------------------------------''''
Private Sub cbOK_Click()
UFKomment.Hide
End Sub

Private Sub CBPaint_Click()

'Редактируем фото в PAINT

'On Error GoTo TheEnd
UFKomment.Enabled = False

shp.Delete
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "c:\windows\system32\mspaint.exe " & """" & avFile & """", 3, True
SendKeys "%{TAB}"

avFile = CompressFoto(avFile, GetQTY)
Call FotoInsert.SetShp

UFKomment.Enabled = True
'Application.Wait Now + TimeValue("00:00:02")
'Application.Windows(TargetWbName).Activate
'TargetWb.Activate

'TheEnd:
End Sub

Private Sub cbRotary_Click()
shp.Select
With Selection
Rot = .ShapeRange.Rotation
Rot = Rot - 90
.ShapeRange.Rotation = Rot
End With
Call Shprot
End Sub

Private Sub cbRotary2_Click()
shp.Select
With Selection
Rot = .ShapeRange.Rotation
Rot = Rot + 90
.ShapeRange.Rotation = Rot
End With
Call Shprot
End Sub

Private Sub Settings_Click()
UFSettings.Show
End Sub

Private Sub UserForm_Activate()
'If l > 0 Then
'UFKomment.Top = UFKtop
'UFKomment.Left = UFKleft
'End If
'l = l + 1
'    If Ch = 1 Then
'    UFKomment.TextBox1.Value = TTL2
'    Else
'    UFKomment.TextBox1.Value = ""
'    End If
'UFKomment.TextBox1.SetFocus
End Sub
