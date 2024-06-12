VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} About 
   Caption         =   "О приложении"
   ClientHeight    =   2520
   ClientLeft      =   104
   ClientTop       =   392
   ClientWidth     =   4016
   OleObjectBlob   =   "About.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()
End Sub

Function Subscr(yes As Boolean)
If Not GetSetting(MyAppName, "Settings", "SubscriptionMail", True) = yes Or GetSetting(MyAppName, "Settings", "FirstRun", True) = True Then
    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    With OutMail
        '.To = "Maxim.Shguljov@liebherr.com; Jurij.Voronkov@liebherr.com"
        .To = "Maxim.Shguljov@liebherr.com"
'tmp        .to = "Maxim.Shguljov@liebherr.com"
        .CC = ""
        .BCC = ""
        If yes = True Then
        .Subject = "Инспектор. Подписка. " & Ver
        .Body = "Присылать обновления"
        Else
        .Subject = "Инспектор. Отписка. " & Ver
        .Body = "Не присылать обновления"
        End If
        .Send
    End With
    On Error GoTo 0
SaveSetting MyAppName, "Settings", "SubscriptionMail", yes
SaveSetting MyAppName, "Settings", "FirstRun", False

    Set OutMail = Nothing
    Set OutApp = Nothing
End If
End Function

Private Sub CommandButton1_Click()
MsgBox ("в разработке")
'   Shell "explorer \\liebherr.i\lru\odinzovo\EMT Service\Справочник сервисного инженера\Бланки_Акты_Протоколы\Бланки тех отчетов ЕМТ MIN\", vbNormalFocus
End Sub


Private Sub Label1_Click()

End Sub

Private Sub MailTo_Click()
'ThisWorkbook.SendMail "Maxim Shguljov"
    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    With OutMail
        .To = "Maxim.Shguljov@liebherr.com; Jurij.Voronkov@liebherr.com"
        .CC = ""
        .BCC = ""
        .Subject = "Инспектор v " & Ver
        .Body = "Проблема, вопрос, предложение..."
        '.Attachments.Add ActiveWorkbook.FullName
        'You can add other files also like this
        '.Attachments.Add ("C:\test.txt")
        '.Send   'or use .Display
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Private Sub UserForm_Initialize()
Me.Caption = "Инспектор v " & Ver
Me.CommandButton1.Caption = "Ознакомиться с инструкцией, скачать обновления" & vbNewLine & "(требуется подключение VPN Liebherr)"
Me.CheckBox1.Value = GetSetting(MyAppName, "Settings", "Subscription", True)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
SaveSetting MyAppName, "Settings", "Subscription", Me.CheckBox1.Value
Call Subscr(GetSetting(MyAppName, "Settings", "Subscription", True))
End Sub
