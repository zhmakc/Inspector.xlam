Attribute VB_Name = "ModPict"
Option Explicit

Private Const S_OK = 0

Private Enum MONITOR_DPI_TYPE
  MDT_EFFECTIVE_DPI = 0
  MDT_ANGULAR_DPI = 1
  MDT_RAW_DPI = 2
  MDT_DEFAULT = MDT_EFFECTIVE_DPI
End Enum

Private Enum MONITOR_DEFAULTS
    MONITOR_DEFAULTTONULL = &H0&
    MONITOR_DEFAULTTOPRIMARY = &H1&
    MONITOR_DEFAULTTONEAREST = &H2&
    MONITOR_DEFAULT = 2

End Enum

#If VBA7 Then
    Private Declare PtrSafe Function GetDpiForMonitor Lib "shcore" (ByVal hMonitor As Long, ByVal dpiType As MONITOR_DPI_TYPE, ByRef dpiX As Long, ByRef dpiY As Long) As Long
    Private Declare PtrSafe Function GetScaleFactorForMonitor Lib "shcore" (ByVal hMonitor As Long, ByRef DEVICE_SCALE_FACTOR As Long) As Long
    Private Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Integer) As Integer
'    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare Function GetDpiForMonitor Lib "shcore" (ByVal hMonitor As Long, ByVal dpiType As MONITOR_DPI_TYPE, ByRef dpiX As Long, ByRef dpiY As Long) As Long
    Private Declare Function GetScaleFactorForMonitor Lib "shcore" (ByVal hMonitor As Long, ByRef DEVICE_SCALE_FACTOR As Long) As Long
    Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
    Private Declare Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Integer) As Integer
#End If


Function GetDpi()
    Dim dpiX As Long
    Dim dpiY As Long
Dim s As String
    
    If GetScaleFactorForMonitor(MonitorFromWindow(Application.hwnd, MONITOR_DEFAULTTONEAREST), dpiX) = S_OK Then
        GetDpi = dpiX / 100 * 96
'        MsgBox GetDpi
    Else: MsgBox "Ошибка GetDPI." & vbCrLf & "Сообщите разработчику."
    End If

End Function
   

Sub rerr()

    If GetDpiForMonitor(MonitorFromWindow(Application.hwnd, MONITOR_DEFAULT), MDT_DEFAULT, dpiX, dpiY) = S_OK Then
        'lblResults.Caption =
        s = "dpiX = " & CStr(dpiX) & " (" & CStr(CSng(dpiX * 100) / 96) & "%)" ' _
            '& vbNewLine _
            '& "dpiY = " & CStr(dpiY) & " (" & CStr(CSng(dpiY * 100) / 96) & "%)"

'Exit Sub
        Dim hMonitor As Long
        Dim dpiXangular As Long
        Dim dpiYangular As Long
        Dim dpiXRaw As Long
        Dim dpiYRaw As Long
        Dim dpiXEff As Long
        Dim dpiYEff As Long

        Dim iRet As Long

        hMonitor = MonitorFromWindow(Application.hwnd, MONITOR_DEFAULT)
        
        iRet = GetDpiForMonitor(hMonitor, _
                            MDT_ANGULAR_DPI, _
                            dpiXangular, _
                            dpiYangular)
        iRet = GetDpiForMonitor(hMonitor, _
                            MDT_RAW_DPI, _
                            dpiXRaw, _
                            dpiYRaw)
        iRet = GetDpiForMonitor(hMonitor, _
                            MDT_EFFECTIVE_DPI, _
                            dpiXEff, _
                            dpiYEff)

        s = s & vbCrLf & vbCrLf & _
            "App handle:        " & CStr(Application.hwnd) & vbCrLf & _
            "Monitor handle:    " & CStr(hMonitor) & vbCrLf & _
            "DPI angular:       " & CStr(dpiXangular) & "    " & CStr(dpiYangular) & vbCrLf & _
            "DPI raw:           " & CStr(dpiXRaw) & "    " & CStr(dpiYRaw) & vbCrLf & _
            "DPI eff:           " & CStr(dpiXEff) & "    " & CStr(dpiYEff)


'MsgBox Application.hWnd
'MsgBox MonitorFromWindow(Application.hWnd, MONITOR_DEFAULT)
'MsgBox MONITOR_DEFAULT
'MsgBox s


    Else
        s = "GetDpiForMonitor error " & CStr(Err.LastDllError)
    End If

MsgBox s

End Sub



Sub DisplayMonitorInfo()
Dim w As Long, h As Long
Dim tt As Variant
For tt = 0 To 10
w = GetSystemMetrics32(tt) ' width in points
h = GetSystemMetrics32(2) ' height in points
MsgBox Format(w, "#,##0") & " x " & Format(h, "#,##0"), vbInformation, "Monitor Size (width x height)"
Next tt
End Sub
























Sub asdsdfs() 'показывают информацию по картинке
'Dim Img 'As ImageFile
Dim p As Property
'Set Img = CommonDialog1.ShowAcquireImage
'
'For Each p In Img.Properties
'    Dim s As String''

'    s = p.Name & "(" & p.PropertyID & ") = "
'    If p.IsVector Then
'        s = s & "[vector data not emitted]"
'    ElseIf p.Type = RationalImagePropertyType Then
'        s = s & p.Value.Numerator & "/" & p.Value.Denominator
'    ElseIf p.Type = StringImagePropertyType Then
'        s = s & """" & p.Value & """"
'    Else
'        s = s & p.Value
'    End If

'    Debug.Print s
'Next
'Exit Function


's = ""
'Dim v As Vector

'Set Img = CreateObject("WIA.ImageFile")


's = "Width = " & Img.Width & vbCrLf & _
'    "Height = " & Img.Height & vbCrLf & _
'    "Depth = " & Img.PixelDepth & vbCrLf & _
'    "HorizontalResolution = " & Img.HorizontalResolution & vbCrLf & _
'    "VerticalResolution = " & Img.VerticalResolution & vbCrLf & _
'   "FrameCount = " & Img.FrameCount & vbCrLf

'If Img.IsIndexedPixelFormat Then
'    s = s & "Pixel data contains palette indexes" & vbCrLf
'End If
'
'If Img.IsAlphaPixelFormat Then
'    s = s & "Pixel data has alpha information" & vbCrLf
'End If

'If Img.IsExtendedPixelFormat Then
'    s = s & "Pixel data has extended color information (16 bit/channel)" & vbCrLf
'End If

'If Img.IsAnimated Then
'    s = s & "Image is animated" & vbCrLf
'End If

'If Img.Properties.Exists("40091") Then
'    Set v = Img.Properties("40091").Value
'    s = s & "Title = " & v.String & vbCrLf
'End If

'If Img.Properties.Exists("40092") Then
'    Set v = Img.Properties("40092").Value
'    s = s & "Comment = " & v.String & vbCrLf
'End If

'If Img.Properties.Exists("40093") Then
'    Set v = Img.Properties("40093").Value
'    s = s & "Author = " & v.String & vbCrLf
'End If

'If Img.Properties.Exists("40094") Then
'    Set v = Img.Properties("40094").Value
'    s = s & "Keywords = " & v.String & vbCrLf
'End If

'If Img.Properties.Exists("40095") Then
'    Set v = Img.Properties("40095").Value
'    s = s & "Subject = " & v.String & vbCrLf
'End If

'MsgBox s
End Sub


