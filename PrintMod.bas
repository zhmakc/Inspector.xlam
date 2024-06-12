Attribute VB_Name = "PrintMod"
Public Function PageNumber()
    'для работы этой функции нужен страничный режим. Переключаем.
    ActiveWindow.View = xlPageBreakPreview

    Dim VPB As Excel.VPageBreak
    Dim HPB As Excel.HPageBreak
    Dim intVPBC As Integer
    Dim intHPPC As Integer
    Dim lngPage As Long
    
    lngPage = 1
    
    If ActiveSheet.PageSetup.Order = xlDownThenOver Then
        intHPPC = ActiveSheet.HPageBreaks.Count + 1
        intVPBC = 1
    Else
        intVPBC = ActiveSheet.VPageBreaks.Count + 1
        intHPPC = 1
    End If

    For Each VPB In ActiveSheet.VPageBreaks
        If VPB.Location.Column > ActiveCell.Column Then
            Exit For
        End If
        
        lngPage = lngPage + intHPPC
    Next VPB
    
    For Each HPB In ActiveSheet.HPageBreaks
        If HPB.Location.Row > ActiveCell.Row Then
            Exit For
        End If
            
        lngPage = lngPage + intVPBC
    Next HPB
    PageNumber = lngPage
    'MsgBox "Номер страницы активной ячейки = " & lngPage
End Function

Public Sub NewPage()
    Dim PN As Long
    
    PN = PageNumber 'узнаем номер страницы
    
    'Проверка на последнюю страницу
    If ActiveSheet.HPageBreaks.Count + 1 = PN Then
        MsgBox "Это последняя страница"
        Exit Sub
    End If
    
    'Устанавливаем новую границу листа.
    Set ActiveSheet.HPageBreaks(PN).Location = Range(ActiveCell(1, 1).Address)
    
End Sub

Sub ResetPage()
ActiveWindow.View = xlPageBreakPreview
On Error Resume Next
    ActiveSheet.HPageBreaks(1).DragOff Direction:=xlDown, RegionIndex:=1
On Error GoTo 0
        Application.PrintCommunication = False
    With ActiveSheet.PageSetup
'        .LeftHeader = "&""-,полужирный""&36&KFFC000____________________________"
'        .CenterHeader = "&""Liebherr,обычный""&28LIEBHERR"
'        .RightHeader = ""
'        .LeftFooter = ""
'        .CenterFooter = ""
'        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0.78740157480315)
'        .RightMargin = Application.InchesToPoints(0.196850393700787)
'        .TopMargin = Application.InchesToPoints(0.905511811023622)
'        .BottomMargin = Application.InchesToPoints(0.590551181102362)
'        .HeaderMargin = Application.InchesToPoints(0.196850393700787)
'        .FooterMargin = Application.InchesToPoints(0.196850393700787)
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .PrintQuality = 600
'        .CenterHorizontally = False
'        .CenterVertically = False
'        .Orientation = xlPortrait
'        .Draft = False
'        .PaperSize = xlPaperA4
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
'        .PrintErrors = xlPrintErrorsDisplayed
'        .OddAndEvenPagesHeaderFooter = False
'        .DifferentFirstPageHeaderFooter = False
'        .ScaleWithDocHeaderFooter = True
'        .AlignMarginsHeaderFooter = True
'        .EvenPage.LeftHeader.Text = ""
'        .EvenPage.CenterHeader.Text = ""
'        .EvenPage.RightHeader.Text = ""
'        .EvenPage.LeftFooter.Text = ""
'        .EvenPage.CenterFooter.Text = ""
'        .EvenPage.RightFooter.Text = ""
'        .FirstPage.LeftHeader.Text = ""
'        .FirstPage.CenterHeader.Text = ""
'        .FirstPage.RightHeader.Text = ""
'        .FirstPage.LeftFooter.Text = ""
'        .FirstPage.CenterFooter.Text = ""
'        .FirstPage.RightFooter.Text = ""
    End With
'End Sub
End Sub
