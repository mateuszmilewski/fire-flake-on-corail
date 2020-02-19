VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TheLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'
'Copyright (c) 2017 FORREST
' Mateusz Milewski mateusz.milewski@opel.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.


Private colors As ColorsHandler
Private jezyk As E_LANG




Public Sub setLang(eTmp As E_LANG)
    
    jezyk = eTmp
End Sub


Public Sub init(ByRef sh As Worksheet)
    
    ActiveWindow.Zoom = 80
    ActiveWindow.DisplayGridlines = False
End Sub


' TheLayout class
Public Function fillFirstRawLabels(r As Range) As Range

    
    r.Value = "PART"
    r.Offset(0, FFOC.E_COMMON_PLT_CODE - 1).Value = "Plant Code"
    r.Offset(0, FFOC.E_COMMON_PLT_NAME - 1).Value = "Plant Name"
    r.Offset(0, FFOC.E_COMMON_PART_NAME - 1).Value = "Part Name"
    r.Offset(0, FFOC.E_COMMON_SUPPLIER - 1).Value = "Supplier"
    r.Offset(0, FFOC.E_COMMON_RESP - 1).Value = "Resp"
    r.Offset(0, FFOC.E_COMMON_CMNT1 - 1).Value = "Comment #1"
    r.Offset(0, FFOC.E_COMMON_CMNT2 - 1).Value = "Comment #2"
    r.Offset(0, FFOC.E_COMMON_FIRST_RUNOUT - 1).Value = "FIRST RUNOUT"
    r.Offset(0, FFOC.E_COMMON_BACKLOG - 1).Value = "Backlog"
    r.Offset(0, FFOC.E_COMMON_Blockages_in_progress - 1).Value = "BIP"
    r.Offset(0, FFOC.E_COMMON_STOCK - 1).Value = "Stock"
    r.Offset(0, FFOC.E_COMMON_RECV - 1).Value = "Recv"
    
    
    Set fillFirstRawLabels = r.Offset(0, FFOC.E_COMMON_FIRST_RQM - 1)
End Function


Public Function getDatesFromDropperAndFillLabels(r As Range, c As Collection) As Range

    Dim d1 As Date
    Dim d2 As Date
    
    Dim firstTime As Boolean
    firstTime = True

    Dim el As DropperHandler
    For Each el In c
        
        If CBool(firstTime) Then
            
            d1 = el.getFirstDate
            d2 = el.getLastDate
        Else
            
            If d1 > el.getFirstDate Then
                d1 = el.getFirstDate
            End If
            
            If d2 < el.getLastDate Then
                d2 = el.getLastDate
            End If
        End If
        firstTime = False
    Next el
    
    
    Dim tempDate As Date
    tempDate = d1
    
    If CLng(tempDate) = 0 Then tempDate = Date
    Do
        
        r.Offset(-1, 0).NumberFormat = "@"
        r.Offset(-1, 0).Value = "" + checkDateFormatAndReturnProperValue(tempDate) ' CStr(DateValue(tempDate))
        
        If CStr(tempDate) <> CStr(r.Offset(-1, 0).Value) Then
        
            ' swap once again month and day
            ' r.Offset(-1, 0).Value = Mid(r.Offset(-1, 0).Value, 4, 2) & "/" & Left(r.Offset(-1, 0).Value, 2) & "/" & Right(r.Offset(-1, 0).Value, 4)
            r.Offset(-1, 0).Value = ""
            r.Offset(-1, 0).Value = Left(tempDate, 2) & "/"
            r.Offset(-1, 0).Value = r.Offset(-1, 0).Value & Mid(tempDate, 4, 2) & "/"
            r.Offset(-1, 0).Value = r.Offset(-1, 0).Value & Right(tempDate, 4)
        End If
        
        r.Offset(-1, 3).Value = CStr(Weekday(tempDate, vbMonday))
        
        ' RQM, ORDER, DELIVERY, BALANCE labels
        r.Offset(0, 0).Value = CStr(ThisWorkbook.Sheets(FFOC.G_SH_NM_REG).Range("L1"))
        r.Offset(0, 1).Value = CStr(ThisWorkbook.Sheets(FFOC.G_SH_NM_REG).Range("L2"))
        r.Offset(0, 2).Value = CStr(ThisWorkbook.Sheets(FFOC.G_SH_NM_REG).Range("L3"))
        r.Offset(0, 3).Value = CStr(ThisWorkbook.Sheets(FFOC.G_SH_NM_REG).Range("L4"))
        
        Set r = r.Offset(0, G_ITEM_OFFSET)
    
        tempDate = tempDate + 1
    Loop Until CDate(tempDate) > CDate(d2)
    
    
    Set getDatesFromDropperAndFillLabels = r
    
End Function


Private Function checkDateFormatAndReturnProperValue(d As Date) As String


    tmp = "" + CStr(d)

    If tmp Like "??/??/????" Then
    
        checkDateFormatAndReturnProperValue = "" + Left(tmp, 2) + "/" + Mid(tmp, 4, 2) + "/" + Right(tmp, 4)
    Else
    
        checkDateFormatAndReturnProperValue = CStr(d)
    End If
    
    
End Function




' ponizej znajduja sie suby i funkcje,
' ktore pracuja na juz wykonanym raporcie,
' ktory narazie przechowuje tylko i wylacznie dane
' ------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------



' TOP ROW WITH DATES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub addDayNamesAndMergeTopDateRows(r As Range)
    
    addDayNamesAndCw r
    mergeProperly r
End Sub


Private Sub addDayNamesAndCw(r As Range)

    Dim regSh As Worksheet
    Set regSh = ThisWorkbook.Sheets("register")
    
    ' r = E3
    Dim tmp As Range
    Set tmp = r
    
    
    
    Set tmp = tmp.Offset(0, 3)
    
    Do
    
        ' quick CW assigning
        ' ------------------------------------------------------------------------
        tmp.Offset(-1, -2).Value = Application.WorksheetFunction.IsoWeekNum(CDate(tmp.Offset(0, -3)))
        ' ------------------------------------------------------------------------
    
        ' weekday
        ' ------------------------------------------------------------------------
        tmp.Value = CStr(regSh.Cells(Int(tmp.Value), Int(jezyk)).Value)
        ' ------------------------------------------------------------------------
        
        
        Set tmp = tmp.Offset(0, G_ITEM_OFFSET)
    Loop Until Trim(tmp) = ""

End Sub

Private Sub mergeProperly(ByRef r As Range)

    Dim sh  As Worksheet
    Set sh = r.Parent
    Dim tmp As Range, tmp1 As Range, tmp2 As Range
    
    Dim toggle As Boolean
    toggle = False
    
    Do
        Set tmp = r
        Set tmp1 = tmp
        Set tmp2 = r.Offset(0, 2)
        
        Set r = r.Offset(0, 3)
        
        tmp1.HorizontalAlignment = xlCenter
        tmp1.Offset(-1, 1).HorizontalAlignment = xlLeft
        
        ' cw
        tmp1.Offset(-1, 0).Value = "CW: "
        sh.Range(tmp1.Offset(-1, 1), tmp2.Offset(-1, 1)).Merge
        
        ' date and weekday
        sh.Range(tmp1, tmp2).Merge
        
        
        
        
        tmp1.Font.Bold = True
        r.Font.Bold = True
        
        
        If toggle Then
            With sh.Range(tmp1.Offset(-1, 0), r)
                .Font.Color = colors.colorMattBlack
                .Font.Size = 10
                .Interior.Color = RGB(255, 255, 255)
            End With

        Else
            
            With sh.Range(tmp1.Offset(-1, 0), r)
                .Font.Color = colors.colorMattBlack
                .Font.Size = 10
                .Interior.Color = colors.colorFireFlakeGrey
            End With
            
        End If
        
        Me.fillSolidGridLines sh.Range(tmp1.Offset(-1, 0), r), colors.colorMattBlack
        
        
        Set r = r.Offset(0, 1)
        
        If toggle Then
            toggle = False
        Else
            toggle = True
        End If
        
    Loop Until Trim(r) = ""

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub adjustLabelRow(ByRef r As Range)


    Dim tmp As Range, innerTmp As Range, obszar As Range, toggle As Boolean
    Set tmp = r.Parent.Range(r, r.End(xlToRight))
    
    
    r.EntireRow.RowHeight = 70 ' set manually - visual check only
    Me.changeTxtOrientation tmp, 90
    
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = r.End(xlDown).Row
    lastCol = r.End(xlToRight).Column
    
    Set obszar = r.Parent.Range(r, r.Parent.Cells(lastRow, lastCol))
    Me.fillThinGridLines obszar, colors.colorMattBlack
    obszar.Font.Size = 9
    
    
    Dim forWholeColumn As Range, fwc As Range
    
    
    ' part number
    Set forWholeColumn = r.Parent.Cells(1, FFOC.E_COMMON_PN + 1)
    Set fwc = forWholeColumn
    fwc.EntireColumn.ColumnWidth = 13
    
    ' plt code
    Set forWholeColumn = r.Parent.Cells(1, FFOC.E_COMMON_PLT_CODE + 1)
    Set fwc = forWholeColumn
    fwc.EntireColumn.ColumnWidth = 5
    ' plt name
    Set forWholeColumn = r.Parent.Cells(1, FFOC.E_COMMON_PLT_NAME + 1)
    Set fwc = forWholeColumn
    fwc.EntireColumn.ColumnWidth = 10
    
    ' supplier
    Set forWholeColumn = r.Parent.Cells(1, FFOC.E_COMMON_SUPPLIER + 1)
    Set fwc = forWholeColumn
    fwc.EntireColumn.ColumnWidth = 13
    
    toggle = True
    Set fwc = r.Parent.Cells(r.Row, FFOC.E_COMMON_FIRST_RQM + 1)
    Do
    
        fwc.EntireColumn.ColumnWidth = 5
        fwc.Offset(0, 1).EntireColumn.ColumnWidth = 5
        fwc.Offset(0, 2).EntireColumn.ColumnWidth = 5
        fwc.Offset(0, 3).EntireColumn.ColumnWidth = 10
        
        fwc.Parent.Range(fwc, fwc.Offset(0, 3)).Font.Size = 11
        fwc.Offset(0, 3).Font.Bold = True
        fwc.Offset(0, 3).Font.Color = colors.colorMattBlack
        
        If toggle Then
            fwc.Parent.Range(fwc, fwc.Offset(lastRow - fwc.Row, 3)).Interior.Color = colors.colorFireFlakeGrey
        Else
            fwc.Parent.Range(fwc, fwc.Offset(lastRow - fwc.Row, 3)).Interior.Color = RGB(255, 255, 255)
        End If
    
        
        Set fwc = fwc.Offset(0, G_ITEM_OFFSET)
        toggle = Not toggle
    Loop Until Trim(fwc) = ""
End Sub



' DYNAMIC COLORS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub colourFontsOnOtherThanBalanceColumns(ByRef r As Range)


    adjustShipFormat r
    adjustRqmCellsFormat r

End Sub




Private Sub adjustRqmCellsFormat(ByRef r As Range)

    Dim tmpr As Range, area As Range
    Set tmpr = r.Offset(0, FFOC.E_COMMON_FIRST_RQM - 1)
    Do
        Set area = prepareArea(tmpr, r)
        area.Offset(0, -3).NumberFormat = "[Blue]0_ ;[Red]-0 ;0_ "
        Set tmpr = tmpr.Offset(0, G_ITEM_OFFSET)
    Loop Until Trim(tmpr) = ""
End Sub

Private Sub adjustShipFormat(ByRef r As Range)

    Dim tmpr As Range, area As Range
    Set tmpr = r.Offset(0, FFOC.E_COMMON_FIRST_SHIP - 1)
    Do
        Set area = prepareArea(tmpr, r)
        area.Offset(0, -1).NumberFormat = "[Blue]0_ ;[Red]-0 ;0_ "
        Set tmpr = tmpr.Offset(0, 4)
    Loop Until Trim(tmpr) = ""
End Sub


Private Function prepareArea(tmpr As Range, r As Range) As Range
    
    Dim tmp As Range
    Set tmp = tmpr
    Set tmp = r.Parent.Range(tmp, tmp.Parent.Cells(r.End(xlDown).Row, tmp.Column))
    Set prepareArea = tmp
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Public Sub backColor(rng As Range, kolor As Long)

    With rng.Interior
        .Pattern = xlSolid
        .Color = kolor
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
 
Public Sub changeTxtOrientation(rng As Range, angle As Integer)

    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = Int(angle)
        
    End With
End Sub


Public Sub fillSolidFrame(rng As Range, kolor As Long)


    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
End Sub

Public Sub fillSolidGridLines(rng As Range, kolor As Long)
    
    
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
End Sub

Public Sub fillThinFrame(rng As Range, kolor As Long)


    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Public Sub fillThinGridLines(rng As Range, kolor As Long)

    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub






' ------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    Set colors = New ColorsHandler
End Sub

Private Sub Class_Terminate()
    Set colors = Nothing
End Sub
