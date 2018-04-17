Attribute VB_Name = "FirstRunoutModule"
Public Sub firstRunoutFormulaFilling(sh As Worksheet, r As Range)


    Dim lastRow As Long, lastColumn As Long
    lastRow = r.End(xlDown).Row
    lastColumn = r.End(xlToRight).Column


    Dim fr_r As Range
    Dim calcArea As Range
    
    
    For x = r.Offset(1, 0).Row To lastRow
        Set fr_r = r.Parent.Cells(x, r.Offset(0, FFOC.E_COMMON_FIRST_RUNOUT).Column - 1)
        Set calcArea = sh.Range(sh.Cells(x, FFOC.E_COMMON_FIRST_BALANCE + 1), sh.Cells(x, lastColumn))
        fr_r.Formula = "=FirstRunout(" & calcArea.AddressLocal & ")"
    Next x
        

End Sub

Private Sub firstRunoutFormulaFillingSide()


    Dim r As Range
    Set r = Range("b4")

    Dim lastRow As Long, lastColumn As Long
    lastRow = r.End(xlDown).Row
    lastColumn = r.End(xlToRight).Column


    Dim fr_r As Range
    Dim calcArea As Range
    
    
    For x = r.Offset(1, 0).Row To lastRow
        Set fr_r = r.Parent.Cells(x, r.Offset(0, FFOC.E_COMMON_FIRST_RUNOUT).Column - 1)
        Set calcArea = Range(Cells(x, FFOC.E_COMMON_FIRST_BALANCE + 1), Cells(x, lastColumn))
        fr_r.Formula = "=FirstRunout(" & calcArea.AddressLocal & ")"
    Next x
        

End Sub




Public Function FirstRunout(area As Range) As String

    FirstRunout = "#"
    Dim r As Range
    For Each r In area
        If r.Parent.Cells(4, r.Column).Value = "BALANCE" Then
            If r.Value < 0 Then
                FirstRunout = CStr(r.Parent.Cells(3, r.Column - 3).Value)
                Exit Function
            End If
        End If
    Next r
End Function

