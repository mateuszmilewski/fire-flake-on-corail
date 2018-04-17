VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'
'Copyright (c) 2018 FORREST
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

Private strStock As String
Private strBalance As String
Private strRqm As String

Private r As Range
Private firstBalance As Range

Private balanceColumnsCollection As Collection

Private sh As Worksheet

Private Sub Class_Initialize()

    ' prepare elements to colour!
    ' -----------------------------------------
    strStock = "Stock"
    strBalance = "BALANCE"
    strRqm = "RQM"
    
    ' anchor PN top left!
    Set r = Range("B4")
    Set sh = r.Parent
    
    Set balanceColumnsCollection = fillBalanceColumnsCollection()
    
    
    For Each colNum In balanceColumnsCollection
    
        x = r.Offset(1, 0).Row
        
        Do
    
            If sh.Cells(x, colNum).Value < 0 Then
                sh.Cells(x, colNum).Interior.Color = RGB(255, 0, 0)
                sh.Cells(x, colNum).Font.Color = RGB(0, 0, 0)
            ElseIf sh.Cells(x, colNum + 4).Value < 0 Or sh.Cells(x, colNum).Value = 0 Then
                sh.Cells(x, colNum).Interior.Color = RGB(255, 199, 206)
                sh.Cells(x, colNum).Font.Color = RGB(0, 0, 0)
            ElseIf CLng(sh.Cells(x, colNum).Value) - CLng(sh.Cells(x, colNum + 1).Value) < 0 Then
                sh.Cells(x, colNum).Interior.Color = RGB(255, 199, 206)
                sh.Cells(x, colNum).Font.Color = RGB(0, 0, 0)
            Else
                sh.Cells(x, colNum).Interior.Color = sh.Cells(r.Row, colNum).Interior.Color
                sh.Cells(x, colNum).Font.Color = RGB(0, 0, 0)
            End If
            
            x = x + 1
        
        Loop Until Trim(sh.Cells(x, colNum).Value) = ""
        
        
    Next colNum
    
    
    ' -----------------------------------------

End Sub

Private Function fillBalanceColumnsCollection() As Collection
    Set fillBalanceColumnsCollection = Nothing
    
    Dim tmpC As Collection
    Set tmpC = New Collection
    
    Dim tmp As Range
    Set tmp = r
    Do
        If tmp.Value = strBalance Then
            tmpC.Add tmp.Column
        End If
        Set tmp = tmp.Offset(0, 1)
    Loop Until Trim(tmp.Value) = ""
    
    
    If tmpC.Count > 0 Then Set fillBalanceColumnsCollection = tmpC
End Function
    

Private Sub Class_Terminate()


    ' clear fields
    ' -----------------------------------------
    Set balanceColumnsCollection = Nothing
    ' -----------------------------------------

End Sub
