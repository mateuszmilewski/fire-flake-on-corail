Attribute VB_Name = "RunOnSelectionChangeModule"
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


Public Sub recalcLayoutAndColors(sh As Worksheet, r As Range)

    With Application
        .enableEvents = False
        .ScreenUpdating = False
        .Calculation = xlCalculationAutomatic
    End With
    
    If sh.Cells(1, 1).Value Like "Report;*" Then
        ' we're in report sheet so we can go on with the logic
        ' --------------------------------------------------------------------
        ''
        '
        
        Dim dc As DynamicColors
        Set dc = New DynamicColors
        '
        ''
        ' --------------------------------------------------------------------
    End If
    
    With Application
        .enableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub
