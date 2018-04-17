VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputListHelper"
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




Public Function getInputList() As Range
    
    Dim lr As Long
    lr = getLastRowFromInputList()
    
    If Trim(ThisWorkbook.Sheets(FFOC.G_SH_NM_IN).Range("A3").Value) <> "" Then
        Set getInputList = ThisWorkbook.Sheets(FFOC.G_SH_NM_IN).Range("A2:A" & CStr(lr))
    Else
        Set getInputList = ThisWorkbook.Sheets(FFOC.G_SH_NM_IN).Range("A2:A" & 2)
    End If
End Function


Private Function getLastRowFromInputList() As Long


    Dim r As Range
    Set r = ThisWorkbook.Sheets(FFOC.G_SH_NM_IN).Range("A2")

    Do
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    getLastRowFromInputList = CLng(r.Row - 1)
End Function

