Attribute VB_Name = "GlobalFooModule"
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

Public Function getPlantName(plt) As String
    getPlantName = ""
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(FFOC.G_SH_NM_PLT_LIST)
    
    
    Dim r As Range
    Set r = sh.Range("A2")
    Do
    
        If Trim(r.Value) = Trim(plt) Then
            getPlantName = Trim(Replace(Replace(r.Offset(0, 1).Value, "Maestro", ""), "Corail", ""))
            Exit Do
        End If
    
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    
    
    If getPlantName = "" Then getPlantName = plt
End Function
