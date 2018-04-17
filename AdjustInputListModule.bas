Attribute VB_Name = "AdjustInputListModule"
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

Public Sub adjustInputList()
    innerAdjustInputList
End Sub


Private Sub innerAdjustInputList()
    
    Dim sh As Worksheet
    Dim i As Worksheet
    Set sh = ThisWorkbook.Sheets(FFOC.G_SH_NM_PLT_LIST)
    Set i = ThisWorkbook.Sheets(FFOC.G_SH_NM_IN)
    
    
    Dim ir As Range
    Set ir = i.Range("A2")
    Dim pltr As Range
    
    
    Do
        Set pltr = sh.Range("A2")
        
        If Len(ir.Value) = 1 Or (Len(ir.Value) = 2 And IsNumeric(ir.Value)) Then
        
        
            Set pltr = sh.Range("A2")
            Do
                If ir.Value = pltr.Value Then
                    ir.Offset(0, 2).Value = pltr.Offset(0, 3).Value
                    Exit Do
                End If
                Set pltr = pltr.Offset(1, 0)
            Loop Until Trim(pltr) = ""
            
            
            If ir.Offset(0, 2).Value = "" Then
                ir.Offset(0, 2).Value = "MANUAL"
            End If
        
        Else
            
            ' najpierw dopasuj nazwe plantu
            Set pltr = sh.Range("A2")
            
            Do
                ptrn = CStr(UCase(Trim(Replace(Replace(CStr(pltr.Offset(0, 1).Value), "Corail", ""), "Maestro", ""))))
            
                If UCase(ir.Value) Like "*" & ptrn & "*" And Trim(ptrn) <> "" Then
                        ir.Value = pltr.Value
                        ir.Offset(0, 2).Value = pltr.Offset(0, 3).Value
                        Exit Do
                End If
                
                Set pltr = pltr.Offset(1, 0)
            Loop Until Trim(pltr) = ""
            
            If ir.Offset(0, 2).Value = "" Then
                ir.Offset(0, 2).Value = "MANUAL"
            End If
            
        End If
        Set ir = ir.Offset(1, 0)
    Loop Until Trim(ir) = ""
End Sub
