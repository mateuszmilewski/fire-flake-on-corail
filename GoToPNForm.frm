VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GoToPNForm 
   Caption         =   "Go To Form"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2790
   OleObjectBlob   =   "GoToPNForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GoToPNForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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


Private Sub BtnSubmit_Click()
    Hide
    
    If Me.TextBoxCorailType.Value <> "" Then
        If Me.TextBoxPlt.Value <> "" Then
            If Me.TextBoxPN.Value <> "" Then
            
            
                ' tylko i wylacznie wtedy mozemy uruchomic logike!
                
                openPartNumberOnProperCorail Me.TextBoxCorailType.Value, Me.TextBoxPlt.Value, Me.TextBoxPN.Value
            End If
        End If
    End If
    
End Sub

Private Sub openPartNumberOnProperCorail(corailType, plt, pn)
    
    
    Dim sh  As Worksheet
    Set sh = ThisWorkbook.Sheets(FFOC.G_SH_NM_PLT_LIST)
    
    Dim r As Range
    Set r = sh.Range("A2")
    Do
    
        If Trim(r.Value) = Trim(plt) And r.Offset(0, 3).Value = corailType Then
            
            If corailType <> "MAESTRO" Then openIEonProperPN r.Offset(0, 2), pn
            If corailType = "MAESTRO" Then openIEonProperPltAndPartNumberInMaestro CStr(r.Offset(0, 2)), CStr(pn)
            Exit Do
        End If
    
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Sub


Private Sub openIEonProperPltAndPartNumberInMaestro(pltLink As String, pn As String)


    
    Dim pltIE As InternetExplorer
    Dim pnIE As InternetExplorer
    
    Set pltIE = New InternetExplorer
    Set pnIE = New InternetExplorer
    
    pltIE.Visible = True
    pnIE.Visible = True
    
    pltIE.navigate CStr(pltLink)
    DoEvents
    Sleep 200
    
    
    maestroSupplForUrl = FFOC.G_MAESTRO_URL_EXT
    ' maestroBaseUrl = "http://maestro.inetpsa.com"
    maestroBaseUrl = ThisWorkbook.Sheets("plt-list").Range("C2").Value
    
    pnIE.navigate maestroBaseUrl & "/" & maestroSupplForUrl & pn
    
    
End Sub

Private Sub openIEonProperPN(link, pn)


    ' "http://ta.control.erp.corail.inetpsa.com/getProductSummaryRead.do?beanId=96661053ZD"
    nxtUrl = "getProductSummaryRead.do?beanId="
    
    Dim ie As InternetExplorer
    Set ie = New InternetExplorer
    ie.Visible = True
    ie.navigate CStr(link) & CStr(nxtUrl) & CStr(pn) & "#"
End Sub
