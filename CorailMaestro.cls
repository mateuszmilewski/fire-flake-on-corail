VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorailMaestro"
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


Implements ICorail

Private ieh As IEHandler
Private req As HTTPRequestHandler
Private parser As ParserHandler

Private Sub Class_Initialize()
    Set ieh = Nothing
    'Set req = New HTTPRequestHandler
    Set parser = New ParserHandler
End Sub

Private Sub Class_Terminate()
    Set ieh = Nothing
    Set parser = Nothing
End Sub


Private Sub ICorail_closeInnerIE()
    
    ieh.closeInnerIE
    
End Sub

Private Sub ICorail_generateInnerHttpRequest(r As Range, plt As String, pn As String)
    Set req = New HTTPRequestHandler
    req.init plt, pn
    req.braceWithDom MAESTRO
End Sub

Private Sub ICorail_generateInnerIE(r As Range, plt As String, pn As String)
    
    Set ieh = New IEHandler
    
    If ieh.checkIfTisAvailable Then
        ieh.pBlocked = True
        
        ieh.getNewInnerIE r, plt, pn, False
        ieh.innerWait
        
        If Not ieh.getFirstTimeOut() Then ieh.braceWithDom MAESTRO
    End If

End Sub

Private Function ICorail_getData() As ConvertedData

    Set ICorail_getData = Nothing
    
    If Not ieh Is Nothing Then
        Set ieh.dom.iElement = tryToGetTableWithDateAndRqmsFromMaestro(ieh)
        Set ieh.dom.table = ieh.dom.iElement
        Set ICorail_getData = parser.htmlTableToRawMatrix(ieh.dom.table, ieh.dom, MAESTRO)
    Else
        Set req.dom.iElement = tryToGetTableWithDateAndRqmsFromMaestroNewForHttpRequest()
        Set req.dom.table = req.dom.iElement
        Set ICorail_getData = parser.htmlTableToRawMatrix(req.dom.table, req.dom, MAESTRO)
    End If
    
End Function

Private Function tryToGetTableWithDateAndRqmsFromMaestroNewForHttpRequest()
    Set tryToGetTableWithDateAndRqmsFromMaestroNewForHttpRequest = Nothing
    
    Set req.dom.iElementCollection = req.dom.doc.getElementsByTagName("table")
    
    podlicz = 0
    
    For Each el In req.dom.iElementCollection
        Set req.dom.iElement = el
        
        ' Debug.Print req.dom.iElement.innerHTML
        
        If podlicz = 9 Then
            Set tryToGetTableWithDateAndRqmsFromMaestroNewForHttpRequest = req.dom.iElement
            Exit Function
        End If
        
        podlicz = podlicz + 1
    Next el
End Function

Private Function tryToGetTableWithDateAndRqmsFromMaestro(ieh As IEHandler)
    Set tryToGetTableWithDateAndRqmsFromMaestro = Nothing
    
    Set ieh.dom.iElementCollection = ieh.dom.doc.getElementsByTagName("table")
    
    podlicz = 0
    
    For Each el In ieh.dom.iElementCollection
        Set ieh.dom.iElement = el
        
        ' Debug.Print ieh.dom.iElement.innerHTML
        
        If podlicz = 9 Then
            Set tryToGetTableWithDateAndRqmsFromMaestro = ieh.dom.iElement
            Exit Function
        End If
        
        podlicz = podlicz + 1
    Next el
    
End Function



