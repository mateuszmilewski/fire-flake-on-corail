VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorailIteration"
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

Private item As CorailItem
Private cItems As Collection

Private pn As String
Private stock As Double
Private recv As Double
Private backlog As Double
' hazards == aleas in this case ;)
Private aleas As Double
' new 21 oct 2019
Private blokcagesIP As Double


Private supplier As String
Private partName As String


Private resp As String

Private cmnt1 As String
Private cmnt2 As String


Public Property Get getSupplier() As String

    getSupplier = CStr(supplier)
End Property

Public Property Get getresp() As String
    getresp = CStr(resp)
End Property

Public Property Get getPartName() As String

    getPartName = CStr(partName)
End Property

Public Property Get getComment1() As String

    getComment1 = CStr(cmnt1)
End Property

Public Property Get getComment2() As String

    getComment2 = CStr(cmnt2)
End Property

Public Sub setComment(ktory, arg As String)
    If ktory = 1 Then cmnt1 = CStr(arg)
    If ktory = 2 Then cmnt2 = CStr(arg)
End Sub




Public Property Get getBacklog() As Double

    getBacklog = CDbl(backlog)
End Property

Public Property Get getAleas() As Double

    getAleas = CDbl(aleas)
End Property

Public Property Get getBlockagesInProgress() As Double
    getBlockagesInProgress = CDbl(blokcagesIP)
End Property


Public Property Get getPN() As String
    
    getPN = pn
End Property

Public Property Get getStock() As Double
    getStock = stock
End Property

Public Sub setStock(arg As Double)
    
    stock = CDbl(arg)
End Sub

Public Property Get getRecv() As Double
    getRecv = recv
End Property

Public Sub setRecv(arg As Double)
    
    recv = CDbl(arg)
End Sub

Public Sub setBacklog(arg As Double)
    
    backlog = CDbl(arg)
End Sub

Public Sub setAleas(arg As Double)
    aleas = CDbl(arg)
End Sub

Public Sub setBlockagesInProgress(arg As Double)
    blokcagesIP = CDbl(arg)
End Sub


Public Sub setSupplier(arg As String)
    supplier = CStr(arg)
End Sub

Public Sub setPartName(arg As String)
    partName = CStr(arg)
End Sub


Public Sub setResp(arg As String)
    resp = CStr(arg)
End Sub



' niezwykle wazna metoda dodawania dla kolekcji ktora w tej klasie jest prywatna
' nie mozemy tak po prostu dodac byle czego!
' korzystamy tylko i wylacznie z tego suba - wymusza on na wpisywaniu tylko okreslonego typu do argumentu
Public Sub addItem(i As CorailItem)
    
    cItems.Add i
End Sub


Public Property Get pItems() As Collection

    Set pItems = cItems

End Property



Private Sub Class_Initialize()
    Set cItems = New Collection
    stock = 0#
    recv = 0#
    backlog = 0#
    
    
    supplier = ""
    partName = ""
    
    resp = ""
End Sub

Private Sub Class_Terminate()
    Set cItems = Nothing
End Sub
