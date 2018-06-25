VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEHandler"
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



Public dom As DOMHandler
Private ie As InternetExplorer
Private url As String
Private identyfikator As Integer
Private blocked As Boolean


Private innerIE As InternetExplorer


Private powloka As ShellHandler

Public txtIsAvail As String



Private firstTimeOutAndStillNoDataAvailable As Boolean

Public Function getFirstTimeOut() As Boolean
    getFirstTimeOut = CBool(firstTimeOutAndStillNoDataAvailable)
End Function


Public Property Get pIE() As InternetExplorer
    Set pIE = ie
End Property

Public Property Let pIE(arg As InternetExplorer)
    Set ie = arg
End Property


Public Property Get pBlocked() As Boolean
    pBlocked = blocked
End Property

Public Property Let pBlocked(arg As Boolean)
    
    blocked = arg
End Property

Private Sub initBlockedParamOnFalse()
    blocked = False
End Sub



Public Property Get pUrl() As String
    
    pUrl = CStr(url)
End Property

Public Property Let pUrl(arg As String)
    
    url = CStr(arg)
End Property



Public Function checkIfTisAvailable() As Boolean
    
    checkIfTisAvailable = True
End Function


Public Sub closeInnerIE()
    
    On Error Resume Next
    innerIE.Quit
    Set innerIE = Nothing
End Sub


Public Sub getNewInnerIE(r As Range, plt As String, pn As String, isCorail As Boolean)

    If isCorail Then
        Set innerIE = New InternetExplorer
        innerIE.Visible = True
        url = generateURL(r, CStr(plt), CStr(pn))
        innerIE.navigate CStr(url)
        Me.innerWait
    Else
        ' most probably maestro :)
        Set innerIE = New InternetExplorer
        innerIE.Visible = True
        If Not sprawdzCzyPlantJestAktywnyISzybkiRefresh(plt) Then aktywujPlant plt, innerIE
        url = generateMaestroURL(r, CStr(plt), CStr(pn))
        Me.innerWait
        innerIE.navigate CStr(url)
        Me.innerWait
    End If
    
End Sub

Private Function sprawdzCzyPlantJestAktywnyISzybkiRefresh(plt As String) As Boolean
    sprawdzCzyPlantJestAktywnyISzybkiRefresh = False
End Function

Private Sub aktywujPlant(plt As String, ie As InternetExplorer)

    Dim maestroPltUrl As String
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(FFOC.G_SH_NM_PLT_LIST)
    
    
    Dim r As Range
    Set r = sh.Range("A2")
    Do
        If Trim(r) = Trim(plt) Then
            maestroPltUrl = r.Offset(0, 2)
            Exit Do
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""

    ie.navigate maestroPltUrl ' , CLng(2048)
    Me.innerWait
End Sub

Private Function generateURL(r As Range, plt As String, pn As String) As String

    ' generateURL = "" & r.Offset(0, 2).Value & "/" & FFOC.G_URL_EXT &
    'generateURL = "http://ta.control.erp.corail.inetpsa.com/getProductSummaryRead.do?beanId=96661053ZD"
    
    generateURL = ""
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(FFOC.G_SH_NM_PLT_LIST)
    
    
    
    ' CORAIL LOOP
    For i = FFOC.G_CORAIL_FIRST_PLT To FFOC.G_CORAIL_LAST_PLT
        
        ' --------------------------------------------------------------------------------------------------------
        
        If Trim(CStr(sh.Range("A" & CStr(i)).Value)) <> "" Then
        
            If plt = CStr(sh.Range("A" & CStr(i)).Value) Then
            
                generateURL = CStr(sh.Range("C" & CStr(i)).Value) & "getProductSummaryRead.do?beanId=" & CStr(pn) & "#"
                Exit For
            End If
        
        Else
            Exit For
        End If
        ' --------------------------------------------------------------------------------------------------------
    Next i
    
    
End Function

Private Function generateMaestroURL(r As Range, plt As String, pn As String) As String

    generateMaestroURL = ""
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(FFOC.G_SH_NM_PLT_LIST)
    generateMaestroURL = CStr(sh.Range("C2")) & "/" & FFOC.G_MAESTRO_URL_EXT & CStr(pn)
    
End Function


Public Sub braceWithDom(e As E_TYPE_OF_CORAIL)

    
    
    
    
    If Not innerIE Is Nothing Then
    
        Me.innerWait
        
        With dom
        
        
            Set .doc = innerIE.document
            Set .idoc = innerIE.document
            
            
            innerWait
            Sleep 100
        
            If e = ORANGE Then
            
                If isDataAvailable(e) Then
            
                    Set .iElementCollectionOfDates = .doc.getElementsByClassName("date")
                    Set .iElementCollectionOfRqms1 = .doc.getElementsByClassName("number fbpcDetailByDay")
                    Set .iElementCollectionOfRqms2 = .doc.getElementsByClassName("number")
                    Set .iElementCollectionOfOrders = .doc.getElementsByClassName("ecwButtonTexteOverOrdered number ecwButtonTexteOver")
                    Set .iElementCollectionOfShipments = .doc.getElementsByClassName("ecwButtonTexteOverShipped number ecwButtonTexteOver")
                    
                    
                    ' this double logic is created becuase in old projects there was undefined issue with assigning problem between objects
                    ' im not sure why it is working in this way, but from time to time i need to be sure that i assign object properly
                    ' so im doubling process of taking elements by tag name of input and give it some extra sleep time...
                    Me.innerWait
                    DoEvents
                    Sleep 200
                    Set .iElementCollection = .doc.getElementsByTagName("input")
                    
                    Me.innerWait
                    DoEvents
                    Sleep 200
                    Set .iElementCollection = .doc.getElementsByTagName("input")
                    
                    podlicz = 0
                    For Each el In .iElementCollection
                    
                    
                        Set .iElement = el
                        If CBool(.iElement.className Like "*INPUT_text_Readonly number*") Or CBool(.iElement.className Like "*ecwControlOver INPUT_text_Readonly number*") Then
                        
                        
    
                            
                            If Not Left(.iElement.parentElement.innerHTML, 4) = "<!--" Then
                            
                            
                            
                                ' Coverage scope
                                ' Stock Depart
                                ' To be consumed
                                ' Recv
                                podlicz = podlicz + 1
                            
                                'Debug.Print .iElement.parentElement.innerHTML
                            
                                If UCase(.iElement.parentElement.innerHTML) Like UCase("*input*span*onclick*openStockPopup*") Then
                            
                                    Set .stockInputElement = .iElement
                                    ' Debug.Print .stockInputElement.Value ' OK
                                    'Exit For
                                End If
                                
                                
                                
                                ' lecimy dalej z szukaniem elementu odpowiadajacego z recv
                                If podlicz = 3 Then
                                    
                                    Set .recvInputElement = .iElement
                                    ' Exit For
                                End If
                                
                                
                                
                                '  Retard en EF/XP - 14th!
                                Debug.Print podlicz & " " & .iElement.innerHTML & " " & .iElement.Value
                                If podlicz = 14 Then
                                
                                    Set .backlogInputElement = .iElement
                                    Debug.Print .backlogInputElement.Value
                                    ' Exit For
                                End If
                                
                                
                            End If
                            
                            
                            
                            
                        End If
                    Next el
                    
                End If

            ElseIf e = BLUE Then
            
            
                Sleep 100
                innerWait
            
                If isDataAvailable(e) Then

                    
                    Set .stockInputElement = .doc.getElementById("stockDepart")
                    Set .recvInputElement = .doc.getElementById("receivedOnSite")
                    Set .backlogInputElement = .doc.getElementById("efXpLate")
                    
                    innerWait
                    Sleep 100
                    Set .stockInputElement = .doc.getElementById("stockDepart")
                    Set .recvInputElement = .doc.getElementById("receivedOnSite")
                    Set .backlogInputElement = .doc.getElementById("efXpLate")
                
                End If
                
            ElseIf e = MAESTRO Then
            
                Sleep 100
                innerWait
                
                If isDataAvailable(e) Then
                
                    
            
                    Set .istockElement = tryGetDataFromMaestro("stock")
                    Set .ibacklogElement = tryGetDataFromMaestro("backlog")
                    
                    ' Debug.Print .istockElement.innerHTML
                    
                    Sleep 100
                    innerWait
                    
                    Set .istockElement = tryGetDataFromMaestro("stock")
                    Set .ibacklogElement = tryGetDataFromMaestro("backlog")
                    
                    ' Debug.Print .istockElement.innerHTML
                End If
                
            ElseIf e = MANUAL Then
            
                Set stockInputElement = Nothing
                Set recvInputElement = Nothing
                Set backlogInputElement = Nothing
                
                
                
            End If
            
            
            
        End With
    Else
        MsgBox "fatal error with innerIE!"
        End
    End If
    
    
End Sub


Private Function isDataAvailable(e As E_TYPE_OF_CORAIL) As Boolean

    isDataAvailable = False
    
    
    txtIsAvail = ""
    
    If e = MAESTRO Then
        
        Set dom.iElement = dom.doc.getElementsByClassName("titrePage")(0)
        
        If Not dom.iElement Is Nothing Then
            
            If dom.iElement.innerHTML = "Erreur" Then
                isDataAvailable = False
                txtIsAvail = "Error. Data is not available in Maestro system."
            ElseIf dom.iElement.innerHTML = "Detail Produit " Then
                isDataAvailable = True
            Else
                isDataAvailable = False
            End If
        Else
            isDataAvailable = True
        End If
        
    ElseIf e = ORANGE Then
        
        isDataAvailable = False
        txtIsAvail = "Error. Data is not available in Orange Corail system."
    ElseIf e = BLUE Then
    
        Set dom.iElement = dom.doc.getElementById("EcwMessageBar")
        
        If Not dom.iElement Is Nothing Then
        
            If dom.iElement.innerHTML Like "*Pas de*produit possible pour cet*" Then
            
                
                Sleep 500
                Application.SendKeys "{ENTER}"
                Sleep 500
                Application.SendKeys "{ENTER}"
                Sleep 500
                
                txtIsAvail = "Error. Data is not available in Corail system."
                isDataAvailable = False
            Else
                isDataAvailable = True
            End If
        Else
            isDataAvailable = True
        End If
    
    End If
End Function

Private Function tryGetDataFromMaestro(what As String) As IHTMLElement

    Set tryGetDataFromMaestro = Nothing
    
    
    If Not innerIE Is Nothing Then
    
        Me.innerWait
        
        With dom
        
        
            Set .doc = innerIE.document
            Set .idoc = innerIE.document
    
            Me.innerWait
            DoEvents
            Sleep 200
            
            If what = "stock" Then
                Set .iElementCollection = .doc.getElementsByTagName("font")
            End If
            If what = "backlog" Then
                Set .iElementCollection = .doc.getElementsByClassName("Donnee")
            End If
            
            podlicz = 0
            For Each el In .iElementCollection
                
                Set iElement = el
                If Not iElement Is Nothing Then
                    ' Debug.Print iElement.innerHTML
                    
                    If podlicz = 1 Then
                    
                        If what = "stock" Then
                            Set tryGetDataFromMaestro = iElement
                            Exit Function
                        End If
                    End If
                    
                    
                    If what = "backlog" Then
                    
                        
                        ' Debug.Print iElement.innerHTML
                        
                        
                        If podlicz = 22 Then
                            Set tryGetDataFromMaestro = iElement
                            Exit Function
                        End If
                    End If
                
                End If
                
                podlicz = podlicz + 1
                
                
            Next el
            
        End With
    End If

End Function



Public Sub setIEWithLoggedCorail()

    ' it will return nothing if cought nothing
    ' ===========================================
    Set ie = powloka.catchLoggedCorail()
    ' ===========================================
End Sub

Public Sub closeAllIEs()

    powloka.zamknijWszystkiePrzegladarki
End Sub




Public Function getUrl()
    
    getUrl = CStr(ie.LocationURL)
End Function



Public Sub getURLByPltId(ajdi As String)
    
    Me.pUrl = getUrlFromShByAjdi(ajdi)
End Sub

Private Function getUrlFromShByAjdi(ajdi)
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(FFOC.G_SH_NM_IN)
    Dim r As Range
    Set r = sh.Range("A2")
    
    Do
    
        
    
        Set r = r.Offset(1, 0)
    Loop While Trim(r) = ""
End Function



Private Sub Class_Initialize()
    
    Set dom = New DOMHandler
    'Set ie = New InternetExplorer
    Set ie = Nothing
    Set innerIE = Nothing
    
    On Error Resume Next
    Set powloka = New ShellHandler
    
    initBlockedParamOnFalse
    
End Sub




Private Sub Class_Terminate()
    
    Set dom = Nothing
    Set ie = Nothing
    Set innerIE = Nothing
    
    Set powloka = Nothing
    
End Sub


Public Sub publicWait(arg As InternetExplorer)
    
    ' this logic is not working with all configs with all pcs
    'With innerIE
    '
    '    Do
    '        DoEvents
    '    Loop While .Busy Or .readyState = READYSTATE_LOADING
    '
    '    Do
    '        DoEvents
    '    Loop Until .readyState = READYSTATE_COMPLETE
    'End With
    
    Dim x As Integer
    x = 0

    Do
        
        Sleep 100
        x = x + 1
        
        If x > 500 Then
            MsgBox "Timeout! Try again!"
            End
            
        End If
    Loop Until loaded(arg)
End Sub

Public Sub innerWait()

    
    firstTimeOutAndStillNoDataAvailable = False
    
    Dim x As Integer
    x = 0

    Do
        
        Sleep 100
        x = x + 1
        
        
        If x > 200 Then
        
            firstTimeOutAndStillNoDataAvailable = True
            txtIsAvail = txtIsAvail & "Unexcepted browser timeout!"
            Exit Do
        
        ElseIf x > 500 Then
            MsgBox "Timeout! Try again!"
            End
            
        End If
    Loop Until loaded(innerIE)
End Sub


Private Function loaded(ByRef ie As InternetExplorer) As Boolean


    Dim obj As HTMLBody
    
    
    Set obj = Nothing
    
    
    
    On Error Resume Next
    Set obj = ie.document.body
    
    'Debug.Print obj.innerHTML
    
    If obj Is Nothing Then
        loaded = False
    Else
    
        loaded = True
    End If

    
    'Set dom.doc = ie.document
    '
    'Dim obj As IHTMLElement
    '
    'Set obj = dom.doc.getElementsByClassName("ecwTitleMessageIhm")(0)
    '
    '
    'If obj Is Nothing Then
    '    loaded = False
    'Else
    '
    '    If UCase(obj.innerHTML) Like UCase("*Product synthesis*") Then
    '        loaded = True
    '    Else
    '        loaded = False
    '    End If
    'End If
    
    
End Function
