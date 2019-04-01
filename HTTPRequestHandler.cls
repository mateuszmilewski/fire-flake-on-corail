VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTTPRequestHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private request As WinHttpRequest
Private url As String
Private pltUrl As String

Public dom As DOMHandler


Private login As String
Private pass As String


Private Function getData(url As String, login As String, pass As String, e As E_TYPE_OF_CORAIL) As String



    'Debug.Print url
    'Debug.Print login
    'Debug.Print pass
    getData = ""
    

    If url <> "" Then
    
    
        If e = BLUE Then
    
            With request
                .Open "GET", url, False
                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                .SetAutoLogonPolicy AutoLogonPolicy_Always
                .SetCredentials login, pass, 0
                .send
                
                
                getData = CStr(.responseText)
            End With
        ElseIf e = MAESTRO Then
            
            With request
            
                '
                startUrl = "http://maestro.inetpsa.com/authentifier.do"
                .Open "GET", startUrl, False
                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                .SetAutoLogonPolicy AutoLogonPolicy_Always
                .SetCredentials login, pass, 0
                .send
                
                
                Debug.Print .statusText
            
                'getData = CStr(.responseText)
                'Debug.Print getData
                
                ' pltUrl = "http://maestro.inetpsa.com/selectCentre.do?methode=select&selectedCentre=88+"
                ' pltUrl = "http://maestro.inetpsa.com/selectCentre.do?methode=select&selectedCentre=88+"
                .Open "GET", pltUrl, False
                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                .SetAutoLogonPolicy AutoLogonPolicy_Always
                .SetCredentials login, pass, 0
                .send
                
                
                Debug.Print .statusText
            
                'getData = CStr(.responseText)
                'Debug.Print getData
                
                
                
                ' url = "http://maestro.inetpsa.com/produit.do?methode=init&selectedcodeProduit=9807026780"
                .Open "GET", url, False
                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                .SetAutoLogonPolicy AutoLogonPolicy_Always
                .SetCredentials login, pass, 0
                .send
                
                
                Debug.Print .statusText
                
                getData = CStr(.responseText)
                
            End With
        End If
    
    Else
        getData = ""
    End If
End Function


Public Sub init(plt As String, pn As String)

    url = ""
    
    
    ' CORAIL
    ' If Trim(UCase(plt)) = "Y" Then url = "http://ty.control.erp.corail.inetpsa.com/getProductSummaryRead.do?beanId=" & CStr(pn)
    
    
    ' MAESTRO
    ' url = "http://maestro.inetpsa.com/produit.do?methode=init&selectedcodeProduit=9807026780"
    
    Dim r As Range
    Set r = ThisWorkbook.Sheets(FFOC.G_SH_NM_PLT_LIST).Range("A2")
    Do

        If Trim(plt) = Trim(r) Then
            
            
            If Trim(r.Offset(0, 3).Value) = "MAESTRO" Then
                url = "http://maestro.inetpsa.com/produit.do?methode=init&selectedcodeProduit=" & CStr(pn)
                pltUrl = Trim(r.Offset(0, 2).Value)
            Else
                url = Trim(r.Offset(0, 4).Value) & CStr(FFOC.G_URL_EXT) & CStr(pn)
                pltUrl = ""
            End If
            Exit Do
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""

    ' Debug.Print url
End Sub


Public Sub braceWithDom(e As E_TYPE_OF_CORAIL)


    
    With dom
    
    
        'Set .doc.body.innerText = getData(url, login, pass)
        'Set .idoc.body.innerText = getData(url, login, pass)
        Set .doc = New HTMLDocument
        Set .idoc = New HTMLDocument
        .doc.body.innerHTML = getData(url, login, pass, e)
        .idoc.body.innerHTML = .doc.body.innerHTML
        

        If e = BLUE Then
                
            Set .stockInputElement = .doc.getElementById("stockDepart")
            Set .recvInputElement = .doc.getElementById("receivedOnSite")
            Set .backlogInputElement = .doc.getElementById("efXpLate")
            
            ' hazards
            Set .hazardsInputElement = .doc.getElementById("aleatoriesSinceFJU")
            
            
            On Error Resume Next
            Set .ipartnameElement = .doc.getElementsByClassName("descriptionProduct")(0).getElementsByTagName("a")(0)
            
            
            On Error Resume Next
            Set .isupplierElement = .doc.getElementById("tableauFluxDePiece")
            Set .isupplierElement = .isupplierElement.getElementsByTagName("table")(3).getElementsByTagName("td")(0)
            
            
                
            'Debug.Print .ipartnameElement.innerText
            'Debug.Print .isupplierElement.innerText
            
            On Error Resume Next
            Set .iapproElement = .doc.getElementById("tableauIdentificationProduct").getElementsByTagName("table")(0).getElementsByTagName("td")(1)
            Set .iapproElement = .iapproElement.getElementsByClassName("INPUT_text_Readonly_Actif INPUT_text_Readonly number")(0)
            Set .iapproElement = .iapproElement.getElementsByTagName("input")(0)
            
            'Debug.Print .iapproElement.outerHTML
            'Debug.Print .iapproElement.innerHTML
            'Debug.Print .iapproElement.Value ' OK
            
            
        ElseIf e = MAESTRO Then
            
                
        
            
            
            Set .stockInputElement = Nothing
            Set .recvInputElement = Nothing
            Set .backlogInputElement = Nothing
            
            Set .istockElement = tryGetDataFromMaestro("stock")
            Set .ibacklogElement = tryGetDataFromMaestro("backlog")
            
        ElseIf e = MANUAL Then
        
            Set .stockInputElement = Nothing
            Set .recvInputElement = Nothing
            Set .backlogInputElement = Nothing
        End If
        
        
        
    End With
    
    
End Sub


Private Sub Class_Initialize()


    'login = "ofzq0l2"
    'pass = pass
    
    'login = CStr(ThisWorkbook.Sheets("register").Range("H4").Value)
    'pass = CStr(ThisWorkbook.Sheets("register").Range("H5").Value)
    
    login = G_LOGIN
    pass = G_PASS
    


    Set dom = New DOMHandler

    Set request = New WinHttpRequest
    
    With request
        .Option(4) = 13056
        .Option(12) = True
    End With
End Sub


Private Sub Class_Terminate()
    Set dom = Nothing
    Set request = Nothing
End Sub


Private Function tryGetDataFromMaestro(what As String) As IHTMLElement

    Set tryGetDataFromMaestro = Nothing
    
    With dom
                
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
End Function
