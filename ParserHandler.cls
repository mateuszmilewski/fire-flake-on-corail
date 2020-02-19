VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParserHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'
'Copyright (c) 2019 FORREST
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


Private dateCatcher As String

Private besoinsPCCatcher As String
Private besoinsBCCatcher As String

Private commandCatcher1 As String
Private commandCatcher2 As String
Private commandCatcher3 As String
Private commandCatcher4 As String
Private cc1 As String
Private cc2 As String
Private cc3 As String
Private cc4 As String
Private expCatcher1 As String
Private expCatcher2 As String
Private expCatcher3 As String
Private expCatcher4 As String
Private ec1 As String
Private ec2 As String
Private ec3 As String
Private ec4 As String


Public Property Get pDateCatcher() As String
    pDateCatcher = dateCatcher
End Property

Public Property Get pBesoinsPCCattcher() As String
    pBesoinsPCCattcher = besoinsPCCatcher
End Property

Public Property Get pBesoinsBCCatcher() As String
    pBesoinsBCCatcher = besoinsBCCatcher
End Property

Public Property Get pCmdCatcher1() As String
    pCmdCatcher1 = commandCatcher1
End Property

Public Property Get pExpCatcher1() As String
    pExpCatcher1 = expCatcher1
End Property

Public Property Get pCmdCatcher2() As String
    pCmdCatcher2 = commandCatcher2
End Property

Public Property Get pExpCatcher2() As String
    pExpCatcher2 = expCatcher2
End Property

Public Property Get pCmdCatcher3() As String
    pCmdCatcher3 = commandCatcher3
End Property

Public Property Get pExpCatcher3() As String
    pExpCatcher3 = expCatcher3
End Property

Public Property Get pCmdCatcher4() As String
    pCmdCatcher4 = commandCatcher4
End Property

Public Property Get pExpCatcher4() As String
    pExpCatcher4 = expCatcher4
End Property




Public Function htmlTableToRawMatrix(tb As HTMLTable, dom As DOMHandler, e As E_TYPE_OF_CORAIL) As ConvertedData

    Dim htmlString As String
    Dim tmp As ConvertedData
    
    
    Dim i As CorailItem
    Dim ii As CorailIteration
    
    Dim innerTb As HTMLTable

    Dim r As HTMLTableRow
    Dim c As HTMLTableCell
    
    Dim s As String
    
    ' new 21 oct 2019
    ' Blockages in progress - decrement from stock
    Dim tmpBIP As String
    
    Dim tmpStock As String
    Dim tmpRecv As String
    Dim tmpBacklog As String
    Dim tmpHazards As String
    
    Dim tmpResp2 As String
    
    
    Dim tmpSupplier As String
    Dim tmpPartName As String
    
    Set ii = New CorailIteration
        
    If e = ORANGE Then
    
    
        ' table with rqm, order, ship, balance
        ' ----------------------------------------------------------------
        
        Set tmp = New ConvertedData
        
        If Not tb Is Nothing Then
            htmlString = tb.innerHTML
            If checkIfHTMLTable(htmlString) Then
        
        
                
                Set innerTb = tb.getElementsByTagName("table")(0)
    
                avoidHeading = True
                For Each r In innerTb.Rows
                
                    Set i = Nothing
                    Set i = New CorailItem
                
                    For Each c In r.Cells
                        
                        
                        If avoidHeading Then
                            Exit For
                        Else
                            ' Debug.Print c.innerHTML
                            
                            
                            If UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pDateCatcher, Chr(34), "")) & "*" Then
                                
                                
                                s = CStr(c.innerHTML)
                                arr = Split(s, ">")
                                
                                s = Left(arr(1), 10)
                                
                                i.parseStringToDate s
                                
                                
                            ElseIf UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pBesoinsPCCattcher, Chr(34), "")) & "*" Then
                                
                                s = CStr(c.innerHTML)
                                arr = Split(s, ">")
                                s = arr(1)
                                s = Replace(UCase(s), "</DIV", "")
                                s = Replace(s, ".0", "")
                                s = Replace(s, ",0", "")
                                
                                i.besoinsPC = CLng(s)
                                
                            ElseIf UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pBesoinsBCCatcher, Chr(34), "")) & "*" Then
                            
                                s = CStr(c.innerHTML)
                                arr = Split(s, ">")
                                s = arr(1)
                                s = Replace(UCase(s), "</DIV", "")
                                s = Replace(s, ".0", "")
                                s = Replace(s, ",0", "")
                                
                                i.besoinsBC = CLng(s)
                                
                            ElseIf (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher1) & "*") Or _
                                (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher2) & "*") Or _
                                (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher3) & "*") Or _
                                UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher4) & "*" Then
                            
                                s = CStr(c.innerHTML)
                                arr = Split(s, ">")
                                s = arr(1)
                                s = Replace(UCase(s), "</DIV", "")
                                s = Replace(s, ".0", "")
                                s = Replace(s, ",0", "")
                                
                                i.order = CLng(s)
                                
                            ElseIf (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher1) & "*") Or _
                                (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher2) & "*") Or _
                                (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher3) & "*") Or _
                                (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher4) & "*") Then
                            
                                s = CStr(c.innerHTML)
                                arr = Split(s, ">")
                                s = arr(1)
                                s = Replace(UCase(s), "</DIV", "")
                                s = Replace(s, ".0", "")
                                s = Replace(s, ",0", "")
                                
                                i.ship = CLng(s)
                                
                            End If
                            
                        End If
                        
                    Next c
                    
                    If Not avoidHeading Then
                    
                        ii.addItem i
                    End If
                    
                    avoidHeading = False
                Next r
        
            End If
        End If
        
        tmpStock = dom.stockInputElement.Value
        tmpStock = Replace(tmpStock, ".0", "")
        tmpStock = Replace(tmpStock, ",0", "")
        
        If Not IsNumeric(tmpStock) Then tmpStock = "0"
        
        tmpRecv = dom.recvInputElement.Value
        tmpRecv = Replace(tmpRecv, ".0", "")
        tmpRecv = Replace(tmpRecv, ",0", "")
        
        If Not IsNumeric(tmpRecv) Then tmpRecv = "0"
        
        tmpBacklog = dom.backlogInputElement.Value
        tmpBacklog = Replace(tmpBacklog, ".0", "")
        tmpBacklog = Replace(tmpBacklog, ",0", "")
        
        If Not IsNumeric(tmpBacklog) Then tmpBacklog = "0"
        
        Set tmp.iteration = ii
        tmp.iteration.setStock CLng(tmpStock)
        tmp.iteration.setRecv CLng(tmpRecv)
        tmp.iteration.setBacklog CLng(tmpBacklog)
        
        tmp.iteration.setComment 1, CStr(ieh.txtIsAvail)
        
        Set htmlTableToRawMatrix = tmp
        
    
    ElseIf e = BLUE Then
    
    
        ' table with rqm, order, ship, balance
        ' ----------------------------------------------------------------
        
        Set tmp = New ConvertedData
        
        If Not tb Is Nothing Then
            htmlString = tb.innerHTML
            If checkIfHTMLTable(htmlString) Then
            
            
                Set innerTb = Nothing
                On Error Resume Next
                Set innerTb = tb.getElementsByTagName("table")(0)
                
                If Not innerTb Is Nothing Then
                
                    avoidHeading = True
                    For Each r In innerTb.Rows
                    
                        Set i = Nothing
                        Set i = New CorailItem
                    
                        For Each c In r.Cells
                            
                            
                            If avoidHeading Then
                                Exit For
                            Else
                                Debug.Print c.innerHTML
                                
                                
                                If UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & Replace(UCase(Me.pDateCatcher), Chr(34), "") & "*" Then
                                    
                                    
                                    s = CStr(c.innerHTML)
                                    arr = Split(s, ">")
                                    
                                    s = Left(arr(1), 10)
                                    
                                    i.parseStringToDate s
                                    
                                    
                                ElseIf UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pBesoinsPCCattcher, Chr(34), "")) & "*" Then
                                    
                                    s = CStr(c.innerHTML)
                                    arr = Split(s, ">")
                                    s = arr(1)
                                    s = Replace(UCase(s), "</DIV", "")
                                    ' s = replaceAllDecimals(s)
                                    
                                    s = adjustDecimalSeparator(s)
                                    i.besoinsPC = CDbl(s)
                                    
                                ElseIf UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pBesoinsBCCatcher, Chr(34), "")) & "*" Then
                                
                                    s = CStr(c.innerHTML)
                                    arr = Split(s, ">")
                                    s = arr(1)
                                    s = Replace(UCase(s), "</DIV", "")
'                                    s = Replace(s, ".0", "")
'                                    s = Replace(s, ",0", "")
                                    
                                    ' s = replaceAllDecimals(s)
                                    
                                    ' Application.DecimalSeparator
                                    
                                    s = adjustDecimalSeparator(s)
                                    i.besoinsBC = CDbl(s)
                                    
                                    
                                    'Debug.Print i.besoinsBC
                                    
                                ElseIf (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher1) & "*") Or _
                                    (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher2) & "*") Or _
                                    (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher3) & "*") Or _
                                    UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher4) & "*" Then
                                
                                    s = CStr(c.innerHTML)
                                    arr = Split(s, ">")
                                    s = arr(1)
                                    s = Replace(UCase(s), "</DIV", "")
'                                    s = Replace(s, ".0", "")
'                                    s = Replace(s, ",0", "")

                                    ' s = replaceAllDecimals(s)
                                    
                                    s = adjustDecimalSeparator(s)
                                    i.order = CDbl(s)
                                    'Debug.Print i.order
                                    
                                ElseIf (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher1) & "*") Or _
                                    (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher2) & "*") Or _
                                    (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher3) & "*") Or _
                                    (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher4) & "*") Then
                                
                                    s = CStr(c.innerHTML)
                                    arr = Split(s, ">")
                                    s = arr(1)
                                    s = Replace(UCase(s), "</DIV", "")
'                                    s = Replace(s, ".0", "")
'                                    s = Replace(s, ",0", "")

                                    ' s = replaceAllDecimals(s)
                                    
                                    s = adjustDecimalSeparator(s)
                                    i.ship = CDbl(s)
                                    
                                    'Debug.Print i.ship
                                    
                                End If
                                
                            End If
                            
                        Next c
                        
                        If Not avoidHeading Then
                        
                            ii.addItem i
                        End If
                        
                        avoidHeading = False
                    Next r
                    
                    
                End If
            End If
        End If
        
        If Not dom.stockInputElement Is Nothing Then
            tmpStock = dom.stockInputElement.Value
            tmpStock = adjustDecimalSeparator(tmpStock)
        End If
        
        If Not IsNumeric(tmpStock) Then tmpStock = "0"
        
        If Not dom.recvInputElement Is Nothing Then
            tmpRecv = dom.recvInputElement.Value
            tmpRecv = adjustDecimalSeparator(tmpRecv)
        End If
        
        If Not IsNumeric(tmpRecv) Then tmpRecv = "0"
        
        If Not dom.backlogInputElement Is Nothing Then
            tmpBacklog = dom.backlogInputElement.Value
            tmpBacklog = adjustDecimalSeparator(tmpBacklog)
        End If
        
        If Not dom.hazardsInputElement Is Nothing Then
            tmpHazards = dom.hazardsInputElement.Value
            tmpHazards = adjustDecimalSeparator(tmpHazards)
        End If
        
        If Not dom.iBlockagesInProgress Is Nothing Then
        
            tmpBIP = dom.iBlockagesInProgress.Value
            tmpBIP = adjustDecimalSeparator(tmpBIP)
            
        End If
        
        
        If Not dom.isupplierElement Is Nothing Then
        
            tmpSupplier = Trim(Right(dom.isupplierElement.innerText, Len(dom.isupplierElement.innerText) - 5))
        End If
        
        
        If Not dom.ipartnameElement Is Nothing Then
        
            tmpPartName = dom.ipartnameElement.innerText
            tmpPartName = Trim(Split(tmpPartName, "-")(1))
        End If
        
        
        If Not dom.iapproElement Is Nothing Then
            tmpResp2 = dom.iapproElement.Value
        End If
        
        
        If Not IsNumeric(tmpBacklog) Then tmpBacklog = "0"
        If Not IsNumeric(tmpBIP) Then tmpBIP = "0"
        If Not IsNumeric(tmpHazards) Then tmpHazards = "0"
        
        Set tmp.iteration = ii
        'tmp.iteration.setStock CDbl(Replace(tmpStock, ".", Application.DecimalSeparator))
        'tmp.iteration.setStock CDbl(Replace(tmpStock, ",", Application.DecimalSeparator))
        tmp.iteration.setStock CDbl(tmpStock)
        
        'tmp.iteration.setRecv CDbl(Replace(tmpRecv, ".", Application.DecimalSeparator))
        'tmp.iteration.setRecv CDbl(Replace(tmpRecv, ",", Application.DecimalSeparator))
        tmp.iteration.setRecv CDbl(tmpRecv)
        
        'tmp.iteration.setBacklog CDbl(Replace(tmpBacklog, ".", Application.DecimalSeparator))
        'tmp.iteration.setBacklog CDbl(Replace(tmpBacklog, ",", Application.DecimalSeparator))
        tmp.iteration.setBacklog CDbl(tmpBacklog)
        
        tmp.iteration.setPartName CStr(tmpPartName)
        tmp.iteration.setSupplier CStr(tmpSupplier)
        tmp.iteration.setResp CStr(tmpResp2)
        ' aleas == hazards
        tmp.iteration.setAleas CStr(tmpHazards)
        
        ' new element for Blockages in progress
        'tmpBIP = CDbl(Replace(tmpBIP, ".", Application.DecimalSeparator))
        'tmpBIP = CDbl(Replace(tmpBIP, ",", Application.DecimalSeparator))
        tmp.iteration.setBlockagesInProgress CDbl(tmpBIP)
        
        
        
        ' tmp.iteration.setComment 1, CStr(ieh.txtIsAvail)
        
        
        Set htmlTableToRawMatrix = tmp
            
    ElseIf e = MAESTRO Then
    
    
        ' table with rqm, order, ship, balance
        ' ----------------------------------------------------------------
        
        Set tmp = New ConvertedData
        
        If Not tb Is Nothing Then
        
            Set innerTb = tb
            htmlString = tb.innerHTML
            If checkIfHTMLTable(htmlString) Then
        
        
                For Each r In innerTb.Rows
                
                    Set i = Nothing
                    Set i = New CorailItem
                    
                    cCounter = 0
                
                    For Each c In r.Cells
                        
                        ' Debug.Print c.innerHTML
                        
                        ' 0: date
                        ' 1: rqm
                        ' 2: order
                        ' 3: it
                        
                        ' always
                        s = CStr(c.innerHTML)
                        
                        If cCounter = 0 Then
                            i.parseStringToDate Trim(s)
                        ElseIf cCounter = 1 Then
                        
                            arr = Split(s, "<")
                            s = Trim(arr(LBound(arr)))
                            's = Replace(s, ".", Application.DecimalSeparator)
                            's = Replace(s, ",", Application.DecimalSeparator)
                            s = adjustDecimalSeparator(s)
                            
                            i.besoinsBC = 0
                            i.besoinsPC = 0
                            If IsNumeric(s) Then i.besoinsBC = CDbl(s)
                            
                        ElseIf cCounter = 2 Then
                        
                            arr = Split(s, "<")
                            s = Trim(arr(LBound(arr)))
                            's = Replace(s, ".", Application.DecimalSeparator)
                            's = Replace(s, ",", Application.DecimalSeparator)
                            s = adjustDecimalSeparator(s)
                            i.order = 0
                            If IsNumeric(s) Then i.order = CLng(s)
                            
                        ElseIf cCounter = 3 Then
                            If s Like "<FONT*" Then
                                arr = Split(s, ">")
                                arr = Split(arr(LBound(arr) + 1), "</")
                            Else
                                arr = Split(s, "<")
                                arr = Split(arr(LBound(arr)), "</")
                            End If
                            s = Trim(arr(LBound(arr)))
                            's = Replace(s, ".", Application.DecimalSeparator)
                            's = Replace(s, ",", Application.DecimalSeparator)
                            s = adjustDecimalSeparator(s)
                            
                            i.ship = 0
                            If IsNumeric(s) Then i.ship = CLng(s)
                        End If
                        
                        cCounter = cCounter + 1
                    Next c
                    
                    ii.addItem i
                Next r
            End If
        End If
        
        If Not dom.istockElement Is Nothing Then
            tmpStock = dom.istockElement.innerHTML

            tmpStock = Trim(Replace(tmpStock, G_NBSP, ""))
            
            If tmpStock Like "*IMG*" Then
                tmpArr = Split(tmpStock, " ")
                tmpStock = Trim(tmpArr(0))
                'tmpStock = Replace(tmpStock, ".", Application.DecimalSeparator)
                'tmpStock = Replace(tmpStock, ",", Application.DecimalSeparator)
                tmpStock = adjustDecimalSeparator(tmpStock)
            End If
        End If
        
        If Not IsNumeric(tmpStock) Then tmpStock = "0"
        
        If Not dom.ibacklogElement Is Nothing Then
            tmpBacklog = dom.ibacklogElement.innerHTML
            's = Replace(s, ".", Application.DecimalSeparator)
            's = Replace(s, ",", Application.DecimalSeparator)
            tmpBacklog = Trim(Replace(tmpBacklog, G_NBSP, ""))
            
            
            If tmpBacklog Like "*FONT*" Then
                
                tmpArr = Split(tmpBacklog, ">")
                tmpBacklog = tmpArr(1)
                tmpBacklog = Replace(tmpBacklog, "</FONT", "")
                'tmpBacklog = Replace(tmpBacklog, ",", Application.DecimalSeparator)
                'tmpBacklog = Replace(tmpBacklog, ",", Application.DecimalSeparator)
                tmpBacklog = adjustDecimalSeparator(tmpBacklog)
            End If
            
            
        End If
        
        If Not IsNumeric(tmpBacklog) Then tmpBacklog = "0"
        
        
        Set tmp.iteration = ii
        tmp.iteration.setStock CLng(tmpStock)
        tmp.iteration.setRecv CLng(0)
        tmp.iteration.setBacklog CLng(tmpBacklog)
        
        ' tmp.iteration.setComment 1, CStr(ieh.txtIsAvail)
        
        
        
        Set htmlTableToRawMatrix = tmp
        
    Else
        MsgBox "brak zdefiniowanego systemu!"
        End
    End If
        
    
End Function


Private Function adjustDecimalSeparator(s) As String



    Dim regHalf As Range
    Set regHalf = ThisWorkbook.Sheets(FFOC.G_SH_NM_REG).Range("Q17")
    
    Dim myExcelDecimalPoint As String
    myExcelDecimalPoint = CStr(Mid(regHalf.Value, 2, 1))
    
    If myExcelDecimalPoint = Application.DecimalSeparator Then
        ' nothing to do!
    Else
        ' there is sth strage
    End If
    
    ' in corail always dots


    adjustDecimalSeparator = ""
    
    If Trim(s) <> "" Then

        If myExcelDecimalPoint = "," Then
            ' Debug.Print "Application.DecimalSeparator ,"
            If s Like "*,*" Then
                ' no operation req
            ElseIf s Like "*.*" Then
                s = Replace(s, ".", myExcelDecimalPoint)
            Else
                ' no operation req - int num
            End If
            
        ElseIf myExcelDecimalPoint = "." Then
            ' Debug.Print "Application.DecimalSeparator ."
            If s Like "*.*" Then
                ' no operation req
            ElseIf s Like "*,*" Then
                s = Replace(s, ",", myExcelDecimalPoint)
            Else
                ' no operation req - int num
            End If
        Else
            MsgBox "wrong decimal separator - please change it in control panel!"
            End
        End If
        
        
        adjustDecimalSeparator = CStr(s)
    Else
        adjustDecimalSeparator = ""
    End If


End Function

' OBSOLETE
Private Function replaceAllDecimals(s) As String
    
    s = Replace(s, ".0", "")
    s = Replace(s, ",0", "")
    
    s = Replace(s, ".1", "")
    s = Replace(s, ",1", "")
    
    s = Replace(s, ".2", "")
    s = Replace(s, ",2", "")
    
    s = Replace(s, ".3", "")
    s = Replace(s, ",3", "")
    
    s = Replace(s, ".4", "")
    s = Replace(s, ",4", "")
    
    s = Replace(s, ".5", "")
    s = Replace(s, ",5", "")
    
    s = Replace(s, ".6", "")
    s = Replace(s, ",6", "")
    
    s = Replace(s, ".7", "")
    s = Replace(s, ",7", "")
    
    s = Replace(s, ".8", "")
    s = Replace(s, ",8", "")
    
    s = Replace(s, ".9", "")
    s = Replace(s, ",9", "")
    
    replaceAllDecimals = CStr(s)
    
End Function



Private Function checkIfHTMLTable(htmlString As String) As Boolean


    Dim tmp As String
    tmp = CStr(htmlString)
    
    If UCase(tmp) Like UCase("*tr*td*td*tr*") Then
    
        checkIfHTMLTable = True
    Else
    
        checkIfHTMLTable = False
    End If
End Function

Private Sub Class_Initialize()
    dateCatcher = "<div class=" & Chr(34) & "date" & Chr(34) & ">"

    besoinsPCCatcher = "<div class=" & Chr(34) & "number fbpcDetailByDay" & Chr(34) & ""
    besoinsBCCatcher = "<div class=" & Chr(34) & "number" & Chr(34) & ">"

    commandCatcher1 = "<div class=" & Chr(34) & "ecwButtonTexteOverOrdered number" & Chr(34) & ""
    commandCatcher2 = "<div class=" & Chr(34) & "ecwButtonTexteOverOrdered number ecwButtonTexteOver" & Chr(34) & ""
    commandCatcher3 = "<div class=" & Chr(34) & "number ecwButtonTexteOverSauOrdered" & Chr(34) & ""
    commandCatcher4 = "<div class=" & Chr(34) & "number ecwButtonTexteOverSauOrdered ecwButtonTexteOver" & Chr(34) & ""
    
    cc1 = commandCatcher1
    cc2 = commandCatcher2
    cc3 = commandCatcher3
    cc4 = commandCatcher4
    
    
    expCatcher1 = "<div class=" & Chr(34) & "ecwButtonTexteOverShipped number" & Chr(34) & ""
    expCatcher2 = "<div class=" & Chr(34) & "ecwButtonTexteOverShipped number ecwButtonTexteOver" & Chr(34) & ""
    expCatcher3 = "<div class=" & Chr(34) & "number ecwButtonTexteOverSauShipped" & Chr(34) & ""
    expCatcher4 = "<div class=" & Chr(34) & "number ecwButtonTexteOverSauShipped ecwButtonTexteOver" & Chr(34) & ""
    
    ec1 = expCatcher1
    ec2 = expCatcher2
    ec3 = expCatcher3
    ec4 = expCatcher4
End Sub
