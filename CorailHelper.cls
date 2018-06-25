VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorailHelper"
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



Private kolekcja As Collection
Private listaWejsciowa As InputListHelper

Private i As Range
Private el As Range
Private e As E_TYPE_OF_CORAIL

Private sh As Worksheet
Private layout As TheLayout





Public Sub run(listOfPartNumbers As InputListHelper)
    
    
    
    Set i = listOfPartNumbers.getInputList()
    
    Set listaWejsciowa = listOfPartNumbers

    
    
    
    For Each el In i
    
        If Not el.EntireRow.Hidden Then
    
            If (Trim(el.Value) <> "") And (Trim(el.Offset(0, 1).Value) <> "") Then
            
    
            
                ' application run independently
                ' ------------------------------------------------------------------
    
                
                
                
                Dim runner As CorailRunner
                Set runner = New CorailRunner
                runner.zdarzenie.proceduraDoPodniesieniaEventu el, kolekcja
                ' runner.zdarzenie.proceduraDoPodniesieniaPrzeniesieniaZawartosciKolekcji kolekcja
                
                'Debug.Print kolekcja.Count 'OK
                
                Set runner = Nothing
                
                
                ' ------------------------------------------------------------------
            End If
        
        End If
    Next el
End Sub



Public Sub putDataOnReportSheet()

    
    If kolekcja.Count > 0 Then

        Dim lblR As Range
        Dim r As Range
        
        
        ' layout zostaje utowrzony jednak narazie wykorzystujemy tylko i wylacznie suby
        ' ktore wypelniaja surowym textem
        ' dopiero nastepne metody/ suby beda tworzyc kolorwy layout
        Set layout = New TheLayout
        layout.setLang Int(ThisWorkbook.Sheets("register").Range("H1").Value)
        ' zapisana referencja nowego arkusza raportu ktory od teraz bedzie przechowywany w zmiennej sh typu Worksheet
        Set sh = ThisWorkbook.Sheets.Add
        
        
        Set r = sh.Range("b5")
        
        
        Set lblR = sh.Range("B4")
        Set lblR = layout.fillFirstRawLabels(lblR)
        Set lblR = layout.getDatesFromDropperAndFillLabels(lblR, kolekcja)
        
        
        Dim wiersz As Long
        wiersz = r.Row
    
    
        Dim el As DropperHandler
        For Each el In kolekcja
        
            el.run sh, CLng(wiersz)
            wiersz = wiersz + 1
        Next el
        
        
        sh.Cells(1, 1).Value = "Report; " & CStr(Now)
        
    Else
        MsgBox "Fatal error - kolekcja jakims cudem jest pusta - upewnij sie ze lista wejsciowa czesci jest prawidlowa. pozdrawiam MM."
    End If
End Sub

Public Sub makeLayout()
    
    If Not layout Is Nothing Then
    
    
    
        ' ustalmy wstepnie layout bez gridu i zoom
        layout.init sh
    
    
        ' mamy sh i proponuje zaczac od nowa z r as range
        Dim r As Range
        Set r = sh.Range("B4")
        
        With layout
            .addDayNamesAndMergeTopDateRows sh.Range("A3").End(xlToRight)
            .adjustLabelRow r
            .colourFontsOnOtherThanBalanceColumns r
        End With
        
    Else
        MsgBox "instancja kalsy TheLayout nie zostala utworzona!"
    End If
End Sub



Public Sub runAllPlants()
    
    ' before we start any logic make sure that all plants are available
    ' ---------------------------------------------------------------------------
    Dim r As Range
    Set r = defineListOfURLs()
    
    ' ---------------------------------------------------------------------------

End Sub

Private Function defineListOfURLs() As Range
    
    Dim r As Range
    Set r = ThisWorkbook.Sheets(FFOC.G_SH_NM_PLT_LIST).Range("C2")
    
    Dim ir As Range
    Set ir = r
    Do
        Set ir = ir.Offset(1, 0)
    Loop Until Trim(ir) = ""
    
    Set r = r.Parent.Range(r, ir.Offset(-1, 0))
    Set defineListOfURLs = r
    
    
    
    ' isMaestroAvail is GLOBAL!!!
    For Each ir In r
        iteracjaOtwierania CStr(ir.Address), isMaestroAvail
    Next ir
    
    
End Function

Public Sub iteracjaOtwierania(adr As String, isMaestroAvail As Boolean)

    Dim ieh As IEHandler
    Set ieh = New IEHandler
    Dim ie As InternetExplorer
    Dim ir As Range
    Set ir = ThisWorkbook.Sheets(FFOC.G_SH_NM_PLT_LIST).Range(adr)
    
    
    If isMaestroAvail And ir.Value Like "*maestro*" Then
    
        Set ie = New InternetExplorer
        ie.Visible = True
        ie.navigate CStr(ir.Value) ' , CLng(2048)
        
        ieh.pIE = ie
        ieh.publicWait ie
        Sleep 200
        
        ' fatal error - innerIE is nothing here
        ' ieh.braceWithDom MAESTRO
        
        Dim doc As HTMLDocument
        Set doc = ie.document
        
        Set ieh.dom.a = doc.getElementsByClassName("Intro")(0)
        
        ' Debug.Print ieh.dom.a.innerHTML
        ieh.dom.a.Click
        ieh.publicWait ie
        
    Else

        
        Set ie = New InternetExplorer
        ie.Visible = True
        ie.navigate CStr(ir.Value) ' , CLng(2048)
        
        With ie
        
            Do
                Sleep 100
            Loop Until corailHelperIELoaded(ie)
        End With

    End If
End Sub

Private Function corailHelperIELoaded(ie As InternetExplorer) As Boolean


    
    Dim obj As HTMLBody
    
    
    Set obj = Nothing
    
    
    
    On Error Resume Next
    Set obj = ie.document.body
    
    'Debug.Print obj.innerHTML
    
    If obj Is Nothing Then
        corailHelperIELoaded = False
    Else
    
        corailHelperIELoaded = True
    End If
    
    
End Function




Private Sub Class_Initialize()
    Set c = Nothing
    Set kolekcja = New Collection
End Sub

Private Sub Class_Terminate()
    Set c = Nothing
    Set kolekcja = Nothing
End Sub
