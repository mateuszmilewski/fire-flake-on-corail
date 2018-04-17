VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorailRunner"
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


' connected with Corail Runner

Public WithEvents zdarzenie As EventHandler
Attribute zdarzenie.VB_VarHelpID = -1

Private kolekcjaDanych As Collection


Private Sub Class_Initialize()

    Set zdarzenie = New EventHandler
    Set kolekcjaDanych = New Collection
    
End Sub

Private Sub Class_Terminate()
    Set zdarzenie = Nothing
    Set kolekcjaDanych = Nothing
End Sub



Private Sub zdarzenie_initCorail(r As Range, kolekcja As Collection)


    Set kolekcjaDanych = kolekcja


    'na bazie jednego interfejsu
    Dim c As ICorail
    ' converted data to wlasciwie dodatkowy boxing zabezpieczajacy iteracje coraila
    Dim dane As ConvertedData
    
    Dim pn As String, plt As String, cmnt1 As String, cmnt2 As String
    
    
    pn = CStr(r.Offset(0, 1).Value)
    plt = CStr(r.Value)
    
    cmnt1 = CStr(r.Offset(0, 3).Value)
    cmnt2 = CStr(r.Offset(0, 4).Value)
    
    
    Dim e As E_TYPE_OF_CORAIL
    e = defineTypeOfCorail(CStr(r.Offset(0, 2).Value))
    
    If e = BLUE Then
        Set c = New CorailBlue
    ElseIf e = ORANGE Then
        Set c = New CorailOrange
    ElseIf e = MANUAL Then
        Set c = New CorailManual
    ElseIf e = MAESTRO Then
        ' :) pls do not comment - i will treat maestro as corail interface OK?
        ' Set c = New CorailManual
        Set c = New CorailMaestro
    Else
        Set c = Nothing
    End If
    
    
    If Not c Is Nothing Then
    
        With c
            
            .generateInnerIE r, CStr(plt), CStr(pn)
            Set dane = .getData()
    
    
    
    
    
            ' add to collection
            ' -----------------------------------------------------------------------
            Dim dropper As New DropperHandler
            dropper.init plt, pn, cmnt1, cmnt2, dane
            kolekcjaDanych.Add dropper
            
            
            ' -----------------------------------------------------------------------
            
            
            .closeInnerIE
    
        End With
    End If


End Sub


Private Function defineTypeOfCorail(txt As String) As E_TYPE_OF_CORAIL
    
    If txt = FFOC.G_BLUE_TXT Then
        defineTypeOfCorail = BLUE
    ElseIf txt = FFOC.G_ORANGE_TXT Then
        defineTypeOfCorail = ORANGE
    ElseIf txt = FFOC.G_MANUAL_TXT Then
        defineTypeOfCorail = MANUAL
    ElseIf txt = FFOC.G_MAESTRO_TXT Then
        defineTypeOfCorail = MAESTRO
    Else
        defineTypeOfCorail = UNDEF
    End If
End Function

Private Sub zdarzenie_returnCollection(kolekcja As Collection)
    ' Set kolekcjaDanych = kolekcja
End Sub
