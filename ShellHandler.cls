VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellHandler"
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


Private sh As shell
Private shWin As ShellWindows
Private item As InternetExplorer


Public Property Get powloka()
    Set powloka = shWin
End Property


Public Function getFirstIE() As InternetExplorer

    
    For Each item In shWin
        On Error Resume Next
        If CStr(item) Like "*Internet*Explorer*" Then
            
            
            Set getFirstIE = item
            Exit Function
        End If
    Next item
    
End Function

Public Function getIEWithURL(strURL) As InternetExplorer
    
    For Each item In shWin
        On Error Resume Next
        If CStr(item) Like "*Internet*Explorer*" Then
            
            If item.LocationURL Like "*" & strURL & "*" Then
            
                Set getIEWithURL = item
                Exit Function
            End If
        End If
    Next item
    
End Function


Public Function catchLoggedCorail() As InternetExplorer


    Set catchLoggedCorail = Nothing
    Exit Function
    
    For Each item In shWin
        On Error Resume Next
        If CStr(item) Like "*Internet*Explorer*" Then
            
            If checkIfItsLoggedCorail(item) Then
                Set catchLoggedCorail = item
                Exit Function
            End If
            
            
        End If
    Next item
End Function


Public Sub zamknijWszystkiePrzegladarki()
    
    Dim tmpCollectionOfIes As New Collection
    Dim tmpIE As InternetExplorer
    
    
    For Each item In shWin
        On Error Resume Next
        If CStr(item) Like "*Internet*Explorer*" Then
        
            Set tmpIE = Nothing
            Set tmpIE = item
            tmpCollectionOfIes.Add tmpIE
        End If
    Next item
    
    
    For Each tmpIE In tmpCollectionOfIes
        tmpIE.Quit
        Set tmpIE = Nothing
    Next tmpIE
End Sub


Public Sub zamknijWszystkiePrzegladarkiZCorail()

    Dim tmpCollectionOfIes As New Collection
    Dim tmpIE As InternetExplorer
    
    
    For Each item In shWin
        On Error Resume Next
        If CStr(item) Like "*Internet*Explorer*" Then
        
        
            
            Set tmpIE = Nothing
            Set tmpIE = item
            
            ' Debug.Print tmpIE.LocationURL
            If UCase(tmpIE.LocationURL) Like UCase("*CORAIL*") Or UCase(tmpIE.LocationURL) Like UCase("*MAESTRO*") Then
                tmpCollectionOfIes.Add tmpIE
            End If
            
            
        End If
    Next item
    
    
    For Each tmpIE In tmpCollectionOfIes
        tmpIE.Quit
        Set tmpIE = Nothing
    Next tmpIE
    
    
    
End Sub


Private Function checkIfItsLoggedCorail(ie As InternetExplorer) As Boolean
    checkIfItsLoggedCorail = False
End Function

Private Sub Class_Initialize()

    Set sh = New shell
    Set shWin = sh.Windows
    
End Sub

Private Sub Class_Terminate()

    Set sh = Nothing
    Set shWin = Nothing
    
End Sub
