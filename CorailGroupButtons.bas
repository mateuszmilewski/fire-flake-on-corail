Attribute VB_Name = "CorailGroupButtons"
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

Public Sub closeAllIEs(ictrl As IRibbonControl)
    innerCloseAllBrowsers
End Sub

Public Sub adjustInputList(ictrl As IRibbonControl)
    MsgBox "to be implemented"
End Sub



Public Sub openPlants(ictrl As IRibbonControl)

    ans = MsgBox("Want to open also Maestro system?", vbYesNo, "Want to open also Maestro system?")
    
    If ans = vbYes Then
        isMaestroAvail = True
    Else
        isMaestroAvail = False
    End If
    
    innerOpenPlants
End Sub


Private Sub innerOpenPlants()
    ' RUN ALL PLANTS AT THE BEGINNING
    ''''''''''''''''''''''''''''''''''''''''''''
    Dim c As CorailHelper
    Set c = New CorailHelper
    c.runAllPlants
    
    MsgBox "ready!"
    ''''''''''''''''''''''''''''''''''''''''''''
End Sub

Public Sub innerCloseAllBrowsers()
    
    Dim ieh As IEHandler
    Set ieh = New IEHandler
    ieh.closeAllIEs
    MsgBox "ready!"
End Sub
