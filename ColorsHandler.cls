VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColorsHandler"
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



Public colorPurpleMain As Long
Public colorPurpleMore As Long
Public colorPurpleDark As Long
Public colorPurpleLight As Long

Public colorMattBlueMain As Long
Public colorMattBlueMore As Long
Public colorMattBlueDark As Long
Public colorMattBlueLight As Long

Public colorMattBlack As Long

Public colorFireFlakeGrey As Long



Private Sub Class_Initialize()

    Me.colorMattBlack = RGB(20, 20, 20)
    
    Me.colorMattBlueDark = RGB(6, 12, 27)
    Me.colorMattBlueLight = RGB(128, 153, 216)
    Me.colorMattBlueMain = RGB(46, 65, 114)
    Me.colorMattBlueMore = RGB(79, 102, 158)
    
    Me.colorPurpleDark = RGB(11, 6, 27)
    Me.colorPurpleLight = RGB(152, 129, 218)
    Me.colorPurpleMain = RGB(65, 48, 117)
    Me.colorPurpleMore = RGB(89, 42, 113)
    
    Me.colorFireFlakeGrey = RGB(200, 200, 200)
    
End Sub

