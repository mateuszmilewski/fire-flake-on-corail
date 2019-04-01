VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DOMHandler"
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




Public i As HTMLInputElement

Public recvInputElement As HTMLInputElement
Public stockInputElement As HTMLInputElement
Public backlogInputElement As HTMLInputElement
Public hazardsInputElement As HTMLInputElement
Public partNumberInput As HTMLInputElement
Public inputButton As HTMLInputButtonElement
Public button As HTMLButtonElement

Public irecvElement As IHTMLElement
Public istockElement As IHTMLElement
Public ibacklogElement As IHTMLElement
Public ihazardsElement As IHTMLElement
Public icoforElement As IHTMLElement
Public ipartnameElement As IHTMLElement
Public isupplierElement As IHTMLElement
Public iapproElement As IHTMLElement


Public a As HTMLAnchorElement
Public s As HTMLSelectElement

Public elementCollection As HTMLElementCollection
Public iElementCollection As IHTMLElementCollection


Public idoc As IHTMLDocument
Public doc As HTMLDocument

Public iElement As IHTMLElement

Public table As HTMLTable



Public iElementCollectionOfDates As IHTMLElementCollection
Public iElementCollectionOfRqms1 As IHTMLElementCollection
Public iElementCollectionOfRqms2 As IHTMLElementCollection
Public iElementCollectionOfOrders As IHTMLElementCollection
Public iElementCollectionOfShipments As IHTMLElementCollection



