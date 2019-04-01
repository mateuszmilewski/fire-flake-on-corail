Attribute VB_Name = "GlobalModule"
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


' sheets
' ---------------------------------------------
Global Const G_SH_NM_REG = "register"
Global Const G_SH_NM_IN = "input"
Global Const G_SH_NM_PLT_LIST = "plt-list"
' ---------------------------------------------



' limites
' ---------------------------------------------

Global Const G_LIMIT_IE = 2

Global Const G_CORAIL_FIRST_PLT = 3
Global Const G_CORAIL_LAST_PLT = 100

' ---------------------------------------------



' Corail Const Types
' ---------------------------------------------


Global Const G_BLUE_TXT = "BLUE"
Global Const G_ORANGE_TXT = "ORANGE"
Global Const G_MANUAL_TXT = "MANUAL"
Global Const G_MAESTRO_TXT = "MAESTRO"

' ---------------------------------------------



' HTML HANDLING CONSTANTS
' ---------------------------------------------
Global Const G_MAIN_FRAME_ID = "frmTop"
Global Const G_INNER_MAIN_FRAME_ID = "frmMain"
' ---------------------------------------------



' to get the frames directly!
' ---------------------------------------------
' getProductSummaryRead.do?beanId=96661053ZD#
Global Const G_URL_EXT = "getProductSummaryRead.do?beanId="
Global Const G_MAESTRO_URL_EXT = "produit.do?methode=init&selectedcodeProduit="
' ---------------------------------------------



' Global Yes / No for MAESTRO
' ---------------------------------------------
Global isMaestroAvail As Boolean
' ---------------------------------------------


' ---------------------------------------------
Global Const G_NBSP = "&nbsp;"
' ---------------------------------------------


' ITEM OFFSET
' ---------------------------------------------
Global Const G_ITEM_OFFSET = 4
Global Const G_ITEM_OFFSET_LEAN = 3
Global Const G_ITEM_OFFSET_EXTENDED = 5
' ---------------------------------------------





Global G_LOGIN As String
Global G_PASS As String
Global G_HAZARDS As Boolean
