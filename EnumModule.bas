Attribute VB_Name = "EnumModule"

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

Public Enum E_TYPE_OF_CORAIL
    BLUE
    ORANGE
    MANUAL
    MAESTRO
    UNDEF
End Enum



Public Enum E_LANG
    PL = 1
    ENG
    FR
End Enum



'r.Value = "PART"
'r.Offset(0, 1).Value = "Plant Code"
'r.Offset(0, 2).Value = "Plant Name"
'r.Offset(0, 3).Value = "Supplier"
'r.Offset(0, 4).Value = "Resp"
'r.Offset(0, 5).Value = "Comment #1"
'r.Offset(0, 6).Value = "Comment #2"
'r.Offset(0, 7).Value = "Backlog"
'r.Offset(0, 8).Value = "Stock"

' NEW: r.offset(0,9) = "Recv"

Public Enum E_COMMON_ORDER
    E_COMMON_PN = 1
    E_COMMON_PLT_CODE
    E_COMMON_PLT_NAME
    E_COMMON_PART_NAME
    E_COMMON_SUPPLIER
    E_COMMON_RESP
    E_COMMON_CMNT1
    E_COMMON_CMNT2
    E_COMMON_FIRST_RUNOUT
    E_COMMON_BACKLOG
    E_COMMON_Blockages_in_progress
    E_COMMON_STOCK
    E_COMMON_RECV
    E_COMMON_FIRST_RQM
    E_COMMON_FIRST_ORDER
    E_COMMON_FIRST_SHIP
    E_COMMON_FIRST_BALANCE
End Enum
