Attribute VB_Name = "V"
'The MIT License (MIT)
'
' Copyright (c) 2019 FORREST
' Mateusz Milewski mateusz.milewski@opel.com aka FORREST
'
' The QT - quickTool
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


' name of this software  due to fact that the main logic was written in a couple of days :P


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-13
' v0.1 init on this project
' 3 cfg sheets: input, register, plt-list
' OOP schema ICorail -> Corail Blue & Orange - a plan
' also plan to have app.run (kind of multi-thread app)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-14
' v0.2 next steps with new classes:
' parser
' rawdata
' shellhandler
' eventhandler connected with corail handler
' sets of corails
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-16
' dopisanie implemenacji odpowiedzialnej za frame:
' Set .frame = .doc.frames(FFOC.G_MAIN_FRAME_ID)
' okazalo sie ze orange corail jest strona w stronie - musialem to jakos obejsc...
'
' new class: DropperHandler
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-20
'v0.4
'duzo zmian
'lacznie z pierwszym udanym polaczeniem z danymi na zywym systemie
'jest to pierwsza podwersja pisana bezposrednio na francuskim sprzecie
'testy natychmiastowe bez koniecznosci przeklikiwania sie pomiedzy mailami
' poprawiony parser
' ujednolicone dzialania pomiedzy corailami blue and orange
' schema:
'CorailHelper -> CorailRunner -> ICorail jako interfejs - orange oraz blue korzystaja z tych samych metod

' Orange, Blue, Manual Corail implements ICorail

''w manual Corail wszystkie metody wlasciwie wygldaja tak samo jak w interfejsie - spowodowane jest to glownie brakiem danych pobiernaych
' wiec generalnie jest pusto i cicho - jedyna zmiana to zaprzestanie wyrzucania bledow krytycznych jesli pod koniec logiki dane wciaz
' sa nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-21
'v0.5
' nowe funkcje:
' 1 open plants
' 2 close all corails and maestros
' 3 after open plants ie is not visible
' 4 initial layout for the tool
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-22
'v0.6
' waiting for IE not working need ta adjust more directly with content of corail site
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-03-06
' v0.7
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' adjust for safe mode in IE
' removal of some logic inside layout changes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' 2018-03-29
' v0.8
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' add export this project module for githib repository...
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' 2018-04-09
' v0.9
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' fix on datadownload - some issues behind taking data and wrong count on balance taking zero from decimal places
' as a "normal" zeros - to be fix on this version
' + dropper handler - added backlog ficzer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-10
' v0.91
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' temp solution for maestro will be treat as manual plan - only filled by zeros and formulas
' some extra fixes on dates and issues on out of range possibility also to be fixed in near future with
' ranges which are too long - some limitation required from end-user.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-11
' v0.92
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' initial implementation for Maestro
' still errors on multi order and requirements numbers - if red font then showing zero - to be fix on 0.93
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' 2018-04-12
' v0.93
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' changes on layout - be more like fire flake
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-16
' v0.94
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' unfinished fire flake layout
' fix in orange requirements data download - to test! - check implementation
' change on parser:
' pCmdCatcher -> pCmdCatcher1 + pCmdCatcher2
' and
' pExpCatcher -> pExpCatcher1 + pCmdCarcher2
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-16 II
' v0.95
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' layout more like fire flake and fill rest of the common data
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-16 III
' v0.01
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' skip to new name -> QT starts to be FF
' to fix;
' no colors on stock
' no filter
' no freeze
' first runout without runout after sorting on top, which is no so right and fine
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' 2018-04-16 IV
' v0.02
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' skip to new name -> QT starts to be FF
' to fix;
' no colors on stock
' no filter
' no freeze
' first runout without runout after sorting on top, which is no so right and fine
' new addons - input comments - converted plt names
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-23
' v0.03
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' fix on part numbers which do not have data in corail or maestro
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' 2018-04-24
' v0.04
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' additional fix on stock
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-06-14
' v0.05
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' new approach on dealing with ending balance on corail on yesterdays dates (split stock on stock depart and reception)
' try to remove connection on format od the date...
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-06-25
' v0.06
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' safe mode for version 0.05
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-06-25
' v0.07
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' final fixes - layout adjustment - prototype for tests.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




' v0.08
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' fix on dropper for manual parts...
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' v0.11
' foo loaded in IEHandler adjusted



' v0.20
' try to use winHttp


' v.21
' added part name + supplier


' v.22 happy copy during tests for tychy plant - performance check on more than 100 parts

' v.23 - monior fixes on layout

' v.24 login - label change - preperation for hourly and weekly


' v.25
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'fix for decimal places in coverage + minor fixes with decimal treatment
'on parser (important) I've used: application.DecimalSeparator
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' v.27
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' feature - added hazards into stock if checkbox value is true
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
