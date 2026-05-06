Attribute VB_Name = "Module1"
'*************************************************************************
'
'  APPLICOM.BAS   Version 3.0   date :14/05/98
'  Copyright (c) 1988-1998 Applicom International. All rights reserved.
'
'  Include Declarations file (APPLICOM  Visual Basic) for Windows applications
'               ******************************************
'               *   WARNING   do not modify this file    *
'               ******************************************
'
'**************************************************************************
Option Explicit

Public Const APP_256BYTES_BASED_LIMITS = 256
Public Const APP_1584BYTES_BASED_LIMITS = 1584

Rem fonctions d'initialisation
Declare Sub initbus Lib "applicom" (stat As Integer)
Declare Sub exitbus Lib "applicom" (stat As Integer)
Declare Function AuSetApplicationMaxSize Lib "applicom" (ByVal buffersize As Integer, stat As Long) As Boolean
Declare Function AuGetApplicomMaxSize Lib "applicom" (sizeDB As Integer, sizeChan As Integer, stat As Long) As Boolean




Rem fonctions mode cyclique
Declare Sub createcyc Lib "applicom" (chan As Integer, fonc As Integer, per As Integer, act As Integer, typf As Integer, codedb As Integer, nes As Integer, nb As Integer, adrl As Long, adrDB As Integer, adrstat As Integer, stat As Integer)
Declare Sub actcyc Lib "applicom" (chan As Integer, fonc As Integer, stat As Integer)
Declare Sub startcyc Lib "applicom" (chan As Integer, fonc As Integer, stat As Integer)
Declare Sub stopcyc Lib "applicom" (chan As Integer, fonc As Integer, stat As Integer)
Declare Sub transcyc Lib "applicom" (chan As Integer, fonc As Integer, nb As Integer, typ As Integer, tabl As Long, stat As Integer)
Declare Sub transcycpack Lib "applicom" (chan As Integer, fonc As Integer, nb As Integer, typ As Integer, tabl As Integer, stat As Integer)
Declare Sub CycExecuted Lib "applicom" (chan As Integer, nb As Integer, talb As Integer, stat As Integer)
Declare Sub AppGetCycParam Lib "applicom" (chan As Integer, fonc As Integer, tabl As Long, stat As Integer)

Rem fonctions database
Declare Sub confdb Lib "applicom" (card As Integer, typ As Integer, nb As Integer, tadr As Integer, cond As Integer, stat As Integer)
Declare Sub getevent Lib "applicom" (card As Integer, nb As Integer, nestyp As Integer, nesAdr As Integer, nesval As Long, nesdat As Integer, nestime As Integer, stat As Integer)
Declare Sub setevent Lib "applicom" (card As Integer, typ As Integer, nb As Integer, adr As Integer, stat As Integer)
Declare Sub incdispword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub incdispdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Long, stat As Integer)
Declare Sub decdispword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub decdispdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Long, stat As Integer)
Declare Sub getdispbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub getdispword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub getdispdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Long, stat As Integer)
Declare Sub getdispfword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Single, stat As Integer)
Declare Sub setdispbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub setdispword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub setdispdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Long, stat As Integer)
Declare Sub setdispfword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Single, stat As Integer)
Declare Sub getbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub getpackbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub getpackbyte Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Byte, stat As Integer)
Declare Sub getword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub getdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Long, stat As Integer)
Declare Sub getfword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Single, stat As Integer)
Declare Sub setbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub setpackbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub setpackbyte Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Byte, stat As Integer)
Declare Sub setword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub setdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Long, stat As Integer)
Declare Sub setfword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Single, stat As Integer)
Declare Sub invbit Lib "applicom" (nb As Integer, tabls As Integer, tabld As Integer, stat As Integer)

Rem fonctions mode attente
Declare Sub readquickbit Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub readpackbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readpackibit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readpackqbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readpackbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub readpackibyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub readpackqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub readbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readibyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readwordbcd Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readiword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readqword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readdword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Long, stat As Integer)
Declare Sub readfword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Single, stat As Integer)

Declare Sub writepackbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writepackqbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writepackbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, adr As Long, tabl As Byte, stat As Integer)
Declare Sub writepackqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub writebyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writeqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writeword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writewordbcd Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writeqword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writedword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Long, stat As Integer)
Declare Sub writefword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Single, stat As Integer)
Declare Sub readcounter Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, adr As Long, tabl As Integer, stat As Integer)
Declare Sub writecounter Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, adr As Long, tabl As Integer, stat As Integer)
Declare Sub readtimer Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, adr As Long, tabl As Integer, stat As Integer)
Declare Sub writetimer Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, adr As Long, tabl As Integer, stat As Integer)
Declare Sub readmes Lib "applicom" (chan As Integer, tstop As Integer, nb As Integer, tim As Integer, tim_ic As Integer, tabl As Byte, stat As Integer)
Declare Sub writemes Lib "applicom" (chan As Integer, nb As Integer, tabl As Byte, stat As Integer)
Declare Sub writereadmes Lib "applicom" (chan As Integer, nb_tx As Integer, buf_tx As Byte, tstop As Integer, max_rx As Integer, time_out As Integer, time_ic As Integer, nb_rx As Integer, buf_rx As Byte, stat As Integer)

Rem fonctions mode différé
Declare Sub readdifquickbit Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub readdifbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Integer, stat As Integer)
Declare Sub readdifibit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifqbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Integer, stat As Integer)
Declare Sub readdifibyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifiword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifqword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifdword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdiffword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub writedifpackbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writedifpackqbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writedifpackbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub writedifpackqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub writedifword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writedifqword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writedifdword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Long, stat As Integer)
Declare Sub writediffword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Single, stat As Integer)
Declare Sub readdifmes Lib "applicom" (chan As Integer, tstop As Integer, nb As Integer, tim As Integer, tim_ic As Integer, stat As Integer)
Declare Sub writedifmes Lib "applicom" (chan As Integer, nb As Integer, tabl As Byte, stat As Integer)
Declare Sub writereaddifmes Lib "applicom" (chan As Integer, nbtx As Integer, buftx As Byte, tstop As Integer, nbrx As Integer, timeout As Integer, timeic As Integer, stat As Integer)
Declare Sub transdif Lib "applicom" (chan As Integer, nb As Integer, typ As Integer, tabl As Long, stat As Integer)
Declare Sub transdifpack Lib "applicom" (chan As Integer, nb As Integer, typ As Integer, tabl As Integer, stat As Integer)
Declare Sub testtransdif Lib "applicom" (chan As Integer, request As Integer, receive As Integer, stat As Integer)

Rem fonctions specifiques UTE
Declare Sub readcounterute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub writecounterute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, preset As Integer, stat As Integer)
Declare Sub resetcounterute Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub readiobitute Lib "applicom" (chan As Integer, nes As Integer, module As Integer, tabl As Integer, status As Integer)
Declare Sub writeiobitute Lib "applicom" (chan As Integer, nes As Integer, modu As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub readmonostableute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub writemonostableute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, preset As Integer, stat As Integer)
Declare Sub readtimerute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub writetimerute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, preset As Integer, stat As Integer)
Declare Sub readdiagute Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub txtute Lib "applicom" (chan As Integer, nes As Integer, txtic As Integer, txil As Integer, bufemi As Byte, bufrec As Byte, txtis As Integer, txtiv As Integer, stat As Integer)

Rem fonctions diverses
Declare Sub AppConnect Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub AppUnconnect Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub bcdbin Lib "applicom" (nb As Integer, tabls As Integer, tabld As Integer, stat As Integer)
Declare Sub binbcd Lib "applicom" (nb As Integer, tabls As Integer, tabld As Integer, stat As Integer)
Declare Sub getmodem Lib "applicom" (chan As Integer, cts As Integer, dcd As Integer, stat As Integer)
Declare Sub setmodem Lib "applicom" (chan As Integer, rts As Integer, dtr As Integer, stat As Integer)
Declare Sub transbitword Lib "applicom" (nb As Integer, tabls As Integer, tabld As Integer, stat As Integer)
Declare Sub transwordbit Lib "applicom" (nb As Integer, tabls As Integer, tabld As Integer, stat As Integer)
Declare Sub unpackdate Lib "applicom" (pdate As Integer, dday As Integer, mmonth As Integer, yyear As Integer, stat As Integer)
Declare Sub unpacktime Lib "applicom" (ptime As Integer, min As Integer, seg As Integer, hhour As Integer, stat As Integer)
Declare Sub watchdog Lib "applicom" (card As Integer, tim As Integer, stat As Integer)
Declare Sub AppGetWatchDog Lib "applicom" (card As Integer, Entree As Integer, Contact As Integer, stat As Integer)
Declare Sub iocounter Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub resetiocounter Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub readdiag Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub readeven Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub readtrace Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub readident Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tabl As Byte, stat As Integer)
Declare Sub accesskey Lib "applicom" (chan As Integer, rts As Integer, dtr As Integer, stat As Integer)
Declare Sub automatic Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub compword Lib "applicom" (nb As Integer, tabls1 As Integer, tabls2 As Integer, tabld As Integer, stat As Integer)
Declare Sub manual Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub statjbus Lib "applicom" (chan As Integer, tabl As Integer, stat As Integer)
Declare Sub AppFmsGetOd Lib "applicom" (chan As Integer, nes As Integer, attribut As Integer, Acces As Integer, Index As Integer, nb As Integer, tabl As Byte, More As Integer, stat As Integer)
Declare Sub AppFmsStatus Lib "applicom" (chan As Integer, nes As Integer, logStat As Integer, PhyStat As Integer, stat As Integer)
Declare Sub AppIniTime Lib "applicom" (card As Integer, stat As Integer)


