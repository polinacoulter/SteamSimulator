Attribute VB_Name = "ScreenSizing_STEAM"
Option Explicit

Global Const idFO_serv = 1
Global Const idFO_system = 2
Global Const idMain_Bearings = 3
Global Const idLube_Oil = 4
Global Const idGland_Steam = 5
Global Const idSalt_Water = 6
Global Const idMain_Condenser = 7
Global Const idMain_Turbines = 8
' 5/7/99
Global Const idFeed_water = 9

Global Const idCondensate = 10
Global Const idForced_Draft = 11
' 19/7/99
Global Const idBunkers = 12

' 31/8/99
Global Const id_BRNR_MANAGEMENT_STBD = 14
Global Const id_BRNR_MANAGEMENT_PORT = 15

' 2/8/99
' === 09/06/09 === Can NOT have two fifteens....(Note: MAXSCRS = 300 so...)
'Global Const id_Control_Air = 15
Global Const id_Control_Air = 233
' === 09/06/09 Added... ===
Global Const idFeed_water_Alone = 234
Global Const idAux_Steam_Alone = 235

' 6/9/99
Global Const id_SSDG = 16

' 16/9/99
Global Const id_EOT = 17

'27/9/99
Global Const id_drum_stbd = 18
Global Const id_drum_port = 19

Global Const id_temp_stbd = 20
Global Const id_temp_port = 21

Global Const id_fuel_stbd = 22
Global Const id_fuel_port = 23

Global Const id_air_stbd = 24
Global Const id_air_port = 25

Global Const id_master_stbd = 26
Global Const id_master_port = 27

Global Const id_plant_master = 28

Global Const id_fuel_air_stbd = 29
Global Const id_fuel_air_port = 30

Global Const id_feed_pump_stbd = 31
Global Const id_feed_pump_port = 32

Global Const id_EOT_MECD = 33

Global Const id_PID_constants = 34

Global Const idSteam_Calc = 35

' 23/11/99
Global Const id_SSTG = 36

'  9/1/00
Global Const id_feed_pump_balance = 37

' 9/8/00
Global Const id_STM_STM = 38

' 13/9/00
Global Const idThrottle_trips = 39

' 14/9/00
Global Const idSalinometer = 40

' 27/11/00
Global Const idVIAX = 41


' === 22/01/02 ===
Global Const id_Mollier = 42

' === 22/01/02 ===
Global Const id_Synchro = 43

' === 24/01/02 ===
Global Const id_controller_status = 44

' === 05/03/02 ===
Global Const id_combustion_controller = 45
Global Const id_controller_stbd = 46
Global Const id_controller_key = 47

Global Const idBallast = 48

