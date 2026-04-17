Attribute VB_Name = "Types"
' data arrays
Global Pressure(5) As Single
Global Enthalpy(5) As Single
Global Temperature(5) As Single

Global Dryness(5) As Single

Global Entropy(5) As Single

Global SIUnits As Boolean ' May need to be deleted
' hold range of x,y axis. used to convert between si and bs
Global MINENTROPY As Single
Global MAXENTROPY As Single
Global MINENTHALPY As Single
Global MAXENTHALPY As Single
'used to prepare grid
Global Const GRID_ROW_ENTROPY = 2
Global Const GRID_ROW_ENTHALPY = 1
Global Const GRID_ROW_PRESSURE = 3
Global Const GRID_ROW_TEMPERATURE = 4
Global Const GRID_ROW_DRYNESS = 5
'used to set graph lables
Declare Function VBAGLabels Lib "GSWAG32.DLL" (ByVal nNLabs&, ByVal szLabs$) As Long


