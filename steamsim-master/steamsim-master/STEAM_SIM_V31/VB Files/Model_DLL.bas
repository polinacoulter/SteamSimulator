Attribute VB_Name = "Model_DLL"


' 5/7/99 Moved here....
Declare Sub set_plot_channels Lib "c:\HAVEN\Model.DLL" (plot_channels As Integer)
Declare Sub set_plot_channels_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (plot_channels As Integer)
Global plot_channels(16) As Integer

Declare Sub M_SNAP Lib "c:\HAVEN\Model.DLL" (SnapNum As Integer, SnapType As Integer)
Declare Sub M_RELSNAP Lib "c:\HAVEN\Model.DLL" (SnapNum As Integer, SnapType As Integer)
Global SnapNum As Integer
Global SnapType As Integer


' 24/7/98
Declare Sub set_instrp Lib "c:\HAVEN\Model.DLL" (Instrp_Id As Integer, Instrp_Value As Single)

' 27/4/98
Declare Sub Init_Tasks Lib "c:\HAVEN\Model.DLL" (task_array As Byte)
Global task_array(100) As Byte

' 21/4/98
Declare Sub set_PORTUGAL Lib "c:\HAVEN\Model.DLL" (Portugal_flag As Byte)
Global Portugal_flag As Byte
