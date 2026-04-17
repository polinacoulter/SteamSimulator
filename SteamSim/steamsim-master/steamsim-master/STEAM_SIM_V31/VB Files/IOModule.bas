Attribute VB_Name = "IOModule"
Declare Sub BUSLINK Lib "c:\HAVEN\Model.DLL" ()

' 20/2/98
Declare Sub set_VB_bus_test Lib "c:\HAVEN\Model.DLL" (VB_bus_test As Byte)


Declare Sub IO Lib "c:\HAVEN\Model.DLL" Alias "io" ()
Declare Sub INITIO Lib "c:\HAVEN\Model.DLL" ()
Declare Sub free_io_board Lib "c:\HAVEN\Model.DLL" ()
Declare Sub set_AO Lib "c:\HAVEN\Model.DLL" (AO As Integer)

' 20/4/98
'Declare Sub get_pots Lib "c:\HAVEN\Model.DLL" (AI As Integer)

Declare Sub set_DO Lib "c:\HAVEN\Model.DLL" (VB_DO As Byte)
Declare Sub save_DO Lib "c:\HAVEN\Model.DLL" ()
Declare Sub restore_DO Lib "c:\HAVEN\Model.DLL" ()
Declare Sub get_DI_Changes Lib "c:\HAVEN\Model.DLL" (DI_Changes As Integer, DI_States As Byte)
Declare Sub reset_DI_Changes Lib "c:\HAVEN\Model.DLL" (nCHANNEL As Integer)
Declare Sub CLOCK Lib "c:\HAVEN\Model.DLL" ()
Declare Sub Set_DIBoardRange Lib "c:\HAVEN\Model.DLL" (DIBaseSelected As Integer, DITopSelected As Integer)

' 18/2/98
' 20/10/99....
'Declare Sub TEST_SSCOPE Lib "c:\HAVEN\Model.DLL" (Fdiff As Single, Scope_controls As Byte)
Declare Sub TEST_SSCOPE Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (Fdiff As Single, Scope_controls As Byte)
Global Fdiff As Single
Global Scope_controls(7) As Byte


' 19/2/98
Declare Sub get_DI_boards Lib "c:\HAVEN\Model.DLL" (DI_board_present As Byte)
Declare Sub set_DI_boards Lib "c:\HAVEN\Model.DLL" (DI_board_present As Byte)
Global DI_board_present(128) As Byte

' 19/2/98
Declare Sub set_DI_board_default Lib "c:\HAVEN\Model.DLL" ()
       


Global RunIO As Byte, old_RunIO As Byte

Global VB_RunIO As Byte
Declare Sub set_RunIO Lib "c:\HAVEN\Model.DLL" (VB_RunIO As Byte)

