Attribute VB_Name = "Wiedmuller_module"
Global Fieldbus As CFieldbus


Global Bus_Names As String
Global Error_string As String
Global Str As String
Global nCom_ports As Long
Global nBaudrate As Long, nCom_port_no As Long

Global NODE_addr As Long
Global bFOUND As Boolean
Global bSOME_ANSWER As Boolean

Global IO_OK As Boolean


Global nInput8 As Long, nInput16 As Long, nInput32 As Long, nOutput8 As Long, nOutput16 As Long, nOutput32 As Long
Global bValue_changed As Boolean

Global station_name As String
Global station_picture_string As String
Global sInfo As String

Global flash_flag As Boolean
Global flash As Boolean

Global out_string_A As String * 1
Global out_string_B As String * 1


