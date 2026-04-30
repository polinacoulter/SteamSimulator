Attribute VB_Name = "AppIO_module"
'****************************************************************************}
'*                                                                          *}
'*  APPIO.BAS  Version 1.0  date :14/05/98                                  *}
'*  Copyright (c) 1992-1998 Applicom International. All rights reserved.    *}
'*                                                                          *}
'*      Definitions file for  APPLICOM IO Windows applications              *}
'*                                                                          *}
'*               ******************************************                 *}
'*               *   WARNING   do not modify this file    *                 *}
'*               ******************************************                 *}
'*                                                                          *}
'****************************************************************************}

Public Declare Function IO_Init Lib "appio" (ByVal wCard As Integer, Status As Integer) As Boolean
Public Declare Function IO_Exit Lib "appio" (ByVal wCard As Integer, Status As Integer) As Boolean

Public Declare Function IO_VerifAuto Lib "appio" (ByVal bAuto As Boolean) As Boolean

Public Declare Function IO_GetDigitalInput Lib "appio" (ByVal wCard As Integer, Status As Integer) As Integer

Public Declare Function IO_RefreshInput Lib "appio" (ByVal wCard As Integer, Status As Integer) As Boolean
Public Declare Function IO_RefreshOutput Lib "appio" (ByVal wCard As Integer, Status As Integer) As Boolean

Public Declare Function IO_VerifParamStatus Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal Offset As Integer, ByVal Nb As Integer, ByVal Types As Integer, Status As Integer) As Boolean
Public Declare Function IO_VerifParamRead Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal Offset As Integer, ByVal Nb As Integer, ByVal Types As Integer, Status As Integer) As Boolean
Public Declare Function IO_VerifParamWrite Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal Offset As Integer, ByVal Nb As Integer, ByVal Types As Integer, Status As Integer) As Boolean

Public Declare Function IO_ReadIBit Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal OffsetBit As Integer, ByVal Nb As Integer, TabBit As Byte, Status As Integer) As Boolean
Public Declare Function IO_ReadIByte Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal Offset As Integer, ByVal Nb As Integer, TabByte As Byte, Status As Integer) As Boolean
Public Declare Function IO_ReadIWord Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal Offset As Integer, ByVal Nb As Integer, TabWord As Integer, Status As Integer) As Boolean
Public Declare Function IO_ReadIDWord Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal Offset As Integer, ByVal Nb As Integer, TabDWord As Long, Status As Integer) As Boolean

Public Declare Function IO_WriteQBit Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal OffsetBit As Integer, ByVal Nb As Integer, TabBit As Byte, Status As Integer) As Boolean
Public Declare Function IO_WriteQByte Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal Offset As Integer, ByVal Nb As Integer, TabByte As Byte, Status As Integer) As Boolean
Public Declare Function IO_WriteQWord Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal Offset As Integer, ByVal Nb As Integer, TabWord As Integer, Status As Integer) As Boolean
Public Declare Function IO_WriteQDWord Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal Offset As Integer, ByVal Nb As Integer, TabDWord As Long, Status As Integer) As Boolean

Public Declare Function IO_GetEquipmentList Lib "appio" (ByVal wCard As Integer, Nb As Integer, TabEquip As Integer, Status As Integer) As Boolean
Public Declare Function IO_GetEquipmentInfo Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, InputSize As Integer, OutputSize As Integer, Status As Integer) As Boolean

Public Declare Function IO_SetWatchDog Lib "appio" (ByVal wCard As Integer, ByVal TimeWD As Integer, Status) As Boolean
Public Declare Function IO_SetReply Lib "appio" (ByVal wCard As Integer, ByVal Reply As Integer, Status As Integer) As Boolean
Public Declare Function IO_GetEquipmentStatus Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, EquipmentStatus As Integer, Status As Integer) As Boolean
Public Declare Function IO_GetGlobalStatus Lib "appio" (ByVal wCard As Integer, GlobalStatus As Integer, Status As Integer) As Boolean

Public Declare Function IO_ReadEquipmentOut Lib "appio" (ByVal wCard As Integer, ByVal Equipment As Integer, ByVal Offset As Integer, ByVal NbByte As Integer, TabByte As Byte, Status As Integer) As Boolean

