Attribute VB_Name = "DataBase"
Option Explicit
'database io ref
'Global Const DB = "Driver={Microsoft Access Driver (*.mdb)};" _
'        & "DBQ=c:\apm\io reference\iodatabase.mdb;" _
'        & "UID=admin;PWD=' ';"
Global Const DB = "Driver={Microsoft Access Driver (*.mdb)};" _
        & "DBQ=c:\steam_sim\text\iodatabase.mdb;" _
        & "UID=admin;PWD=' ';"
        ' database path is held by dbq above
Global DBase As New ADODB.Recordset
Global AIDataBase As New ADODB.Recordset
Global DIDataBase As New ADODB.Recordset





Function FillList(DB As String, cntl As Control, _
    sSQL As String) As Boolean

    FillList = False
    
    If TypeOf cntl Is ListBox _
    Or TypeOf cntl Is ComboBox Then
        DBase.Open sSQL, DB, adOpenDynamic, adLockReadOnly, adCmdText
        If (Not (DBase.BOF)) Then
            Dim strAdd As String
            Dim bHasID As Boolean
            Dim i As Integer
            
            'IO Ref  is meant as ItemData
            bHasID = ((vbLong = DBase.Fields(0).Type) _
              And (DBase.Fields.Count > 1))
            Do
                For i = IIf(bHasID, 1, 0) To DBase.Fields.Count - 1
                On Error Resume Next
                    strAdd = strAdd & CStr(DBase(i))
                    If (i < (DBase.Fields.Count - 1)) Then
                        strAdd = strAdd & vbTab
                    End If
                Next i
                            
                If strAdd <> "" Then cntl.AddItem strAdd
                strAdd = ""
                            
                If (bHasID) Then
                    cntl.ItemData(cntl.NewIndex) = DBase(0)
                End If
                
                DBase.MoveNext
            Loop Until DBase.EOF
            FillList = True
            If TypeOf cntl Is ComboBox Then
                cntl.ListIndex = -1
            End If
        End If
    End If
  DBase.Close
End Function

Public Sub OpenDI_AI_DataBase()
'opens ai and di data base for harware test and logging
Dim SQL As String
    '11/10/99 only open if Host connected to Hardware
    
    '29/10/99 db open anyway If Not PTT Then
        SQL = "Select [i/o ref], [Description] From DI"
        DIDataBase.Open SQL, DB, adOpenDynamic, adLockReadOnly, adCmdText
    
        SQL = "Select [i/o ref], [Description] From AI"
        AIDataBase.Open SQL, DB, adOpenDynamic, adLockReadOnly, adCmdText
    ' End If
    
End Sub

Public Sub CloseDI_AI_Database()
'closes databases
    '11/10/99
    '29/10/99 db open anyway  If Not PTT Then
    
        '11/11/99 no database any more
        'DIDataBase.Close
        'AIDataBase.Close
    '  End If
    
End Sub
