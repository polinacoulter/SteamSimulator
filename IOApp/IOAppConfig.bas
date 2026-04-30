Attribute VB_Name = "IOAppConfig"
Option Explicit

' Loads IOApp.cfg from App.Path on every Start click. The config lists which
' Profibus blocks to skip during reads/writes, per category. Skip-list lookup
' is done at every block iteration in the IOAppProfibus.bas wrappers.
'
' Each skip list is stored as a comma-delimited string with leading and
' trailing commas (e.g., ",52,53,") so InStr can do unambiguous substring
' matches. The literal value "all" is encoded as ",*," in the normalized form.

Public CardA_AI_Skip As String
Public CardA_DI_Skip As String
Public CardA_AO_Skip As String
Public CardA_DO_Skip As String
Public CardB_AI_Skip As String
Public CardB_DI_Skip As String
Public CardB_AO_Skip As String
Public CardB_DO_Skip As String

Public Sub LoadConfig()
    ' Default: nothing skipped.
    CardA_AI_Skip = ","
    CardA_DI_Skip = ","
    CardA_AO_Skip = ","
    CardA_DO_Skip = ","
    CardB_AI_Skip = ","
    CardB_DI_Skip = ","
    CardB_AO_Skip = ","
    CardB_DO_Skip = ","

    Dim path As String
    path = App.path & "\IOApp.cfg"

    If LenB(Dir(path)) = 0 Then
        ' No config file present; use defaults silently.
        Exit Sub
    End If

    On Error Resume Next
    Dim fp As Integer
    fp = FreeFile
    Open path For Input As #fp
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Dim raw As String
    Do While Not EOF(fp)
        Line Input #fp, raw

        Dim line As String
        line = Trim$(raw)

        If Len(line) > 0 And Left$(line, 1) <> "#" Then
            Dim eqPos As Integer
            eqPos = InStr(line, "=")
            If eqPos > 1 Then
                Dim key As String, value As String
                key = LCase$(Trim$(Left$(line, eqPos - 1)))
                value = Trim$(Mid$(line, eqPos + 1))

                Dim normalized As String
                normalized = NormalizeSkipList(value)

                Select Case key
                    Case "carda_ai_skip": CardA_AI_Skip = normalized
                    Case "carda_di_skip": CardA_DI_Skip = normalized
                    Case "carda_ao_skip": CardA_AO_Skip = normalized
                    Case "carda_do_skip": CardA_DO_Skip = normalized
                    Case "cardb_ai_skip": CardB_AI_Skip = normalized
                    Case "cardb_di_skip": CardB_DI_Skip = normalized
                    Case "cardb_ao_skip": CardB_AO_Skip = normalized
                    Case "cardb_do_skip": CardB_DO_Skip = normalized
                End Select
            End If
        End If
    Loop
    Close #fp
End Sub

Private Function NormalizeSkipList(value As String) As String
    ' Returns "," for empty, ",*," for "all", ",N,N,...," for a list.
    Dim trimmed As String
    trimmed = Trim$(value)
    If Len(trimmed) = 0 Then
        NormalizeSkipList = ","
        Exit Function
    End If
    If LCase$(trimmed) = "all" Then
        NormalizeSkipList = ",*,"
        Exit Function
    End If

    Dim parts() As String
    parts = Split(trimmed, ",")

    Dim result As String
    result = ","
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim p As String
        p = Trim$(parts(i))
        If Len(p) > 0 Then
            result = result & CStr(CInt(Val(p))) & ","
        End If
    Next i

    NormalizeSkipList = result
End Function

Public Function IsBlockSkipped(skipList As String, block As Integer) As Boolean
    If InStr(skipList, ",*,") > 0 Then
        IsBlockSkipped = True
        Exit Function
    End If
    IsBlockSkipped = (InStr(skipList, "," & CStr(block) & ",") > 0)
End Function

Public Function ConfigSummary() As String
    ' One-line summary suitable for the log on Start.
    ConfigSummary = "AI A=" & DescribeSkip(CardA_AI_Skip) & _
                    " B=" & DescribeSkip(CardB_AI_Skip) & _
                    "  DI A=" & DescribeSkip(CardA_DI_Skip) & _
                    " B=" & DescribeSkip(CardB_DI_Skip) & _
                    "  AO A=" & DescribeSkip(CardA_AO_Skip) & _
                    " B=" & DescribeSkip(CardB_AO_Skip) & _
                    "  DO A=" & DescribeSkip(CardA_DO_Skip) & _
                    " B=" & DescribeSkip(CardB_DO_Skip)
End Function

Private Function DescribeSkip(skipList As String) As String
    If InStr(skipList, ",*,") > 0 Then
        DescribeSkip = "all"
    ElseIf skipList = "," Or Len(skipList) = 0 Then
        DescribeSkip = "none"
    Else
        ' Strip leading/trailing commas for display.
        Dim s As String
        s = skipList
        If Left$(s, 1) = "," Then s = Mid$(s, 2)
        If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)
        DescribeSkip = s
    End If
End Function
