VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'##### MM/DD/YYYY #####
'## 2019 ##
'02/21 SQLGetRecordset

'## 2018 ##
'03/07 Return 1 output
'03/09 sleep after conection

'## 2016 ##
'09/01 Criacao do modulo

Dim SQLconn As ADODB.Connection
Dim SQLrs As ADODB.Recordset
Dim Excelconn As ADODB.Connection

Public Function Sql_Connect(sConnString As String) As Boolean
    Dim ok As Boolean
    Set SQLconn = New ADODB.Connection    'Create the Connection and Recordset objects.
    
    On Error Resume Next
        SQLconn.Open sConnString    'Open the connection and execute.
    On Error GoTo 0
    ok = CBool(SQLconn.State And adStateOpen)
    If Not ok Then
        LogSheet "Sql_Connect", "wait to connect", "", getLogType(enLog.Warning), A_Write
        Sleep 2000
        On Error Resume Next
            SQLconn.Open sConnString    'Open the connection and execute.
        On Error GoTo 0
        ok = CBool(SQLconn.State And adStateOpen)
    End If
    Sql_Connect = ok
End Function

Public Sub SQL_Cleanup()
    If CBool(SQLconn.State And adStateOpen) Then SQLconn.Close
    Set SQLconn = Nothing
End Sub

Public Function Excel_Connect() As Boolean
    Dim sCon As String
    sCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='"
    '20180119'sCon = sCon & Main_Workbook_Name & "';"
    sCon = sCon & Main_Workbook_Path & "\" & Main_Workbook_Name & "';"
    sCon = sCon & "Extended Properties='Excel 12.0;HDR=YES;IMEX=1';"
    Set Excelconn = New ADODB.Connection    'Create the Connection and Recordset objects.
    On Error Resume Next
    Excelconn.Open sCon    'Open the connection and execute.
    On Error GoTo 0
    Excel_Connect = CBool(Excelconn.State And adStateOpen)
End Function

Public Sub Excel_Cleanup()
    If CBool(Excelconn.State And adStateOpen) Then Excelconn.Close
    Set Excelconn = Nothing
End Sub

Public Sub SQLTakeData(sTable As String, sWhereOrder As String, sSheet As String, iCellColor As Long, iTabColor As Long)
    Dim iCol As Long, iColBase As Long, sCol As String, iFirstRow As Long, sSelect As String
    Set SQLrs = New ADODB.Recordset
    sSelect = "SELECT * FROM " & sTable & space(1) & sWhereOrder & ";"
    
    '20180119'Set SQLrs = SQLconn.Execute(sSelect)
    On Error Resume Next
        Err.Clear
        Set SQLrs = SQLconn.Execute(sSelect)
        If Err.Number <> 0 Then
            Debug.Print Err.Description & vbCrLf
            sSelect = "SELECT * FROM " & sTable & space(1) & sWhereOrder
            '--
            Err.Clear
            Set SQLrs = SQLconn.Execute(sSelect)
            If Err.Number <> 0 Then
                Debug.Print Err.Description & vbCrLf
                Err.Clear
                SQLrs.Open sSelect, SQLconn, , , adCmdText
            End If
            If Err.Number <> 0 Then
                Debug.Print Err.Description & vbCrLf
                Exit Sub
            End If
        End If
    On Error GoTo 0
    
    'Check if have data.
    If Not SQLrs.EOF Then
        'Transfer result.
        iCol = 1
        With Workbooks(Main_Workbook_Name).Sheets(sSheet)
            .Tab.ColorIndex = iTabColor
            iFirstRow = 2
            While .Range(IdToColumn(iCol) & iFirstRow).Value <> ""
                iCol = iCol + 1
                EnableEvents
            Wend
            iFirstRow = 1
            .Range(IdToColumn(iCol) & iFirstRow).Value = sTable
            .Range(IdToColumn(iCol) & iFirstRow).Interior.ColorIndex = iCellColor

            iColBase = iCol
            iFirstRow = 2
            For iCol = 0 To SQLrs.Fields.Count - 1
                .Range(IdToColumn(iColBase + iCol) & iFirstRow).Value = Trim(SQLrs.Fields(iCol).Name)
            Next iCol

            iCol = iColBase
            iFirstRow = 3
            .Range(IdToColumn(iCol) & iFirstRow).CopyFromRecordset SQLrs
        End With
        SQLrs.Close    'Close the recordset
    Else
        'MsgBox "Error: No records returned." & sTable, vbCritical
        LogSheet "SQLTakeData", "No records returned.", sSelect, getLogType(enLog.Warning), A_Write
    End If
    Set SQLrs = Nothing
End Sub

Public Function ExcelTakeData(sSheet As String, sTable As String, colIni As String, colEnd As String, FirstRow As Long) As Collection
    If Not WorksheetExists(sSheet) Then Exit Function
    '
    Dim rs As ADODB.Recordset, sSQL As String
    Set rs = New ADODB.Recordset
    sSQL = "select * from [" & sSheet & "$" & colIni & FirstRow & ":" & colEnd & "] "
    rs.Open sSQL, Excelconn, adOpenKeyset, adOpenStatic, adCmdText

    If Not rs.BOF And Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
        Dim i As Long, v As Variant
        Set ExcelTakeData = New Collection
        Dim mData()
        ReDim mData(0 To rs.Fields.Count - 1)
        For i = 0 To rs.Fields.Count - 1
            mData(i) = rs.Fields(i).Name
        Next i
        ExcelTakeData.add mData()
        Do Until rs.EOF = True
            On Error Resume Next '20190118
            v = rs.Fields(0).Value '20190118
            If 0 = Err.Number Then '20190118
            If Not IsNull(rs.Fields(0).Value) Then
                On Error GoTo 0 '20190118
                For i = 0 To rs.Fields.Count - 1
                    mData(i) = rs.Fields(i).Value
                Next i
                ExcelTakeData.add mData()
            End If
            End If
            rs.MoveNext
        Loop
        On Error GoTo 0 '20190118
    End If
    '
    rs.Close
    Set rs = Nothing
End Function

Public Function SQLOneData(ByVal sSelect As String) As Variant
    Dim RET() As Variant, sCompare As String, ok As Boolean, iField As Long
    ReDim RET(0 To 0)
    
    Set SQLrs = New ADODB.Recordset
    '
    sSelect = Trim(sSelect)
    If "(" = Left(sSelect, 1) And ")" = Right(sSelect, 1) Then
        sSelect = Mid(sSelect, 2, Len(sSelect) - 2)
    End If
    If ";" <> Right(sSelect, 1) Then sSelect = sSelect & ";"
    'If SQLconn.Provider = "Microsoft.ACE.OLEDB.12.0" Then
    '    sSelect = Replace(sSelect, SQ, DQ)
    'Else
        
    'End If
    
    On Error Resume Next
        Err.Clear
        Set SQLrs = SQLconn.Execute(sSelect)
        If Err.Number <> 0 Then
            Stop
            Debug.Print Err.Description & vbCrLf
            '--
            Err.Clear
            
            Set SQLrs = SQLconn.Execute(sSelect)
            If Err.Number <> 0 Then
                Debug.Print Err.Description & vbCrLf
                Err.Clear
                SQLrs.Open sSelect, SQLconn, , , adCmdText
            End If
            If Err.Number <> 0 Then
                Debug.Print Err.Description & vbCrLf
                RET(0) = Err.Description
                GoTo Exiting
            End If
        End If
    On Error GoTo 0
    
    ok = False
    sCompare = UCase("(SELECT")
    If Left(UCase(sSelect), Len(sCompare)) = sCompare Then ok = True
    sCompare = "SELECT"
    If Left(UCase(sSelect), Len(sCompare)) = sCompare Then ok = True
    
    If Not ok Then
        'it is not select
        RET(0) = "DONE"
    Else
        If Not SQLrs.EOF Then
            While Not SQLrs.EOF
                For iField = 0 To -1 + SQLrs.Fields.Count
                    If "" <> RET(UBound(RET)) Then
                        ReDim Preserve RET(0 To 1 + UBound(RET))
                    End If
                    RET(UBound(RET)) = SQLrs.Fields(iField).Value
                Next iField
                
                SQLrs.MoveNext
                EnableEvents
            Wend
            SQLrs.Close 'Close the recordset
        Else
            LogSheet "SQLOneData", "No records returned.", sSelect, getLogType(enLog.Warning), A_Write
            RET(0) = "ERR"
        End If
    End If
    '
Exiting:
    SQLOneData = RET
    Set SQLrs = Nothing
End Function

Public Sub SQLGetRecordset(ByRef sSelect As String, ByRef Faults As ErrObject, ByRef rs As ADODB.Recordset)
    If CBool(SQLconn.State = adStateClosed) Then Stop
    Set rs = New ADODB.Recordset
    Set rs = SQLconn.Execute(sSelect)
    rs.MoveFirst
    If CBool(rs.State = adStateClosed) Then Stop
End Sub
