VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================
' clsDB.cls
' Centralized Database Connection
'==============================

Option Explicit

' -----------------------------
' Public ADODB connections
' -----------------------------
Public con As ADODB.Connection         ' Main database
Public conArchive As ADODB.Connection  ' Archive database

' -----------------------------
' Table change tracker
' -----------------------------
Private m_TableCounts As Object   ' Late binding dictionary

' -----------------------------
' Internal typing flag
' -----------------------------
Private bIsTyping As Boolean

'==============================
' Open database connections
'==============================
Public Sub OpenDB()
    ' --- Main database ---
    If con Is Nothing Then
        Set con = New ADODB.Connection
        con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Vin\Documents\MEQODES-main\MEQODE\mdb\MasterData.mdb;"
        con.Open
    ElseIf con.State = adStateClosed Then
        con.Open
    End If

    ' --- Archive database ---
    If conArchive Is Nothing Then
        Set conArchive = New ADODB.Connection
        conArchive.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Vin\Documents\MEQODES-main\MEQODE\Archives\bookingarchive\bookingarchive.mdb;"
        conArchive.Open
    ElseIf conArchive.State = adStateClosed Then
        conArchive.Open
    End If

    ' Initialize table tracker
    If m_TableCounts Is Nothing Then
        Set m_TableCounts = CreateObject("Scripting.Dictionary")
    End If
End Sub

'==============================
' Close database connections
'==============================
Public Sub CloseDB()
    ' Close main
    If Not con Is Nothing Then
        If con.State = adStateOpen Then con.Close
        Set con = Nothing
    End If

    ' Close archive
    If Not conArchive Is Nothing Then
        If conArchive.State = adStateOpen Then conArchive.Close
        Set conArchive = Nothing
    End If

    ' Clear table tracker
    If Not m_TableCounts Is Nothing Then
        m_TableCounts.RemoveAll
        Set m_TableCounts = Nothing
    End If
End Sub

'==============================
' Sanitize text for SQL queries
'==============================
Public Function SafeText(ByVal s As String) As String
    SafeText = Replace(s, "'", "''")
End Function

'==============================
' Check if table record count changed
'==============================
Public Function HasTableChanged(ByVal tableName As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim currentCount As Long

    If con Is Nothing Then Exit Function
    If con.State <> adStateOpen Then Exit Function

    Set rs = New ADODB.Recordset
    rs.Open "SELECT COUNT(*) AS TotalRec FROM " & tableName, _
            con, adOpenForwardOnly, adLockReadOnly

    currentCount = rs!TotalRec
    rs.Close
    Set rs = Nothing

    ' First time tracking this table
    If Not m_TableCounts.Exists(tableName) Then
        m_TableCounts.Add tableName, currentCount
        HasTableChanged = False
        Exit Function
    End If

    ' Compare with previous count
    If m_TableCounts(tableName) <> currentCount Then
        m_TableCounts(tableName) = currentCount
        HasTableChanged = True
    Else
        HasTableChanged = False
    End If
End Function

'==============================
' Execute SQL safely
'==============================
Public Sub ExecuteSQL(ByVal sql As String, Optional ByVal useArchive As Boolean = False)
    On Error GoTo ErrHandler
    If useArchive Then
        If Not conArchive Is Nothing Then conArchive.Execute sql
    Else
        If Not con Is Nothing Then con.Execute sql
    End If
    Exit Sub
ErrHandler:
    MsgBox "Database error: " & Err.Description & vbCrLf & sql, vbCritical
End Sub

'==============================
' Execute scalar (single-value) query
'==============================
Public Function ExecuteScalar(ByVal sql As String, Optional ByVal useArchive As Boolean = False) As Variant
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If useArchive Then
        rs.Open sql, conArchive, adOpenForwardOnly, adLockReadOnly
    Else
        rs.Open sql, con, adOpenForwardOnly, adLockReadOnly
    End If

    If Not rs.EOF Then
        ExecuteScalar = rs.Fields(0).Value
    Else
        ExecuteScalar = Null
    End If

    rs.Close
    Set rs = Nothing
End Function

