VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pConnStr As String


Property Get connStr() As String
    connStr = pConnStr
End Property


Property Let connStr(value As String)
    Select Case LCase(value)
        Case "steel server"
            value = "Driver={SQL Server};Server=uss-stlsvr.glw.uss.com;Database=USSGLW;Uid=gsm_ro;Pwd=stlmakero;"
    End Select

    pConnStr = value
End Property


Public Function isConnValid() As Boolean
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection

    On Error GoTo CleanFail

    With conn
        .ConnectionString = (pConnStr)
        .Open
        .CursorLocation = adUseClient
    End With

CleanExit:
    isConnValid = True
    conn.Close
    Exit Function

CleanFail:
    isConnValid = False
    conn.Close
    Exit Function
End Function



Public Function runQuery(query As String) As ADODB.Recordset
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    With conn
        .ConnectionString = (pConnStr)
        .Open
        .CursorLocation = adUseClient
    End With

    rs.Open query, conn, adOpenStatic, adLockOptimistic
    
    Set rs.ActiveConnection = Nothing

    Set runQuery = rs
    
    conn.Close
End Function






Public Function typeNumToDesc(typeNum As Integer) As String
    Dim typeDesc As String

    Select Case typeNum
        Case 2
            typeDesc = "adSmallInt"
        Case 3
            typeDesc = "adInteger"
        Case 4
            typeDesc = "adSingle"
        Case 5
            typeDesc = "adDouble"
        Case 6
            typeDesc = "adCurrency"
        Case 7
            typeDesc = "adDate"
        Case 9
            typeDesc = "adIDispatch"
        Case 11
            typeDesc = "adBoolean"
        Case 12
            typeDesc = "adVariant"
        Case 14
            typeDesc = "adDecimal"
        Case 17
            typeDesc = "adUnsignedTinyInt"
        Case 20
            typeDesc = "adBigInt"
        Case 72
            typeDesc = "adGUID"
        Case 128
            typeDesc = "adBinary"
        Case 129
            typeDesc = "adChar"
        Case 130
            typeDesc = "adWChar"
        Case 131
            typeDesc = "adNumeric"
        Case 135
            typeDesc = "adDBTimeStamp"
        Case 200
            typeDesc = "adVarChar"
        Case 201
            typeDesc = "adLongVarChar"
        Case 202
            typeDesc = "adVarWChar"
        Case 203
            typeDesc = "adLongVarWChar"
        Case 204
            typeDesc = "adVarBinary"
        Case 205
            typeDesc = "adLongVarBinary"
        Case Else
            typeDesc = "unknown"
    End Select

    typeNumToDesc = typeDesc
End Function


