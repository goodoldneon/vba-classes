VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordsetToRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'# Optional settings.
' .title = {String}
' .headerTop = {Boolean}
' .headerTopAlign = {"center", "left", or "right"}
' .headerLeft = {Boolean}
' .headerLeftAlign = {"center", "left", or "right"}
' .headerFontSize = {Integer}


Private pTitle As String
Private pRotate As Boolean
Private pBorder As Boolean
Private pFieldNames As Boolean

Private Type header
    position As String
    left As Boolean
    leftAlign As String
    top As Boolean
    topAlign As String
    fontSize As Integer
End Type

Private pHeader As header


'# Initialize (run on instantiation).
Private Sub Class_Initialize()
    pBorder = False
    pRotate = False
    pTitle = ""
    pFieldNames = True
    
    With pHeader
        .position = "top"
        .left = False
        .leftAlign = ""
        .top = False
        .topAlign = ""
        .fontSize = 12
    End With
End Sub



'# Get and Let.
Public Property Get title() As String
    title = pTitle
End Property
Public Property Let title(value As String)
    pTitle = value
End Property

Public Property Get rotate() As Boolean
    rotate = pRotate
End Property
Public Property Let rotate(value As Boolean)
    pRotate = value
End Property

Public Property Get fieldNames() As Boolean
    fieldNames = pFieldNames
End Property
Public Property Let fieldNames(value As Boolean)
    pFieldNames = value
End Property

Public Property Get headerTop() As Boolean
    headerTop = pHeader.top
End Property
Public Property Let headerTop(value As Boolean)
    pHeader.topAlign = "center"
    pHeader.top = value
End Property

Public Property Get headerTopAlign() As String
    headerTopAlign = pHeader.topAlign
End Property
Public Property Let headerTopAlign(value As String)
    value = LCase(value)
    
    Select Case value
        Case "center"
        Case "left"
        Case "right"
        Case Else
            value = "center"
    End Select

    pHeader.topAlign = value
End Property

Public Property Get headerLeft() As Boolean
    headerLeft = pHeader.left
End Property
Public Property Let headerLeft(value As Boolean)
    pHeader.leftAlign = "center"
    pHeader.left = value
End Property

Public Property Get headerLeftAlign() As String
    headerLeftAlign = pHeader.leftAlign
End Property
Public Property Let headerLeftAlign(value As String)
    value = LCase(value)
    
    Select Case value
        Case "center"
        Case "left"
        Case "right"
        Case Else
            value = "center"
    End Select

    pHeader.leftAlign = value
End Property

Public Property Get headerFontSize() As Integer
    headerFontSize = pHeader.fontSize
End Property
Public Property Let headerFontSize(value As Integer)
    pHeader.fontSize = value
End Property

Public Property Get border() As Boolean
    border = pBorder
End Property
Public Property Let border(value As Boolean)
    pBorder = value
End Property
'^ Get and Let.



Public Sub resetToDefaults()
    Call Class_Initialize
End Sub



Public Sub create(ByVal rs As ADODB.Recordset, rng As Range)
    Dim rngHead As Range: Set rngHead = rng
    Dim rngBody As Range: Set rngBody = rng.Offset(1, 0)
    Dim formatStr As String
    Dim i As Integer
    Dim j As Integer


    If pRotate Then
        With rs
            'Handle title.
            If pTitle <> "" Then    'If a title has been set.
                Range(rngHead, rngHead.Offset(0, .RecordCount)).Select    'Used for .HorizontalAlignment = xlCenterAcrossSelection.
                rngHead.value = pTitle
                Call styleTitle(Range(rngHead, rngHead.Offset(0, .RecordCount)))

                Set rngHead = rngHead.Offset(1, 0)  'Move rngHead to where the headers will start (1 row down).
            End If

            Set rngBody = rngHead.Offset(1, 0)  'Set rngBody to where the data will start (1 row down).
            
            'Handle headers and data.
            For i = 0 To rs.Fields.Count - 1    'Loop thru rows (recordset fields).
                rs.MoveFirst
                
                rngHead.Offset(i, 0).value = rs.Fields(i).Name
                If i = 0 And pHeader.top Then
                    Call styleHeader(rngHead.Offset(i, 0), pHeader.topAlign)    'Style top header.
                ElseIf pHeader.left Then
                    Call styleHeader(rngHead.Offset(i, 0), pHeader.leftAlign)   'Style left header.
                End If

                'Loop thru columns (recordset data).
                For j = 1 To rs.RecordCount
                    rngHead.Offset(i, j).value = rs(i)  'Data.

                    If i = 0 And pHeader.top Then
                        Call styleHeader(rngHead.Offset(i, j), pHeader.topAlign)    'Style top header.
                    End If

                    rs.MoveNext
                Next j

                'Format cells numbers.
                Select Case rs.Fields(i).Type
                    Case adNumeric
                        'Format same number of decimal places (NumericScale)
                        formatStr = "0."
                        For j = 1 To rs.Fields(i).NumericScale
                            formatStr = formatStr & "0"
                        Next j
                        Range(rngBody.Offset(i - 1, 0), rngBody.Offset(i - 1, rs.RecordCount)).NumberFormat = formatStr
                    Case adDBTimeStamp  'Datetime
                        formatStr = "m/d/yy h:mm;@"
                        Range(rngBody.Offset(i - 1, 0), rngBody.Offset(i - 1, rs.RecordCount)).NumberFormat = formatStr
                End Select
            Next
            
            If pBorder Then
                Call styleBorder(Range(rngHead, rngHead.Offset(rs.Fields.Count - 1, rs.RecordCount)))
            End If
        End With
    Else
        'Print headers.
        With rs
            If pTitle <> "" Then
                'Build the title bar.
                Range(rngHead, rngHead.Offset(0, .RecordCount)).Select    'Used for .HorizontalAlignment = xlCenterAcrossSelection.
                rngHead.value = pTitle
                Call styleTitle(Range(rngHead, rngHead.Offset(0, rs.Fields.Count - 1)))

                Set rngHead = rngHead.Offset(1, 0)
            End If

            'Build the top header row.
            If pFieldNames Then
                For i = 0 To rs.Fields.Count - 1
                    rngHead.Offset(0, i).value = rs.Fields(i).Name
                Next
                
                Set rngBody = rngHead.Offset(1, 0)
            Else
                Set rngBody = rngHead
            End If

            'Style headers.
            'Left header goes first to make sure the top header style can override the first column header.
            If pHeader.left Then
                Call styleHeader(Range(rngHead, rngHead.Offset(rs.RecordCount, 0)), pHeader.leftAlign)
            End If
            
            If pHeader.top Then
                Call styleHeader(Range(rngHead, rngHead.Offset(0, rs.Fields.Count - 1)), pHeader.topAlign)
            End If
        End With

        If rs.RecordCount > 0 Then  'If there's at least one record.
            If IsNull(rs(0)) = False Then   'If the first field in the first record isn't null.
                'Print data.
                rngBody.CopyFromRecordset rs

                'Border around data.
                If pBorder Then
                    Call styleBorder(Range(rngHead, rngHead.Offset(rs.RecordCount, rs.Fields.Count - 1)))
                End If

                'Format cell numbers.
                For i = 0 To rs.Fields.Count - 1
                    Select Case rs.Fields(i).Type
                        Case adNumeric
                            'Format same number of decimal places (NumericScale)
                            formatStr = "0."
                            For j = 1 To rs.Fields(i).NumericScale
                                formatStr = formatStr & "0"
                            Next j
                        Case adDBTimeStamp  'Datetime
                            formatStr = "m/d/yy h:mm;@"
                        Case Else
                            formatStr = "General"
                    End Select
                    
                    Range(rngBody.Offset(0, i), rngBody.Offset(rs.RecordCount - 1, i)).NumberFormat = formatStr
                Next
            End If
        End If
    End If
End Sub


Public Sub styleTitle(rng As Range)
    With rng
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0.6
        .Borders.LineStyle = xlContinuous
        .Borders.ThemeColor = 4
        .Borders.Weight = xlThin
        .Font.Bold = True
        .HorizontalAlignment = xlCenterAcrossSelection  'Looks identical to merging.
    End With
End Sub


Public Sub styleHeader(rng As Range, hAlign As String)
    
    With rng
        .HorizontalAlignment = xlCenter
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0.8
        .Borders.LineStyle = xlContinuous
        .Borders.ThemeColor = 4
        .Borders.Weight = xlThin
        .Font.Bold = True
        .Font.size = pHeader.fontSize

        Select Case LCase(hAlign)
            Case "center"
                .HorizontalAlignment = xlCenter
            Case "left"
                .HorizontalAlignment = xlLeft
            Case "right"
                .HorizontalAlignment = xlRight
        End Select
    End With
End Sub


Public Sub styleBorder(rng As Range)
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.ThemeColor = 4
        .Borders.Weight = xlThin
    End With
End Sub


