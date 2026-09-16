Attribute VB_Name = "QualityReportEngine"
Option Explicit

Sub Fill_From_QualityReport_AlwaysOverwriteB3()
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim wsSrc As Worksheet                  ' Active/source sheet (destination)
    Dim wsQR As Worksheet                   ' "Quality Report" sheet (source)
    Dim key As String
    Dim f As Range                          ' header match in row 2
    Dim matchCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim v As Variant
    
    Set wsSrc = ActiveSheet
    Set wsQR = ThisWorkbook.Worksheets("Quality Report")
    
    ' 1) Get lookup key from A1 on active sheet
    key = Trim(CStr(wsSrc.Range("A1").Value))
    If Len(key) = 0 Then
        MsgBox "A1 is blank on the active sheet—nothing to match.", vbInformation
        GoTo CleanExit
    End If
    
    ' 2) Find exact (case-insensitive) header match in row 2 of Quality Report
    Set f = wsQR.Rows(2).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, _
                              SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If f Is Nothing Then
        MsgBox "No match found for '" & key & "' in 'Quality Report' row 2.", vbExclamation
        GoTo CleanExit
    End If
    matchCol = f.Column
    
    ' 3) Determine data range: row 3 to last used row in the matched column
    lastRow = wsQR.Cells(wsQR.Rows.Count, matchCol).End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "No data below row 2 in the matched column.", vbInformation
        GoTo CleanExit
    End If
    
    ' Optional: start from a clean slate
    ' wsSrc.Range("B3").ClearContents
    
    ' 4) Loop down the matched column from row 3
    For r = 3 To lastRow
        v = wsQR.Cells(r, matchCol).Value
        
        ' Only act when numeric and > 0
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then
                ' Always overwrite B3 with the corresponding Column B value from the same row
                wsSrc.Range("B3").Value = wsQR.Cells(r, 2).Value   ' Column 2 = "B"
                Copy_DQ_Metrics
            End If
        End If
    Next r
    
    ' Result: B3 ends up with the value from the LAST row in the loop that met (>0)

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Sub Copy_DQ_Metrics()
    Dim ws As Worksheet
    Dim lastRowA As Long, lastRowB As Long, lastRow As Long
    Dim lastColRow9 As Long
    Dim srcRange As Range, destTopLeft As Range, destRange As Range

    Set ws = ActiveSheet  ' Work on the active worksheet

    ' Find the last used row based on columns A and B
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastRow = Application.Max(lastRowA, lastRowB)

    ' If nothing below row 8, exit gracefully
    If lastRow < 9 Then
        MsgBox "No data found in A9:B9 and below.", vbInformation
        Exit Sub
    End If

    ' Define the source range: A9 to B(lastRow)
    Set srcRange = ws.Range("A9:B" & lastRow)

    ' Find the last used column in row 9, then set destination 2 column to the right
    lastColRow9 = ws.Cells(9, ws.Columns.Count).End(xlToLeft).Column
    Set destTopLeft = ws.Cells(9, lastColRow9 + 2)

    ' Build destination range with same size as the source
    Set destRange = destTopLeft.Resize(srcRange.Rows.Count, srcRange.Columns.Count)

    ' Paste VALUES only
    destRange.Value = srcRange.Value

    ' Optional: Clear any formats in destination (uncomment if you want values-only & no formats)
    ' destRange.ClearFormats

    'MsgBox "Copied " & srcRange.Address(0, 0) & " to " & destRange.Address(0, 0) & " as values.", vbInformation
End Sub


