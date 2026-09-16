Attribute VB_Name = "PostcodeValidation"
Option Explicit
'==========================
' One click: export invalid postcodes to "InvalidPostcodes"
'==========================
Sub ExportInvalidPostcodes()
    Dim lo As ListObject
    Dim wsOut As Worksheet
    Dim r As ListRow
    Dim postCode As String
    Dim i As Long, colIndex As Long
    Dim outRow As Long
    
    ' Get the ReportData table from the ReportData worksheet
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets("ReportData").ListObjects("ReportData")
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Table 'ReportData' not found on worksheet 'ReportData'.", vbExclamation
        Exit Sub
    End If
    If lo.DataBodyRange Is Nothing Then
        MsgBox "Table 'ReportData' has no data rows.", vbInformation
        Exit Sub
    End If
    
    ' Find the PostCode column (case-insensitive)
    colIndex = 0
    For i = 1 To lo.ListColumns.Count
        If LCase(lo.ListColumns(i).Name) = "postcode" Then
            colIndex = i
            Exit For
        End If
    Next i
    If colIndex = 0 Then
        MsgBox "Column 'PostCode' not found in 'ReportData'.", vbExclamation
        Exit Sub
    End If
    
    ' Create or clear output sheet
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("InvalidPostcodes")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add
        wsOut.Name = "InvalidPostcodes"
    Else
        'Clear previous results only - preserve rows 1 and 2 for buttons/title
        wsOut.Rows("3:" & wsOut.Rows.Count).Clear
    End If
    
    ' Copy header
    lo.HeaderRowRange.Copy Destination:=wsOut.Range("A3")
    
    ' Copy invalid rows
    outRow = 4
    For Each r In lo.ListRows
        postCode = CStr(r.Range.Cells(1, colIndex).Value)
        If Not IsValidUKPostcode(postCode) Then
            r.Range.Copy Destination:=wsOut.Cells(outRow, 1)
            outRow = outRow + 1
        End If
    Next r
    
    MsgBox (outRow - 4) & " invalid postcode row(s) copied to 'InvalidPostcodes'.", vbInformation
End Sub

'==========================
' Regex validator for GB/NI postcodes (incl. GIR 0AA)
' – Normalises input (trim, uppercase, collapse spaces)
' – Inserts missing space before last 3 chars (e.g., SW1A1AA -> SW1A 1AA)
' – Treats blank/Null as invalid (change if you want blanks allowed)
'==========================
Private Function IsValidUKPostcode(ByVal raw As Variant) As Boolean
    Dim s As String
    Dim re As Object
    
    If IsNull(raw) Then
        IsValidUKPostcode = False
        Exit Function
    End If
    
    s = Trim$(CStr(raw))
    If Len(s) = 0 Then
        IsValidUKPostcode = True
        Exit Function
    End If
    
    ' Normalise
    s = UCase$(s)
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    If Len(s) > 3 And InStr(s, " ") = 0 Then
        s = Left$(s, Len(s) - 3) & " " & Right$(s, 3)
    End If
    
    Set re = CreateObject("VBScript.RegExp")
    With re
        .IgnoreCase = True
        .Global = False
        ' GB/NI formats:
        '  - GIR 0AA
        '  - Outward: A9 | A9A | A99 | AA9 | AA9A | AA99 | A9A/AA9A with allowed letters
        '  - Inward:  9AA
        .Pattern = "^(GIR 0AA|(?:[A-PR-UWYZ][0-9]{1,2}" & _
                   "|[A-PR-UWYZ][A-HK-Y][0-9]{1,2}" & _
                   "|[A-PR-UWYZ][0-9][A-HJKPSTUW]" & _
                   "|[A-PR-UWYZ][A-HK-Y][0-9][ABEHMNPRVWXY]) ?" & _
                   "[0-9][ABD-HJLNP-UW-Z]{2})$"
    End With
    
    IsValidUKPostcode = re.Test(s)
End Function

