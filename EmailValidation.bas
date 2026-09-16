Attribute VB_Name = "EmailValidation"
Option Explicit

Sub ExportInvalidEmails()
    Dim lo As ListObject
    Dim wsOut As Worksheet
    Dim r As ListRow
    Dim i As Long, emailCol As Long, outRow As Long
    
    ' === Try to find a table actually named "ReportData" anywhere ===
    Set lo = FindTableByNameCI("ReportData")
    If lo Is Nothing Then
        ' If we didn't find it, tell the user what's available
        Dim msg As String: msg = "Table 'ReportData' not found." & vbCrLf & vbCrLf & _
                                 "Tables found in this workbook:" & vbCrLf & ListAllTableNames()
        MsgBox msg, vbExclamation
        Exit Sub
    End If
    
    If lo.DataBodyRange Is Nothing Then
        MsgBox "Table '" & lo.Name & "' has no data rows.", vbInformation
        Exit Sub
    End If
    
    ' === Locate column "Email" (case-insensitive) ===
    emailCol = 0
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, "Email", vbTextCompare) = 0 Then
            emailCol = i: Exit For
        End If
    Next
    If emailCol = 0 Then
        MsgBox "Column 'Email' not found in table '" & lo.Name & "'.", vbExclamation
        Exit Sub
    End If
    
    ' === Prepare output sheet ===
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("InvalidEmails")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsOut.Name = "InvalidEmails"
    Else
            'Clear previous results only - preserve rows 1 and 2 for buttons/title
            wsOut.Rows("3:" & wsOut.Rows.Count).Clear
    End If
    
    ' === Copy header ===
    lo.HeaderRowRange.Copy wsOut.Range("A3")
    
    ' === Copy rows where Email1 is blank or fails validation ===
    outRow = 4
    
For Each r In lo.ListRows
    Dim e As String: e = Trim$(CStr(r.Range.Columns(emailCol).Value))
    ' – Blank values are excluded here because they are handled by the Missing Data check
    If Len(e) > 0 Then
        If Not IsValidEmail_Practical(e) Then
            r.Range.Copy wsOut.Cells(outRow, 1)
            outRow = outRow + 1
        End If
    End If
Next

    
    MsgBox (outRow - 4) & " invalid or missing Email1 row(s) copied to 'InvalidEmails'." & vbCrLf & _
           "Source table: " & lo.Name & " (" & lo.Parent.Name & ")", vbInformation
End Sub

'==== Helpers ====

' Case-insensitive table finder across the whole workbook
Private Function FindTableByNameCI(ByVal tableName As String) As ListObject
    Dim ws As Worksheet, t As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each t In ws.ListObjects
            If StrComp(t.Name, tableName, vbTextCompare) = 0 Then
                Set FindTableByNameCI = t
                Exit Function
            End If
        Next t
    Next ws
End Function

' List all table names (for diagnostics)
Private Function ListAllTableNames() As String
    Dim ws As Worksheet, t As ListObject, s As String
    For Each ws In ThisWorkbook.Worksheets
        For Each t In ws.ListObjects
            s = s & "• " & t.Name & "  (sheet: " & ws.Name & ")" & vbCrLf
        Next t
    Next ws
    If Len(s) = 0 Then s = "(No tables found)"
    ListAllTableNames = s
End Function

'==========================
' Pragmatic email syntax validator WITHOUT RegExp
'==========================
Private Function IsValidEmail_Practical(ByVal raw As Variant) As Boolean
    Dim s As String, atPos As Long
    Dim localPart As String, domain As String
    Dim i As Long, ch As String
    Dim labels() As String, lbl As String
    Dim tld As String, penult As String
    Dim badTLDs As Variant
    
    ' ---- Basic guards ----
    If IsNull(raw) Then Exit Function
    s = Trim$(CStr(raw))
    If Len(s) = 0 Then Exit Function                  ' caller ignores blanks
    If InStr(s, " ") > 0 Then Exit Function           ' no spaces anywhere
    If InStr(s, "..") > 0 Then Exit Function          ' no consecutive dots
    
    ' exactly one @
    If (Len(s) - Len(Replace(s, "@", ""))) <> 1 Then Exit Function
    
    ' split local@domain
    atPos = InStr(1, s, "@")
    localPart = Left$(s, atPos - 1)
    domain = Mid$(s, atPos + 1)
    
    ' length guards
    If Len(s) > 254 Then Exit Function
    If Len(localPart) = 0 Or Len(localPart) > 64 Then Exit Function
    If Len(domain) = 0 Then Exit Function
    
    ' ---- Local part checks ----
    '  - allowed characters: A-Z, a-z, 0-9, . _ % + -
    '  - no leading/trailing dot; already blocked ".."
    If Left$(localPart, 1) = "." Or Right$(localPart, 1) = "." Then Exit Function
    
    For i = 1 To Len(localPart)
        ch = Mid$(localPart, i, 1)
        If Not IsAlphaNumeric(ch) Then
            If InStr("._%+-", ch) = 0 Then Exit Function
        End If
    Next i
    
    ' ---- Domain checks ----
    ' Must contain at least one dot
    If InStr(domain, ".") = 0 Then Exit Function
    
    labels = Split(domain, ".")
    ' Each label: only letters/digits/hyphen; no leading/trailing hyphen; not empty
    For i = LBound(labels) To UBound(labels)
        lbl = labels(i)
        If Len(lbl) = 0 Then Exit Function
        If Left$(lbl, 1) = "-" Or Right$(lbl, 1) = "-" Then Exit Function
        
        Dim j As Long
        For j = 1 To Len(lbl)
            ch = Mid$(lbl, j, 1)
            If Not IsAlphaNumeric(ch) And ch <> "-" Then Exit Function
        Next j
    Next i
    
    ' TLD (last label) must be >= 2 letters ONLY
    tld = LCase$(labels(UBound(labels)))
    If Len(tld) < 2 Then Exit Function
    For i = 1 To Len(tld)
        ch = Mid$(tld, i, 1)
        If Not (ch >= "a" And ch <= "z") Then Exit Function
    Next i
    
    ' Optional: previous label (penultimate) for .uk checks / typo guards
    If UBound(labels) >= 1 Then penult = LCase$(labels(UBound(labels) - 1))
    
    ' ---- Common typo guards ----
    badTLDs = Array("cuk", "con", "comm", "cmo")
    For i = LBound(badTLDs) To UBound(badTLDs)
        If tld = badTLDs(i) Then Exit Function
    Next i
    ' Specific UK typo: ".co.cuk"
    If penult = "co" And tld = "cuk" Then Exit Function
    
    ' Passed all checks
    IsValidEmail_Practical = True
End Function

' Helper: alphanumeric A-Z / a-z / 0-9
Private Function IsAlphaNumeric(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    ch = UCase$(ch)
    Dim a As Integer: a = Asc(ch)
    IsAlphaNumeric = (a >= 48 And a <= 57) Or (a >= 65 And a <= 90)
End Function


Sub TestEmailRegex()
    Debug.Print "simple@example.com", IsValidEmail_Practical("simple@example.com") ' True
    Debug.Print "user+tag@domain.co.uk", IsValidEmail_Practical("user+tag@domain.co.uk") ' True
    Debug.Print "val@dkbcrop.co.cuk", IsValidEmail_Practical("val@dkbcrop.co.cuk") ' False (bad TLD)
    Debug.Print "bad@domain", IsValidEmail_Practical("bad@domain") ' False (no dot)
    Debug.Print " first.last @example.com ", IsValidEmail_Practical(" first.last @example.com ") ' False (space)
    Debug.Print "", IsValidEmail_Practical("") ' False (blank)
End Sub

