Option Explicit

' ============================================
' ENHANCED CERTIFICATE GENERATOR FOR MS WORD
' ============================================
' This macro creates certificates from a template with Arabic text support
' Features:
' - Creates copy without modifying template
' - Arabic date and number formatting
' - Automatic font detection and fallback
' - PDF export with proper naming
' - Error handling and validation

Private Const PASSWORD As String = "123"
Private Const PREFERRED_AR_FONT As String = "Noto Naskh Arabic"
Private Const FALLBACK_AR_FONT As String = "Traditional Arabic"
Private Const NAME_SIZE As Single = 16
Private Const DATE_SIZE As Single = 12

' ============================================
' MAIN ENTRY POINT
' ============================================
Public Sub Certificate_OnOpen()
    Dim answer As String
    answer = InputBox("Enter trainee name (type 123 to unlock template):", "Generate Certificate")

    ' Handle cancel or empty input
    If StrPtr(answer) = 0 Or Trim(answer) = "" Then Exit Sub

    ' Check for unlock password
    If Trim(answer) = PASSWORD Then
        UnlockTemplate
        Exit Sub
    End If

    ' Generate certificate
    CreateCertificateFromTemplate Trim(answer)
End Sub

' ============================================
' UNLOCK TEMPLATE FOR EDITING
' ============================================
Private Sub UnlockTemplate()
    On Error Resume Next
    If ThisDocument.ProtectionType <> wdNoProtection Then
        ThisDocument.Unprotect Password:=PASSWORD
    End If
    On Error GoTo 0
    MsgBox "✅ Template unlocked for editing.", vbInformation, "Template Unlocked"
End Sub

' ============================================
' CREATE CERTIFICATE FROM TEMPLATE
' ============================================
Private Sub CreateCertificateFromTemplate(ByVal traineeName As String)
    Dim newDoc As Document
    Dim saveFolder As String, pdfPath As String
    Dim arabicDate As String, arFont As String
    Dim originalDoc As Document

    On Error GoTo FailSafe

    ' Store reference to template document
    Set originalDoc = ThisDocument

    ' === Pick Arabic font ===
    arFont = PickAvailableArabicFont(PREFERRED_AR_FONT, FALLBACK_AR_FONT)
    If Trim(arFont) = "" Then arFont = FALLBACK_AR_FONT

    ' === Create new document based on template ===
    ' This creates a COPY - original template remains unchanged
    Set newDoc = Documents.Add(Template:=originalDoc.FullName, NewTemplate:=False, DocumentType:=0)
    newDoc.Activate

    ' === Prepare save folder ===
    saveFolder = Environ$("USERPROFILE") & "\Desktop\Certificates\"
    If Dir(saveFolder, vbDirectory) = "" Then MkDir saveFolder

    ' === Generate Arabic date ===
    arabicDate = ToArabicDate(Date) & " م"

    ' === Replace placeholders in the NEW document (not template) ===
    ReplaceControlText newDoc, "Trainee_Name", traineeName, True, arFont
    ReplaceControlText newDoc, "Cert_Date", arabicDate, False, arFont

    ' === Additional content controls (if any) ===
    ' You can add more replacements here, for example:
    ' ReplaceControlText newDoc, "Course_Name", "Advanced Training", False, arFont
    ' ReplaceControlText newDoc, "Instructor_Name", "Dr. Ahmed", False, arFont

    ' === Generate PDF path ===
    pdfPath = saveFolder & "Certificate_" & SanitizeFileName(traineeName) & "_" & Format(Date, "yyyymmdd") & ".pdf"

    ' === Protect the NEW document (not template) ===
    newDoc.Protect Type:=wdAllowOnlyReading, NoReset:=True, Password:=PASSWORD

    ' === Export to PDF ===
    newDoc.ExportAsFixedFormat _
        OutputFileName:=pdfPath, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        CreateBookmarks:=wdExportCreateNoBookmarks

    ' === Close the NEW document without saving ===
    newDoc.Close SaveChanges:=False

    ' === Return to template ===
    originalDoc.Activate

    ' === Success message ===
    MsgBox "✅ Certificate created successfully!" & vbCrLf & vbCrLf & _
           "Name: " & traineeName & vbCrLf & _
           "Date: " & arabicDate & vbCrLf & vbCrLf & _
           "Saved to:" & vbCrLf & pdfPath, _
           vbInformation, "Certificate Generated"
    Exit Sub

FailSafe:
    On Error Resume Next
    If Not newDoc Is Nothing Then
        newDoc.Close SaveChanges:=False
    End If
    If Not originalDoc Is Nothing Then
        originalDoc.Activate
    End If
    MsgBox "⚠️ Error generating certificate!" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Error"
End Sub

' ============================================
' REPLACE CONTENT CONTROL TEXT
' ============================================
Private Sub ReplaceControlText(doc As Document, ctrlName As String, valueText As String, isName As Boolean, arFont As String)
    Dim cc As ContentControl
    Dim foundControl As Boolean
    
    foundControl = False
    
    For Each cc In doc.ContentControls
        ' Match by Title or Tag (case-insensitive)
        If LCase(Trim(cc.Title)) = LCase(Trim(ctrlName)) Or _
           LCase(Trim(cc.Tag)) = LCase(Trim(ctrlName)) Then
            
            foundControl = True
            
            ' Set the text based on control type
            Select Case cc.Type
                Case wdContentControlText, wdContentControlRichText
                    cc.Range.Text = valueText
                Case wdContentControlDropdownList, wdContentControlComboBox
                    ' For dropdowns, try to set the value
                    On Error Resume Next
                    cc.Range.Text = valueText
                    On Error GoTo 0
                Case Else
                    ' For other types, use placeholder
                    On Error Resume Next
                    cc.SetPlaceholderText , , valueText
                    On Error GoTo 0
            End Select
            
            ' Apply formatting
            ApplyFontToRange cc.Range, isName, arFont
        End If
    Next cc
    
    ' Debug: Alert if control not found
    If Not foundControl Then
        Debug.Print "Warning: Content control '" & ctrlName & "' not found in document"
    End If
End Sub

' ============================================
' PICK AVAILABLE ARABIC FONT
' ============================================
Private Function PickAvailableArabicFont(preferred As String, fallback As String) As String
    Dim f As Variant
    
    ' First, try preferred font
    For Each f In Application.FontNames
        If LCase(f) = LCase(preferred) Then
            PickAvailableArabicFont = preferred
            Exit Function
        End If
    Next
    
    ' Second, try any Arabic/Naskh font
    For Each f In Application.FontNames
        If InStr(LCase(f), "arabic") > 0 Or _
           InStr(LCase(f), "naskh") > 0 Or _
           InStr(LCase(f), "noto") > 0 Then
            PickAvailableArabicFont = f
            Exit Function
        End If
    Next
    
    ' Finally, use fallback
    PickAvailableArabicFont = fallback
End Function

' ============================================
' APPLY FONT FORMATTING TO RANGE
' ============================================
Private Sub ApplyFontToRange(rng As Range, isName As Boolean, arFont As String)
    On Error Resume Next
    
    With rng.Font
        ' Set both BiDi and regular font names
        .NameBi = arFont
        .Name = arFont
        .Size = IIf(isName, NAME_SIZE, DATE_SIZE)
        .Bold = isName
        .Italic = False
        .Underline = wdUnderlineNone
    End With
    
    ' Set language and alignment for RTL text
    rng.LanguageID = wdArabic
    rng.ParagraphFormat.Alignment = wdAlignParagraphRight
    
    ' Ensure BiDi is enabled
    rng.LanguageDetected = False
    
    On Error GoTo 0
End Sub

' ============================================
' CONVERT DATE TO ARABIC FORMAT
' ============================================
Private Function ToArabicDate(d As Date) As String
    Dim months As Variant
    Dim dayStr As String, yearStr As String
    
    ' Arabic month names
    months = Array("", "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", _
                   "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر")
    
    ' Convert day and year to Arabic numerals
    dayStr = ReplaceDigits(CStr(Day(d)))
    yearStr = ReplaceDigits(CStr(Year(d)))
    
    ' Format: Day Month Year
    ToArabicDate = dayStr & " " & months(Month(d)) & " " & yearStr
End Function

' ============================================
' REPLACE WESTERN DIGITS WITH ARABIC NUMERALS
' ============================================
Private Function ReplaceDigits(txt As String) As String
    Dim i As Integer
    Dim digits As Variant
    Dim result As String
    
    ' Arabic-Indic numerals (٠-٩)
    digits = Array("٠", "١", "٢", "٣", "٤", "٥", "٦", "٧", "٨", "٩")
    
    result = txt
    For i = 0 To 9
        result = Replace(result, CStr(i), digits(i))
    Next i
    
    ReplaceDigits = result
End Function

' ============================================
' SANITIZE FILENAME (REMOVE INVALID CHARACTERS)
' ============================================
Private Function SanitizeFileName(s As String) As String
    Dim bad As Variant
    Dim c As Variant
    Dim result As String
    
    ' Invalid filename characters
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    result = s
    For Each c In bad
        result = Replace(result, c, "-")
    Next c
    
    ' Remove leading/trailing spaces
    result = Trim(result)
    
    ' Limit length to avoid path issues (max 100 chars)
    If Len(result) > 100 Then
        result = Left(result, 100)
    End If
    
    SanitizeFileName = result
End Function

' ============================================
' OPTIONAL: BATCH CERTIFICATE GENERATION
' ============================================
Public Sub GenerateBatchCertificates()
    Dim namesRange As Range
    Dim cell As Range
    Dim traineeName As String
    Dim count As Integer
    
    ' Prompt user to select range of names
    On Error Resume Next
    Set namesRange = Application.InputBox("Select range of trainee names:", "Batch Generation", Type:=8)
    On Error GoTo 0
    
    If namesRange Is Nothing Then Exit Sub
    
    count = 0
    For Each cell In namesRange.Cells
        traineeName = Trim(cell.Value)
        If traineeName <> "" And traineeName <> PASSWORD Then
            CreateCertificateFromTemplate traineeName
            count = count + 1
        End If
    Next cell
    
    MsgBox "✅ Generated " & count & " certificates successfully!", vbInformation
End Sub
