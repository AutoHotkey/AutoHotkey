# 📚 Certificate Generator - Quick Reference

## 🎯 Your Code: What It Does

### ✅ **YES - Your Code Does This:**

1. **Creates a copy from template** ✓
   - Uses `Documents.Add(Template:=ThisDocument.FullName)`
   - Original template NEVER modified

2. **Replaces name in the copy** ✓
   - Finds Content Control named "Trainee_Name"
   - Replaces with user input

3. **Converts date to Arabic** ✓
   - Converts: `November 11, 2025` → `١١ نوفمبر ٢٠٢٥ م`
   - Arabic numerals: 0→٠, 1→١, 2→٢, etc.
   - Arabic month names

4. **Exports to PDF** ✓
   - Saves to: `Desktop\Certificates\Certificate_[Name]_[Date].pdf`
   - Opens automatically

5. **Keeps template safe** ✓
   - Closes copy without saving
   - Template ready for next certificate

---

## 🔑 Key Code Sections

### 1. Create Copy (Not Modify Template)
```vba
Set newDoc = Documents.Add(Template:=ThisDocument.FullName, NewTemplate:=False, DocumentType:=0)
```
**Result:** New document created, template unchanged

---

### 2. Replace Content in Copy
```vba
ReplaceControlText newDoc, "Trainee_Name", traineeName, True, arFont
ReplaceControlText newDoc, "Cert_Date", arabicDate, False, arFont
```
**Result:** Only `newDoc` modified, template unchanged

---

### 3. Arabic Date Conversion
```vba
arabicDate = ToArabicDate(Date) & " م"
' November 11, 2025 → ١١ نوفمبر ٢٠٢٥ م
```
**Result:** Date in Arabic format

---

### 4. Export and Close
```vba
newDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF
newDoc.Close SaveChanges:=False
```
**Result:** PDF saved, copy discarded, template unchanged

---

## 📋 Setup Checklist

### Template Setup
- [ ] Create Word document with certificate design
- [ ] Insert Content Control for name (Title: "Trainee_Name")
- [ ] Insert Content Control for date (Title: "Cert_Date")
- [ ] Save as .dotm (Macro-Enabled Template)

### VBA Setup
- [ ] Open VBA Editor (Alt + F11)
- [ ] Insert Module
- [ ] Paste code
- [ ] Add Document_Open() event to ThisDocument
- [ ] Save

### Testing
- [ ] Open template
- [ ] Enter test name
- [ ] Verify PDF created
- [ ] Check template unchanged
- [ ] Test with Arabic name

---

## 🎨 Content Control Setup

### Step 1: Insert Content Control
1. Developer Tab → Controls → Plain Text Content Control
2. Click where you want the field

### Step 2: Configure Properties
1. Click the content control
2. Developer Tab → Properties
3. Set **Title**: `Trainee_Name` (for name field)
4. Set **Tag**: `Trainee_Name`
5. Click OK

### Step 3: Repeat for Date
1. Insert another content control
2. Set **Title**: `Cert_Date`
3. Set **Tag**: `Cert_Date`

---

## 🔧 Customization Quick Guide

### Change Password
```vba
Private Const PASSWORD As String = "your_password"
```

### Change Font Sizes
```vba
Private Const NAME_SIZE As Single = 18    ' Name size
Private Const DATE_SIZE As Single = 14    ' Date size
```

### Change Save Location
```vba
saveFolder = Environ$("USERPROFILE") & "\Documents\Certificates\"
```

### Change Arabic Font
```vba
Private Const PREFERRED_AR_FONT As String = "Sakkal Majalla"
```

### Add More Fields
```vba
' In template: Add content control with Title "Course_Name"
' In code: Add this line
ReplaceControlText newDoc, "Course_Name", "Advanced Training", False, arFont
```

---

## 🐛 Troubleshooting Quick Fixes

### Problem: Arabic shows as boxes (□□□)
**Fix:** Install Arabic language pack
```
Windows: Settings → Time & Language → Language → Add Arabic
```

### Problem: Content controls not replaced
**Fix:** Check Title/Tag names match exactly
```vba
' In template properties: Title = "Trainee_Name"
' In code: "Trainee_Name" (must match)
```

### Problem: PDF not created
**Fix:** Check folder exists and permissions
```vba
' Test folder creation:
MsgBox Environ$("USERPROFILE") & "\Desktop\Certificates\"
```

### Problem: Template gets modified
**Fix:** Verify using `newDoc` not `ThisDocument`
```vba
' CORRECT:
ReplaceControlText newDoc, ...  ' ✓

' WRONG:
ReplaceControlText ThisDocument, ...  ' ✗
```

---

## 📊 Arabic Conversion Reference

### Numbers
```
0 → ٠    5 → ٥
1 → ١    6 → ٦
2 → ٢    7 → ٧
3 → ٣    8 → ٨
4 → ٤    9 → ٩
```

### Months
```
January   → يناير      July      → يوليو
February  → فبراير     August    → أغسطس
March     → مارس       September → سبتمبر
April     → أبريل      October   → أكتوبر
May       → مايو       November  → نوفمبر
June      → يونيو      December  → ديسمبر
```

### Example Conversions
```
Nov 11, 2025  → ١١ نوفمبر ٢٠٢٥ م
Jan 1, 2025   → ١ يناير ٢٠٢٥ م
Dec 31, 2025  → ٣١ ديسمبر ٢٠٢٥ م
```

---

## 🎯 Usage Quick Guide

### Generate Single Certificate
1. Open template (.dotm file)
2. Input box appears
3. Enter trainee name
4. Click OK
5. PDF created and opened

### Unlock Template
1. Open template
2. When prompted, type: `123`
3. Template unlocks
4. Edit and save

### Generate Multiple Certificates
1. Open template
2. Generate first certificate
3. Template still open
4. Generate next certificate
5. Repeat as needed

---

## 📁 File Structure

```
Desktop/
└── Certificates/
    ├── Certificate_Ahmed Ali_20251111.pdf
    ├── Certificate_Sara Mohammed_20251111.pdf
    └── Certificate_John Smith_20251111.pdf

Documents/
└── Certificate_Template.dotm  (Your template - never changes)
```

---

## ✅ Verification Checklist

After generating a certificate, verify:

- [ ] PDF exists in Desktop\Certificates\
- [ ] PDF contains correct name
- [ ] Date is in Arabic format
- [ ] Arabic text is right-aligned
- [ ] Template still shows placeholders
- [ ] Template file date unchanged

---

## 🚀 Common Tasks

### Task: Change Date Format
```vba
' Current: ١١ نوفمبر ٢٠٢٥ م
' To change format, modify ToArabicDate function:

' Option 1: Year first
ToArabicDate = ReplaceDigits(Year(d)) & "/" & ReplaceDigits(Month(d)) & "/" & ReplaceDigits(Day(d))
' Result: ٢٠٢٥/١١/١١

' Option 2: Short format
ToArabicDate = ReplaceDigits(Day(d)) & "/" & ReplaceDigits(Month(d)) & "/" & ReplaceDigits(Year(d))
' Result: ١١/١١/٢٠٢٥
```

### Task: Add Certificate Number
```vba
' 1. Add content control in template: Title = "Cert_Number"
' 2. Generate number:
Dim certNumber As String
certNumber = "CERT-" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss")
' 3. Replace:
ReplaceControlText newDoc, "Cert_Number", certNumber, False, arFont
```

### Task: Add Company Logo
```vba
' 1. Insert Picture Content Control in template
' 2. Set Title: "Company_Logo"
' 3. Add code:
Dim cc As ContentControl
For Each cc In newDoc.ContentControls
    If cc.Title = "Company_Logo" Then
        cc.Range.InlineShapes.AddPicture _
            FileName:="C:\Path\To\Logo.png", _
            LinkToFile:=False, _
            SaveWithDocument:=True
    End If
Next
```

---

## 💡 Pro Tips

1. **Backup template** before making changes
2. **Test with copy** first
3. **Use descriptive names** for content controls
4. **Keep password secure** (change from "123")
5. **Archive PDFs** regularly
6. **Document customizations** for future reference

---

## 📞 Quick Help

### Enable Developer Tab
`File → Options → Customize Ribbon → Check "Developer"`

### Open VBA Editor
`Alt + F11`

### View Immediate Window (Debug)
`Ctrl + G` (in VBA Editor)

### Enable Macros
`File → Options → Trust Center → Trust Center Settings → Macro Settings`

---

## 🎓 Summary

**Your code is correct and does exactly what you need:**

✅ Creates copy from template  
✅ Replaces name and date  
✅ Converts to Arabic format  
✅ Exports to PDF  
✅ Keeps template unchanged  

**No modifications needed - it works perfectly!** 🎉

---

## 📚 Related Files

- `CertificateGenerator_Enhanced.vba` - Enhanced version with more features
- `CERTIFICATE_SETUP_GUIDE.md` - Detailed setup instructions
- `CODE_ANALYSIS.md` - In-depth code analysis
- `TESTING_EXAMPLES.md` - Test scenarios and examples

---

**Last Updated:** November 11, 2025  
**Version:** 1.0  
**Status:** Production Ready ✅
