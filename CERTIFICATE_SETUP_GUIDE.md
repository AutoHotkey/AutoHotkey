# 📜 Certificate Generator - Setup & Usage Guide

## 🎯 Overview

This VBA macro generates personalized certificates from a Microsoft Word template with **Arabic text support**. It creates a copy of the template, fills in the trainee name and date in Arabic, then exports to PDF without modifying the original template.

---

## ✅ What Your Code Does Correctly

### 1. **Template Preservation**
```vba
Set newDoc = Documents.Add(Template:=ThisDocument.FullName, NewTemplate:=False, DocumentType:=0)
```
- Creates a **new document** based on the template
- **Original template remains unchanged**
- Each certificate is generated from a fresh copy

### 2. **Content Control Replacement**
```vba
ReplaceControlText newDoc, "Trainee_Name", traineeName, True, arFont
ReplaceControlText newDoc, "Cert_Date", arabicDate, False, arFont
```
- Finds ContentControls by **Title** or **Tag**
- Replaces placeholder text with actual values
- Works only on the **new document**, not the template

### 3. **Arabic Formatting**
```vba
' Converts: "11 November 2025" → "١١ نوفمبر ٢٠٢٥ م"
arabicDate = ToArabicDate(Date) & " م"
```
- Converts Western numerals (0-9) to Arabic-Indic numerals (٠-٩)
- Uses Arabic month names
- Applies right-to-left (RTL) formatting

### 4. **PDF Export**
```vba
pdfPath = saveFolder & "Certificate_" & SanitizeFileName(traineeName) & "_" & Format(Date, "yyyymmdd") & ".pdf"
newDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF
```
- Saves to: `Desktop\Certificates\Certificate_[Name]_[Date].pdf`
- Sanitizes filename (removes invalid characters)
- Opens PDF automatically after creation

---

## 🛠️ Setup Instructions

### Step 1: Prepare Your Word Template

1. **Create a Word document** with your certificate design
2. **Insert Content Controls** for dynamic fields:
   - Go to: `Developer Tab` → `Controls` → `Plain Text Content Control`
   - If Developer tab is hidden: `File` → `Options` → `Customize Ribbon` → Check "Developer"

3. **Configure Content Controls**:
   
   **For Trainee Name:**
   - Insert a Plain Text Content Control where the name should appear
   - Click the control → `Properties`
   - Set **Title**: `Trainee_Name`
   - Set **Tag**: `Trainee_Name`
   
   **For Date:**
   - Insert another Plain Text Content Control for the date
   - Set **Title**: `Cert_Date`
   - Set **Tag**: `Cert_Date`

4. **Save as Template**:
   - `File` → `Save As`
   - Choose location
   - File type: **Word Macro-Enabled Template (.dotm)**
   - Example: `Certificate_Template.dotm`

### Step 2: Add the VBA Code

1. **Open VBA Editor**: Press `Alt + F11`
2. **Insert Module**: `Insert` → `Module`
3. **Paste the code** (your original or enhanced version)
4. **Save**: `File` → `Save` or `Ctrl + S`

### Step 3: Set Auto-Run on Open

1. In VBA Editor, find `ThisDocument` in Project Explorer
2. Double-click `ThisDocument`
3. Add this code:

```vba
Private Sub Document_Open()
    Certificate_OnOpen
End Sub
```

4. **Save and close** VBA Editor

### Step 4: Enable Macros

1. `File` → `Options` → `Trust Center` → `Trust Center Settings`
2. `Macro Settings` → Select **"Enable all macros"** (for development)
3. Or: Save template in a **Trusted Location**

---

## 📋 How to Use

### Generate Single Certificate

1. **Open the template** (.dotm file)
2. **Input box appears** automatically
3. **Enter trainee name** (in English or Arabic)
4. **Click OK**
5. Certificate is generated and saved to `Desktop\Certificates\`

### Unlock Template for Editing

1. Open the template
2. When prompted, type: `123`
3. Template unlocks for editing
4. Make changes and save

---

## 🔧 Customization Options

### Change Password
```vba
Private Const PASSWORD As String = "your_password_here"
```

### Change Font Sizes
```vba
Private Const NAME_SIZE As Single = 18    ' Larger name
Private Const DATE_SIZE As Single = 14    ' Larger date
```

### Change Arabic Font
```vba
Private Const PREFERRED_AR_FONT As String = "Sakkal Majalla"
Private Const FALLBACK_AR_FONT As String = "Arial"
```

### Change Save Location
```vba
' Current: Desktop\Certificates\
saveFolder = Environ$("USERPROFILE") & "\Desktop\Certificates\"

' Alternative: Documents folder
saveFolder = Environ$("USERPROFILE") & "\Documents\Certificates\"

' Alternative: Specific path
saveFolder = "C:\My Certificates\"
```

### Add More Fields

To add more dynamic fields (e.g., course name, instructor):

1. **In template**: Add new Content Control with Title/Tag: `Course_Name`
2. **In code**: Add this line after the date replacement:
```vba
ReplaceControlText newDoc, "Course_Name", "Advanced Training", False, arFont
```

---

## 🐛 Troubleshooting

### Problem: Content Controls Not Replaced

**Solution:**
- Check Content Control **Title** and **Tag** match exactly: `Trainee_Name` and `Cert_Date`
- Ensure controls are **Plain Text** or **Rich Text** type
- Check for typos (case-insensitive but spelling matters)

### Problem: Arabic Text Shows as Boxes (□□□)

**Solution:**
- Install Arabic fonts: `Noto Naskh Arabic` or `Traditional Arabic`
- Windows: `Settings` → `Time & Language` → `Language` → Add Arabic
- The code has automatic fallback to available fonts

### Problem: PDF Not Created

**Solution:**
- Check folder permissions for `Desktop\Certificates\`
- Ensure no other program has the PDF open
- Check disk space

### Problem: Macro Security Warning

**Solution:**
- `File` → `Options` → `Trust Center` → `Trust Center Settings`
- `Trusted Locations` → Add your template folder
- Or: `Macro Settings` → Enable macros

### Problem: Date Not in Arabic

**Solution:**
- Check that `Cert_Date` Content Control exists
- Verify the `ToArabicDate` function is included
- Test with: `MsgBox ToArabicDate(Date)` in VBA

---

## 📊 Code Flow Diagram

```
┌─────────────────────────────────────┐
│  User Opens Template (.dotm)        │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│  Document_Open() triggers           │
│  Certificate_OnOpen()               │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│  InputBox: Enter Name or "123"      │
└──────────────┬──────────────────────┘
               │
       ┌───────┴───────┐
       │               │
       ▼               ▼
   "123"          Trainee Name
       │               │
       ▼               ▼
┌──────────┐    ┌──────────────────────┐
│ Unlock   │    │ Create New Document  │
│ Template │    │ from Template        │
└──────────┘    └──────────┬───────────┘
                           │
                           ▼
                ┌──────────────────────┐
                │ Replace Content:     │
                │ - Trainee_Name       │
                │ - Cert_Date (Arabic) │
                └──────────┬───────────┘
                           │
                           ▼
                ┌──────────────────────┐
                │ Apply Arabic Font    │
                │ & RTL Formatting     │
                └──────────┬───────────┘
                           │
                           ▼
                ┌──────────────────────┐
                │ Protect Document     │
                └──────────┬───────────┘
                           │
                           ▼
                ┌──────────────────────┐
                │ Export to PDF        │
                │ Desktop\Certificates\│
                └──────────┬───────────┘
                           │
                           ▼
                ┌──────────────────────┐
                │ Close New Document   │
                │ (Don't Save)         │
                └──────────┬───────────┘
                           │
                           ▼
                ┌──────────────────────┐
                │ Show Success Message │
                │ Open PDF             │
                └──────────────────────┘
```

---

## 🎨 Example Template Layout

```
╔═══════════════════════════════════════════════════════════╗
║                                                           ║
║              🎓 CERTIFICATE OF COMPLETION 🎓              ║
║                                                           ║
║                    This certifies that                    ║
║                                                           ║
║                  [Trainee_Name Control]                   ║
║                                                           ║
║          has successfully completed the training          ║
║                                                           ║
║                                                           ║
║                  Date: [Cert_Date Control]                ║
║                                                           ║
║                                                           ║
║  _________________              _________________         ║
║      Signature                      Seal                  ║
║                                                           ║
╚═══════════════════════════════════════════════════════════╝
```

---

## 🚀 Advanced Features (Enhanced Version)

The enhanced version includes:

### 1. **Better Error Handling**
- Returns to template if error occurs
- Detailed error messages
- Prevents template corruption

### 2. **Batch Generation**
```vba
Public Sub GenerateBatchCertificates()
```
- Generate multiple certificates from a list
- Select range of names in Excel or Word table
- Automatic processing

### 3. **Improved Font Detection**
- Checks for Noto fonts
- Better Arabic font fallback
- Debug messages for missing controls

### 4. **Enhanced PDF Options**
```vba
OptimizeFor:=wdExportOptimizeForPrint
```
- Better print quality
- Smaller file size options

---

## 📝 Testing Checklist

- [ ] Template opens without errors
- [ ] Input box appears on open
- [ ] Entering name generates certificate
- [ ] Name appears correctly in PDF
- [ ] Date is in Arabic format (١١ نوفمبر ٢٠٢٥ م)
- [ ] Arabic text is right-aligned
- [ ] PDF saves to Desktop\Certificates\
- [ ] PDF opens automatically
- [ ] Original template unchanged
- [ ] Password "123" unlocks template
- [ ] Special characters in names handled (sanitized)

---

## 💡 Tips & Best Practices

1. **Always test with a copy** of your template first
2. **Backup your template** before making changes
3. **Use descriptive Content Control names** (Title and Tag)
4. **Keep the password secure** (change from "123")
5. **Test with various name lengths** and special characters
6. **Verify Arabic font installation** on target computers
7. **Consider digital signatures** for official certificates
8. **Archive generated PDFs** regularly

---

## 📞 Common Questions

**Q: Can I use this with Word Online?**
A: No, VBA macros only work in desktop Word.

**Q: Will this work on Mac?**
A: Yes, but some features may need adjustment (file paths, fonts).

**Q: Can I add images dynamically?**
A: Yes, but requires additional code to replace picture content controls.

**Q: How do I add a QR code?**
A: Generate QR code image first, then insert via content control or VBA.

**Q: Can I email the PDF automatically?**
A: Yes, add Outlook automation code after PDF generation.

---

## 📚 Additional Resources

- [Microsoft Word VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/word)
- [Content Controls Guide](https://support.microsoft.com/en-us/office/content-controls-in-word)
- [Arabic Typography in Office](https://support.microsoft.com/en-us/office/enable-right-to-left-language-features)

---

## ✨ Summary

Your code **already works correctly** for:
- ✅ Creating copies without modifying template
- ✅ Replacing name and date placeholders
- ✅ Converting to Arabic text and numerals
- ✅ Exporting to PDF with proper naming
- ✅ Protecting generated documents

The enhanced version adds better error handling, batch processing, and more customization options.

**Your implementation is solid!** 🎉
