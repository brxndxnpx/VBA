# AcroApp

An Acrobat class that is used to combine PDFs and convert files (e.g. images) to PDFs.

- Requires Adobe Acrobat DC to be installed on the user's machine.

---

## Properties

| Property    | Type      | Description                                 |
|-------------|-----------|---------------------------------------------|
| IsInstalled | `Boolean` | Indicates if Adobe Acrobat DC is installed. |

## Methods/Functions

|                                 | Description                                      |
|---------------------------------|--------------------------------------------------|
| [`PDFCombine`](#pdfcombine)     | Merges an array of file paths into a single PDF. |
| [`ConvertToPDF`](#converttopdf) | Converts a file to a PDF.                        |

---

### [`PDFCombine`](AcroApp.cls#L74)

Combines an array of file paths into a single PDF.

The files must be PDFs.

**Parameters**
- `FileName` `ByRef`
    - The finalized PDF file name. The extension isn't required.
- `Items` `ByRef`
    - The array of PDFs to merge. Consists of full file paths.
- `OutputDirectory` `ByRef`
    - The output directory to save the merged PDF to.

**Returns**
- A `String`; The combined PDF's file path.

---

### [`ConvertToPDF`](AcroApp.cls#L158)

Converts a file to a PDF.

**Parameters**
- `Path` `ByRef`
    - The full file path of the file to convert.

**Returns**
- A `String`; The converted file's PDF path.


---

## Usage

This sub routine will...
1. Create sample text files in the user's temp folder.
2. Convert those files to PDFs.
3. Merge the converted PDFs into a single PDF document.
4. Open the merged PDF.
    - You will likely see a warning prompt because it's following a hyperlink.
5. Delete the individual text and PDFs files.
6. Display a message to notify the user to delete the merged PDF in their temp folder.
7. Open their temp folder for them.

```vb
Private Sub Demo()
    Dim AC          As New AcroApp
    Dim filename_   As String
    Dim FS          As Object
    Dim filepath_   As String
    Dim writeText_  As String
    Dim i           As Long
    Dim textStream  As Object
    Dim temp_       As String
    Dim mergedPDF_  As String
    Dim files_      As Variant
    Dim pdfs_       As Variant
    
    ' Set the name of the file that's to be created
    'filename_ = "Sample Text File.txt"
    filename_ = "Sample Text File"
    
    ' Create a file system object to build the file path and create the text file
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    ' Get the user's temp folder
    temp_ = Environ$("TEMP")
    
    ' Create an array to house the newly created sample files.
    '   This will be used to convert the files to PDFs
    ReDim files_(0 To 5)
    
    ' Create an array to house the newly created sample PDFs.
    '   This will be used to merge the PDFs into a single PDF document
    ReDim pdfs_(0 To 5)
    
    For i = 0 To 5
        ' Set the name of the file that's to be created
        filename_ = "Sample Text File " & i & ".txt"
        
        ' Get the path of the new file to be created
        filepath_ = FS.BuildPath(temp_, filename_)
        
        ' Sample text to write
        writeText_ = "File " & i & vbNewLine & vbNewLine & _
            "Hello" & vbNewLine & "World!" & vbNewLine & vbNewLine & "How are you?"
        
        ' Create an empty text file in the user's temp folder
        Set textStream = FS.CreateTextFile(filepath_, True)
      
        ' Write text to the file
        textStream.WriteLine writeText_
        textStream.Close
        
        ' Append the text file to the files array (to delete later)
        files_(i) = filepath_
        
        ' Convert the text file to a PDF
        ' Append the path to the PDFs array (to delete later)
        pdfs_(i) = AC.ConvertToPDF(filepath_)
    Next
    
    mergedPDF_ = FS.BuildPath(temp_, "Sample PDF File.pdf")
    
    ' Merge the PDFs into 1 document
    AC.PDFCombine "Sample PDF File", pdfs_, temp_
    
    ' Open the PDF
    ActiveWorkbook.FollowHyperlink mergedPDF_
    
    ' Delete the files
    For i = LBound(files_) To UBound(files_)
        FS.DeleteFile files_(i)
        FS.DeleteFile pdfs_(i)
    Next
    
    ' Open the user's temp file so they can delete the Sample PDF File.pdf file.
    Shell "Explorer.exe" & " " & temp_, vbNormalFocus
    
    ' Prompt the user to delete the Sample PDF File.pdf file.
    MsgBox "The macro deleted the individual files but you will have to delete the 'Sample PDF File.pdf' in the temp folder." & vbNewLine & _
        "The macro couldn't delete it because it is opened."
    
End Sub
```
