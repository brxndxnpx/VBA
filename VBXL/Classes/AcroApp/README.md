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

### [`PDFCombine`](AcroApp.cls#L76)

Combines an array of file paths into a single PDF.

The files must be PDFs.

**Parameters**
- `FileName` `ByVal`
    - The finalized PDF file name. The extension isn't required.
- `Items` `ByRef`
    - The array of PDFs to merge. Consists of full file paths.
- `OutputDirectory` `ByRef`
    - The output directory to save the merged PDF to.

**Returns**
- A `String`; The combined PDF's file path.

---

### [`ConvertToPDF`](AcroApp.cls#L162)

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
    Dim FS          As Object
    Dim basename_   As String
    Dim filename_   As String
    Dim filepath_   As String
    Dim text_       As String
    Dim stream_     As Object
    Dim folder_     As String
    Dim mergedPDF_  As String
    Dim txt_files_  As Variant
    Dim pdf_files_  As Variant
    Dim i           As Long
    
    ' Set the name of the file that's to be created
    basename_ = "Sample Example File"
    
    ' Create a file system object to build the file path and create the text file
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    ' Get the user's temp folder
    folder_ = Environ$("TEMP")
    
    ' Create an array to house the newly created sample files.
    '   This will be used to convert the files to PDFs
    ReDim txt_files_(0 To 5)
    
    ' Create an array to house the newly created sample PDFs.
    '   This will be used to merge the PDFs into a single PDF document
    ReDim pdf_files_(0 To 5)
    
    For i = 0 To 5
        ' Set the name of the file that's to be created
        filename_ = basename_ & " " & i & ".txt"
        
        ' Get the path of the new file to be created
        filepath_ = FS.BuildPath(folder_, filename_)
        
        ' Sample text to write
        text_ = "File " & i & vbNewLine & vbNewLine & _
            "Hello" & vbNewLine & "World!" & vbNewLine & vbNewLine & "How are you?"
        
        ' Create an empty text file in the user's temp folder
        Set stream_ = FS.CreateTextFile(filepath_, True)
      
        ' Write text to the file
        stream_.WriteLine text_
        stream_.Close
        
        ' Append the text file to the files array (to delete later)
        txt_files_(i) = filepath_
        
        ' Convert the text file to a PDF
        ' Append the path to the PDFs array (to delete later)
        pdf_files_(i) = AC.ConvertToPDF(filepath_)
    Next
    
    ' Merge the PDFs into 1 document
    mergedPDF_ = AC.PDFCombine(basename_, pdf_files_, folder_)
    
    ' Open the PDF
    ActiveWorkbook.FollowHyperlink mergedPDF_
    
    ' Delete the files
    For i = LBound(txt_files_) To UBound(txt_files_)
        FS.DeleteFile txt_files_(i)
        FS.DeleteFile pdf_files_(i)
    Next
    
    ' Open the user's temp file so they can delete the Sample PDF File.pdf file.
    Shell "Explorer.exe " & folder_, vbNormalFocus
    
    ' Prompt the user to delete the Sample PDF File.pdf file.
    MsgBox "The macro deleted the individual text and PDF files created, " & vbNewLine _
        & "but you will have to delete the '" & basename_ & "'.pdf' in the temp folder." & vbNewLine _
        & "The macro couldn't delete it because it is opened."
    
End Sub
```
