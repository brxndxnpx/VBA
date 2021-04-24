# Classes

Late-bound objects created for easing VBA development.

I tried to decouple these objects as much as I could for individual classes or modules to be usable.
- There may still be dependencies since some classes work better with some modules.
- e.g. The `FileSystem` class works well with the `Environment` module for easy path access.


# AcroApp

An Acrobat class that is used to combine PDFs and convert files (e.g. images) to PDFs.

- Requires Adobe Acrobat DC to be installed on the user's machine.

---

## Properties

| Property    | Type      | Description                                 |
|-------------|-----------|---------------------------------------------|
| IsInstalled | `Boolean` | Indicates if Adobe Acrobat DC is installed. |

## Methods/Functions

| Method/Functions | Type     | Description                                      | Returns                                    |
|------------------|----------|--------------------------------------------------|--------------------------------------------|
| MergeDocuments   | Method   | Merges an array of file paths into a single PDF. |                                            |
| ConvertToPDF     | Function | Converts a file to a PDF.                        | `String`: The converted PDF's file path. |

- `MergeDocuments` can be turned into a function to also return the file path of the merged document.

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
    Dim AC As New AcroApp
    Dim filename_ As String
    Dim FS As Object
    Dim filepath_ As String
    Dim writeText_ As String
    Dim i As Long
    Dim textStream As Object
    Dim temp_ As String
    Dim mergedPDF_ As String
    Dim files_ As Variant
    Dim pdfs_ As Variant
    
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
    AC.MergeDocuments "Sample PDF File", pdfs_, temp_
    
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


# ColorInfo
A color class that contains metadata for RGB, hex, and Microsoft Office's integer (`Long` datatype) color code.

The color information will change in sync with the last updated property.
- For example, if the R value in RGB were to change, then the `HexCode` and `ColorCode` properties will also change to match the new color's value.
- If the `HexCode` were to change then the RGB values and `ColorCode` values will also change to match the new color's value.

This class/object is primarily used to convert colors, e.g. from RGB to hex.
- Of course, the methods/functions in the class object can also be set as functions in a module.

---

## Properties

| Property  | Type     | Description                         |
|-----------|----------|-------------------------------------|
| ColorCode | `Long`   | The Excel color code for the color. |
| HexCode   | `String` | The hex code for the color.         |
| R         | `Long`   | The R part of the RGB color         |
| G         | `Long`   | The G part of the RGB color         |
| B         | `Long`   | The B part of the RGB color         |


---

## Usage

### Setting The Color

You can set the color in 3 ways.
1. Hex code, e.g. `#FFFFFF` for white.
0. RGB values, e.g. `255`, `255`, `255` for white.
    - Unfortunately VBA doesn't support passing values through constructors.
    - This is set using the `ColorInfo.SetRGBValues()` method.
0. Microsoft Office color codes, e.g. `16777215` for white.

#### 1. Using Hex Code
```vb
Private Sub Demo()
    Dim Color As New ColorInfo
    
    ' Set the color by hex code
    Color.HexCode = "#002060" 'Dark Blue
End Sub
```

#### 2. Using RGB Values
```vb
Private Sub Demo()
    Dim Color As New ColorInfo
 
    ' Set the color by RGB values
    Color.SetRGBValues 255, 0, 0
End Sub
```

#### 3. Using Microsoft Office Color Code
```vb
Private Sub Demo()
    Dim Color As New ColorInfo

    ' Set the color by MS color code
    Color.ColorCode = 5288448 'Green
End Sub
```

---

### Getting The Color

You can get the color by accessing it's property once a color is set.

```vb
Private Sub Demo()
    Dim Color As New ColorInfo

    ' Set the color by RGB values
    Color.SetRGBValues 255, 0, 0

    Debug.Print Color.HexCode
End Sub
```


# Dictionary

A late-binding wrapper class for the `Scripting.Dictionary` object.

Allows using the `Scripting.Dictionary` object with intellisense without having to add a reference.
- Other users may not have the `Scripting` library added in their Visual Basic Editor. 
- This will alleviate that issue by using late-binding.

You can also store other objects in the dictionary.

---

## Properties

| Property | Description                                      |
|----------|--------------------------------------------------|
| Keys     | The keys in the dictionary.                      |
| Items    | The items in the dictionary.                     |
| Item     | Gets an item in the dictionary. `Default Member` |
| Count    | The number of items in the dictionary.           |


## Methods/Functions

| Method/Functions | Type     | Description                                    | Returns                            |
|------------------|----------|------------------------------------------------|------------------------------------|
| Add              | Method   | Adds an item to the dictionary.                |                                    |
| Replace          | Method   | Replaces an item at a key with a another item. |                                    |
| GetItem          | Function | Gets an item by it's key.                      | Type will vary. The item.          |
| GetKey           | Function | Gets a key by it's item.                       | Type will vary. The item's key.    |
| Exists           | Function | Checks if a key exists.                        | `Boolean`: True if the key exists. |
| Remove           | Method   | Removes an item.                               |                                    |
| RemoveAll        | Method   | Removes all items.                             |                                    |

---

## Notes

### Utilizing The Default Member
You should import this object into Excel by using the Import feature to be able to access the `Item` property as the default member.

Having the `Item` property as the default member allows you to access items in the collection without referencing the `Item` property directly.

You may use this often in arrays and collections without realizing it. Example:

```vb
' Immediate Window Example

' Using default member
?ThisWorkbook.Sheets(1).Name

' Referencing the Item member directly
?ThisWorkbook.Sheets.Item(1).Name
```

### Not Utilizing The Default Member
If you decide to copy and paste the code for this object into VBA directly, then exclude the following lines:

The attribute lines at the top of the file.

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
```

The two attribute lines in the `Item` getter.

```vb
Attribute Item.VB_Description = "Gets or sets the element at the specified index."
Attribute Item.VB_UserMemId = 0
```

---

## Usage

```vb
Public Sub Demo()
    Dim d as New Dictionary

    d.Add 1, "Hello"
    d.Add "3", "World"

    Debug.Print d(1)    ' Prints Hello
    Debug.Print d("3")  ' Prints World
End Sub
```

# FileSystem

A late-binding wrapper class for the `Scripting.FileSystemObject` object.

Allows using the `Scripting.FileSystemObject` object with intellisense without having to add a reference.
- Other users may not have the `Scripting` library added in their Visual Basic Editor. 
- This will alleviate that issue by using late-binding.

Recommended to be used in conjunction with the [Environment.bas](../../../VBXL/Modules/Environment/Environment.bas) module for easier file path access.

Recommended to be used in conjunction with the [DynamicLinkLibraries.bas](../../../VBXL/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas) module for downloading documents via URL.

Dependencies:
[DynamicLinkLibraries.bas](../../../VBXL/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas)
- The `DownloadDocument` method requires the [DynamicLinkLibraries.bas](../../../VBXL/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas) to be included in the project.
    - This method can be removed otherwise.


---

## Methods/Functions

| Method/Functions   | Type     | Description                                                                                               | Returns                                                                       |
|--------------------|----------|-----------------------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| FileExists         | Function | Checks if a file exists.                                                                                  | `Boolean`: True if exists.                                                    |
| FolderExists       | Function | Checks if a folder exists.                                                                                | `Boolean`: True if exists.                                                    |
| GetFilesInFolder   | Function | Gets all of the file paths in a directory. Has an optional parameter for retrieving items in sub-folders. | `Variant`: An array of objects of the `Scripting.FileSystemObject.File` type. |
| SelectFile         | Function | Selects a file and gets it's path.                                                                        | `String`: The object's path.                                                  |
| SelectFolder       | Function | Selects a folder and gets it's path.                                                                      | `String`: The object's path.                                                  |
| GetFileName        | Function | Gets a file's name.                                                                                       | `String`: The object's name.                                                  |
| GetFolderName      | Function | Gets a folder's name.                                                                                     | `String`: The object's name.                                                  |
| GetFile            | Function | Gets a file as an `Object`.                                                                               | `Object`: The file as a `Scripting.FileSystemObject.File` type.               |
| GetFolder          | Function | Gets a folder as an `Object`.                                                                             | `Object`: The folder as a `Scripting.FileSystemObject.Folder` type.           |
| GetExtension       | Function | Gets the file extension from a file path                                                                  | `String`: The file's extension.                                               |
| BuildPath          | Function | Combines paths. This is the same as the `Environment.PathCombine()` function.                             | `String`: The combined file path.                                             |
| AddFolder          | Function | Adds a folder to an **existing** folder.                                                                  | `String`: The new folder's path.                                              |
| CreateFolder       | Method   | Creates a folder.                                                                                         |                                                                               |
| DeleteFile         | Function | Deletes a file.                                                                                           | `Boolean`: True if the file was successfully deleted.                         |
| DeleteFolder       | Function | Deletes a folder.                                                                                         | `Boolean`: True if the folder was successfully deleted.                       |
| IsFolderEmpty      | Function | Checks if a folder is empty.                                                                              | `Boolean`: True if there aren't any files in the folder.                      |
| MoveFile           | Method   | Moves a file.                                                                                             |                                                                               |
| MoveFolder         | Method   | Moves a folder.                                                                                           |                                                                               |
| CreateTextFile     | Function | Creates a text file.                                                                                      | `Object`: A `TextStream` object.                                              |
| WriteToTextFile    | Function | Writes to an existing/instantiated text file object.                                                      |                                                                               |
| DownloadDocument * | Function | Downloads a document from a path or URL.                                                                  |                                                                               |
| GetSpecialFolder   | Function | Gets a special folder's path by it's index.                                                               |                                                                               |
| CopyFile           | Function | Copies a file.                                                                                            |                                                                               |
| CopyFolder         | Function | Copies a folder.                                                                                          |                                                                               |
| IsFileOpen         | Function | Checks if a file is opened/locked by another application.                                                 |                                                                               |

The `DownloadDocument` method requires the [DynamicLinkLibraries.bas](../../../VBXL/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas) to be included in the project.
- This method can be removed otherwise.

The `GetSpecialFolder` uses a numbered index.
- e.g. 0 = System Root; 1 = System Folder; 2 = Temp Folder. 
- The native `Environ()` or `Environ$()` can usually be used instead.

---

## Usage

This will get all files on the user's desktop (not files in sub-folders) and prints them to the immediate window.
- The `Desktop` function used below is referenced in [Environment.bas](../../../VBXL/Modules/Environment/Environment.bas).

With [Environment.bas](../../../VBXL/Modules/Environment/Environment.bas).

```vb
Private Sub Demo()
    Dim FS As New FileSystem
    Dim files_ As Variant
    Dim i As Long
    
    ' Gets the file names on the user's desktop
    files_ = FS.GetFilesInFolder(Environment.Desktop, False)
    
    For i = LBound(files_) To UBound(files_)
        Debug.Print files_(i).Name
    Next
End Sub
```

Without [Environment.bas](../../../VBXL/Modules/Environment/Environment.bas).

```vb
Private Sub Demo()
    Dim FS As New FileSystem
    Dim files_ As Variant
    Dim i As Long
    Dim desktop_ As String

    ' Checks if the user is using OneDrive
    '   If true, then use the Desktop folder in the user's OneDrive folder.
    '   Otherwise use the default windows Desktop folder.
    desktop_ = FS.BuildPath(IIf(Environ$("OneDrive") <> vbNullString, Environ$("OneDrive"), Environ$("UserProfile")), "Desktop")
    
    ' Gets the file names on the user's desktop
    files_ = FS.GetFilesInFolder(desktop_, False)
    
    For i = LBound(files_) To UBound(files_)
        Debug.Print files_(i).Name
    Next
End Sub
```

# StringBuilder

Inspired by the C# .NET `System.Text.StringBuilder` class, but simplified for VBA.

A string building class that utilizes the `Scripting.Dictionary` object.
- Each item in the dictionary is an individual line.
- The dictionary's key is auto incremented and managed by the `StringBuilder` object.

This class is useful for building lengthy complex or formatted text, like an XML document.

---

## Properties

| Property | Description                 |
|----------|-----------------------------|
| Lines    | The keys in the dictionary. |

## Methods/Functions

| Method/Functions | Description                                                               |
|------------------|---------------------------------------------------------------------------|
| Append           | Appends text to the current line.                                         |
| AppendLine       | Appends text to the next line.                                            |
| AppendArray      | Appends text from an array. Each dimension of the array will be appended. |
| ToString         | Returns the lines as a single string.                                     |
| ToArray          | Returns the lines as an array.                                            |
| PrintEachLine    | Prints out each line to the immediate window. Useful for debugging.       |

---

## Usage

```vb
Private Sub Demo()
    Dim s As New StringBuilder
    
    s.Append "Hello"
    s.Append " "
    s.Append "World"
    
    s.AppendLine "How are you?"

    Debug.Print s.ToString
    ' Prints:
    '   Hello World
    '   How are you?
    
    ' Used for debugging
    s.PrintEachLine
    ' Prints:
    '   Hello World
    '   How are you?
End Sub
```

```vb
Private Sub Demo()
    Dim s As New StringBuilder
    Dim title_ As String: title_ = "Check this out"
    Dim h1_ As String: h1_ = "HELLO WORLD!!!"
    Dim btnText_ As String: btnText_ = "Click Me!!!!"
    
    s.AppendLine "<!DOCTYPE html>"
    s.AppendLine "<html>"
    s.AppendLine vbTab & "<head>"
    s.AppendLine vbTab & vbTab & "<title>"
    s.Append title_
    s.Append "</title>"
    s.AppendLine vbTab & "</head>"
    s.AppendLine vbTab & "<body>"
    s.AppendLine vbTab & vbTab & "<h1>"
    s.Append h1_
    s.Append "</h1>"
    s.AppendLine vbTab & vbTab & "<button>"
    s.Append btnText_
    s.Append "</button>"
    s.AppendLine vbTab & "</body>"
    s.AppendLine "</html>"

    Debug.Print s.ToString
    ' Prints:
    '   <!DOCTYPE html>
    '   <html>
    '       <head>
    '           <title>Check this out</title>
    '       </head>
    '       <body>
    '           <h1>HELLO WORLD!!!</h1>
    '           <button>Click Me!!!!</button>
    '       </body>
    '   </html>

End Sub
```

