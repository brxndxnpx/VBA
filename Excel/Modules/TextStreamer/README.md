# TextStreamer

Performs basic reads and writes text to files without creating a `TextStream` object.

Excel\Modules\Environment\Environment.bas

Recommended to be used in conjunction with the [Environment.bas](../../../Excel/Modules/Environment/Environment.bas) module for easier file path access.

Recommended to be used in conjunction with the [FileSystem.cls](../../../Excel/Classes/FileSystem/FileSystem.cls) class for easier file access.

Recommended to be used in conjunction with the [StringBuilder.cls](../../../Excel/Classes/StringBuilder/StringBuilder.cls) class for easier text readability.
- You can use the `StringBuilder.AppendArray()` to inject the text array into the `StringBuilder` object.

---

## Methods/Functions

| Method/Functions | Description                                               |
|------------------|-----------------------------------------------------------|
| ReadFile         | Reads the text from a file. Returns the text as an array. |
| WriteFile        | Writes to a file. Will overwrite the file.           |


---


## Usage

The below snippet will...
1. Create an empty text file on the users desktop.
2. Write text to the file.
3. Read the text from the file.
4. Delete the text file.

```vb
Private Sub Demo()
    Dim filename_ As String
    Dim FS As Object
    Dim desktop_ As String
    Dim filepath_ As String
    Dim writeText_ As String
    Dim readText_ As Variant ' The output of ReadFile() is an array
    Dim i As Long
    
    ' Set the name of the file that's to be created
    filename_ = "Sample Text File.txt"
    
    ' Sample text to write
    writeText_ = "Hello" & vbNewLine & "World!" & vbNewLine & vbNewLine & "How are you?"
    
    ' Create a file system object to build the file path and create the text file
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    ' The the user's desktop
    desktop_ = FS.BuildPath(IIf(Environ$("OneDrive") <> vbNullString, Environ$("OneDrive"), Environ$("UserProfile")), "Desktop")
    
    ' Get the path of the new file to be created
    filepath_ = FS.BuildPath(desktop_, filename_)
    
    ' Create an empty text file on the user's desktop
    FS.CreateTextFile filepath_, True
    
    ' Write to the text file
    WriteFile filepath_, writeText_
    
    ' Read the text file
    readText_ = ReadFile(filepath_)
    
    ' Write the results to the immediate window
    For i = LBound(readText_) To UBound(readText_)
        Debug.Print readText_(i)
    Next
    
    ' Delete the file
    FS.DeleteFile filepath_
End Sub
```