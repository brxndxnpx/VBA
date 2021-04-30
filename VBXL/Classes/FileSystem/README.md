# FileSystem

A late-binding wrapper class for the `Scripting.FileSystemObject` object.

Allows using the `Scripting.FileSystemObject` object with intellisense without having to add a reference.
- Other users may not have the `Scripting` library added in their Visual Basic Editor. 
- This will alleviate that issue by using late-binding.

Recommended to be used in conjunction with...
- [Environment.bas](/VBXL/Modules/Environment/Environment.bas) module for easier file path access.

---

## Methods/Functions

| Method/Functions | Type     | Description                                                                                               | Returns                                                      |
|------------------|----------|-----------------------------------------------------------------------------------------------------------|--------------------------------------------------------------|
| FileExists       | Function | Checks if a file exists.                                                                                  | `Boolean`: True if exists.                                   |
| FolderExists     | Function | Checks if a folder exists.                                                                                | `Boolean`: True if exists.                                   |
| GetFilesInFolder | Function | Gets all of the file paths in a directory. Has an optional parameter for retrieving items in sub-folders. | `Variant`: An array of objects of the `Scripting.File` type. |
| SelectFile       | Function | Selects a file and gets it's path.                                                                        | `String`: The object's path.                                 |
| SelectFolder     | Function | Selects a folder and gets it's path.                                                                      | `String`: The object's path.                                 |
| GetFileName      | Function | Gets a file's name.                                                                                       | `String`: The object's name.                                 |
| GetFolderName    | Function | Gets a folder's name.                                                                                     | `String`: The object's name.                                 |
| GetFile          | Function | Gets a file as an `Object`.                                                                               | `Object`: The file as a `Scripting.File` type.               |
| GetFolder        | Function | Gets a folder as an `Object`.                                                                             | `Object`: The folder as a `Scripting.Folder` type.           |
| GetExtension     | Function | Gets the file extension from a file path                                                                  | `String`: The file's extension.                              |
| BuildPath        | Function | Combines paths. This is the same as the `Environment.PathCombine()` function.                             | `String`: The combined file path.                            |
| AddFolder        | Function | Adds a folder to an **existing** folder.                                                                  | `String`: The new folder's path.                             |
| CreateFolder     | Method   | Creates a folder.                                                                                         |                                                              |
| DeleteFile       | Function | Deletes a file.                                                                                           | `Boolean`: True if the file was successfully deleted.        |
| DeleteFolder     | Function | Deletes a folder.                                                                                         | `Boolean`: True if the folder was successfully deleted.      |
| IsFolderEmpty    | Function | Checks if a folder is empty.                                                                              | `Boolean`: True if there aren't any files in the folder.     |
| MoveFile         | Method   | Moves a file.                                                                                             |                                                              |
| MoveFolder       | Method   | Moves a folder.                                                                                           |                                                              |
| CreateTextFile   | Function | Creates a text file.                                                                                      | `Object`: A `TextStream` object.                             |
| WriteToTextFile  | Function | Writes to an existing/instantiated text file object.                                                      |                                                              |
| DownloadDocument | Function | Downloads a document from a path or URL.                                                                  |                                                              |
| GetSpecialFolder | Function | Gets a special folder's path by it's index.                                                               |                                                              |
| CopyFile         | Function | Copies a file.                                                                                            |                                                              |
| CopyFolder       | Function | Copies a folder.                                                                                          |                                                              |
| IsFileOpen       | Function | Checks if a file is opened/locked by another application.                                                 |                                                              |


The `GetSpecialFolder` uses a numbered index.
- e.g. 0 = System Root; 1 = System Folder; 2 = Temp Folder. 
- The native `Environ()` or `Environ$()` can usually be used instead.

---



## Usage

This will get all files on the user's desktop (not files in sub-folders) and prints them to the immediate window.

Without [Environment.bas](/VBXL/Modules/Environment/Environment.bas).

```vb
Private Sub Demo()
    Dim FS           As New FileSystem
    Dim Files        As Variant
    Dim i            As Long
    Dim FolderPath   As String

    ' Checks if the user is using OneDrive
    '   If true, then use the Desktop folder in the user's OneDrive folder.
    '   Otherwise use the default windows Desktop folder.
    FolderPath = FS.BuildPath(IIf(Environ$("OneDrive") <> vbNullString, Environ$("OneDrive"), Environ$("UserProfile")), "Desktop")
    
    ' Gets the file names on the user's desktop
    Files = FS.GetFilesInFolder(FolderPath, False)
    
    For i = LBound(Files) To UBound(Files)
        Debug.Print Files(i).Name
    Next
End Sub
```

With [Environment.bas](/VBXL/Modules/Environment/Environment.bas).
- The `DesktopPath` function used below is referenced in [Environment.bas](/VBXL/Modules/Environment/Environment.bas).

```vb
Private Sub Demo()
    Dim FS      As New FileSystem
    Dim Files   As Variant
    Dim i       As Long
    
    ' Gets the file names on the user's desktop
    Files = FS.GetFilesInFolder(Environment.DesktopPath, False)
    
    For i = LBound(Files) To UBound(Files)
        Debug.Print Files(i).Name
    Next
End Sub
```
