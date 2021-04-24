# FileSystem

A late-binding wrapper class for the `Scripting.FileSystemObject` object.

Allows using the `Scripting.FileSystemObject` object with intellisense without having to add a reference.
- Other users may not have the `Scripting` library added in their Visual Basic Editor. 
- This will alleviate that issue by using late-binding.

Recommended to be used in conjunction with the [Environment.bas](../../../Excel/Modules/Environment/Environment.bas) module for easier file path access.

Recommended to be used in conjunction with the [DynamicLinkLibraries.bas](../../../Excel/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas) module for downloading documents via URL.

Dependencies:
[DynamicLinkLibraries.bas](../../../Excel/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas)
- The `DownloadDocument` method requires the [DynamicLinkLibraries.bas](../../../Excel/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas) to be included in the project.
    - This method can be removed otherwise.


---

## Methods/Functions

| Method/Functions   | Description                                                                                                                                  |
|--------------------|----------------------------------------------------------------------------------------------------------------------------------------------|
| FileExists         | Checks if a file exists.                                                                                                                     |
| FolderExists       | Checks if a folder exists.                                                                                                                   |
| GetFilesInFolder   | Gets all of the file paths in a directory. Has an optional parameter for retrieving items in sub-folders.                                    |
| SelectFile         | Selects a file and gets it's path.                                                                                                           |
| SelectFolder       | Selects a folder and gets it's path.                                                                                                         |
| GetFileName        | Gets a file's name.                                                                                                                          |
| GetFolderName      | Gets a folder's name.                                                                                                                        |
| GetFile            | Gets a file as an `Object`.                                                                                                                  |
| GetFolder          | Gets a folder as an `Object`.                                                                                                                |
| GetExtension       | Gets the file extension from a file path                                                                                                     |
| BuildPath          | Combines paths. This is the same as the `Environment.PathCombine()` function.                                                                |
| AddFolder          | Adds a folder to an **existing** folder. Returns the new folder's path.                                                                      |
| CreateFolder       | Creates a folder.                                                                                                                            |
| DeleteFile         | Deletes a file. Returns true if the file was successfully deleted.                                                                           |
| DeleteFolder       | Deletes a folder. Returns true if the folder was successfully deleted.                                                                       |
| IsFolderEmpty      | Checks if a folder is empty.                                                                                                                 |
| MoveFile           | Moves a file.                                                                                                                                |
| MoveFolder         | Moves a folder.                                                                                                                              |
| CreateTextFile     | Creates a text file.                                                                                                                         |
| WriteToTextFile    | Writes to an existing/instantiated text file object.                                                                                         |
| DownloadDocument * | Downloads a document from a path or URL.                                                                                                     |
| GetSpecialFolder   | Gets a special folder's path by it's index. e.g. 0 = System Root; 1 = System Folder; 2 = Temp Folder. * Environ can usually be used instead. |
| CopyFile           | Copies a file.                                                                                                                               |
| CopyFolder         | Copies a folder.                                                                                                                             |
| IsFileOpen         | Checks if a file is opened/locked by another application.                                                                                    |

The `DownloadDocument` method requires the [DynamicLinkLibraries.bas](../../../Excel/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas) to be included in the project.
- This method can be removed otherwise.

---

## Usage

This will get all files on the user's desktop (not files in sub-folders) and prints them to the immediate window.
- The `Desktop` function used below is referenced in [Environment.bas](../../../Excel/Modules/Environment/Environment.bas).

With [Environment.bas](../../../Excel/Modules/Environment/Environment.bas).

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

Without [Environment.bas](../../../Excel/Modules/Environment/Environment.bas).

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