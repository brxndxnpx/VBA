Option Explicit

Private Const AppPID As String = "Scripting.FileSystemObject"

Public Enum IOMode
    ForReading = 1
    ForWriting = 2
    ForAppending = 8
End Enum

Public Enum SpecialFolderConst
    WindowsFolder = 0
    SystemFolder = 1
    TemporaryFolder = 2
End Enum

Public Enum Tristate
    TristateMixed = -2
    TristateTrue = -1
    TristateFalse = 0
    TristateUseDefault = -2
End Enum

Public Enum StandardStreamTypes
    StdIn = 0
    StdOut = 1
    StdErr = 2
End Enum

Public Enum CompareMethod
    BinaryCompare = 0
    TextCompare = 1
    DatabaseCompare = 2
End Enum

Private Type TFileSystem
    FS As Object
End Type

Private this  As TFileSystem

Private Property Get FS()
    If this.FS Is Nothing Then Set this.FS = CreateObject(AppPID)
    Set FS = this.FS
End Property

''' Summary:
'''     Clears objects from memory
Private Sub Class_Terminate()
    Set this.FS = Nothing
End Sub

''' Summary:
'''     Checks if a file exists
Public Function FileExists(ByRef Path As String) As Boolean: FileExists = FS.FileExists(Path): End Function

''' Summary:
'''     Checks if a folder exists
Public Function FolderExists(ByRef Path As String) As Boolean: FolderExists = FS.FolderExists(Path): End Function

''' Summary:
'''     Gets all of the file paths in a directory
Public Function GetFilesInFolder(Optional ByRef FolderPath As String = vbNullString, Optional ByRef includeSubfolders As Boolean = True)
    Dim oFolder As Object
    '    , output As Variant

    '   Check if there was a given folder path
    If FolderPath = vbNullString Then
        FolderPath = SelectFolder()

        If VarType(FolderPath) <> vbString Then
            If FolderPath = False Then GetFilesInFolder = False: Exit Function
        Else
            If FolderPath = "False" Then GetFilesInFolder = False: Exit Function
        End If
    End If

    Set oFolder = FS.GetFolder(FolderPath)

    '    output = GetFolderContents(oFolder, output, includeSubfolders)
    '    GetFilesInFolder = output
    
    GetFilesInFolder = GetFolderContents(oFolder, GetFilesInFolder, includeSubfolders)
End Function

''' Summary:
'''     Gets the contents of a folder recursively
Private Function GetFolderContents(ByVal oFolder As Object, ByRef output As Variant, ByRef includeSubfolders As Boolean)
    Dim oSubFolder As Object, oFile As Object, i As Long
    
    If includeSubfolders = True Then
        For Each oSubFolder In oFolder.SubFolders
            GetFolderContents oSubFolder, output, includeSubfolders
        Next
    End If
    
    If IsEmpty(output) Then
        ReDim output(1 To oFolder.Files.Count): i = 1
    Else
        ReDim Preserve output(LBound(output) To UBound(output) + oFolder.Files.Count): i = UBound(output) + 1
    End If
    
    For Each oFile In oFolder.Files
        Set output(i) = oFile: i = i + 1
    Next

    GetFolderContents = output
End Function

''' Summary:
'''     Selects a file and gets it's path
Public Function SelectFile(Optional ByRef initialPath As String = vbNullString) As Variant: SelectFile = DisplayFileDialog(initialPath, msoFileDialogFilePicker): End Function

''' Summary:
'''     Selects a folder and gets it's path
Public Function SelectFolder(Optional ByRef initialPath As String = vbNullString) As Variant: SelectFolder = DisplayFileDialog(initialPath, msoFileDialogFolderPicker): End Function

''' Summary:
'''     Displays a file dialog
''' Returns:
'''     Returns false if canceled
Private Function DisplayFileDialog(ByVal initialPath As String, ByVal dialogType As MsoFileDialogType)
    Dim fDialog As FileDialog: Set fDialog = Application.FileDialog(dialogType)
    
    With fDialog
        .Title = "Select a File"
        .AllowMultiSelect = False
        If initialPath <> vbNullString Then .InitialFileName = initialPath
        If .Show <> -1 Then DisplayFileDialog = False Else: DisplayFileDialog = .SelectedItems(1)
    End With
End Function

''' Summary:
'''     Gets a file's name
Public Function GetFileName(ByRef xPath As String, ByVal xIncludeExtension As Boolean) As String
    Dim xFile As Object, xName As String, xFileExtension As String
    Set xFile = GetFile(xPath): xName = xFile.Name
    
    If Not xIncludeExtension Then
        xFileExtension = GetExtension(xPath, True)
        xName = Replace(xName, xFileExtension, vbNullString)
    End If
    
    GetFileName = xName
End Function

''' Summary:
'''     Gets a folder's name
Public Function GetFolderName(ByRef xPath As String) As String
    On Error GoTo NoFolder_

    Dim xFolder As Object, xName As String
    Set xFolder = GetFolder(xPath): xName = xFolder.Name
    GetFolderName = xName
    
    If True = False Then
NoFolder_:
        GetFolderName = Right$(xPath, Len(xPath) - InStrRev(xPath, "\"))
    End If
    
    On Error GoTo 0
End Function

''' Summary:
'''     Gets a file as an object
Public Function GetFile(Optional ByVal Path As String = vbNullString) As Variant
    Dim f     As Variant
    If Path = vbNullString Then
        f = SelectFile()
        If VarType(f) = vbBoolean Then GetFile = False: Exit Function
    Else
        f = Path
    End If
    
    Set GetFile = FS.GetFile(f)
End Function

''' Summary:
'''     Gets a folder as an object
Public Function GetFolder(Optional ByVal Path As String = vbNullString) As Variant
    Dim f     As Variant
    If Path = vbNullString Then
        f = SelectFolder()
        If VarType(f) = vbBoolean Then GetFolder = False
    Else
        f = Path
    End If

    Set GetFolder = FS.GetFolder(f)
End Function

''' Summary:
'''     Gets the file extension from a file path
Public Function GetExtension(ByRef Path As Variant, Optional ByRef IncludePeriod As Boolean = True) As String
    Dim ext As String: ext = FS.GetExtensionName(Path)
    If ext = vbNullString Then GetExtension = vbNullString Else: GetExtension = IIf(IncludePeriod, "." & ext, ext)
End Function

''' Summary
'''     Combines paths
Public Function BuildPath(ParamArray Paths() As Variant) As String
    Dim i     As Long, output As String
    For i = LBound(Paths) To UBound(Paths)
        If Paths(i) <> vbNullString Then output = FS.BuildPath(output, Paths(i))
    Next
    BuildPath = output
End Function

''' Summary:
'''     Adds a folder to an existing folder
''' Returns:
'''     Returns the new folder's path
Public Function AddFolder(ByRef ParentFolder As String, ByRef childFolder As String) As String
    Dim newFolder As String: newFolder = BuildPath(ParentFolder, childFolder)
    On Error Resume Next: MkDir newFolder: On Error GoTo 0
    AddFolder = newFolder
End Function

''' Summary:
'''     Creates a folder path
Public Sub CreateFolder(ByRef FolderName As String)
    On Error Resume Next: FS.CreateFolder FolderName: On Error GoTo 0
End Sub

''' Summary:
'''     Deletes a file
''' Returns:
'''     Returns true if the file was successfully deleted
Public Sub DeleteFile(ByRef Path As String, Optional ByRef Force As Boolean = False) 'As Boolean
    On Error Resume Next
    FS.DeleteFile GetFile(Path), Force
    On Error GoTo 0
End Sub

''' Summary:
'''     Deletes a folder
''' Returns:
'''     Returns true if the folder was successfully deleted
Public Sub DeleteFolder(ByRef Path As String, Optional ByRef Force As Boolean = False) 'As Boolean
    On Error Resume Next
    FS.DeleteFolder GetFolder(Path), Force
    On Error GoTo 0
End Sub

''' Summary:
'''     Checks if a folder is empty
Public Function IsFolderEmpty(ByRef Path As String) As Boolean
    If GetFolder(Path).Files.Count = 0 Then IsFolderEmpty = True Else: IsFolderEmpty = False
End Function

''' Summary:
'''     Moves a file
Public Sub MoveFile(ByRef Source As String, ByRef Destination As String): FS.MoveFile Source, Destination: End Sub

''' Summary:
'''     Moves a folder
Public Sub MoveFolder(ByRef Source As String, ByRef Destination As String): FS.MoveFolder Source, Destination: End Sub

''' Summary:
'''     Creates a text file
Public Function CreateTextFile(ByRef FilePath As String, Optional ByRef FileText As String, Optional ByRef Overwrite As Boolean = True, Optional ByRef Unicode As Boolean = False, Optional ByVal CloseAfter As Boolean = True) As Object
    Dim xTextFile As Object: Set xTextFile = FS.CreateTextFile(FilePath, Overwrite, Unicode)
    If FileText <> vbNullString Then xTextFile.WriteLine FileText
    If CloseAfter Then xTextFile.Close Else: Set CreateTextFile = xTextFile
End Function

''' Summary:
'''     Writes to an existing/instantiated text file object
Public Function WriteToTextFile(ByVal TextFile As Object, ByRef FileText As String) As Object
    TextFile.WriteLine FileText
    Set WriteToTextFile = TextFile
End Function

''' Summary:
'''     Downloads a document from a path or URL
Public Function DownloadDocument(ByRef DocumentName As String, ByRef DocumentPath As String, ByRef saveLocation As String)
    Dim targetLocation As String: targetLocation = BuildPath(saveLocation, DocumentName)
    
    If GetExtension(targetLocation) = vbNullString Then
        targetLocation = targetLocation & GetExtension(DocumentPath)
    End If

    '   Downloads/saves the PDF to a location
    URLDownloadToFileA 0, DocumentPath, targetLocation, 0, 0
End Function

''' Summary:
'''     Gets a special folder's path by it's index.
'''         e.g. 0 = System Root; 1 = System Folder; 2 = Temp Folder
'''         * Environ can usually be used instead.
Public Function GetSpecialFolder(ByRef xIndex As Long) As String
    GetSpecialFolder = FS.GetSpecialFolder(xIndex)
End Function

Public Sub CopyFile(ByRef Source As String, ByRef Destination As String, Optional ByRef Overwrite As Boolean = True)
    FS.CopyFile Source, Destination, Overwrite
End Sub

Public Sub CopyFolder(ByRef Source As String, ByRef Destination As String, Optional ByRef Overwrite As Boolean = True)
    FS.CopyFolder Source, Destination, Overwrite
End Sub

Public Function IsFileOpen(ByVal FilePath As String) As Boolean




End Function

'Function BuildPath(Path As String, Name As String) As String
'    Member of Scripting.FileSystemObject
'    Generate a path from an existing path and a name

'Sub CopyFile(Source As String, Destination As String, [OverWriteFiles As Boolean = True])
'    Member of Scripting.FileSystemObject
'    Copy a file

'Sub CopyFolder(Source As String, Destination As String, [OverWriteFiles As Boolean = True])
'    Member of Scripting.FileSystemObject
'    Copy a folder

'Function CreateFolder(Path As String) As Folder
'    Member of Scripting.FileSystemObject
'    Create a folder


'
'Function CreateTextFile(FileName As String, [Overwrite As Boolean = True], [Unicode As Boolean = False]) As TextStream
'    Member of Scripting.FileSystemObject
'    Create a file as a TextStream
'
'
'Sub DeleteFile(FileSpec As String, [Force As Boolean = False])
'    Member of Scripting.FileSystemObject
'    Delete a file

'Sub DeleteFolder(FolderSpec As String, [Force As Boolean = False])
'    Member of Scripting.FileSystemObject
'    Delete a folder


'Function DriveExists(DriveSpec As String) As Boolean
'    Member of Scripting.FileSystemObject
'    Check if a drive or a share exists

'Property Drives As Drives
'    Read -only
'    Member of Scripting.FileSystemObject
'    Get drives collection


'Function GetAbsolutePathName(Path As String) As String
'    Member of Scripting.FileSystemObject
'    Return the canonical representation of the path


'Function GetBaseName(Path As String) As String
'    Member of Scripting.FileSystemObject
'    Return base name from a path


'Function GetDrive(DriveSpec As String) As Drive
'    Member of Scripting.FileSystemObject
'    Get drive or UNC share


'Function GetDriveName(Path As String) As String
'    Member of Scripting.FileSystemObject
'    Return drive from a path

'Function GetExtensionName(Path As String) As String
'    Member of Scripting.FileSystemObject
'    Return extension from path

'Function GetFile(FilePath As String) As File
'    Member of Scripting.FileSystemObject
'    Get file

'Function GetFileName(Path As String) As String
'    Member of Scripting.FileSystemObject
'    Return the file name from a path

'Function GetFileVersion(FileName As String) As String
'    Member of Scripting.FileSystemObject
'    Retrieve the file version of the specified file into a string

'Function GetFolder(FolderPath As String) As Folder
'    Member of Scripting.FileSystemObject
'    Get folder

'Function GetParentFolderName(Path As String) As String
'    Member of Scripting.FileSystemObject
'    Return path to the parent folder


'Function GetSpecialFolder(SpecialFolder As SpecialFolderConst) As Folder
'    Member of Scripting.FileSystemObject
'    Get location of various system folders


'Function GetStandardStream(StandardStreamType As StandardStreamTypes, [Unicode As Boolean = False]) As TextStream
'    Member of Scripting.FileSystemObject
'    Retrieve the standard input, output or error stream

'Function GetTempName() As String
'    Member of Scripting.FileSystemObject
'    Generate name that can be used to name a temporary file

'Sub MoveFile(Source As String, Destination As String)
'    Member of Scripting.FileSystemObject
'    Move a file

'Sub MoveFolder(Source As String, Destination As String)
'    Member of Scripting.FileSystemObject
'    Move a folder

'Function OpenTextFile(FileName As String, [IOMode As IOMode = ForReading], [Create As Boolean = False], [Format As Tristate = TristateFalse]) As TextStream
'    Member of Scripting.FileSystemObject
'    Open a file as a TextStream

