# Environment

Environment functions pertaining to the user and the user's machine.


---

## Methods/Functions

| Method/Functions     | Description                                                                             |
|----------------------|-----------------------------------------------------------------------------------------|
| PathCombine          | Combines paths by utilizing `Scripting.FileSystemObject`.                               |
| DesktopPath          | The user's desktop path.                                                                |
| DocumentsPath        | The user's documents path.                                                              |
| DownloadsPath        | The user's downloads path.                                                              |
| UserProfilePath      | The user's profile path.                                                                |
| OneDrivePath         | The user's OneDrive.                                                                    |
| TempPath             | The user's temporary files path.                                                        |
| AppDataPath          | The user's application data path.                                                       |
| HomePath             | The user's home path. This is the same as UserProfilePath but without the drive letter. |
| SystemRootPath       | The window's root path.                                                                 |
| ProgramFiles32Path   | The user's 32x program files.                                                           |
| ProgramFiles64Path   | The user's 64x program files.                                                           |
| ExcelLibraryPath     | Excel's default library path.                                                           |
| ExcelUserLibraryPath | Excel's default user library path.                                                      |
| ExcelStartupPath     | Excel's default startup path. This is where your PERSONAL.XLSB file is stored.          |
| OfficeRibbonPath     | Microsoft office Ribbon path.                                                           |
| CPUProcessorInfo     | The CPU/Processor info - whether the user uses 32/64 bit.                               |

---


## Usage

```vb
Private Sub Demo()
    Debug.Print PathCombine(DesktopPath, "A Folder On The Desktop")
    Debug.Print DesktopPath
    Debug.Print DocumentsPath
    Debug.Print DownloadsPath
    Debug.Print UserProfilePath
    Debug.Print OneDrivePath
    Debug.Print TempPath
    Debug.Print AppDataPath
    Debug.Print HomePath
    Debug.Print SystemRootPath
    Debug.Print ProgramFiles32Path
    Debug.Print ProgramFiles64Path
    Debug.Print ExcelLibraryPath
    Debug.Print ExcelUserLibraryPath
    Debug.Print ExcelStartupPath
    Debug.Print OfficeRibbonPath
    Debug.Print CPUProcessorInfo
End Sub
```