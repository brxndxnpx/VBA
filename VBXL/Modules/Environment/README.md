# Environment

Environment functions pertaining to the user and the user's machine.


---

## Methods/Functions

| Method/Functions     | Description                                                                    |
|----------------------|--------------------------------------------------------------------------------|
| PathCombine          | Combines paths by utilizing `Scripting.FileSystemObject`.                      |
| Desktop              | The user's desktop path.                                                       |
| Documents            | The user's documents path.                                                     |
| Downloads            | The user's downloads path.                                                     |
| UserProfile          | The user's profile path.                                                       |
| OneDrive             | The user's OneDrive.                                                           |
| Temp                 | The user's temporary files path.                                               |
| AppData              | The user's application data path.                                              |
| HomePath             | The user's home path.                                                          |
| SystemRoot           | The window's root path.                                                        |
| ProgramFiles32       | The user's 32x program files.                                                  |
| ProgramFiles64       | The user's 64x program files.                                                  |
| ExcelLibraryPath     | Excel's default library path.                                                  |
| ExcelUserLibraryPath | Excel's default user library path.                                             |
| CPUProcessor         | The CPU/Processor info - whether the user uses 32/63 bit.                      |
| ExcelStartupPath     | Excel's default startup path. This is where your PERSONAL.XLSB file is stored. |
| OfficeRibbonPath     | Microsoft office Ribbon path.                                                  |

---


## Usage

```vb
Private Sub Demo()
    Debug.Print PathCombine(Desktop, "A Folder On The Desktop")
    Debug.Print Desktop
    Debug.Print Documents
    Debug.Print Downloads
    Debug.Print UserProfile
    Debug.Print OneDrive
    Debug.Print Temp
    Debug.Print AppData
    Debug.Print HomePath
    Debug.Print SystemRoot
    Debug.Print ProgramFiles32
    Debug.Print ProgramFiles64
    Debug.Print ExcelLibraryPath
    Debug.Print ExcelUserLibraryPath
    Debug.Print CPUProcessor
    Debug.Print ExcelStartupPath
    Debug.Print OfficeRibbonPath
End Sub
```