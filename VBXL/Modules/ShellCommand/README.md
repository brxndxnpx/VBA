# ShellCommand

Basic shell commands.

---

## Methods/Functions

| Method/Functions   | Description                                                                  |
|--------------------|------------------------------------------------------------------------------|
| `StartApplication` | Starts up an application or script and passes optional arguments.            |
| `OpenFileExplorer` | Opens the file explorer. A path can be provided to open a specific location. |

---

## Usage

Opens the user's desktop folder in the file explorer.

```vb
Private Sub Demo()
    StartApplication "Explorer.exe", Environ$("UserProfile") & "\Desktop"

    OpenFileExplorer Environ$("UserProfile") & "\Desktop"
End Sub
```