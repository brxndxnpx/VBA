Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     Execute shell commands
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Summary:
'''     Starts up an application or script and passes optional arguments
Public Sub StartApplication(ByRef AppName As String, ByVal Args As String, Optional ByVal WindowStyle As VbAppWinStyle = vbMinimizedFocus)
    If Args = vbNullString Then 
        Shell AppName, WindowStyle 
    Else
        Shell AppName & " " & Args, WindowStyle
    End If
End Sub

''' Summary:
'''     Starts up the file explorer application
Public Sub OpenFileExplorer(Optional ByVal Path As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus)
    Shell "Explorer.exe" & " " & Path, WindowStyle
End Sub

