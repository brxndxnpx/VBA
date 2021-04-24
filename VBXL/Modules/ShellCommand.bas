Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     Execute shell commands
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Summary:
'''     Starts up an application or script and passes optional arguments
Public Sub StartApplication(ByRef xAppName As String, ByVal xArgs As String, Optional ByVal xWindowStyle As VbAppWinStyle = vbMinimizedFocus)
    If xArgs = vbNullString Then Shell xAppName, xWindowStyle Else: Shell xAppName & " " & xArgs, xWindowStyle
End Sub

''' Summary:
'''     Starts up the file explorer application
Public Sub StartFileExplorer(Optional ByVal xPath As String, Optional ByVal xWindowStyle As VbAppWinStyle = vbNormalFocus)
    Shell "Explorer.exe" & " " & xPath, xWindowStyle
End Sub

