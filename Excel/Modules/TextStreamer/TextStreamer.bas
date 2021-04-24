Option Explicit
Option Private Module

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     Reads and writes text to files.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Summary
'''     Reads the text from a file and stores the text in a dictionary
'''     Returns the text as an array
Public Function ReadFile(ByVal FilePath As String) As Variant
    Dim HFile As Long: HFile = FreeFile
    Dim Lines() As String
    
    If FilePath = vbNullString Then ReadFile = Empty: Exit Function
    
    Open FilePath For Input As #HFile
    Lines = Split(Input$(LOF(HFile), #HFile), vbNewLine)
    Close #HFile

    ReadFile = Lines
End Function

''' Summary
'''     Writes to a file. Will overwrite the file.
Public Sub WriteFile(ByVal FilePath As String, ByVal TextToWrite As String)
    Dim HFile As Long: HFile = FreeFile

    If TextToWrite = vbNullString Or FilePath = vbNullString Then Exit Sub
    
    Open FilePath For Output Access Write As HFile
    Print #HFile, TextToWrite
    Close HFile
End Sub
