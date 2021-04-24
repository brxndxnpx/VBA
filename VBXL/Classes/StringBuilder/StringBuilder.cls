Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''   A string building class that utilizes the Scripting.Dictionary object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const AppPID As String = "Scripting.Dictionary"


Private Type TStringBuilder
    Values As Object
End Type

Private this  As TStringBuilder

Private Property Get Values()
    If this.Values Is Nothing Then Set this.Values = CreateObject(AppPID)
    Set Values = this.Values
End Property

Private Sub Class_Terminate()
    Set this.Values = Nothing
End Sub


''' Summary:
'''     The line text as an array
Public Property Get Lines(): Lines = Values.Count: End Property


''' Summary:
'''     Appends text to the current line
Public Sub Append(ByRef Text As String)
    ' Set the Lines in a variable to prevent unnessary re-calling
    Dim lines_ As Long: lines_ = Lines
    
    If lines_ = 0 Then
        Values.Add 1, Text
    Else
        Dim xItemValue As String: xItemValue = Values(lines_)
        xItemValue = xItemValue & Text
        
        Values.Remove lines_
        Values.Add lines_, xItemValue
    End If
End Sub

''' Summary:
'''     Appends text to the next line
Public Sub AppendLine(Optional ByRef Text As String = vbNullString)
    ' Set the Lines in a variable to prevent unnessary re-calling
    Dim lines_ As Long: lines_ = Lines
    
    If lines_ = 0 Then
        Values.Add 1, Text
    Else
        Values.Add lines_ + 1, Text
    End If
End Sub

''' Summary:
'''     Appends text from an array. Each dimension of the array will be appended.
Public Sub AppendArray(ByRef TextArray As Variant, Optional ByVal UseNewLinePerDimension As Boolean = True)
    Dim x     As Long, xText As String
    
    If IsEmpty(TextArray) Then Exit Sub
    If UBound(TextArray) = -1 Then Exit Sub
    
    For x = LBound(TextArray) To UBound(TextArray)
        xText = TextArray(x)
        If UseNewLinePerDimension Then AppendLine xText Else: Append xText
    Next x
End Sub

''' Summary:
'''     Returns the lines as a single string
Public Function ToString() As String
    Dim x     As Long
    Dim xItems As Variant
    Dim xOutput As String
    
    If Lines = 0 Then ToString = vbNullString: Exit Function
    
    xItems = Values.Items
    
    For x = LBound(xItems) To UBound(xItems)
        If xOutput = vbNullString Then
            xOutput = xItems(x)
        Else
            xOutput = xOutput & vbNewLine & xItems(x)
        End If
    Next x

    ToString = xOutput
End Function

''' Summary:
'''     Returns the lines as an array
Public Function ToArray()
    If Lines = 0 Then ToArray = Empty: Exit Function
    ToArray = Values.Items
End Function

''' Summary:
'''     Prints out each line to the immediate window.
'''     Useful for debugging.
Public Sub PrintEachLine()
    Dim x     As Long
    For x = 1 To Values.Count
        Debug.Print Values(x)
    Next x
End Sub


