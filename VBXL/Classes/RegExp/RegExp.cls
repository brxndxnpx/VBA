Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     A class for executing regular expressions utilizing the VBScript.RegExp object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Summary:
'''     Executes regular expressions against a string.
''' Parameters:
'''     ByVal Value As String: The text to parse.
'''     ByVal Pattern As String: The regular expression.
'''     ByVal IncludeQuotes As Boolean: Whether or not to include double quotes in matches.
'''     ByVal UseGlobal As Boolean: Whether or not to the global regex setting.
''' Returns:
'''     A Variant; Returns the results in an array.
Public Function Execute(ByVal Value As String, ByVal Pattern As String, _
    ByVal Optional IncludeQuotes As Boolean, ByVal Optional UseGlobal As Boolean = True)
    
  Dim i As Long
  Dim Groups, Group, Values

  With CreateObject("VBScript.RegExp")
    .Global = UseGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .TEST(Value) Then
      Set Groups = .Execute(Value)
      ReDim Values(1 To Groups.Count)
      For Each Group In Groups
        i = i + 1
        Values(i) = Group.Value
        If IncludeQuotes Then If Len(Group.submatches(0)) Or Group.Value = """""" Then Values(i) = Group.submatches(0)
      Next
    End If
  End With

  Execute = Values
End Function