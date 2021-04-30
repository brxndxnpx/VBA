Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     User defined functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''' Summary:
'''     Quotes a text using the character provided. Will use the " character by default.
''' Parameters:
'''     ByVal Source As String: The text to reverse.
'''     Optional Character As String: The character to use as the quote. Will use the " character by default.
''' Returns:
'''     A String; The text returned in a pair of quotes.
Public Function QUOTE(ByVal Source As String, Optional Character As String = """") As String
    If Len(Character) > 1 Then
        Dim Str     As New StringBuilder
        Dim Spaces  As Long
        Dim i       As Long
        
        Spaces = CHARCOUNT(Character, " ")
        
        Str.Append Left(Character, 1)
        If Spaces > 0 Then
            For i = 1 To Spaces: Str.Append " ": Next
            Str.Append Source
            For i = 1 To Spaces: Str.Append " ": Next
        Else
            Str.Append Source
        End If
        Str.Append Right(Character, 1)
        
        QUOTE = Str.ToString
    Else
        QUOTE = Character & Source & Character
    End If
End Function


''' Summary:
'''     Reverses text.
''' Parameters:
'''     ByVal Source As String: The text to reverse.
''' Returns:
'''     A String; The reversed text.
Public Function REVERSE(ByVal Source As String) As String
    REVERSE = StrReverse(Source)
End Function


''' Summary:
'''     Counts the number of characters in a string.
''' Parameters:
'''     ByVal Source As String: The text to examine.
'''     ByVal Character As String: The text to count.
''' Returns:
'''     A Long; The number of characters in the next.
Public Function CHARCOUNT(ByVal Source As String, ByVal Character As String) As Long
    CHARCOUNT = Len(Source) - Len(Replace(Source, Character, ""))
End Function

