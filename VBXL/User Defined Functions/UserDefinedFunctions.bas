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
        Dim Spaces  As String
        Dim i       As Long
        
        For i = 1 To CHARCOUNT(Character, " ")
            Spaces = Spaces & " "
        Next
        
        QUOTE = Left(Character, 1) & Spaces & Source & Spaces & Right(Character, 1)
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


''' Summary:
'''     Checks if an object or value is in a list.
''' Parameters:
'''     ByRef Source As Variant: The object or value to check.
'''     ParamArray Predicate() As Variant: The list to check if the object or value is contained in.
''' Returns:
'''     A Boolean; True if the object or value is contained in the list
Public Function ISIN(ByRef Source As Variant, ParamArray Predicate() As Variant) As Boolean
    Dim i As Long
    Dim SearchObjects As Boolean
    SearchObjects = IsObject(Source)
    
    For i = LBound(Predicate) To UBound(Predicate)
        If SearchObjects And IsObject(Predicate(i)) Then
            If Predicate(i) Is Source Then ISIN = True: Exit Function
        Else
            If Predicate(i) = Source Then ISIN = True: Exit Function
        End If
    Next
End Function

