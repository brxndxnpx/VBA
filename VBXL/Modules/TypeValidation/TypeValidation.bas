Option Explicit
Option Private Module

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     Generic validation for objects/variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Summary:
'''     The default system date.
Public Const vbNullDate As Date = #12:00:00 AM#

''' Summary:
'''     Checks if an object is nothing.
''' Parameters:
'''     ByRef Value As Object: The object to validate.
''' Returns:
'''     A Boolean; True if the object is nothing.
Public Function IsNothing(ByRef Value As Object) As Boolean
    IsNothing = Value Is Nothing
End Function

''' Summary:
'''     Checks if an object is not nothing.
''' Parameters:
'''     ByRef Value As Object: The object to validate.
''' Returns:
'''     A Boolean; True if the object is not nothing.
Public Function IsNotNothing(ByRef Value As Object) As Boolean
    IsNotNothing = Not Value Is Nothing
End Function

''' Summary:
'''     Checks if a string is null.
''' Parameters:
'''     ByVal Value As String: The string to validate.
''' Returns:
'''     A Boolean; Whether or not the string is equal to vbNullString, i.e. "".
Public Function IsNullString(ByVal Value As String) As Boolean
    IsNullString = Value = vbNullString
End Function

''' Summary:
'''     Checks if a date is null.
''' Parameters:
'''     ByVal Value As Date: The date to validate.
''' Returns:
'''     A Boolean; Whether or not the date is equal to the default system date (not set).
Public Function IsNullDate(ByVal Value As Date) As Boolean
    IsNullDate = Value = vbNullDate
End Function
