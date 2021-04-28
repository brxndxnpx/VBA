Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     A class for preset Content-Type headers (read-only).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Summary:
'''     A static Content-type header for 'application/x-www-form-urlencoded'.
Public Property Get ApplicationUrlEncodedForm() As String: ApplicationUrlEncodedForm = "application/x-www-form-urlencoded": End Property

''' Summary:
'''     A static Content-type header for 'application/json'.
Public Property Get ApplicationJson() As String: ApplicationJson = "application/json": End Property

