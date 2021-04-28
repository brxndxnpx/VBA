Option Explicit
Option Private Module

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     DLL declarations/imports
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'   Declare Window Api To Download Files
'       URLDownloadToFile
'       VBA7 = 64bit

#If Win64 Then
    Public Declare PtrSafe Function URLDownloadToFile Lib "UrlMon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As LongPtr
#Else
    Public Declare Function URLDownloadToFile Lib "UrlMon" Alias "URLDownloadToFileA" ( ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If