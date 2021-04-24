Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     An Acrobat class that is used to combine PDFs and convert files (e.g. images) to PDFs.
'''     Requires Adobe Acrobat DC to be installed on the user's machine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Enum AcrobatApplicationKey
    PDSaveIncremental = 0
    PDSaveFull = 1
    PDSaveCopy = 2
    PDSaveLinearized = 4
    PDSaveWithPSHeader = 8
    PDSaveBinaryOK = 16
    PDSaveCollectGarbage = 32
End Enum

Private Const AppPID_App As String = "AcroExch.App"
Private Const AppPID_AVDoc As String = "AcroExch.AVDoc"
Private Const AppPID_PDDoc As String = "AcroExch.PDDoc"

Private xAcrobatApp As Object
Private xAcrobatAVDoc As Object
Private xAcrobatPDDoc As Object

Private IsInstalled_ As Boolean

''' Summary:
'''     If Adobe Acrobat DC is installed
Public Property Get IsInstalled() As Boolean: IsInstalled = IsInstalled_: End Property

''' Summary:
'''     Initializes the class object
'''     Creates new Acrobat objects
'''     Checks if Adobe Acrobat DC is installed
Private Sub Class_Initialize()
    IsInstalled_ = True

    On Error GoTo ErrHandler_
    
    Set xAcrobatApp = CreateObject(AppPID_App)
    Set xAcrobatAVDoc = CreateObject(AppPID_AVDoc)
    Set xAcrobatPDDoc = CreateObject(AppPID_PDDoc)

    If True = False Then
ErrHandler_:
        IsInstalled_ = False
    End If
    
    On Error GoTo 0
    
    Set xAcrobatApp = Nothing
    Set xAcrobatAVDoc = Nothing
    Set xAcrobatPDDoc = Nothing
End Sub

''' Summary:
'''     Clears objects from memory
Private Sub Class_Terminate()
    Set xAcrobatApp = Nothing
    Set xAcrobatAVDoc = Nothing
    Set xAcrobatPDDoc = Nothing
End Sub

''' Summary:
'''     Merges an array of file paths into a single PDF
Public Sub MergeDocuments(ByRef finalizedName As String, ByRef documentsToMerge As Variant, ByRef saveToDirectory As String)
    '    If Right(saveToDirectory, 1) <> "\" Then saveToDirectory = saveToDirectory & "\"
    
    Dim i     As Long, numberOfPages As Long, numberOfNewPages As Long, docPath As String, documentParts() As Object, FS As New FileSystem
    
    Set xAcrobatApp = CreateObject(AppPID_App)
    Set xAcrobatAVDoc = CreateObject(AppPID_AVDoc)
    Set xAcrobatPDDoc = CreateObject(AppPID_PDDoc)
    
    If FS.GetExtension(finalizedName) = vbNullString Then finalizedName = finalizedName & ".pdf"
    
    ' If xAcrobatPDDoc.Create = False Then
    '     '        MsgBox "Did not create Combined PDF file.", vbCritical + vbOKOnly, "File Not Created"
    ' End If
    
    ReDim documentParts(1 To UBound(documentsToMerge))
        
    For i = 1 To UBound(documentsToMerge)
        'Check If PDF Exists
        docPath = Trim$(documentsToMerge(i))
        If Not FS.FileExists(docPath) Then GoTo NextDoc_
            
        'Opens PDFs
        Set documentParts(i) = CreateObject(AppPID_PDDoc)
        documentParts(i).Open docPath
            
        If i Then
            'Merge PDFs to AcroDocs
            numberOfNewPages = documentParts(i).GetNumPages()
                
            If Not xAcrobatPDDoc.InsertPages(numberOfPages - 1, documentParts(i), 0, numberOfNewPages, True) Then
                '                    MsgBox "Cannot insert pages of" & vbLf & saveToDirectory & documentParts(i), vbExclamation, "Canceled"
            End If
    
            'Gets Page # - PartDocs(i)
            numberOfPages = numberOfPages + numberOfNewPages
            documentParts(i).Close
            Set documentParts(i) = Nothing
        Else
            'Gets Page # - AcroDoc
            numberOfPages = xAcrobatPDDoc.GetNumPages()
        End If
            
NextDoc_:
    Next
     
    'Save Merged Doc To DestFile
    If i > UBound(documentsToMerge) Then
        If Not xAcrobatPDDoc.Save(PDSaveFull, FS.BuildPath(saveToDirectory, finalizedName)) Then
            '                MsgBox "Cannot save the resulting document" & vbLf & saveToDirectory & finalizedName, vbExclamation, "Canceled"
        End If
    End If
    
    'Errors
    If Err.Number Then
        MsgBox Err.Description, vbCritical, "Error #" & Err.Number
    ElseIf i > UBound(documentsToMerge) Then
        'MsgBox "The resulting file is created:" & vbLf & p & DestFile, vbInformation, "Done"
    End If
         
    If Not xAcrobatPDDoc Is Nothing Then xAcrobatPDDoc.Close
    xAcrobatApp.Exit
        
    Set xAcrobatPDDoc = Nothing
    Set xAcrobatApp = Nothing
End Sub

''' Summary:
'''     Converts a file to a PDF.
Public Function ConvertToPDF(ByRef Path As String) As String
    Dim objJSO As Object, boResult As Boolean, expFormat As String, nPath As String, FS As New FileSystem
    Dim ext   As String: ext = FS.GetExtension(Path, False)
   
    Set xAcrobatApp = CreateObject(AppPID_App)
    Set xAcrobatAVDoc = CreateObject(AppPID_AVDoc)
    '    Set xAcrobatPDDoc = CreateObject(AppPID_PDDoc)
    
    boResult = xAcrobatAVDoc.Open(Path, vbNullString)
    Set xAcrobatPDDoc = xAcrobatAVDoc.GetPDDoc
   
    '   Set the javascript object
    Set objJSO = xAcrobatPDDoc.GetJSObject
    
    '   Check conversion type
    Select Case LCase$(ext)
        Case "eps"
            expFormat = "com.adobe.acrobat.eps"
        Case "html", "htm"
            expFormat = "com.adobe.acrobat.html"
        Case "jpeg", "jpg", "jpe"
            expFormat = "com.adobe.acrobat.jpeg"
        Case "jpf", "jpx", "jp2", "j2k", "j2c", "jpc"
            expFormat = "com.adobe.acrobat.jp2k"
        Case "docx"
            expFormat = "com.adobe.acrobat.docx"
        Case "doc"
            expFormat = "com.adobe.acrobat.doc"
        Case "png"
            expFormat = "com.adobe.acrobat.png"
        Case "ps"
            expFormat = "com.adobe.acrobat.ps"
        Case "rft"
            expFormat = "com.adobe.acrobat.rft"
        Case "xlsx"
            expFormat = "com.adobe.acrobat.xlsx"
        Case "xls"
            expFormat = "com.adobe.acrobat.spreadsheet"
        Case "txt"
            expFormat = "com.adobe.acrobat.accesstext"
        Case "tiff", "tif"
            expFormat = "com.adobe.acrobat.tiff"
        Case "xml"
            expFormat = "com.adobe.acrobat.xml-1-00"
        Case Else
            expFormat = vbNullString
    End Select
    
    If expFormat <> vbNullString Then            'And Err.Number = 0 Then
        '   Set the path of the new file. Note that Adobe instead of xls uses xml files.
        '   That's why here the xls extension changes to xml.
        nPath = IIf(LCase$(ext) <> "xls", Replace(Path, LCase$(ext), "pdf"), Replace(Path, "xml", "pdf"))
        
        '   Save PDF file to the new format; Close the PDF file without saving the changes; Close the Acrobat application.
        boResult = objJSO.SaveAs(nPath): boResult = xAcrobatAVDoc.Close(True): boResult = xAcrobatApp.Exit
    Else
        '   Close the PDF file without saving the changes; Close the Acrobat application.
        boResult = xAcrobatAVDoc.Close(True): boResult = xAcrobatApp.Exit
    End If
       
    Set xAcrobatApp = Nothing
    Set xAcrobatAVDoc = Nothing
    Set xAcrobatPDDoc = Nothing

    ConvertToPDF = nPath
End Function
