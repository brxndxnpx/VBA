Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     Methods and functions for the Excel.Workbook object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Summary:
'''     Checks if a Worksheet contains any errors, i.e. #DIV/0!, #REF!, etc
''' Parameters:
'''     ByRef WorkSh As Worksheet: The worksheet to examine.
''' Returns:
'''     A Boolean; Whether or not the sheet contains errors.
Public Function WorksheetHasErrors(ByRef WorkSh As Worksheet) As Boolean
    WorksheetHasErrors = Not WorkSh.UsedRange.Find("#", LookAt:=xlPart, LookIn:=xlValues) Is Nothing
End Function


''' Summary:
'''     Gets all of the ranges (cells) that has errors in a Worksheet, i.e. #DIV/0!, #REF!, etc
''' Parameters:
'''     ByRef WorkSh As Worksheet: The Worksheet to examine.
'''     Optional ErrorContainer As Variant: The error container. You can reuse this container to append errors to it.
''' Returns:
'''     A Variant; All the errors found in a worksheet.
Public Function WorksheetErrors(ByRef WorkSh As Worksheet, Optional ErrorContainer As Variant) As Variant
    Dim Search As Range
    Dim Origin As Range
    
    With WorkSh.UsedRange
        Set Search = .Find("#", LookAt:=xlPart, LookIn:=xlValues)
        If Not Search Is Nothing Then
            Set Origin = Search
            
            Do
                If IsEmpty(ErrorContainer) Or IsMissing(ErrorContainer) Then
                    ReDim ErrorContainer(1 To 1)
                Else
                    ReDim Preserve ErrorContainer(1 To UBound(ErrorContainer) + 1)
                End If
                
                Set ErrorContainer(UBound(ErrorContainer)) = Search
                Set Search = .FindNext(Search)
            Loop While Not Search Is Nothing And Search.Address <> Origin.Address
        End If
    End With

    WorksheetErrors = ErrorContainer
End Function


''' Summary:
'''     Gets all of the errors in a Workbook, i.e. #DIV/0!, #REF!, etc
''' Parameters:
'''     ByRef WorkSh As Worksheet: The worksheet to examine.
''' Returns:
'''     A Variant; All of the errors in the workbook.
Public Function WorkbookErrors(ByRef WorkBk As Workbook) As Variant
    Dim ErrorContainer As Variant
    Dim WorkSh As Worksheet
    
    For Each WorkSh In WorkBk.Sheets
        ErrorContainer = WorksheetErrors(WorkSh, ErrorContainer)
    Next
    
    WorkbookErrors = ErrorContainer
End Function


''' Summary:
'''     Unhides all sheets in a Workbook.
''' Parameters:
'''     Optional ByRef WorkBk As Workbook: The workbook to target. Will target the active workbook is no value is provided.
Public Sub UnhideAllSheets(Optional ByRef WorkBk As Workbook)
    Dim WorkSh As Worksheet

    If WorkBk Is Nothing Then Set WorkBk = ActiveWorkbook

    For Each WorkSh In WorkBk.Sheets
        WorkSh.Visible = xlSheetVisible
    Next
End Sub


''' Summary:
'''     Unhide Worksheet(s).
''' Parameters:
'''     ParamArray Items() As Variant: The Worksheet(s) to unhide.
Public Sub UnhideSheets(ParamArray Items() As Variant)
    Dim i       As Long
    Dim WorkSh  As Worksheet
    
    For i = LBound(Items) To UBound(Items)
        Set WorkSh = Items(i)
        WorkSh.Visible = xlSheetVisible
    Next
End Sub


''' Summary:
'''     Hide Worksheet(s).
''' Parameters:
'''     ParamArray Items() As Variant: The Worksheet(s) to hide.
Public Sub HideSheets(ParamArray Items() As Variant)
    Dim i       As Long
    Dim WorkSh  As Worksheet
    
    For i = LBound(Items) To UBound(Items)
        Set WorkSh = Items(i)
        WorkSh.Visible = xlSheetHidden
    Next
End Sub









