Option Explicit
Option Private Module

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     Array helper functions.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub TestArrayExtensions()
    Dim Source As Variant
    
    Dim example_String   As String
    Dim example_Integer  As Long
    Dim example_Object   As Object
    
    example_String = "HELLO WORLD"
    example_Integer = 1090
    Set example_Object = CreateObject("Scripting.Dictionary")
    
    ArryAppend Source, example_String, example_Integer, example_Object
    ArryRemove Source, 2
    ArryDebug Source
End Sub

''' Summary:
'''     Appends items to an array.
''' Parameters:
'''     ByRef Source: The array to append.
'''     ByRef Items(): The items to append to the source.
Public Sub ArryAppend(ByRef Source As Variant, ParamArray Items() As Variant)
    Dim i As Long
    
    For i = LBound(Items) To UBound(Items)
        ' Resize the array to fit the new value/object
        ArryResize Source
        
        ' Set the new value/object
        If IsObject(Items(i)) Then Set Source(UBound(Source)) = Items(i) Else: Source(UBound(Source)) = Items(i)
    Next
End Sub

''' Summary:
'''     Resizes an array. Will instantiate a new array if the array is empty.
''' Parameters:
'''     ByRef Source: The array to resize.
'''     ByVal Optional AddedUBound: The number of additional upper bound dimensions to add to the source.
'''     ByVal Optional PreserveData: Whether or not to preserve the data in the source.
Public Sub ArryResize(ByRef Source As Variant, Optional ByVal AddedUBound As Long = 1, Optional ByVal PreserveData As Boolean = True)
    If IsEmpty(Source) Then
        ReDim Source(1 To AddedUBound)
    Else
        If UBound(Source) = -1 Then
            ReDim Source(1 To AddedUBound)
        Else
            If Not PreserveData Then
                ReDim Source(1 To UBound(Source) + AddedUBound)
            Else
                ReDim Preserve Source(1 To UBound(Source) + AddedUBound)
            End If
        End If
    End If
End Sub

''' Summary:
'''     Removes an item from an array and resizes it.
''' Parameters:
'''     ByRef Source: The array to reference.
'''     ByVal Index: The index to remove.
Public Sub ArryRemove(ByRef Source As Variant, ByVal Index As Long)
    Dim i As Long
    
    For i = Index To UBound(Source) - 1
        If IsObject(Source(i + 1)) Then
            Set Source(i) = Source(i + 1)
        Else
            Source(i) = Source(i + 1)
        End If
    Next
    
    ReDim Preserve Source(LBound(Source) To UBound(Source) - 1)
End Sub

''' Summary:
'''     Counts the items in an array.
''' Parameters:
'''     ByRef Source: The array to reference.
Public Function ArryCount(ByRef Source As Variant)
    If IsEmpty(Source) Then
        ArryCount = 0
    Else
        ArryCount = IIf(UBound(Source) = -1, 0, UBound(Source))
    End If
End Function

''' Summary:
'''     Debug.Prints the values of the items in the array along with it's data type.
''' Parameters:
'''     ByRef Source: The array to reference.
Public Sub ArryDebug(ByRef Source As Variant)
    If IsEmpty(Source) Then Debug.Print "Array Is Empty": Exit Sub
    
    Dim x As Long
    For x = LBound(Source) To UBound(Source)
        If IsObject(Source(x)) Then
            Debug.Print "Object", TypeName(Source(x))
        Else
            Debug.Print TypeName(Source(x)), Source(x)
        End If
    Next
    
End Sub




