Option Explicit
Option Private Module

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     Array helper functions.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
'''     ByVal Optional AddedBounds: The number of additional upper bound dimensions to add to the source.
'''     ByVal Optional PreserveData: Whether or not to preserve the data in the source.
Public Sub ArryResize(ByRef Source As Variant, Optional ByVal AddedBounds As Long = 1, Optional ByVal PreserveData As Boolean = True)
    If IsEmpty(Source) Then
        ' Set the array size to the Option Base setting; 1 or 0.        
        ReDim Source(IIf(AddedBounds = 1, LBound(Array()), AddedBounds))
    Else
        If UBound(Source) = -1 Then
            ReDim Source(LBound(Source) To AddedBounds)
        Else
            If Not PreserveData Then
                ReDim Source(LBound(Source) To UBound(Source) + AddedBounds)
            Else
                ReDim Preserve Source(LBound(Source) To UBound(Source) + AddedBounds)
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
''' Returns:
'''     A Long; The number of items in the array.
Public Function ArryCount(ByRef Source As Variant)
    If IsEmpty(Source) Then
        ArryCount = 0
    Else
        If UBound(Source) = -1 Then
            ArryCount = 0
            Exit Function
        End If

        ArryCount = IIf(LBound(Source) = 0, UBound(Source) + 1, UBound(Source))
        ' ArryCount = IIf(UBound(Source) = -1, 0, UBound(Source))
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





