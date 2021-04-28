VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     A array based class that is intended to make managing arrays easier by resizing whenever an item is added or removed.
'''     Can be used for creating complex nested arrays.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Summary:
'''     The items in the list.
Public Items  As Variant

''' Summary:
'''     An item at an index.
Public Property Get Item(ByRef Index As Variant) As Variant
Attribute Item.VB_Description = "Gets or sets the element at the specified index."
Attribute Item.VB_UserMemId = 0

    If Index > UBound(Items) Or Index < LBound(Items) Then Exit Property
    
    If IsObject(Items(Index)) Then
        Set Item = Items(Index)
    Else
        Item = Items(Index)
    End If
End Property

''' Summary
'''     The lower bound of the items array.
Public Property Get LowerBound() As Long
    LowerBound = IIf(IsEmpty(Items), -1, LBound(Items))
End Property

''' Summary
'''     The upper bound of the items array.
Public Property Get UpperBound() As Long
    UpperBound = IIf(IsEmpty(Items), -1, UBound(Items))
End Property

''' Summary
'''     The number of items in the list.
Public Property Get Count() As Long
    If IsEmpty(Items) Then 
        Count = 0 
    Else
        Count = IIf(LBound(Items) = 0, UBound(Items) + 1, UBound(Items))
    End If
End Property


''' Summary:
'''     Adds an item to the list.
''' Parameters:
'''     ParamArray Values() As Variant: The items to add.
Public Sub Add(ParamArray Values() As Variant)
    Dim i   As Long ' For iterating through the Values
    Dim x   As Long ' For iterating through the Items
    
    '   Resize the array and set x to the next item index
    If IsEmpty(Items) Then
        ReDim Items(UBound(Values))
        x = LBound(Items)
    Else
        x = UBound(Items) + 1
        ReDim Preserve Items(0 To UBound(Items) + IIf(LBound(Values) = 0, UBound(Values) + 1, UBound(Values)))
    End If

    '   Insert the new item(s)
    For i = LBound(Values) To UBound(Values)
        If IsObject(Values(i)) Then
            Set Items(x) = Values(i)
        Else
            Items(x) = Values(i)
        End If
        x = x + 1
    Next
End Sub

''' Summary:
'''     Adds an item to the list.
'''     Will nest data into an array if multiple parameters are passed.
''' Parameters:
'''     ParamArray Values() As Variant: The items to add to the next record.
Public Sub AddAsArray(ParamArray Values() As Variant)
    Dim ItemArray  As Variant
    Dim i          As Long
      
    ' Check if there's only 1 item in Values()
    If LBound(Values) = UBound(Values) Then
        If IsObject(Values(LBound(Values))) Then
            Set ItemArray = Values(LBound(Values))
        Else
            ItemArray = Values(LBound(Values))
        End If
        GoTo InsertItem
    End If
      
    ' If there are several items then create an array
    ReDim ItemArray(UBound(Values))
    For i = LBound(Values) To UBound(Values)
        If IsObject(Values(i)) Then Set ItemArray(i) = Values(i) Else: ItemArray(i) = Values(i)
    Next
      
InsertItem:
    '   Resize the array
    If IsEmpty(Items) Then
        ReDim Items(0)
    Else
        ReDim Preserve Items(0 To UBound(Items) + 1)
    End If

    '   Insert the new item
    If IsObject(ItemArray) Then
        Set Items(UBound(Items)) = ItemArray
    Else
        Items(UBound(Items)) = ItemArray
    End If
End Sub


''' Summary:
'''     Removes an item by an index.
''' Parameters:
'''     ByVal Index As Long: The index to remove.
Public Sub Remove(ByVal Index As Long)
    Dim i As Long
    
    ' Validation
    If Index > UBound(Items) Or Index < LBound(Items) Then
        Exit Sub
    End If

    ' Move the items in the array
    For i = Index To UBound(Items) - 1
        If IsObject(Items(i + 1)) Then
            Set Items(i) = Items(i + 1)
        Else
            Items(i) = Items(i + 1)
        End If
    Next
    
    ' Resize the items array
    ReDim Preserve Items(LBound(Items) To UBound(Items) - 1)
End Sub


''' Summary:
'''     Removes a range of items by an index range.
''' Parameters:
'''     ByVal Index As Long: The index to start at.
'''     ByVal NumberOfItems As Long: The number of items to remove.
'''         This includes the item at the Index
Public Sub RemoveRange(ByVal Index As Long, ByVal NumberOfItems As Long)
    Dim i As Long
    
    ' Validation
    If Index > UBound(Items) Or Index < LBound(Items) Then Exit Sub
    If NumberOfItems <= 0 Then Exit Sub
    
    ' If the number of items to remove exceeds the index of the last item
    '   Then remove all items after the index
    If Index + NumberOfItems > UBound(Items) Then
        ReDim Preserve Items(LBound(Items) To Index - 1)
        Exit Sub
    End If
    
    ' Move the items in the array
    For i = Index To UBound(Items) - NumberOfItems
        If IsObject(Items(i + NumberOfItems)) Then
            Set Items(i) = Items(i + NumberOfItems)
        Else
            Items(i) = Items(i + NumberOfItems)
        End If
    Next

    ' Resize the items array
    ReDim Preserve Items(LBound(Items) To UBound(Items) - NumberOfItems)
End Sub

''' Summary:
'''     Clears the items in the List
Public Sub Clear()
    Items = Empty
End Sub