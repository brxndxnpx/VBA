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
'''     A array based class that allows grouping data or objects into a single record in an array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Summary:
'''     The items in the List.
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
'''     Adds an item to the List
'''     Will nest data into an array if multiple parameters are passed
''' Parameters:
'''     ParamArray Values() As Variant: The items to add to the next record.
Public Sub Add(ParamArray Values() As Variant)
    Dim NewItem    As Variant
    Dim VarOffset  As Long
    Dim i          As Long
      
    If LBound(Values) = UBound(Values) Then
        If IsObject(Values(LBound(Values))) Then 
            Set NewItem = Values(LBound(Values)) 
        Else
            NewItem = Values(LBound(Values))
        End If
        GoTo InsertItem
    End If
      
    VarOffset = IIf(LBound(Values) = 0, 1, 0)
    ReDim NewItem(1 To UBound(Values) + VarOffset)
    For i = LBound(Values) To UBound(Values)
        If IsObject(Values(i)) Then Set NewItem(i + VarOffset) = Values(i) Else: NewItem(i + VarOffset) = Values(i)
    Next
      
InsertItem:
    If IsEmpty(Items) Then 
        ReDim Items(0) 
    Else
        ReDim Preserve Items(0 To UBound(Items) + 1)
    End If

    If IsObject(NewItem) Then 
        Set Items(UBound(Items)) = NewItem 
    Else
        Items(UBound(Items)) = NewItem
    End If
End Sub

''' Summary:
'''     Removes an item by an index.
''' Parameters:
'''     ByRef Index As Variant: The index to remove.
Public Sub Remove(ByRef Index As Long)
    Dim x     As Long
    Dim i     As Long
    
    If Index > UBound(Items) Or Index < LBound(Items) Then 
        Exit Sub
    End If

    For i = Index To UBound(Items) - 1
        If IsObject(Items(i + 1)) Then
            Set Items(i) = Items(i + 1)
        Else
            Items(i) = Items(i + 1)
        End If
    Next

    ReDim Preserve Items(LBound(Items) To UBound(Items) - 1)
End Sub


''' Summary:
'''     Clears the items in the List
Public Sub Clear()
    Items = Empty
End Sub