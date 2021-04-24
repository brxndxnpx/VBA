VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TMMAP_VBXL.Class Objects.General"
Option Explicit

''' Summary:
'''     The List's name to help identify it
Public Name   As String

''' Summary:
'''     The items in the List
Public Items  As Variant

''' Summary:
'''     The item at an index
'@DefaultMember
Public Property Get Item(ByRef xIndex As Variant) As Variant
Attribute Item.VB_Description = "Gets or sets the element at the specified index."
Attribute Item.VB_UserMemId = 0

    If xIndex > UBound(Items) Or xIndex < LBound(Items) Then Exit Property
    
    If IsObject(Items(xIndex)) Then
        Set Item = Items(xIndex)
    Else
        Item = Items(xIndex)
    End If
End Property

''' Summary
'''     Gets a count of the numbers in the list array
Public Property Get Count() As Long
    If IsEmpty(Items) Then Count = 0 Else: Count = UBound(Items)
End Property

''' Summary:
'''     Adds an item to the List
'''     Will nest data into an array if multiple parameters are passed
Public Sub Add(ParamArray Values() As Variant)
    Dim NewItem As Variant, VarOffset As Long, i As Long
      
    If LBound(Values) = UBound(Values) Then
        If IsObject(Values(LBound(Values))) Then Set NewItem = Values(LBound(Values)) Else: NewItem = Values(LBound(Values))
        GoTo InsertItem
    End If
      
    VarOffset = IIf(LBound(Values) = 0, 1, 0)
    ReDim NewItem(1 To UBound(Values) + VarOffset)
    For i = LBound(Values) To UBound(Values)
        If IsObject(Values(i)) Then Set NewItem(i + VarOffset) = Values(i) Else: NewItem(i + VarOffset) = Values(i)
    Next
      
InsertItem:
    If IsEmpty(Items) Then ReDim Items(1 To 1) Else: ReDim Preserve Items(1 To UBound(Items) + 1)
    If IsObject(NewItem) Then Set Items(UBound(Items)) = NewItem Else: Items(UBound(Items)) = NewItem
End Sub


''' Summary:
'''     Removes an item
Public Sub Remove(ByRef xItem As Variant)
    Dim x     As Long
    Dim i     As Long
    
    For x = LBound(Items) To UBound(Items)
        If IsObject(xItem) Then
            If xItem Is Items(x) Then
                ArryRemove Items, x
                Exit Sub
            End If
        Else
            If xItem = Items(x) Then
                ArryRemove Items, x
                Exit Sub
            End If
        End If
    Next
End Sub

''' Summary:
'''     Clears the items in the List
Public Sub Clear()
    Items = Empty
End Sub

''' Summary:
'''     Find's the first item that matches
Public Function Find(ByVal Value As Variant)
    Dim i As Long
    
    For i = 1 To UBound(Items)
        If Items(i) = Value Then Find = Value: Exit Function
    Next i
    
    Find = Null
End Function


'Public Function Exists(ByRef Value As Variant)
'    Dim i As Long
'
'    For i = 1 To UBound(Items)
'        If IsObject(Value) Then
'            If Items(i) Is Value Then Set Exists = True: Exit Function
'        Else
'            If Items(i) = Value Then Exists = True: Exit Function
'        End If
'    Next i
'
'    Exists = False
'End Function


'
'
'
'''' Summary:
''''     Filters an array by matching a column's value to a see if it matches a criteria
'Function Filter(criteriaVal As String, conditionCol As Long)
'    Dim output() As Variant
'
'    Dim xRow As Long, xCol As Long, Rows As Long, cols As Long
'
'
'    cols = UBound(arrayVar, 2)
'    Rows = 1
'
'    ReDim output(1 To cols, 1 To Rows)
'
'    For xRow = 1 To UBound(arrayVar, 1)
'        If arrayVar(xRow, conditionCol) = criteriaVal Then
'            ReDim Preserve output(1 To cols, 1 To Rows)
'            For xCol = 1 To UBound(arrayVar, 2)
'                output(xCol, Rows) = arrayVar(xRow, xCol)
'            Next
'            Rows = Rows + 1
'        End If
'    Next
'
'
'    FilterArray = TransposeArray(output())
'
'
'End Function
'
'
'''' Summary:
''''     Filters an array by matching a column's value to a see if it matches a criteria
'Function Transpose(arrayVar() As Variant)
'    Dim output() As Variant
'    Dim X, Y As Long
'
'    ReDim output(1 To UBound(arrayVar, 2), 1 To UBound(arrayVar, 1))
'
'    For X = 1 To UBound(arrayVar, 2)
'        For Y = 1 To UBound(arrayVar, 1)
'            output(X, Y) = arrayVar(Y, X)
'        Next
'    Next
'
'    TransposeArray = output()
'End Function
'
'
'







