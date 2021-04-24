VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''   A late binding wrapper class for the Scripting.Dictionary object.
'''   Allows using the Scripting.Dictionary object with intellisense without having to add a reference.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const AppPID As String = "Scripting.Dictionary"

Private Type TDictionary
    Dictionary As Object
End Type

Private this  As TDictionary

Private Property Get Dictionary()
    If this.Dictionary Is Nothing Then Set this.Dictionary = CreateObject(AppPID)
    Set Dictionary = this.Dictionary
End Property

Private Sub Class_Terminate()
    Set this.Dictionary = Nothing
End Sub

''' Summary:
'''     The keys in the dictionary
Public Property Get Keys() As Variant: Keys = Dictionary.Keys: End Property

''' Summary:
'''     The items in the dictionary
Public Property Get Items() As Variant: Items = Dictionary.Items: End Property

''' Summary:
'''     Gets an item in the dictionary
Public Property Get Item(ByRef xKey As Variant) As Variant
Attribute Item.VB_Description = "Gets or sets the element at the specified index."
Attribute Item.VB_UserMemId = 0

    Dim xBackupOutput As Object
    If Not Exists(xKey) Then Set Item = xBackupOutput: Exit Property

    If IsObject(Dictionary(xKey)) Then
        Set Item = Dictionary(xKey)
    Else
        Item = Dictionary(xKey)
    End If
End Property

''' Summary:
'''     The number of items in the dictionary
Public Property Get Count() As Long: Count = IIf(Dictionary Is Nothing, 0, Dictionary.Count): End Property

''' Summary:
'''     Adds an item to the dictionary
Public Sub Add(ByRef xKey As Variant, ByRef xItem As Variant) ' As Dictionary
    If Exists(xKey) Then
        Replace xKey, xItem
    Else
        Dictionary.Add xKey, xItem
    End If
End Sub

''' Summary:
'''     Replaces an item at a key with a another item
Public Sub Replace(ByRef xKey As Variant, ByRef xItem As Variant)
    On Error Resume Next
    Dictionary.Remove xKey
    On Error GoTo 0
    Dictionary.Add xKey, xItem
End Sub

''' Summary:
'''     Gets an item by it's key
Public Function GetItem(ByRef Key As Variant) As Variant
    If Not Exists(Key) Then Exit Function

    If IsObject(Dictionary(Key)) Then
        Set GetItem = Dictionary(Key)
    Else
        GetItem = Dictionary(Key)
    End If
End Function

''' Summary:
'''     Gets a key by it's item
Public Function GetKey(ByRef Item As Variant) As Variant
    Dim x     As Long
    Dim itms  As Variant: itms = Items
    Dim kys   As Variant: kys = Keys

    For x = LBound(itms) To UBound(itms)
        If IsObject(Item) Then
            If itms(x) Is Item Then Set GetKey = kys(x): Exit Function
        Else
            If itms(x) = Item Then GetKey = kys(x): Exit Function
        End If
    Next x
End Function

''' Summary:
'''     Checks if a key exists
Public Function Exists(ByRef Key As Variant) As Boolean
    Exists = Dictionary.Exists(Key)
End Function

''' Summary:
'''     Removes an item
Public Sub Remove(ByRef Key As Variant)
    Dictionary.Remove Key
End Sub

''' Summary:
'''     Removes all items
Public Sub RemoveAll()
    Dictionary.RemoveAll
End Sub

