# Dictionary

A late-binding wrapper class for the `Scripting.Dictionary` object.

Allows using the `Scripting.Dictionary` object with intellisense without having to add a reference.
- Other users may not have the `Scripting` library added in their Visual Basic Editor. 
- This will alleviate that issue by using late-binding.

You can also store other objects in the dictionary.

---

## Properties

| Property | Description                                      |
|----------|--------------------------------------------------|
| Keys     | The keys in the dictionary.                      |
| Items    | The items in the dictionary.                     |
| Item     | Gets an item in the dictionary. `Default Member` |
| Count    | The number of items in the dictionary.           |


## Methods & Functions

| Method/Functions | Type     | Description                                    | Returns                            |
|------------------|----------|------------------------------------------------|------------------------------------|
| Add              | Method   | Adds an item to the dictionary.                |                                    |
| Replace          | Method   | Replaces an item at a key with a another item. |                                    |
| GetItem          | Function | Gets an item by it's key.                      | Type will vary. The item.          |
| GetKey           | Function | Gets a key by it's item.                       | Type will vary. The item's key.    |
| Exists           | Function | Checks if a key exists.                        | `Boolean`: True if the key exists. |
| Remove           | Method   | Removes an item.                               |                                    |
| RemoveAll        | Method   | Removes all items.                             |                                    |

---

## Notes

### Utilizing The Default Member
You should import this object into Excel by using the Import feature to be able to access the `Item` property as the default member.

Having the `Item` property as the default member allows you to access items in the collection without referencing the `Item` property directly.

You may use this often in arrays and collections without realizing it. Example:

```vb
' Immediate Window Example

' Using default member
?ThisWorkbook.Sheets(1).Name

' Referencing the Item member directly
?ThisWorkbook.Sheets.Item(1).Name
```

### Not Utilizing The Default Member
If you decide to copy and paste the code for this object into VBA directly, then exclude the following lines:

The attribute lines at the top of the file.

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
```

The two attribute lines in the `Item` getter.

```vb
Attribute Item.VB_Description = "Gets or sets the element at the specified index."
Attribute Item.VB_UserMemId = 0
```

---

## Usage

```vb
Public Sub Demo()
    Dim d As New Dictionary

    d.Add 1, "Hello"
    d.Add "3", "World"

    Debug.Print d(1)    ' Prints Hello
    Debug.Print d("3")  ' Prints World
End Sub
```