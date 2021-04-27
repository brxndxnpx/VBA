# List

A array based class that allows grouping data or objects into a single record in an array.

Primarily used for creating complex nested arrays.


## Properties

| Property     | Type      | Description                         |
|--------------|-----------|-------------------------------------|
| `Items`      | `Variant` | The items in the List               |
| `Item`       | `Variant` | An item at an index                 |
| `LowerBound` | `Long`    | The lower bound of the items array. |
| `UpperBound` | `Long`    | The upper bound of the items array. |
| `Count`      | `Long`    | The number of items in the list.    |

## Methods & Functions

|                     | Type | Description                                                                               |
|---------------------|------|-------------------------------------------------------------------------------------------|
| [`Add`](#add)       |      | Adds an item to the List. Will nest data into an array if multiple parameters are passed. |
| [`Remove`](#remove) |      | Removes an item by an index.                                                              |
| [`Clear`](#clear)   |      | Clears the items in the List.                                                             |

---

### [`Add`](List.cls#L63)

Adds an item to the List. 

Will nest data into an array if multiple parameters are passed.

**Parameters**
- `Values()` `ByRef` `ParamArray` `Variant`
    - The items to add to the next record.
    - If multiple items are passed then they will be grouped into a new array and placed into NEXT record.


### [`Remove`](List.cls#L101)

Removes an item by an index.

**Parameters**
- `Index` `ByRef` `Long`
    - The index to remove.

### [`Clear`](List.cls#L123)

Summary

**Parameters**
- `Param1` `ByRef` `String`
    - Description
- `Param2` `ByVal` `String` [`Optional`]
    - Description


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
Attribute VB_Name = "List"
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
Private Sub Demo()
    Dim container  As New List
    Dim obj        As Worksheet
    Dim x          As Long
    Dim y          As Long
    
    Set obj = ActiveWorkbook.Sheets(1)
    
    ' Add a string, a number, an object, and an array to the list
    container.Add "Hello World"
    container.Add 5
    container.Add obj
    
    ' This will created a nested array
    container.Add "Hello", "World", "How", "Are", "You", "?" 
    container.Add "I am the last item"
    
    ' Print the items in the list to the immediate window
    Debug.Print "Items:", container.Count
    
    For x = container.LowerBound To container.UpperBound
        If IsObject(container(x)) Then
            ' If the item is an Object
            Debug.Print x, "Object", TypeName(container(x))
        Else
            
            If VarType(container(x)) >= vbArray Then
                ' If the item is an Array
                Debug.Print x, "Array"
                For y = LBound(container(x)) To UBound(container(x))
                    Debug.Print "", y, TypeName(container(x)(y)), container(x)(y)
                Next
            Else
                ' If the item is a primitive type (String, Integer, Long, etc)
                Debug.Print x, TypeName(container(x)), container(x)
            End If
        End If
    Next

    ' Prints:
    '   Items:         5 
    '    0            String        Hello World
    '    1            Integer        5 
    '    2            Object        Worksheet
    '    3            Array
    '                  1            String        Hello
    '                  2            String        World
    '                  3            String        How
    '                  4            String        Are
    '                  5            String        You
    '                  6            String        ?
    '    4            String        I am the last item

End Sub
```
