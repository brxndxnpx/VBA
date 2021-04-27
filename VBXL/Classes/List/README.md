# List

A array based class that is intended to make managing arrays easier by resizing whenever an item is added or removed.

Can be used for creating complex nested arrays.


## Properties

| Property     | Type      | Description                         |
|--------------|-----------|-------------------------------------|
| `Items`      | `Variant` | The items in the list.              |
| `Item`       | `Variant` | An item at an index.                |
| `LowerBound` | `Long`    | The lower bound of the items array. |
| `UpperBound` | `Long`    | The upper bound of the items array. |
| `Count`      | `Long`    | The number of items in the list.    |

## Methods & Functions

|                               | Type | Description                                                                               |
|-------------------------------|------|-------------------------------------------------------------------------------------------|
| [`Add`](#add)                 |      | Adds item(s) to the list. Will add each item as a new record.                             |
| [`AddAsArray`](#addasarray)   |      | Adds an item to the list. Will nest the items into an array if multiple items are passed. |
| [`Remove`](#remove)           |      | Removes an item by an index.                                                              |
| [`RemoveRange`](#removerange) |      | Removes a range of items by an index range.                                               |
| [`Clear`](#clear)             |      | Clears the items in the List.                                                             |

---

### [`Add`](List.cls#L64)

Adds item(s) to the list.
- Will add each item as a new record.

**Parameters**
- `Values()` `ByRef` `ParamArray` `Variant`
    - The item(s) to add.


### [`AddAsArray`](List.cls#L93)

Adds an item to the list. 
- Will nest the items into an array if multiple items are passed.

**Parameters**
- `Values()` `ByRef` `ParamArray` `Variant`
    - The item(s) to add to the next record.
    - If multiple items are passed then they will be grouped into a child array.


### [`Remove`](List.cls#L134)

Removes an item by an index.

**Parameters**
- `Index` `ByVal` `Long`
    - The index to remove.


### [`RemoveRange`](List.cls#L162)

Removes a range of items by an index range.

**Parameters**
- `Index` `ByVal` `Long`
    - The index to start at.
- `NumberOfItems` `ByVal` `Long`
    - The number of items to remove.
    - This includes the item at the Index


### [`Clear`](List.cls#L193)

Clears the list.

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



### Adding Items

```vb
Private Sub Demo()
    Dim container As New List
    Dim x As Long
    Dim y As Long
    
    ' Add a string, a number, an object
    container.Add "Hello World"
    container.Add 5
    container.Add ActiveWorkbook.Sheets(1)
    
    ' Add several records at once
    '   This will add each item as a new record
    container.Add "Hello", "World", "How", "Are", "You", "?"
    
    ' Add a array child
    '   This will created a nested array
    container.AddAsArray "Hello", "World", "How", "Are", "You", "Again", "?"
    
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
    '   Items:         11 
    '    0            String        Hello World
    '    1            Integer        5 
    '    2            Object        Worksheet
    '    3            String        Hello
    '    4            String        World
    '    5            String        How
    '    6            String        Are
    '    7            String        You
    '    8            String        ?
    '    9            Array
    '                  0            String        Hello
    '                  1            String        World
    '                  2            String        How
    '                  3            String        Are
    '                  4            String        You
    '                  5            String        Again
    '                  6            String        ?
    '    10           String        I am the last item    

End Sub
```

### Removing Items

I'll be using 2 functions/methods in the examples below to prevent writing the same code in each example.

The code is the exact same as the example above but split into 2 functions/methods.

`GenerateList()`: Generates a list with the same data as above.

```vb
' Instantiate a new List object and populate it with dummy data
Private Function GenerateList() As List
    Dim container As New List

    ' Add a string, a number, an object
    container.Add "Hello World"
    container.Add 5
    container.Add ActiveWorkbook.Sheets(1)
    
    ' Add several records at once
    '   This will add each item as a new record
    container.Add "Hello", "World", "How", "Are", "You", "?"
    
    ' Add a array child
    '   This will created a nested array
    container.AddAsArray "Hello", "How", "Are", "You", "Again", "?"
    
    container.Add "I am the last item"
    
    PrintListItems container
    Debug.Print ""
    
    Set GenerateList = container
End Function
```

`PrintListItems()`: Prints the items in the list just like the example above.
```vb
' Print the items in the list to the immediate window
Private Sub PrintListItems(ByRef container As List)
    Dim x As Long
    Dim y As Long
    
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
End Sub
```

---

This snippet removes the first and last items using the `Remove()` method.

```vb
Private Sub Demo()
    Dim container As List
    Set container = GenerateList
    
    ' Remove the first and last items
    container.Remove container.LowerBound
    container.Remove container.UpperBound
    
    PrintListItems container

    ' Prints:
    '   Items:         9 
    '    0            Integer        5 
    '    1            Object        Worksheet
    '    2            String        Hello
    '    3            String        World
    '    4            String        How
    '    5            String        Are
    '    6            String        You
    '    7            String        ?
    '    8            Array
    '                  0            String        Hello
    '                  1            String        How
    '                  2            String        Are
    '                  3            String        You
    '                  4            String        Again
    '                  5            String        ?

    ' Items Before Removals:
    '   Items:         11 
    '    0            String        Hello World
    '    1            Integer        5 
    '    2            Object        Worksheet
    '    3            String        Hello
    '    4            String        World
    '    5            String        How
    '    6            String        Are
    '    7            String        You
    '    8            String        ?
    '    9            Array
    '                  0            String        Hello
    '                  1            String        World
    '                  2            String        How
    '                  3            String        Are
    '                  4            String        You
    '                  5            String        Again
    '                  6            String        ?
    '    10           String        I am the last item      
End Sub
```

This snippet removes a range of items using the `RemoveRange()` method.
- This example removes 5 items starting at the index.
    - i.e. `container.RemoveRange 3, 1` would remove 1 item - the item at the 3rd index.
    - This would be the same as `container.Remove 3`

```vb
Private Sub Demo()
    Dim container As List
    Set container = GenerateList
    
    ' Remove the a range of items starting at the 3rd index
    '   Removes 5 items; The first item is included the item at the index.
    '   i.e. container.RemoveRange 3, 1 would remove 1 item - the item at the 3rd index
    container.RemoveRange 3, 5
    
    PrintListItems container

    ' Prints:
    '   Items:         6 
    '    0            String        Hello World
    '    1            Integer        5 
    '    2            Object        Worksheet
    '    3            String        ?
    '    4            Array
    '                  0            String        Hello
    '                  1            String        How
    '                  2            String        Are
    '                  3            String        You
    '                  4            String        Again
    '                  5            String        ?
    '    5            String        I am the last item

    ' Items Before Removals:
    '   Items:         11 
    '    0            String        Hello World
    '    1            Integer        5 
    '    2            Object        Worksheet
    '    3            String        Hello
    '    4            String        World
    '    5            String        How
    '    6            String        Are
    '    7            String        You
    '    8            String        ?
    '    9            Array
    '                  0            String        Hello
    '                  1            String        World
    '                  2            String        How
    '                  3            String        Are
    '                  4            String        You
    '                  5            String        Again
    '                  6            String        ?
    '    10           String        I am the last item         
End Sub
```