# Arry

Array helper functions.

Works with arrays with a base of 0 or 1.
- If the array dimensions aren't already set (the array is `Empty`), the resized array will have a base of 0.

You can override the 0 base index by using the `Option Base` statement at the top of the module.
- See MSDN documentation [here](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-base-statement).


## Methods & Functions

|                             | Description                                                                                                         |
|-----------------------------|---------------------------------------------------------------------------------------------------------------------|
| [`ArryAppend`](#arryappend) | Appends items to an array.                                                                                          |
| [`ArryResize`](#arryresize) | Resizes an array.<br>Will instantiate a new array if the array is empty.                                            |
| [`ArryRemove`](#arryremove) | Removes an item from an array and resizes it.                                                                       |
| [`ArryCount`](#arrycount)   | Counts the items in an array.                                                                                       |
| [`ArryDebug`](#arrydebug)   | Uses `Debug.Print` to print the values of the items in the array along with it's data type to the immediate window. |

---

### [`ArryAppend`](Arry.bas#L14)

Appends items to an array.

**Parameters**
- `Source` `ByRef`
    - The array to append.
- `Items()` `ByRef` `ParamArray`
    - The item(s) to append to the source.

---

### [`ArryResize`](Arry.bas#L32)

Resizes an array. Will instantiate a new array if the array is empty.

**Parameters**
- `Source` `ByRef`
    - The array to resize.
- `AddedBounds` `ByVal` [`Optional`]
    - The number of additional upper bound dimensions to add to the source.
- `PreserveData` `ByVal` [`Optional`]
    - Whether or not to preserve the data in the source.

---

### [`ArryRemove`](Arry.bas#L54)

Removes an item from an array and resizes it.

**Parameters**
- `Source` `ByRef`
    - The array to reference.
- `Index` `ByVal`
    - The index to remove.

---


### [`ArryCount`](Arry.bas#L74)

Counts the items in an array.

**Parameters**
- `Source` `ByRef`
    - The array to reference.

**Returns**
- `Long`: The number of items in the array. 

---

### [`ArryDebug`](Arry.bas#L92)

Uses `Debug.Print` the print values of the items in the array along with it's data type to the immediate window.

**Parameters**
- `Source` `ByRef`
    - The array to reference.

---

## Usage

Passing an array that is `Empty`.

```vb
Private Sub Demo()
    Dim Source           As Variant
    Dim example_String   As String
    Dim example_Integer  As Long
    Dim example_Object   As Object
    
    ' Set the example variables
    example_String = "HELLO WORLD"
    example_Integer = 1090
    Set example_Object = CreateObject("Scripting.Dictionary")
    
    ' Append the variables to the array (Source)
    ArryAppend Source, example_String, example_Integer, example_Object

    ' Remove the item at the 2nd index: example_Integer (zero based indexing)
    ArryRemove Source, 1

    ' Print the items to the immediate window
    ArryDebug Source

    ' Prints:
    '   String        HELLO WORLD
    '   Object        Dictionary    
End Sub
```

Passing an array with a base index of 1.

```vb
Private Sub Demo()
    Dim Source           As Variant
    Dim example_String   As String
    Dim example_Integer  As Long
    Dim example_Object   As Object
    
    ' Set the example variables
    example_String = "HELLO WORLD"
    example_Integer = 1090
    Set example_Object = CreateObject("Scripting.Dictionary")
    
    ReDim Source(1 To 1)
    Source(1) = "Oh NO!"

    ' Append the variables to the array (Source)
    ArryAppend Source, example_String, example_Integer, example_Object

    ' Remove the item at the 2nd index: example_Integer
    ArryRemove Source, 2

    ' Print the items to the immediate window
    ArryDebug Source

    ' Prints:
    '   String        HELLO WORLD
    '   Long           1090 
    '   Object        Dictionary
End Sub
```
