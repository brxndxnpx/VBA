# Arry

Array helper functions.

## Methods & Functions

|                             | Description                                                                                                         |
|-----------------------------|---------------------------------------------------------------------------------------------------------------------|
| [`ArryAppend`](#arryappend) | Appends items to an array.                                                                                          |
| [`ArryResize`](#arryresize) | Resizes an array.<br>Will instantiate a new array if the array is empty.                                            |
| [`ArryRemove`](#arryremove) | Removes an item from an array and resizes it.                                                                       |
| [`ArryCount`](#arrycount)   | Counts the items in an array.                                                                                       |
| [`ArryDebug`](#arrydebug)   | Uses `Debug.Print` to print the values of the items in the array along with it's data type to the immediate window. |

---

### `ArryAppend`

Appends items to an array.

**Parameters**
- `Source` `ByRef`
    - The array to append.
- `Items()` `ByRef`
    - The item(s) to append to the source.

---

### `ArryResize`

Resizes an array. Will instantiate a new array if the array is empty.

**Parameters**
- `Source` `ByRef`
    - The array to resize.
- `AddedBounds` `ByVal` _`Optional`_
    - The number of additional upper bound dimensions to add to the source.
- `PreserveData` `ByVal` _`Optional`_
    - Whether or not to preserve the data in the source.

---

### `ArryRemove`

Removes an item from an array and resizes it.

**Parameters**
- `Source` `ByRef`
    - The array to reference.
- `Index` `ByVal`
    - The index to remove.

---


### `ArryCount`

Counts the items in an array.

**Parameters**
- `Source` `ByRef`
    - The array to reference.

**Returns**
- The number of items in the array. `Long`

---

### `ArryDebug`

Uses `Debug.Print` the print values of the items in the array along with it's data type to the immediate window.

**Parameters**
- `Source` `ByRef`
    - The array to reference.
