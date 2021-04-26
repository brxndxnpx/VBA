# Arry

Array helper functions.

## Methods/Functions

### `ArryAppend`

**Method**

Appends items to an array.

**Parameters**
- `Source` `ByRef`
    - The array to append.
- `Items()` `ByRef`
    - The item(s) to append to the source.

---

### `ArryResize`

**Method**

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

**Method**

Removes an item from an array and resizes it.

**Parameters**
- `Source` `ByRef`
    - The array to reference.
- `Index` `ByVal`
    - The index to remove.

---


### `ArryCount`

**Function**

Counts the items in an array.

**Parameters**
- `Source` `ByRef`
    - The array to reference.

**Returns**

Type: `Long`

The number of items in the array.


---

### `ArryDebug`

**Method**

Uses `Debug.Print` the print values of the items in the array along with it's data type to the immediate window.

**Parameters**
- `Source` `ByRef`
    - The array to reference.
