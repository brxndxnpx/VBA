# Arry

Array helper functions.

## Methods/Functions

| Methods    | Type     | Description                                                                      | Parameters                                                                                                                                                                                                                                                                                 | Returns                                           |
|------------|----------|----------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|---------------------------------------------------|
| ArryAppend | Method   | Appends items to an array.                                                       | `Source` `ByRef`<br>&emsp;The array to append.<br>`Items()` `ByRef`<br>&emsp;The items to append to the source.                                                                                                                                                                            |                                                   |
| ArryResize | Method   | Resizes an array.<br>Will instantiate a new array if the array is empty.         | `Source` `ByRef`<br>&emsp;The array to resize.<br>`AddedBounds` `ByVal` [`Optional`] <br>&emsp;The number of additional upper<br>&emsp;bound dimensions to add to the source.<br>`PreserveData` `ByVal` [`Optional`] <br>&emsp;Whether or not to preserve<br>&emsp;the data in the source. |                                                   |
| ArryRemove | Method   | Removes an item from an<br>array and resizes it.                                 | `Source` `ByRef`<br>&emsp;The array to reference.<br>`Index` `ByVal`<br>&emsp;The index to remove.                                                                                                                                                                                         |                                                   |
| ArryCount  | Function | Counts the items in an array.                                                    | `Source` `ByRef`<br>&emsp;The array to reference.                                                                                                                                                                                                                                          | `Long`<br>&emsp;The number of items in the array. |
| ArryDebug  | Method   | `Debug.Print` the values of the<br>items in the array along with it's data type. | `Source` `ByRef`<br>&emsp;The array to reference.                                                                                                                                                                                                                                          |                                                   |

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
