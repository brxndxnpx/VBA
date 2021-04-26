# Arry

Array helper functions.

## Methods/Functions

| Method/Functions | Type     | Description                                                                   | Parameters                                                                                                                                                                                                                             | Returns |
|------------------|----------|-------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|---------|
| ArryAppend       | Method   | Appends items to an array.                                                    | `ByRef` `Source`<br>&emsp;The array to append.<br>`ByRef` `Items()`<br>&emsp;The items to append to the source.                                                                                                                        |         |
| ArryResize       | Method   | Resizes an array. Will instantiate a new array if the array is empty.         | `ByRef` `Source`: The array to resize.<br>`ByVal` _`Optional`_ `AddedUBound`: The number of additional upper bound dimensions to add to the source.<br>`ByVal Optional PreserveData`: Whether or not to preserve the data in the source. |         |
| ArryRemove       | Method   | Removes an item from an array and resizes it.                                 | `ByRef` `Source`: The array to reference.<br>`ByVal` `Index`: The index to remove.                                                                                                                                                     |         |
| ArryCount        | Function | Counts the items in an array.                                                 | `ByRef` `Source`: The array to reference.                                                                                                                                                                                              |         |
| ArryDebug        | Method   | `Debug.Print` the values of the items in the array along with it's data type. | `ByRef` `Source`: The array to reference.                                                                                                                                                                                              |         |

### `ArryAppend`

| Type   | Description                |
|--------|----------------------------|
| Method | Appends items to an array. |

Parameters
- `Source` `ByRef`
    - The array to append.
- `Items()` `ByRef` _`Optional`_
    - The item(s) to append to the source.


### `ArryResize`

| Type   | Description                                                           |
|--------|-----------------------------------------------------------------------|
| Method | Resizes an array. Will instantiate a new array if the array is empty. |

### `ArryRemove`

| Type   | Description                                   |
|--------|-----------------------------------------------|
| Method | Removes an item from an array and resizes it. |

### `ArryCount`

| Type     | Description                   | Returns                           | Return Type |
|----------|-------------------------------|-----------------------------------|-------------|
| Function | Counts the items in an array. | The number of items in the array. | `Long`      |

### `ArryDebug`

| Type   | Description                                                                   |
|--------|-------------------------------------------------------------------------------|
| Method | `Debug.Print` the values of the items in the array along with it's data type. |






| Method/Functions | Type     | Description                                                                   | Parameters                                                                                                                                                                                                                             | Returns |
|------------------|----------|-------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|---------|
| ArryAppend       | Method   | Appends items to an array.                                                    | `ByRef` `Source`<br>&emsp;The array to append.<br>`ByRef` `Items()`<br>&emsp;The items to append to the source.                                                                                                                        |         |
| ArryResize       | Method   | Resizes an array. Will instantiate a new array if the array is empty.         | `ByRef` `Source`: The array to resize.<br>`ByVal` `Optional` `AddedUBound`: The number of additional upper bound dimensions to add to the source.<br>`ByVal Optional PreserveData`: Whether or not to preserve the data in the source. |         |
| ArryRemove       | Method   | Removes an item from an array and resizes it.                                 | `ByRef` `Source`: The array to reference.<br>`ByVal` `Index`: The index to remove.                                                                                                                                                     |         |
| ArryCount        | Function | Counts the items in an array.                                                 | `ByRef` `Source`: The array to reference.                                                                                                                                                                                              |         |
| ArryDebug        | Method   | `Debug.Print` the values of the items in the array along with it's data type. | `ByRef` `Source`: The array to reference.                                                                                                                                                                                              |         |
