# TypeValidation

Generic validation for objects/variables.

## Constants

|              | Type   | Description                                |
|--------------|--------|--------------------------------------------|
| `vbNullDate` | `Date` | The default system date (`#12:00:00 AM#`). |

## Methods & Functions

|                                 | Type      | Description                                                    |
|---------------------------------|-----------|----------------------------------------------------------------|
| [`IsNothing`](#isnothing)       | `Boolean` | Checks if an object is nothing.                                |
| [`IsNotNothing`](#isnotnothing) | `Boolean` | Checks if an object is not nothing.                            |
| [`IsNullString`](#isnullstring) | `Boolean` | Checks if a string is null or empty.                           |
| [`IsNullDate`](#isnulldate)     | `Boolean` | Checks if a date is null or not set (the default system date). |

---

### [`IsNothing`](TypeValidation.bas#L19)

Checks if an object is nothing.

**Parameters**
- `Value` `ByRef` `Object`
    - The object to validate.

**Returns**
- A `Boolean`; True if the object is nothing.

### [`IsNotNothing`](TypeValidation.bas#L29)

Checks if an object is not nothing.

**Parameters**
- `Value` `ByRef` `Object`
    - The object to validate.

**Returns**
- A `Boolean`; True if the object is not nothing.


### [`IsNullString`](TypeValidation.bas#L39)

Checks if a string is null.

**Parameters**
- `Value` `ByVal` `String`
    - The string to validate.

**Returns**
- A `Boolean`; Whether or not the string is equal to vbNullString, i.e. "".


### [`IsNullDate`](TypeValidation.bas#L49)

Checks if a date is null.

**Parameters**
- `Value` `ByVal` `Date`
    - The date to validate.

**Returns**
- A `Boolean`; Whether or not the date is equal to the default system date (not set).


---


## Usage

```vb
Private Sub Demo()
    Dim exampleDate As Date
    Dim exampleString As String
    Dim exampleObj As Object
    
    ' Check if the values are null/empty or nothing
    Debug.Print IsNullDate(exampleDate)
    Debug.Print IsNullString(exampleString)
    Debug.Print IsNothing(exampleObj)
    
    ' Prints:
    '   True
    '   True
    '   True

    ' Set the values
    exampleDate = Date
    exampleString = "Hello World"
    Set exampleObj = ActiveWorkbook.Sheets(1)
    
    ' Check if the values are null/empty or nothing
    Debug.Print IsNullDate(exampleDate)
    Debug.Print IsNullString(exampleString)
    Debug.Print IsNothing(exampleObj)

    ' Prints:
    '   False
    '   False
    '   False
End Sub
```
