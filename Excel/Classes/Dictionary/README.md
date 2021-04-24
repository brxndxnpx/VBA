# Dictionary

A late binding wrapper class for the `Scripting.Dictionary` object.

Allows using the `Scripting.Dictionary` object with intellisense.

---

## Properties

| Property | Type     |                                                     |
|----------|----------|-----------------------------------------------------|
| Keys     | `Long`   | The keys in the dictionary.                         |
| Items    | `String` | The items in the dictionary.                        |
| Item     | `Long`   | Gets an item in the dictionary. **Default Member**. |
| Count    | `Long`   | The number of items in the dictionary.              |


## Methods/Functions

| Method/Functions |                                                                  |
|------------------|------------------------------------------------------------------|
| Add              | Adds an item to the dictionary.                                  |
| Replace          | Replaces an item at a key with a another item.                   |
| GetItem          | Gets an item by it's key. Used if `Item` isn't working properly. |
| GetKey           | Gets a key by it's item.                                         |
| Exists           | Checks if a key exists.                                          |
| Remove           | Removes an item.                                                 |
| RemoveAll        | Removes all items.                                               |


---

## Usage

### Setting The Color

You can set the color in 3 ways.
1. Hex code, e.g. `#FFFFFF` for white.
0. RGB values, e.g. `255`, `255`, `255` for white.
    - Unfortunately VBA doesn't support passing values through constructors.
    - This is set using the `ColorInfo.SetRGBValues()` method.
0. Microsoft Office color codes, e.g. `16777215` for white.

#### 1. Using Hex Code
```vb
Private Sub Demo()
    Dim Color As New ColorInfo
    
    ' Set the color by hex code
    Color.HexCode = "#002060" 'Dark Blue
End Sub
```

#### 2. Using RGB Values
```vb
Private Sub Demo()
    Dim Color As New ColorInfo
 
    ' Set the color by RGB values
    Color.SetRGBValues 255, 0, 0
End Sub
```

#### 3. Using Microsoft Office Color Code
```vb
Private Sub Demo()
    Dim Color As New ColorInfo

    ' Set the color by MS color code
    Color.ColorCode = 5288448 'Green
End Sub
```

---

### Getting The Color

You can get the color by accessing it's property once a color is set.

```vb
Private Sub Demo()
    Dim Color As New ColorInfo

    ' Set the color by RGB values
    Color.SetRGBValues 255, 0, 0

    Debug.Print Color.HexCode
End Sub
```

