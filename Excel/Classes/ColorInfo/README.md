# ColorInfo

A color class that contains metadata for RGB, hex, and Microsoft Office's integer (`Long` datatype) color code.
The color information will change in sync with the last updated property.
- For example, if the R value in RGB were to change, then the `HexCode` and `ColorCode` properties will also change to match the new color's value.
- If the `HexCode` were to change then the RGB values and `ColorCode` values will also change to match the new color's value.

This class/object is primarily used to convert colors, e.g. from RGB to hex.
- Of course, the methods/functions in the class object can also be set as functions in a module.

---

## Properties

| Property  | Type     |                                     |
|-----------|----------|-------------------------------------|
| ColorCode | `Long`   | The Excel color code for the color. |
| HexCode   | `String` | The hex code for the color.         |
| R         | `Long`   | The R part of the RGB color         |
| G         | `Long`   | The G part of the RGB color         |
| B         | `Long`   | The B part of the RGB color         |


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

