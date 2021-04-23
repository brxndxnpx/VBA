# ColorInfo

A color class that contains metadata for RGB, hex, and Microsoft Office's integer (long) color code.

---

## Usage

### Setting The Color


#### Using Hex Code
```vb
Private Sub Demo()
    Dim Color As New ColorInfo
    
    ' Set the color by hex code
    Color.HexCode = "#002060" 'Dark Blue
End Sub
```

#### Using RGB Values
```vb
Private Sub Demo()
    Dim Color As New ColorInfo
 
    ' Set the color by RGB values
    Color.SetRGBValues 255, 0, 0
End Sub
```

#### Using Microsoft Office Color Code
```vb
Private Sub Demo()
    Dim Color As New ColorInfo

    ' Set the color by MS color code
    Color.ColorCode = 5288448 'Green
End Sub
```




