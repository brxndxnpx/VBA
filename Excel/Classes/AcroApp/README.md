# ColorInfo

An Acrobat class that is used to combine PDFs and convert files (e.g. images) to PDFs.

- Requires Adobe Acrobat DC to be installed on the user's machine.

---

## Properties

| Property    | Type      | Description                                 |
|-------------|-----------|---------------------------------------------|
| IsInstalled | `Boolean` | Indicates if Adobe Acrobat DC is installed. |

## Methods/Functions

| Method/Functions | Description                                      |
|------------------|--------------------------------------------------|
| MergeDocuments   | Merges an array of file paths into a single PDF. |
| ConvertToPDF     | Converts a file to a PDF.                        |


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

