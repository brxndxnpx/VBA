# ColorSwatches

Enumerated color values. 

- Works for Userforms and Worksheet objects i.e. Cell/Range colors.
- Can be used for creating custom color themes for Userforms/Workbooks.
- Based on [Material Design](https://www.materialpalette.com/colors) color swatches.

---

## Usage

Set the fill color for `Cell(1, 1)`.

```vb
Private Sub Demo()
    ' Set the fill color for Cell(1, 1)
    ActiveWorkbook.Sheets(1).Cells(1, 1).Interior.Color = ColorSwatch.Amber600
End Sub
```
