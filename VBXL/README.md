# VBXL
VBA modules, classes, and code snippets pertaining to Excel.

## Classes

Late-bound objects created for easing VBA development.

I tried to decouple these objects as much as I could for individual classes or modules to be usable.
- There may still be dependencies since some classes work better with some modules.
- e.g. The [FileSystem.cls](/VBXL/Classes/FileSystem/FileSystem.cls) class works well with the [Environment.bas](/VBXL/Modules/Environment/Environment.bas) module for easy path access.

### [AcroApp](/VBXL/Classes/AcroApp/)

An Acrobat class that is used to combine PDFs and convert files (e.g. images) to PDFs.

- Requires Adobe Acrobat DC to be installed on the user's machine.


### [ColorInfo](/VBXL/Classes/ColorInfo/)

A color class that contains metadata for RGB, hex, and Microsoft Office's integer (`Long` datatype) color code.

The color information will change in sync with the last updated property.
- For example, if the R value in RGB were to change, then the `HexCode` and `ColorCode` properties will also change to match the new color's value.
- If the `HexCode` were to change then the RGB values and `ColorCode` values will also change to match the new color's value.

This class/object is primarily used to convert colors, e.g. from RGB to hex.
- Of course, the methods/functions in the class object can also be set as functions in a module.


### [Dictionary](/VBXL/Classes/Dictionary/)

A late-binding wrapper class for the `Scripting.Dictionary` object.

Allows using the `Scripting.Dictionary` object with intellisense without having to add a reference.
- Other users may not have the `Scripting` library added in their Visual Basic Editor. 
- This will alleviate that issue by using late-binding.

You can also store other objects in the dictionary.


### [FileSystem](/VBXL/Classes/FileSystem/)

A late-binding wrapper class for the `Scripting.FileSystemObject` object.

Allows using the `Scripting.FileSystemObject` object with intellisense without having to add a reference.
- Other users may not have the `Scripting` library added in their Visual Basic Editor. 
- This will alleviate that issue by using late-binding.

Recommended to be used in conjunction with...
- [Environment.bas](/VBXL/Modules/Environment/Environment.bas) module for easier file path access.

### [List](/VBXL/Classes/List/)

A array based class that is intended to make managing arrays easier by automatically resizing whenever an item is added or removed.

Was primarily made to give the user friendliness a `Collection` gives, but without sacrificing speed when iterating through items.
- An `Array` is faster than a `Collection`.

Can be used for creating complex nested arrays.


### [OutlookApp](/VBXL/Classes/OutlookApp/)

A late bound object made to utilize Outlook functionalities from Excel.
- Uses the Outlook application on the current user's machine.
    - The `Outlook.Application` object.
- Uses the accounts on the user's Outlook application.

### [StringBuilder](/VBXL/Classes/StringBuilder/)

Inspired by the C# .NET `System.Text.StringBuilder` class, but simplified for VBA.

A string building class that utilizes the `Scripting.Dictionary` object.
- Each item in the dictionary is an individual line.
- The dictionary's key is auto incremented and managed by the `StringBuilder` object.

This class is useful for building lengthy complex or formatted text, like an XML document.

### [SqlAccessor](/VBXL/Classes/SqlAccessor/)

A late binding object made to execute simple database queries and return the dataset as an array.
- Uses the `ADODB.Command`, `ADODB.Recordset`, and `ADODB.Connection` objects.
    - See MSDN Documentation [here](https://docs.microsoft.com/en-us/sql/ado/guide/data/creating-and-executing-a-simple-command?view=sql-server-ver15).


---

## Modules

Static modules created for easing VBA development.

I tried to decouple these modules as much as I could for individual classes or modules to be usable.
- There may still be dependencies since some classes work better with some modules.
- e.g. The [FileSystem.cls](/VBXL/Classes/FileSystem/FileSystem.cls) class works well with the [Environment.bas](/VBXL/Modules/Environment/Environment.bas) module for easy path access.

### [Arry](/VBXL/Modules/Arry/)

Array helper functions.

Works with arrays with a base of 0 or 1.
- If the array dimensions aren't already set (the array is `Empty`), the resized array will have a base of 0.

You can override the 0 base index by using the `Option Base` statement at the top of the module.
- See MSDN documentation [here](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-base-statement).


### [ColorSwatches](/VBXL/Modules/ColorSwatches/)

Enumerated color values. 
- Works for Userforms and Worksheet objects i.e. Cell/Range colors.
- Can be used for creating custom color themes for Userforms/Workbooks.
- Based on [Material Design](https://www.materialpalette.com/colors) color swatches.
- Color codes prefixed with "A" are _accent_ colors.
    - i.e. `ColorSwatch.AmberA700`

### [Environment](/VBXL/Modules/Environment/)

Environment functions pertaining to the user and the user's machine.

### [ObjectInspector](/VBXL/Modules/ObjectInspector/)

Used to inspect objects and retrieve their methods, functions, and properties.

### [ShellCommand](/VBXL/Modules/ShellCommand/)

Basic shell commands.

### [TextStreamer](/VBXL/Modules/TextStreamer/)

Performs basic reads and writes text to files without creating a `Scripting.TextStream` object.

Recommended to be used in conjunction with...
- [Environment.bas](/VBXL/Modules/Environment/Environment.bas) module for easier file path access.
- [FileSystem.cls](/VBXL/Classes/FileSystem/FileSystem.cls) class for easier file access.
- [StringBuilder.cls](/VBXL/Classes/StringBuilder/StringBuilder.cls) class for easier text readability.
    - You can use the `StringBuilder.AppendArray()` to inject the text array into the `StringBuilder` object when reading a file.

### [TypeValidation](/VBXL/Modules/TypeValidation/)

Generic validation for objects/variables.        