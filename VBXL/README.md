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

---

### [ColorInfo](/VBXL/Classes/ColorInfo/)

A color class that contains metadata for RGB, hex, and Microsoft Office's integer (`Long` datatype) color code.

The color information will change in sync with the last updated property.
- For example, if the R value in RGB were to change, then the `HexCode` and `ColorCode` properties will also change to match the new color's value.
- If the `HexCode` were to change then the RGB values and `ColorCode` values will also change to match the new color's value.

This class/object is primarily used to convert colors, e.g. from RGB to hex.
- Of course, the methods/functions in the class object can also be set as functions in a module.

---

### [Dictionary](/VBXL/Classes/Dictionary/)

A late-binding wrapper class for the `Scripting.Dictionary` object.

Allows using the `Scripting.Dictionary` object with intellisense without having to add a reference.
- Other users may not have the `Scripting` library added in their Visual Basic Editor. 
- This will alleviate that issue by using late-binding.

You can also store other objects in the dictionary.

---

### [FileSystem](/VBXL/Classes/FileSystem/)

A late-binding wrapper class for the `Scripting.FileSystemObject` object.

Allows using the `Scripting.FileSystemObject` object with intellisense without having to add a reference.
- Other users may not have the `Scripting` library added in their Visual Basic Editor. 
- This will alleviate that issue by using late-binding.

Recommended to be used in conjunction with...
- [Environment.bas](/VBXL/Modules/Environment/Environment.bas) module for easier file path access.
- [DynamicLinkLibraries.bas](/VBXL/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas) module for downloading documents via URL.

Dependencies:
[DynamicLinkLibraries.bas](/VBXL/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas)
- The `DownloadDocument` method requires the [DynamicLinkLibraries.bas](/VBXL/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas) to be included in the project.
    - This method can be removed otherwise.

---

### [StringBuilder](/VBXL/Classes/StringBuilder/)

Inspired by the C# .NET `System.Text.StringBuilder` class, but simplified for VBA.

A string building class that utilizes the `Scripting.Dictionary` object.
- Each item in the dictionary is an individual line.
- The dictionary's key is auto incremented and managed by the `StringBuilder` object.

This class is useful for building lengthy complex or formatted text, like an XML document.


---

## Modules

Static modules created for easing VBA development.

I tried to decouple these modules as much as I could for individual classes or modules to be usable.
- There may still be dependencies since some classes work better with some modules.
- e.g. The [FileSystem.cls](/VBXL/Classes/FileSystem/FileSystem.cls) class works well with the [Environment.bas](/VBXL/Modules/Environment/Environment.bas) module for easy path access.

### [Arry](/VBXL/Modules/Arry/)

Array helper functions.

### [ColorSwatches](/VBXL/Modules/ColorSwatches/)

Enumerated color values. 

- Works for Userforms and Worksheet objects i.e. Cell/Range colors.
- Can be used for creating custom color themes for Userforms/Workbooks.
- Based on [Material Design](https://www.materialpalette.com/colors) color swatches.


- Color codes prefixed with "A" are _accent_ colors.
    - i.e. `ColorSwatch.AmberA700`


### [DynamicLinkLibraries](/VBXL/Modules/DynamicLinkLibraries/)

Contains references to Window's .dll files to access Window's API functions.

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