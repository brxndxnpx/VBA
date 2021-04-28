# Modules

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