# Classes

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
- [DynamicLinkLibraries.bas](/VBXL/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas) module for downloading documents via URL.

Dependencies:
[DynamicLinkLibraries.bas](/VBXL/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas)
- The `DownloadDocument` method requires the [DynamicLinkLibraries.bas](/VBXL/Modules/DynamicLinkLibraries/DynamicLinkLibraries.bas) to be included in the project.
    - This method can be removed otherwise.

### [List](/VBXL/Classes/List/)

A array based class that allows grouping data or objects into a single record in an array.

Primarily used for creating complex nested arrays.

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

