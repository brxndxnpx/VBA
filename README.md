# VBA
VBA modules, classes, and code snippets for Microsoft Office Applications (Excel, Word, Outlook, etc).

- Each class or module should be decoupled enough to be usable on it's own.
    - If you find an issue or a class/module that depends on another class/module please let me know.

# Background

I started this repo to log and share my collection of VBA snippets. I got tired of searching through my old projects to find snippets of code I've previously used. I finally decided to create a public repository for those classes, modules, and snippets that I reuse on the regular and to document each item.

If you have any requests or ideas for modules/class then let me know :smile:.

If you find any issues or figure a better way to do something then let me know about that too!

# Contents

## VBXL (Excel)

**Classes**
- [AcroApp](/VBXL/Classes/AcroApp/)
- [ColorInfo](/VBXL/Classes/ColorInfo/)
- [Dictionary](/VBXL/Classes/Dictionary/)
- [FileSystem](/VBXL/Classes/FileSystem/)
- [JsonConverter](/VBXL/Classes/JsonConverter/)
- [List](/VBXL/Classes/List/)
- [OutlookApp](/VBXL/Classes/OutlookApp/)
- [RegExp](/VBXL/Classes/RegExp/)
- [SqlAccessor](/VBXL/Classes/SqlAccessor/)
- [StringBuilder](/VBXL/Classes/StringBuilder/)
- [WebRequest](/VBXL/Classes/WebRequest/)

**Modules**
- [Arry](/VBXL/Modules/Arry/)
- [ColorSwatches](/VBXL/Modules/ColorSwatches/)
- [Environment](/VBXL/Modules/Environment/)
- [ObjectInspector](/VBXL/Modules/ObjectInspector/)
- [ShellCommand](/VBXL/Modules/ShellCommand/)
- [TextStreamer](/VBXL/Modules/TextStreamer/)
- [TypeValidation](/VBXL/Modules/TypeValidation/)

## VBOL (Outlook)

- TBD

## VBWD (Word)

- TBD



<!-- 
## Notes

After coming across this [StackOverflow](https://stackoverflow.com/questions/26409117/why-use-integer-instead-of-long#:~:text=Traditionally%2C%20VBA%20programmers%20have%20used,re%20declared%20as%20type%20Integer) thread, I no longer use `Integer` types in the code provided here - unless it is an `Array(Long)` or `Variant(Long)`.



- Storing a handful of `Long` data types won't cause performance or memory issues, but iterating 

According to this (_dated)_ [MSDN documentation](https://docs.microsoft.com/en-us/previous-versions/office/developer/office2000/aa164506(v=office.10)?redirectedfrom=MSDN)...


> The Integer and Long data types can both hold positive or negative values. The difference between them is their size: Integer variables can hold values between -32,768 and 32,767, while Long variables can range from -2,147,483,648 to 2,147,483,647. Traditionally, VBA programmers have used integers to hold small numbers, because they required less memory. In recent versions, however, VBA converts all integer values to type Long, even if they're declared as type Integer. So there's no longer a performance advantage to using Integer variables; in fact, Long variables may be slightly faster because VBA does not have to convert them.


It's important to note that the documentation above may be incorrect as of now.
- As one of the comment states:
> Integers _still_ require less memory to store - a large array of integers will need significantly less RAM than an Long array with the same dimensions. But because the processor needs to work with 32 bit chunks of memory, VBA converts Integers to Longs _temporarily_ when it performs calculations -->