# StringBuilder

Inspired by the C# .NET `System.Text.StringBuilder` class, but simplified for VBA.

A string building class that utilizes the `Scripting.Dictionary` object.
- Each item in the dictionary is an individual line.
- The dictionary's key is auto incremented and managed by the `StringBuilder` object.

This class is useful for building lengthy complex or formatted text, like an XML document.

---

## Properties

| Property | Description                 |
|----------|-----------------------------|
| Lines    | The keys in the dictionary. |

## Methods/Functions

| Method/Functions | Description                                                               |
|------------------|---------------------------------------------------------------------------|
| Append           | Appends text to the current line.                                         |
| AppendLine       | Appends text to the next line.                                            |
| AppendArray      | Appends text from an array. Each dimension of the array will be appended. |
| ToString         | Returns the lines as a single string.                                     |
| ToArray          | Returns the lines as an array.                                            |
| PrintEachLine    | Prints out each line to the immediate window. Useful for debugging.       |

---

## Usage

```vb
Private Sub Demo()
    Dim s As New StringBuilder
    
    s.Append "Hello"
    s.Append " "
    s.Append "World"
    
    s.AppendLine "How are you?"

    Debug.Print s.ToString
    ' Prints:
    '   Hello World
    '   How are you?
    
    ' Used for debugging
    s.PrintEachLine
    ' Prints:
    '   Hello World
    '   How are you?
End Sub
```

```vb
Private Sub Demo()
    Dim s As New StringBuilder
    Dim title_ As String: title_ = "Check this out"
    Dim h1_ As String: h1_ = "HELLO WORLD!!!"
    Dim btnText_ As String: btnText_ = "Click Me!!!!"
    
    s.AppendLine "<!DOCTYPE html>"
    s.AppendLine "<html>"
    s.AppendLine vbTab & "<head>"
    s.AppendLine vbTab & vbTab & "<title>"
    s.Append title_
    s.Append "</title>"
    s.AppendLine vbTab & "</head>"
    s.AppendLine vbTab & "<body>"
    s.AppendLine vbTab & vbTab & "<h1>"
    s.Append h1_
    s.Append "</h1>"
    s.AppendLine vbTab & vbTab & "<button>"
    s.Append btnText_
    s.Append "</button>"
    s.AppendLine vbTab & "</body>"
    s.AppendLine "</html>"

    Debug.Print s.ToString
    ' Prints:
    '   <!DOCTYPE html>
    '   <html>
    '       <head>
    '           <title>Check this out</title>
    '       </head>
    '       <body>
    '           <h1>HELLO WORLD!!!</h1>
    '           <button>Click Me!!!!</button>
    '       </body>
    '   </html>

End Sub
```