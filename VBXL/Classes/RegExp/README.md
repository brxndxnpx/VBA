# RegExp

A class for executing regular expressions utilizing the `VBScript.RegExp` object.

## Methods & Functions

|                       | Type     | Description                                    |
|-----------------------|----------|------------------------------------------------|
| [`Execute`](#execute) | `Variant` | Executes regular expressions against a string. |

---

### `Execute`

Executes regular expressions against a string.

**Parameters**

- `ByVal Value As String`
    - The text to parse.
- `ByVal Pattern As String`
    - The regular expression.
- `ByVal IncludeQuotes As Boolean`
    - Whether or not to include double quotes in matches.
- `ByVal UseGlobal As Boolean`
    - Whether or not to the global regex setting.

**Returns**
- `Variant`: Returns the results in an array.

---

## Usage

This example will parse this json text and tokenize each match.

```json
{
    "anime": "Avatar: The Last Airbender",
    "character": "Sokka",
    "quote": "[to himself as he chops at the ice] I'm just a guy with a boomerang. I didn't ask for all this flying and magic."
}
```


```vb
Private Sub Demo()
    Dim Re        As New RegExp
    Dim Results   As Variant
    Dim Json      As String
    Dim i         As Long
    
    ' Regex pattern
    Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    
    Json = "{""anime"":""Avatar: The Last Airbender"",""character"":""Sokka"",""quote"":" & _
        """[to himself as he chops at the ice] I'm just a guy with a boomerang. I didn't ask for all this flying and magic.""}"
    
    ' Exlude quotations
    Results = Re.Execute(Json, Pattern, True)
    For i = LBound(Results) To UBound(Results)
        Debug.Print Results(i)
    Next
    
    ' Prints:
    '   {
    '   anime
    '   :
    '   Avatar: The Last Airbender
    '   ,
    '   character
    '   :
    '   Sokka
    '   ,
    '   quote
    '   :
    '   [to himself as he chops at the ice] I'm just a guy with a boomerang. I didn't ask for all this flying and magic.
    '   }    

    Debug.Print
    
    ' Include quotations
    Results = Re.Execute(Json, Pattern, False)
    For i = LBound(Results) To UBound(Results)
        Debug.Print Results(i)
    Next

    ' Prints:
    '   {
    '   "anime"
    '   :
    '   "Avatar: The Last Airbender"
    '   ,
    '   "character"
    '   :
    '   "Sokka"
    '   ,
    '   "quote"
    '   :
    '   "[to himself as he chops at the ice] I'm just a guy with a boomerang. I didn't ask for all this flying and magic."
    '   }
End Sub
```
