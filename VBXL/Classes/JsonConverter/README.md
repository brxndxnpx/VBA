# JsonConverter

A class created for parsing JSON strings. 
- Stores the object in a dictionary.
- Uses dot notation to access nested properties

## Methods & Functions

|                               | Type                              | Description                                              |
|-------------------------------|-----------------------------------|----------------------------------------------------------|
| [`Deserialize`](#deserialize) | `Object` (`Scripting.Dictionary`) | Converts JSON into a dictionary resembling the object.   |
| [`PrintPaths`](#printpaths)   |                                   | Prints each path and it's value to the immediate window. |

---

### `Deserialize`

Converts JSON into a dictionary resembling the object.

**Parameters**
- `ByVal Json As String`
    - The json text to convert.
- `[Optional] ByVal Key As String`
    - The optional base key.

**Returns**
- `Object`: A `Scripting.Dictionary` object resembling the json data.

---

### `PrintPaths`

Prints each path and it's value to the immediate window.

**Parameters**
- `[Optional] ByRef Obj As Object`
    - A `Scripting.Dictionary` object.
    - Will use the previously parse json object if left empty.


---

## Usage

Json Examples:

```json
{
    "anime": "Avatar: The Last Airbender",
    "character": "Sokka",
    "quote": "[to himself as he chops at the ice] I'm just a guy with a boomerang. I didn't ask for all this flying and magic."
}
```

```json
{
 "profile" : {
    "username" : "Zezimaaaaa",
    "lastLoginDate" : "2007-04-21T10:00:00.000Z",
    "level": 101,
    "levels" : {
        "woodCutting": {
            "level": 99,
            "experience": 91847684236
        },
        "fireMaking": {
            "level": 98,
            "experience": 101668484668
        },
        "smithing": {
            "level": 96,
            "experience": 87315654231
        }
    },
    "inventory": [
        {
            "itemId": 42,
            "itemName": "Magic Log",
            "quantity": 20
        },
        {
            "itemId": 86,
            "itemName": "Rune Bar",
            "quantity": 7
        },
        {
            "itemId": 42,
            "itemName": "Dragon Woodcutting Axe",
            "quantity": 1
        }
    ]
  }
}
```

```vb
Private Sub Demo()
    Dim JConverter   As New JsonConverter
    Dim Json         As String
    Dim Results1     As Object
    Dim Results2     As Object
    
    ' Parse two json strings
    
    ' First json:
    Json = "{ ""profile"" : { ""username"" : ""Zezimaaaaa"", ""lastLoginDate"" : ""2007-04-21T10:00:00.000Z"", " & _
        """level"": 101, ""levels"" : {" & _
        """woodCutting"": { ""level"": 99, ""experience"": 91847684236 }," & _
        """fireMaking"": { ""level"": 98, ""experience"": 101668484668 }," & _
        """smithing"": { ""level"": 96, ""experience"": 87315654231 } }," & _
        """inventory"": [" & _
        "{ ""itemId"": 42, ""itemName"": ""Magic Log"", ""quantity"": 20 }," & _
        "{ ""itemId"": 86, ""itemName"": ""Rune Bar"", ""quantity"": 7 }," & _
        "{ ""itemId"": 42, ""itemName"": ""Dragon Woodcutting Axe"", ""quantity"": 1 } ] } }"
    
    ' Parse the text. Setting the first key to 'game'
    Set Results1 = JConverter.Deserialize(Json, "game")
    
    ' Second json:
    Json = "{""anime"":""Avatar: The Last Airbender"",""character"":""Sokka"",""quote"":" & _
        """[to himself as he chops at the ice] I'm just a guy with a boomerang. I didn't ask for all this flying and magic.""}"
    
    ' Parse the text. Not setting the first key at all
    Set Results2 = JConverter.Deserialize(Json)
    
    ' Print the results of the last json object (no arguments passed)
    Debug.Print "Results2:"
    JConverter.PrintPaths
    ' Prints:
    '   Results2:
    '   anime                                     Avatar: The Last Airbender
    '   anime.character                           Sokka
    '   anime.quote                               [to himself as he chops at the ice] I'm just a guy with a boomerang. 
    '                                               I didn't ask for all this flying and magic.
    
    Debug.Print

    ' Print the results of a particular json object
    Debug.Print "Results1:"
    JConverter.PrintPaths Results1

    ' Prints:
    '   Results1:
    '   game.profile.username                     Zezimaaaaa
    '   game.profile.lastLoginDate                2007-04-21T10:00:00.000Z
    '   game.profile.level                        101
    '   game.profile.levels.woodCutting.level     99
    '   game.profile.levels.woodCutting.experience              91847684236
    '   game.profile.levels.fireMaking.level      98
    '   game.profile.levels.fireMaking.experience 101668484668
    '   game.profile.levels.smithing.level        96
    '   game.profile.levels.smithing.experience   87315654231
    '   game.profile.inventory(0).itemId          42
    '   game.profile.inventory(0).itemName        Magic Log
    '   game.profile.inventory(0).quantity        20
    '   game.profile.inventory(1).itemId          86
    '   game.profile.inventory(1).itemName        Rune Bar
    '   game.profile.inventory(1).quantity        7
    '   game.profile.inventory(2).itemId          42
    '   game.profile.inventory(2).itemName        Dragon Woodcutting Axe
    '   game.profile.inventory(2).quantity        1

End Sub
```

