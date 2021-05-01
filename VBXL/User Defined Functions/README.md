# User Defined Functions

## Functions

|                           | Type      | Description                                  |
|---------------------------|-----------|----------------------------------------------|
| [`QUOTE`](#quote)         | `String`  | Quotes a text using the character provided.  |
| [`REVERSE`](#reverse)     | `String`  | Reverses text.                               |
| [`CHARCOUNT`](#charcount) | `Long`    | Counts the number of characters in a string. |
| [`ISIN`](#isin)           | `Boolean` | Checks if an object or value is in a list.   |


## `QUOTE`

Quotes a text using the character provided.
- Will use the `"` character by default.

You can specify the open and close characters by providing more than one character, i.e. `[]`.
- It will first character as the open bracket and the last character as the closing bracket.


**Parameters**
- `ByVal Source As String`
    - The text to reverse.
- `[Optional] Character As String`
    - The character to use as the quote. Will use the " character by default.

**Returns**
- `String`: The text returned in a pair of quotes.

### Usage

Using the default quotes `"` by omitting the `Character` parameter.

<p align="center"><sub>VBA</sub></p>

```vb
Public Sub Demo()
    Debug.Print QUOTE("hello")
End Sub
```

<p align="center"><sub>Formula</sub></p>  

```vb
=QUOTE("hello")
```

<p align="center"><sub>Output</sub></p>  

```
"hello"
```

---

Providing the quotes to use.


<p align="center"><sub>VBA</sub></p>

```vb
Public Sub Demo()
    Debug.Print QUOTE("hello", "'")
End Sub
```

<p align="center"><sub>Formula</sub></p>  

```vb
=QUOTE("hello", "'")
```


<p align="center"><sub>Output</sub></p>  

```
'hello'
```


---

Using brackets as the quote characters.

<p align="center"><sub>VBA</sub></p>

```vb
Public Sub Demo()
    Debug.Print QUOTE("hello", "[]")
    Debug.Print QUOTE("hello", "()")
    Debug.Print QUOTE("hello", "{}")
End Sub
```


<p align="center"><sub>Formula</sub></p>  

```vb
=QUOTE("hello", "[]")
=QUOTE("hello", "()")
=QUOTE("hello", "{}")
```


<p align="center"><sub>Output</sub></p>  

```
[hello]
(hello)
{hello}
```

---

Using spaced brackets as the quote characters.

<p align="center"><sub>VBA</sub></p>


```vb
Public Sub Demo()
    Debug.Print QUOTE("hello", "[ ]")
    Debug.Print QUOTE("hello", "(  )")
    Debug.Print QUOTE("hello", "{   }")
End Sub
```


<p align="center"><sub>Formula</sub></p>  

```vb
=QUOTE("hello", "[ ]")
=QUOTE("hello", "(  )")
=QUOTE("hello", "{   }")
```


<p align="center"><sub>Output</sub></p>  

```
[ hello ]
(  hello  )
{   hello   }
```


---


## `REVERSE`

Reverses text.

**Parameters**

- `ByVal Source As String`
    - The text to reverse.

**Returns**
- `String`: The reversed text.

### Usage

<p align="center"><sub>VBA</sub></p>


```vb
Public Sub Demo()
    Debug.Print REVERSE("hello")
End Sub
```

<p align="center"><sub>Formula</sub></p>  

```vb
=REVERSE("hello")
```

<p align="center"><sub>Output</sub></p>  

```
olleh
```


---

## `CHARCOUNT`

Counts the number of characters in a string.

**Parameters**

- `ByVal Source As String`
    - The text to examine.
- `ByVal Character As String`
    - The text to count.

**Returns**
- `Long`: The number of characters in the next.


### Usage

<p align="center"><sub>VBA</sub></p>

```vb
Public Sub Demo()
    Debug.Print CHARCOUNT("hello", "l")
End Sub
```

<p align="center"><sub>Formula</sub></p>  

```vb
=CHARCOUNT("hello", "l")
```

<p align="center"><sub>Output</sub></p>  

```
2
```

---

## `ISIN`

Checks if an object or value is in a list.

Useful for preventing repeating formulas when using `AND` or `OR`.

**Parameters**
- `ByRef Source As Variant`
    - The object or value to check.
- `ParamArray Predicate() As Variant`
    - The list to check if the object or value is contained in.

**Returns**
- `Boolean`: True if the object or value is contained in the list.


### Usage

Demonstrating how the `ISIN` function could be used.

<p align="center"><sub>VBA</sub></p>

```vb
Public Sub Demo()
    Dim Value As String
    
    ' A dummy string to demonstrate how the ISIN function could be used.
    Value = "CM 233124 AACW"

    Debug.Print "NOT Using ISIN"
    If Left(Value, 2) = "DM" Or Left(Value, 2) = "DI" Or _
        Left(Value, 2) = "CM" Or Left(Value, 2) = "SM" Then
        Debug.Print True
    Else
        Debug.Print False
    End If

    Debug.Print
    Debug.Print "Using ISIN"

    ' You only have to write the LEFT function once instead of 4 times
    If ISIN(Left(Value, 2), "DM", "DI", "CM", "SM") Then
        Debug.Print True
    Else
        Debug.Print False
    End If
End Sub
```


<p align="center"><sub>Formula</sub></p>  

```vb
=ISIN("CM 233124 AACW", "DM", "DI", "CM", "SM")
```

<p align="center"><sub>Output</sub></p>  

```
True
```















