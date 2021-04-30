# User Defined Functions

## Functions

|                           | Type     | Description                                  |
|---------------------------|----------|----------------------------------------------|
| [`QUOTE`](#quote)         | `String` | Quotes a text using the character provided.  |
| [`REVERSE`](#reverse)     | `String` | Reverses text.                               |
| [`CHARCOUNT`](#charcount) | `Long`   | Counts the number of characters in a string. |


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

