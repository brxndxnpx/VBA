
# Type Declaration Characters

[MSDN Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary)


| Type           | Short Hand Declaration | Long Hand Declaration | Documentation |
|----------------|:----------------------:|-----------------------|---------------|
| `String`       |           $            | `Dim a As String`     |               |
| `Long`         |           &            | `Dim a As Long`       |               |
| `LongLong`     |           ^            | `Dim a As LongLong`   |               |
| `Integer`      |           %            | `Dim a As Integer`    |               |
| `Single`       |           !            | `Dim a As Single`     |               |
| `Currency`     |           @            | `Dim a As Currency`   |               |
| `Byte`         |          None          | `Dim a As Byte`       |               |
| `Date`         |         #Date#         | `Dim a As Date`       |               |
| `Double`       |           #            | `Dim a As Double`     |               |
| `Boolean`      |          None          | `Dim a As Boolean`    |               |
| `Variant`      |          None          | `Dim a As Variant`    |               |
| `Decimal`      |          None          | `Dim a As Decimal`    |               |
| `LongPtr`      |          None          | `Dim a As LongPtr`    |               |
| `Object`       |          None          | `Dim a As Object`     |               |
| `User-defined` |                        |                       |               |



```vb
Dim e&  ' Using shorthand type-declaration
Dim e As Long ' Using longhand type-declaration
```

---

## String

[MSDN Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/string-data-type)

The type-declaration character for String is the dollar ($) sign.


```vb
Private Sub Demo()
    DemoFunc "hello"
End Sub

Private Function DemoFunc(Value$)
    Debug.Print Value
    ' Prints: hello
End Function
```

```vb
' This function returns a String
Function DemoFunc$(Value$)
    DemoFunc = Value
End Function

' This is the same as the function above
Function DemoFunc(Value As String) As String
    DemoFunc = Value
End Function
```





## Long

[MSDN Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/long-data-type)

The type-declaration character for Long is the ampersand (&).

```vb
Private Sub Demo()
    DemoFunc 20
End Sub

Private Function DemoFunc(Value$)
    Debug.Print Value
    ' Prints: 20
End Function
```

## LongLong

[MSDN Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/longlong-data-type)

The type-declaration character for LongLong is the caret (^).
    
```vb
Private Sub Demo()
    DemoFunc 20
End Sub

Private Function DemoFunc(Value^)
    Debug.Print Value
    ' Prints: 20
End Function
```

## Single

[MSDN Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/single-data-type)

The type-declaration character for Single is the exclamation point (!).
    

```vb
Private Sub Demo()
    DemoFunc 20
End Sub

Private Function DemoFunc(Value!)
    Debug.Print Value
    ' Prints: 20
End Function
```

## Currency

[MSDN Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/currency-data-type)
The type-declaration character for Currency is the at (@) sign.
  

```vb
Private Sub Demo()
    DemoFunc 20
End Sub

Private Function DemoFunc(Value@)
    Debug.Print Value
    ' Prints: 20
End Function
```

## Date

[MSDN Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/date-data-type)

Any recognizable literal date values can be assigned to Date variables. Date literals must be enclosed within number signs (#), for example, #January 1, 1993# or #1 Jan 93#.
    
```vb
Private Sub Demo()
    DemoFunc #04-01-2021#
End Sub

Private Function DemoFunc(Value As Date)
    Debug.Print Value
    ' Prints: 04-01-2021
End Function
```

## Double

[MSDN Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/double-data-type)

The type-declaration character for Double is the number (#) sign.
    

```vb
Private Sub Demo()
    DemoFunc 20
End Sub

Private Function DemoFunc(Value#)
    Debug.Print Value
    ' Prints: 20
End Function
```

## Integer

[MSDN Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/integer-data-type)

The type-declaration character for Integer is the percent (%) sign.
    


```vb
Private Sub Demo()
    DemoFunc 20
End Sub

Private Function DemoFunc(Value%)
    Debug.Print Value
    ' Prints: 20
End Function
```

## Boolean

[MSDN Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/boolean-data-type)

Boolean variables display as either:
  True or False (when Print is used), or
  #TRUE# or #FALSE# (when Write # is used).
    

```vb
Private Sub Demo()
    DemoFunc True
End Sub

Private Function DemoFunc(Value As Boolean)
    Debug.Print Value
    ' Prints: True
End Function
```

## Variant

[MSDN Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/variant-data-type)

The Variant data type has no type-declaration character.
    

```vb
Private Sub Demo()
    DemoFunc True
End Sub

Private Function DemoFunc(Value)
    Debug.Print Value
    ' Throws a type mismatched error
End Function
```    