Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     A class created for parsing JSON strings. Stores the object in a dictionary.
'''     Stores the object in a dictionary.
'''     Uses dot notation to access nested properties
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private p As Long

Private Type TJsonConverter
    Tokens As Variant
    Obj As Object
End Type

Private this As TJsonConverter

'   TODO: Add a serialize function
'''' Summary
''''     Converts an object into a JSON string.
'Public Function Serialize(Source As Object) As String
'
'
'End Function


''' Summary
'''     Converts JSON into a dictionary resembling the object.
''' Parameters:
'''     ByVal Json As String: The json text to convert.
'''     Optional ByVal Key As String: The optional base key.
''' Returns:
'''     An Object; A Scripting.Dictionary object resembling the json data.
Public Function Deserialize(ByVal Json As String, Optional ByVal Key As String = "") As Object
    p = 1

    ' Tokenize the json
    this.Tokens = Tokenize(Json)

    ' Create a dictionary to house the object and properties
    Set this.Obj = CreateObject("Scripting.Dictionary")

    ' Start reading and storing the tokenized json into the dictionary
    If this.Tokens(p) = "{" Then ParseObj Key Else ParseArr Key

    Set Deserialize = this.Obj
End Function


Private Function ParseObj(Key As String)
    Do: p = p + 1
        Select Case this.Tokens(p)
            Case "]"
            Case "[":  ParseArr Key
            Case "{"
                       If this.Tokens(p + 1) = "}" Then
                           p = p + 1
                           this.Obj.Add Key, "null"
                       Else
                           ParseObj Key
                       End If
                
            Case "}":  Key = ReducePath(Key): Exit Do
            'Case ":":  key = key & "." & this.Tokens(p - 1)
            Case ":":  Key = Key & IIf(Key = vbNullString, vbNullString, ".") & this.Tokens(p - 1)
            
            Case ",":  Key = ReducePath(Key)
            Case Else: If this.Tokens(p + 1) <> ":" Then this.Obj.Add Key, this.Tokens(p)
        End Select
    Loop
End Function

Private Function ParseArr(Key As String)
    Dim e As Long
    Do: p = p + 1
        Select Case this.Tokens(p)
            Case "}"
            Case "{":  ParseObj Key & ArrayID(e)
            Case "[":  ParseArr Key
            Case "]":  Exit Do
            Case ":":  Key = Key & ArrayID(e)
            Case ",":  e = e + 1
            Case Else: this.Obj.Add Key & ArrayID(e), this.Tokens(p)
        End Select
    Loop
End Function

''' Summary:
'''     Tokenizes a json string.
''' Parameters:
'''     ByVal Value As String: The json text to tokenize.
''' Returns:
'''     A Variant; Returns the results in an array with each record being a token.
Private Function Tokenize(s As String)
    Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    Tokenize = ExecuteRegex(s, Pattern, True)
End Function

''' Summary:
'''     Executes regular expression queries.
''' Parameters:
'''     ByVal Value As String: The text to parse.
'''     ByVal Pattern As String: The regular expression.
'''     ByVal Group1Bias As String:
'''     ByVal UseGlobal As String: Whether or not to the global regex setting.
''' Returns:
'''     A Variant; Returns the results in an array.
Private Function ExecuteRegex(ByVal Value As String, ByVal Pattern As String, _
    Optional ByVal Group1Bias As Boolean, Optional ByVal UseGlobal As Boolean = True)
    
  Dim i As Long
  Dim Groups, Group, Values

  With CreateObject("vbscript.regexp")
    .Global = UseGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .TEST(Value) Then
      Set Groups = .Execute(Value)
      ReDim Values(1 To Groups.Count)
      For Each Group In Groups
        i = i + 1
        Values(i) = Group.Value
        If Group1Bias Then If Len(Group.submatches(0)) Or Group.Value = """""" Then Values(i) = Group.submatches(0)
      Next
    End If
  End With

  ExecuteRegex = Values
End Function

Private Function ArrayID(e) As String
    ArrayID = "(" & e & ")"
End Function

Private Function ReducePath(Key As String) As String
    If InStr(Key, ".") Then ReducePath = Left(Key, InStrRev(Key, ".") - 1) Else ReducePath = Key
End Function


''' Summary:
'''     Prints each path and it's value to the immediate window.
''' Parameters:
'''     Optional ByRef Obj As Object: A Scripting.Dictionary object.
'''         Will use the previously parse json object if left empty.
Sub PrintPaths(Optional ByRef Source As Object = Nothing)
    Dim Key
    
    If Source Is Nothing Then Set Source = this.Obj
    
    For Each Key In Source
        If Len(Key) < 14 Then
            Debug.Print Key, "", "", Source(Key)
        ElseIf Len(Key) < 28 Then
            Debug.Print Key, "", Source(Key)
        Else
            Debug.Print Key, Source(Key)
        End If
    Next
End Sub




' Function GetFilteredValues(Obj, match)
'     Dim c As Long
'     Dim i As Long
'     Dim v
'     Dim w

'     v = Obj.Keys
'     ReDim w(1 To Obj.Count)
'     For i = 0 To UBound(v)
'         If v(i) Like match Then
'             c = c + 1
'             w(c) = Obj(v(i))
'         End If
'     Next
'     ReDim Preserve w(1 To c)
'     GetFilteredValues = w
' End Function

' Function GetFilteredTable(Obj, cols)
'     Dim c As Long
'     Dim i As Long
'     Dim j As Long
'     Dim v, w, z

'     v = Obj.Keys
'     z = GetFilteredValues(Obj, cols(0))
'     ReDim w(1 To UBound(z), 1 To UBound(cols) + 1)
'     For j = 1 To UBound(cols) + 1
'          z = GetFilteredValues(Obj, cols(j - 1))
'          For i = 1 To UBound(z)
'             w(i, j) = z(i)
'          Next
'     Next
'     GetFilteredTable = w
' End Function

' Function OpenTextFile(f) As String
'     With CreateObject("ADODB.Stream")
'         .Charset = "utf-8"
'         .Open
'         .LoadFromFile f
'         OpenTextFile = .ReadText
'     End With
' End Function






