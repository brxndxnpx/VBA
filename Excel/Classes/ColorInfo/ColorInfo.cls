Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     A color class that contains metadata for RGB, hex, and Microsoft Office's integer (long) color code.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Type TColorInfo
    IsConfiguring   As Boolean
    R               As Long
    G               As Long
    B               As Long
    HexCode         As String
    ColorCode       As Long
End Type

Private this  As TColorInfo


''' Summary:
'''     The Excel color code for the color.
Public Property Get ColorCode() As Long: ColorCode = this.ColorCode: End Property
Public Property Let ColorCode(ByVal Value As Long): this.ColorCode = Value: SetFromColorCode: End Property

''' Summary:
'''     The R part of the RGB color
Public Property Get R() As Long: R = this.R: End Property
Public Property Let R(ByVal Value As Long): this.R = Value: SetFromRGB: End Property

''' Summary:
'''     The G part of the RGB color
Public Property Get G() As Long: G = this.G: End Property
Public Property Let G(ByVal Value As Long): this.G = Value: SetFromRGB: End Property

''' Summary:
'''     The B part of the RGB color
Public Property Get B() As Long: B = this.B: End Property
Public Property Let B(ByVal Value As Long): this.B = Value: SetFromRGB: End Property

''' Summary:
'''     The hex code for the color.
Public Property Get HexCode() As String: HexCode = this.HexCode: End Property
Public Property Let HexCode(ByVal Value As String): this.HexCode = Value: SetFromHexCode: End Property


Public Sub SetRGBValues(ByVal r_ As Long, ByVal g_ As Long, ByVal b_ As Long)
    this.R = r_
    this.G = g_
    this.B = b_
    
    SetFromRGB
End Sub

Private Sub SetFromRGB()
    If this.IsConfiguring Then Exit Sub
    this.IsConfiguring = True

    SetHexCode
    SetColorCode
    
    this.IsConfiguring = False
End Sub

Private Sub SetFromHexCode()
    If this.IsConfiguring Then Exit Sub
    this.IsConfiguring = True
    
    SetRGB HexCode
    SetColorCode
    
    this.IsConfiguring = False
End Sub

Private Sub SetFromColorCode()
    If this.IsConfiguring Then Exit Sub
    this.IsConfiguring = True
    
    SetRGB ColorCode
    SetHexCode
    
    this.IsConfiguring = False
End Sub

''' Summary:
'''     Sets the hex code
'''     * Depends on RGB
Private Sub SetHexCode()
    HexCode = "#" & Right$("00" & Hex$(R), 2) & Right$("00" & Hex$(G), 2) & Right$("00" & Hex$(B), 2)
End Sub

''' Summary:
'''     Sets the color code
'''     * Depends on RGB
Private Sub SetColorCode()
    ColorCode = RGB(R, G, B)
End Sub

''' Summary:
'''     Sets the RGB values depending on either the hex code or the color code
Private Sub SetRGB(ByRef Value As Variant)
    If TypeName(Value) = "String" Then
        Dim hexColor As String: hexColor = Right$("000000" & Replace(Value, "#", vbNullString), 6)
    
        R = Val("&H" & Mid$(hexColor, 1, 2))
        G = Val("&H" & Mid$(hexColor, 3, 2))
        B = Val("&H" & Mid$(hexColor, 5, 2))
    ElseIf TypeName(Value) = "Long" Then
        R = (Value Mod 256)
        G = (Value \ 256) Mod 256
        B = (Value \ 65536) Mod 256
    End If
End Sub


