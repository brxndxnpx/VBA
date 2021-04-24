Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     Demonstrates using the ColorInfo class by setting the color by the hex code, RGB values, or the Microsoft Office color code.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub DEMO_ColorInfo()
    DEMO_ColorByHex "#002060" 'Dark Blue
    
    DEMO_ColorByRGB 255, 0, 0 'Dark Red

    DEMO_ColorByColorCode 5288448 'Green
End Sub

''' Summary:
'''     Demonstrates using the ColorInfo class by setting the color by the hex code.
Private Sub DEMO_ColorByHex(Optional ByVal hexCode_ As String = "#FFFFFF")
    Dim Color As New ColorInfo
    
    ' Set the color by hex code
    Color.HexCode = hexCode_
    
    DebugResult_ColorInfo "DEMO_ColorByHex", Color
End Sub

''' Summary:
'''     Demonstrates using the ColorInfo class by setting the color by the RGB values.
Private Sub DEMO_ColorByRGB(Optional ByVal r_ As Long = 255, Optional ByVal g_ As Long = 255, Optional ByVal b_ As Long = 255)
    Dim Color As New ColorInfo
    
    ' Set the color by RGB values
    Color.SetRGBValues r_, g_, b_
    
    DebugResult_ColorInfo "DEMO_ColorByRGB", Color
End Sub

''' Summary:
'''     Demonstrates using the ColorInfo class by setting the color by the Microsoft Office color code.
Private Sub DEMO_ColorByColorCode(Optional ByVal colorCode_ As Long = 16777215)
    Dim Color As New ColorInfo
    
    ' Set the color by MS color code
    Color.ColorCode = colorCode_
    
    DebugResult_ColorInfo "DEMO_ColorByColorCode", Color
End Sub

''' Summary:
'''     Writes the results to the immediate window.
Private Sub DebugResult_ColorInfo(ByVal callerMember_ As String, ByRef color_ As ColorInfo)
    Debug.Print callerMember_
    Debug.Print "RGB: " & color_.R & ", " & color_.G; ", " & color_.B
    Debug.Print "ColorCode: " & color_.ColorCode
    Debug.Print "HexCode: " & color_.HexCode
    Debug.Print vbNewLine
End Sub






