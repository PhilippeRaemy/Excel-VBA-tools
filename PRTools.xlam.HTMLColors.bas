Attribute VB_Name = "HTMLColors"
Option Explicit

Private ColorRef As Scripting.Dictionary
Private Inited As Boolean
Private Reds As Scripting.Dictionary
Private Greens As Scripting.Dictionary
Private Blues As Scripting.Dictionary

Public Sub test(r As Integer, g As Integer, b As Integer)
    Dim c As Range, Col As HtmlColor
    Set c = selection.Cells(1, 1)
    c.Interior.Color = RGB(r, g, b)
    c.value = "RGB(" & r & ", " & g & ", " & b & ")"
    Set c = selection.Cells(1, 2)
    Set Col = Rgb2Color(RGB(r, g, b))
    c.value = Col.ToString()
    c.Interior.Color = RGB(Col.Red, Col.Green, Col.Blue)
End Sub

Public Sub ttest(rr As Integer, gg As Integer, bb As Integer, stp As Integer)
Dim r As Integer, g As Integer, b As Integer

    For r = rr - 10 * stp To rr + 10 * stp Step stp
        For g = gg - 10 * stp To gg + 10 * stp Step stp
            For b = bb - 10 * stp To bb + 10 * stp Step stp
                test r, g, b
                selection.Cells(2, 1).Select
            Next
        Next
    Next

End Sub

Private Function Rgb2ColorRaw(ByVal RGB As Long) As HtmlColor
    Set Rgb2ColorRaw = New HtmlColor
    Rgb2ColorRaw.Red = RGB Mod 256
    RGB = Int(RGB / 256)
    Rgb2ColorRaw.Green = RGB Mod 256
    RGB = Int(RGB / 256)
    Rgb2ColorRaw.Blue = RGB Mod 256
End Function

Public Function GetRGB(ByVal ColorName As String) As Long
    Dim Col As HtmlColor
    Init
    If ColorRef.Exists(ColorName) Then
        Set Col = ColorRef(ColorName)
        GetRGB = RGB(Col.Red, Col.Green, Col.Blue)
    Else
        GetRGB = RGB(128, 128, 128)
    End If
End Function
Private Function Rgb2Color(ByVal RGB As Long) As HtmlColor
    Dim Searched As HtmlColor, Col As Variant
    Dim Closest As Long, Distance As Long
    Init
    Set Searched = Rgb2ColorRaw(RGB)
    Closest = -1
    For Each Col In ColorRef.Items
        Distance = (Col.Red - Searched.Red) ^ 2 _
                + (Col.Green - Searched.Green) ^ 2 _
                + (Col.Blue - Searched.Blue) ^ 2
        ' Debug.Print Distance, Closest, Col.ToString(),
        If Closest > Distance Or Closest < 0 Then
            Closest = Distance
            Set Rgb2Color = Col
            ' Debug.Print "Closest",
            If Closest = 0 Then Exit Function ' found exact!
        End If
        ' Debug.Print
    Next Col
    
End Function
Public Function Rgb2ColorName(ByVal RGB As Long) As String
    Rgb2ColorName = Rgb2Color(RGB).Name
End Function



Private Sub AddColor(ByVal Name As String, ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)
    Dim Col As HtmlColor
    Set Col = New HtmlColor
    Col.Name = Name
    Col.Red = Red
    Col.Green = Green
    Col.Blue = Blue
    ColorRef.Add Name, Col
End Sub

Private Sub Init()
    If ColorRef Is Nothing Then
        Set ColorRef = New Scripting.Dictionary
        AddColor "Pink", 255, 192, 203
        AddColor "LightPink", 255, 182, 193
        AddColor "HotPink", 255, 105, 180
        AddColor "DeepPink", 255, 20, 147
        AddColor "PaleVioletRed", 219, 112, 147
        AddColor "MediumVioletRed", 199, 21, 133
        AddColor "LightSalmon", 255, 160, 122
        AddColor "Salmon", 250, 128, 114
        AddColor "DarkSalmon", 233, 150, 122
        AddColor "LightCoral", 240, 128, 128
        AddColor "IndianRed", 205, 92, 92
        AddColor "Crimson", 220, 20, 60
        AddColor "FireBrick", 178, 34, 34
        AddColor "DarkRed", 139, 0, 0
        AddColor "Red", 255, 0, 0
        AddColor "OrangeRed", 255, 69, 0
        AddColor "Tomato", 255, 99, 71
        AddColor "Coral", 255, 127, 80
        AddColor "DarkOrange", 255, 140, 0
        AddColor "Orange", 255, 165, 0
        AddColor "Yellow", 255, 255, 0
        AddColor "LightYellow", 255, 255, 224
        AddColor "LemonChiffon", 255, 250, 205
        AddColor "LightGoldenrodYellow", 250, 250, 210
        AddColor "PapayaWhip", 255, 239, 213
        AddColor "Moccasin", 255, 228, 181
        AddColor "PeachPuff", 255, 218, 185
        AddColor "PaleGoldenrod", 238, 232, 170
        AddColor "Khaki", 240, 230, 140
        AddColor "DarkKhaki", 189, 183, 107
        AddColor "Gold", 255, 215, 0
        AddColor "Cornsilk", 255, 248, 220
        AddColor "BlanchedAlmond", 255, 235, 205
        AddColor "Bisque", 255, 228, 196
        AddColor "NavajoWhite", 255, 222, 173
        AddColor "Wheat", 245, 222, 179
        AddColor "BurlyWood", 222, 184, 135
        AddColor "Tan", 210, 180, 140
        AddColor "RosyBrown", 188, 143, 143
        AddColor "SandyBrown", 244, 164, 96
        AddColor "Goldenrod", 218, 165, 32
        AddColor "DarkGoldenrod", 184, 134, 11
        AddColor "Peru", 205, 133, 63
        AddColor "Chocolate", 210, 105, 30
        AddColor "SaddleBrown", 139, 69, 19
        AddColor "Sienna", 160, 82, 45
        AddColor "Brown", 165, 42, 42
        AddColor "Maroon", 128, 0, 0
        AddColor "DarkOliveGreen", 85, 107, 47
        AddColor "Olive", 128, 128, 0
        AddColor "OliveDrab", 107, 142, 35
        AddColor "YellowGreen", 154, 205, 50
        AddColor "LimeGreen", 50, 205, 50
        AddColor "Lime", 0, 255, 0
        AddColor "LawnGreen", 124, 252, 0
        AddColor "Chartreuse", 127, 255, 0
        AddColor "GreenYellow", 173, 255, 47
        AddColor "SpringGreen", 0, 255, 127
        AddColor "MediumSpringGreen", 0, 250, 154
        AddColor "LightGreen", 144, 238, 144
        AddColor "PaleGreen", 152, 251, 152
        AddColor "DarkSeaGreen", 143, 188, 143
        AddColor "MediumSeaGreen", 60, 179, 113
        AddColor "SeaGreen", 46, 139, 87
        AddColor "ForestGreen", 34, 139, 34
        AddColor "Green", 0, 128, 0
        AddColor "DarkGreen", 0, 100, 0
        AddColor "MediumAquamarine", 102, 205, 170
        AddColor "Aqua", 0, 255, 255
        AddColor "Cyan", 0, 255, 255
        AddColor "LightCyan", 224, 255, 255
        AddColor "PaleTurquoise", 175, 238, 238
        AddColor "Aquamarine", 127, 255, 212
        AddColor "Turquoise", 64, 224, 208
        AddColor "MediumTurquoise", 72, 209, 204
        AddColor "DarkTurquoise", 0, 206, 209
        AddColor "LightSeaGreen", 32, 178, 170
        AddColor "CadetBlue", 95, 158, 160
        AddColor "DarkCyan", 0, 139, 139
        AddColor "Teal", 0, 128, 128
        AddColor "LightSteelBlue", 176, 196, 222
        AddColor "PowderBlue", 176, 224, 230
        AddColor "LightBlue", 173, 216, 230
        AddColor "SkyBlue", 135, 206, 235
        AddColor "LightSkyBlue", 135, 206, 250
        AddColor "DeepSkyBlue", 0, 191, 255
        AddColor "DodgerBlue", 30, 144, 255
        AddColor "CornflowerBlue", 100, 149, 237
        AddColor "SteelBlue", 70, 130, 180
        AddColor "RoyalBlue", 65, 105, 225
        AddColor "Blue", 0, 0, 255
        AddColor "MediumBlue", 0, 0, 205
        AddColor "DarkBlue", 0, 0, 139
        AddColor "Navy", 0, 0, 128
        AddColor "MidnightBlue", 25, 25, 112
        AddColor "Lavender", 230, 230, 250
        AddColor "Thistle", 216, 191, 216
        AddColor "Plum", 221, 160, 221
        AddColor "Violet", 238, 130, 238
        AddColor "Orchid", 218, 112, 214
        AddColor "Fuchsia", 255, 0, 255
        AddColor "Magenta", 255, 0, 255
        AddColor "MediumOrchid", 186, 85, 211
        AddColor "MediumPurple", 147, 112, 219
        AddColor "BlueViolet", 138, 43, 226
        AddColor "DarkViolet", 148, 0, 211
        AddColor "DarkOrchid", 153, 50, 204
        AddColor "DarkMagenta", 139, 0, 139
        AddColor "Purple", 128, 0, 128
        AddColor "Indigo", 75, 0, 130
        AddColor "DarkSlateBlue", 72, 61, 139
        AddColor "SlateBlue", 106, 90, 205
        AddColor "MediumSlateBlue", 123, 104, 238
        AddColor "White", 255, 255, 255
        AddColor "Snow", 255, 250, 250
        AddColor "Honeydew", 240, 255, 240
        AddColor "MintCream", 245, 255, 250
        AddColor "Azure", 240, 255, 255
        AddColor "AliceBlue", 240, 248, 255
        AddColor "GhostWhite", 248, 248, 255
        AddColor "WhiteSmoke", 245, 245, 245
        AddColor "Seashell", 255, 245, 238
        AddColor "Beige", 245, 245, 220
        AddColor "OldLace", 253, 245, 230
        AddColor "FloralWhite", 255, 250, 240
        AddColor "Ivory", 255, 255, 240
        AddColor "AntiqueWhite", 250, 235, 215
        AddColor "Linen", 250, 240, 230
        AddColor "LavenderBlush", 255, 240, 245
        AddColor "MistyRose", 255, 228, 225
        AddColor "Gainsboro", 220, 220, 220
        AddColor "LightGrey", 211, 211, 211
        AddColor "Silver", 192, 192, 192
        AddColor "DarkGray", 169, 169, 169
        AddColor "Gray", 128, 128, 128
        AddColor "DimGray", 105, 105, 105
        AddColor "LightSlateGray", 119, 136, 153
        AddColor "SlateGray", 112, 128, 144
        AddColor "DarkSlateGray", 47, 79, 79
        AddColor "Black", 0, 0, 0
    End If
End Sub
