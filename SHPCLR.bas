Attribute VB_Name = "mdlSHPCLR"
Option Explicit

    Public Const HalfPi   As Double = 1.5757963267949
    Public Const Pi       As Double = 3.14159265358979
    Public Const TwoPi    As Double = 6.28318520717959
    
    Public Enum bvlSHAPE
        [Rectangular]
        [Circular]
        [Rounded Rectangle]
        [Capsule]
    End Enum

    Public Enum bvlSLOPE
        [Straight]
        [Concave]
        [Convex]
    End Enum

    Const HSLMAX As Integer = 240
    Const RGBMAX As Integer = 255
    Const UNDEFINED As Integer = (HSLMAX * 2 / 3)
    
    Type HSLCol
        Hue As Integer
        Sat As Integer
        Lum As Integer
    End Type
    
    ' In this declaration for the RECT type I have used the vernacular
    ' of compass points so as not to confuse Left, Right etc. with
    ' VB functions or logical coordinates system.
    Type RECT
        West As Long
        North As Long
        East As Long
        South As Long
    End Type
    
    Type POINTAPI
        x As Long
        y As Long
    End Type
    
    Declare Function CreateRectRgn Lib "gdi32" _
                    (ByVal X1 As Long, ByVal Y1 As Long, _
                     ByVal X2 As Long, ByVal Y2 As Long) As Long
                     
    Declare Function Rectangle Lib "gdi32" _
                    (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
                     ByVal X2 As Long, ByVal Y2 As Long) As Long
    
    Declare Function CreateRoundRectRgn Lib "gdi32" _
                    (ByVal X1 As Long, ByVal Y1 As Long, _
                     ByVal X2 As Long, ByVal Y2 As Long, _
                     ByVal X3 As Long, ByVal Y3 As Long) As Long
    
    Declare Function RoundRect Lib "gdi32" _
                    (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
                     ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
                     ByVal Y3 As Long) As Long
    
    Declare Function CreateEllipticRgn Lib "gdi32" _
                    (ByVal X1 As Long, ByVal Y1 As Long, _
                     ByVal X2 As Long, ByVal Y2 As Long) As Long
    
    Declare Function SetWindowRgn Lib "user32" _
                    (ByVal hWnd As Long, ByVal hRgn As Long, _
                     ByVal bRedraw As Boolean) As Long
    
    Declare Function CombineRgn Lib "gdi32" _
                    (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, _
                     ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
    Public Const RGN_AND = 1
    Public Const RGN_OR = 2
    Public Const RGN_XOR = 3
    Public Const RGN_DIFF = 4
    Public Const RGN_COPY = 5
    Public Const RGN_MIN = RGN_AND
    Public Const RGN_MAX = RGN_COPY
    
    Declare Function PtInRegion Lib "gdi32" _
                    (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
    
    Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    
    Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
    
    Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    
    Declare Function GetSysColor Lib "user32" _
                    (ByVal nIndex As Long) As Long
    
    Declare Function BitBlt Lib "gdi32" _
                    (ByVal hDestDC As Long, _
                     ByVal x As Long, ByVal y As Long, _
                     ByVal nWidth As Long, ByVal nHeight As Long, _
                     ByVal hSrcDC As Long, _
                     ByVal xSrc As Long, ByVal ySrc As Long, _
                     ByVal dwRop As Long) As Long
    
    Declare Function StretchBlt Lib "gdi32" _
                    (ByVal hdc As Long, _
                     ByVal x As Long, ByVal y As Long, _
                     ByVal nWidth As Long, ByVal nHeight As Long, _
                     ByVal hSrcDC As Long, _
                     ByVal xSrc As Long, ByVal ySrc As Long, _
                     ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
                     ByVal dwRop As Long) As Long
    
    Public Const DT_BOTTOM = &H8
    Public Const DT_CALCRECT = &H400
    Public Const DT_CENTER = &H1
    Public Const DT_LEFT = &H0
    Public Const DT_RIGHT = &H2
    Public Const DT_SINGLELINE = &H20
    Public Const DT_TOP = &H0
    Public Const DT_VCENTER = &H4
    
    Declare Function DrawText Lib "user32" Alias "DrawTextA" _
                    (ByVal hdc As Long, ByVal lpStr As String, _
                     ByVal nCount As Long, lpRect As RECT, _
                     ByVal wFormat As Long) As Long
    
    Declare Function GetCursorPos Lib "user32" _
                    (lpPoint As POINTAPI) As Long
    
    Declare Function WindowFromPoint Lib "user32" _
                    (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    
    Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" _
                    (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
    Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
    End Type



Function WHOLE_NUMBER(ByRef NUMBER As Variant) As Boolean
    WHOLE_NUMBER = IIf(NUMBER = CLng(NUMBER), True, False)
End Function

Function iMax(A As Integer, B As Integer) As Integer
    'Return the Larger of two values
    iMax = IIf(A > B, A, B)
End Function

Function iMin(A As Integer, B As Integer) As Integer
    'Return the smaller of two values
    iMin = IIf(A < B, A, B)
End Function

'==============================================================='
'                                                               '
' Credit to Dan Redding of Blue Knot Software for this routine. '
'                                                               '
'==============================================================='
Function RGBtoHSL(RGBCol As Long) As HSLCol '***
    'Returns an HSLCol datatype containing Hue, Luminescence
    'and Saturation; given an RGB Color value
    Dim R As Integer, G As Integer, B As Integer
    Dim cMax As Integer, cMin As Integer
    Dim RDelta As Double, GDelta As Double, _
    BDelta As Double
    Dim H As Double, S As Double, L As Double
    Dim cMinus As Long, cPlus As Long
    
    R = RGBCol And &HFF
    G = (RGBCol And &H100FF00) / &H100
    B = (RGBCol And &HFF0000) / &H10000
    
    cMax = iMax(iMax(R, G), B) 'Highest and lowest
    cMin = iMin(iMin(R, G), B) 'color values
    
    cMinus = cMax - cMin 'Used to simplify the
    cPlus = cMax + cMin 'calculations somewhat.
    
    'Calculate luminescence (lightness)
    L = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX)
    


    If cMax = cMin Then 'achromatic (r=g=b, greyscale)
        S = 0 'Saturation 0 for greyscale
        H = UNDEFINED 'Hue undefined for greyscale
    Else
        'Calculate color saturation


        If L <= (HSLMAX / 2) Then
            S = ((cMinus * HSLMAX) + 0.5) / cPlus
        Else
            S = ((cMinus * HSLMAX) + 0.5) / (2 * RGBMAX - cPlus)
        End If
        
        'Calculate hue
        RDelta = (((cMax - R) * (HSLMAX / 6)) + 0.5) / cMinus
        GDelta = (((cMax - G) * (HSLMAX / 6)) + 0.5) / cMinus
        BDelta = (((cMax - B) * (HSLMAX / 6)) + 0.5) / cMinus
        


        Select Case cMax
            Case CLng(R)
            H = BDelta - GDelta
            Case CLng(G)
            H = (HSLMAX / 3) + RDelta - BDelta
            Case CLng(B)
            H = ((2 * HSLMAX) / 3) + GDelta - RDelta
        End Select
        
        If H < 0 Then H = H + HSLMAX
    End If
    
    RGBtoHSL.Hue = CInt(H)
    RGBtoHSL.Lum = CInt(L)
    RGBtoHSL.Sat = CInt(S)

End Function

'==============================================================='
'                                                               '
' Credit to Dan Redding of Blue Knot Software for this routine. '
'                                                               '
'==============================================================='
Function HSLtoRGB(HueLumSat As HSLCol) As Long '***
    Dim R As Long, G As Long, B As Long
    Dim H As Long, L As Long, S As Long
    Dim Magic1 As Integer, Magic2 As Integer
    H = HueLumSat.Hue
    L = HueLumSat.Lum
    S = HueLumSat.Sat


    If S = 0 Then 'Greyscale
        R = (L * RGBMAX) / HSLMAX 'luminescence,
        'converted to the proper range
        G = R 'All RGB values same in greyscale
        B = R


        If H <> UNDEFINED Then
            'This is technically an error.
            'The RGBtoHSL routine will always return
            '
            'Hue = UNDEFINED (in this case 160)
            'when Sat = 0.
            'if you are writing a color mixer and
            'letting the user input color values,
            'you may want to set Hue = UNDEFINED
            'in this case.
        End If
    Else
        'Get the "Magic Numbers"


        If L <= HSLMAX / 2 Then
            Magic2 = (L * (HSLMAX + S) + _
            (HSLMAX / 2)) / HSLMAX
        Else
            Magic2 = L + S - ((L * S) + _
            (HSLMAX / 2)) / HSLMAX
        End If
        Magic1 = 2 * L - Magic2
        'get R, G, B; change units from HSLMAX r
        '     ange
        'to RGBMAX range
        R = (HuetoRGB(Magic1, Magic2, H + (HSLMAX / 3)) _
        * RGBMAX + (HSLMAX / 2)) / HSLMAX
        G = (HuetoRGB(Magic1, Magic2, H) _
        * RGBMAX + (HSLMAX / 2)) / HSLMAX
        B = (HuetoRGB(Magic1, Magic2, H - (HSLMAX / 3)) _
        * RGBMAX + (HSLMAX / 2)) / HSLMAX
    End If
    HSLtoRGB = RGB(CInt(R), CInt(G), CInt(B))
End Function

'==============================================================='
'                                                               '
' Credit to Dan Redding of Blue Knot Software for this routine. '
'                                                               '
'==============================================================='
Function HuetoRGB(mag1 As Integer, mag2 As Integer, _
    Hue As Long) As Long '***
    'Utility function for HSLtoRGB
    'Range check


    If Hue < 0 Then
        Hue = Hue + HSLMAX
    ElseIf Hue > HSLMAX Then
        Hue = Hue - HSLMAX
    End If
    'Return r, g, or b value from parameters
    '
    Select Case Hue 'Values get progressively larger.
        'Only the first true condition will exec
        '     ute
        Case Is < (HSLMAX / 6)
        HuetoRGB = (mag1 + (((mag2 - mag1) * Hue + _
        (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Is < (HSLMAX / 2)
        HuetoRGB = mag2
        Case Is < (HSLMAX * 2 / 3)
        HuetoRGB = (mag1 + (((mag2 - mag1) * _
        ((HSLMAX * 2 / 3) - Hue) + _
        (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Else
        HuetoRGB = mag1
    End Select
End Function

'======================================'
'                                      '
' Return an RGB color value, even if a '
' system colour identifier was passed. '
'                                      '
'======================================'
Function RGBColour(lColour As Long) As Long
    If lColour >= 0 Then 'Already is an RGB Colour
        RGBColour = lColour
    Else 'This is a System colour, get the current RGB colour value
        RGBColour = GetSysColor(lColour And &HFFFFFF)
    End If
End Function

'========================================'
'                                        '
' Test a String to determine if it       '
' contains Valid Cartesian Co-ordinates. '
'                                        '
'========================================'
Public Function Cartesian(ByRef strXY As String, _
                           ByRef Xvalue As Integer, _
                           ByRef Yvalue As Integer) As Boolean
    Dim Cntr As Integer, Cmma As Integer
    'Enable Error Handling
    On Error GoTo Cartesian_ERROR
    'Set default function value
    Cartesian = False
    'Remove leading, trailing and embedded spaces.
    strXY = Trim(strXY)
    For Cntr = 1 To Len(strXY)
        If Mid(strXY, Cntr, 1) = " " _
        Then strXY = Left(strXY, Cntr - 1) & Mid(strXY, Cntr + 1)
    Next Cntr
    'Attempt to fix any format errors such as missing brackets,
    'missing comma or just a single coordinate entered
    If strXY = "" Then strXY = "(0,0)"
    If Left(strXY, 1) = "," Then strXY = "(0," & Mid(strXY, 2)
    If Left(strXY, 1) <> "(" Then strXY = "(" & strXY
    If InStr(strXY, ",") = 0 Then
        If Right(strXY, 1) = ")" Then
            strXY = Mid(strXY, 1, Len(strXY) - 1) & ",0"
        Else
            strXY = strXY & ",0)"
        End If
    End If
    If Right(strXY, 1) = "," Then strXY = strXY & "0)"
    If Right(strXY, 1) <> ")" Then strXY = strXY & ")"
    'Test that the string now represents cartesian co-ordinates in the
    'format "(x,y)", where both x and y are Integer values
    Cmma = InStr(strXY, ",")
    If Left(strXY, 1) = "(" And Right(strXY, 1) = ")" And Cmma > 2 Then
        If IsNumeric(Mid(strXY, 2, Cmma - 2)) _
        And IsNumeric(Mid(strXY, Cmma + 1, Len(strXY) - Cmma - 1)) Then
            Xvalue = CInt(Mid(strXY, 2, Cmma - 2))
            Yvalue = CInt(Mid(strXY, Cmma + 1, Len(strXY) - Cmma - 1))
            Cartesian = True
            strXY = "(" & LTrim(Str(Xvalue)) & "," & LTrim(Str(Yvalue)) & ")"
        Else
            Err.Raise vbObjectError + 2
        End If
    Else
        Err.Raise vbObjectError + 1
    End If
    Exit Function
Cartesian_ERROR:
    Select Case Err.NUMBER
        Case 6:
            MsgBox "INVALID CARTESIAN CO-ORDINATES VALUES." & vbCrLf & vbCrLf & _
                   "Values must be Integers in the" & vbCrLf & _
                   "range -32768 to +32767 only.", vbExclamation, _
                   "Cartesian Co-ordinates Validation"
        Case vbObjectError + 1
            MsgBox "INVALID CARTESIAN CO-ORDINATES FORMAT." & vbCrLf & vbCrLf & _
                   "Co-ordinates must be in the format" & vbCrLf & _
                   """(x,y)"" where x and y are Integers" & vbCrLf & _
                   "in the range -32768 to +32767.", vbExclamation, _
                   "Cartesian Co-ordinates Validation"
        Case vbObjectError + 2
            MsgBox "NON-NUMERIC CARTESIAN CO-ORDINATES." & vbCrLf & vbCrLf & _
                   "Co-ordinates must numeric Integers" & vbCrLf & _
                   "in the range -32768 to +32767.", vbExclamation, _
                   "Cartesian Co-ordinates Validation"
        Case Else
            MsgBox Err.Description, vbExclamation, _
                   "Cartesian Co-ordinates Validation"
    End Select
    Err.Clear
End Function

Public Function MouseIsOver(OBJhWnd As Long) As Boolean
    Dim CPos As POINTAPI
    Call GetCursorPos(CPos)
    MouseIsOver = IIf(OBJhWnd = WindowFromPoint(CPos.x, CPos.y), True, False)
End Function

'====================================='
'                                     '
' Create a Shaped Bevel for an Object '
' and return a handle to the region.  '
'                                     '
'====================================='
Public Function Bevel_REGION(ByRef drbPIC As PictureBox, _
                             ByRef drbBvlHeight As Integer, _
                             Optional ByVal drbCOLOUR As OLE_COLOR = &H8000000F, _
                             Optional ByVal drbSHAPE As bvlSHAPE = [Rectangular], _
                             Optional ByVal drbSLOPE As bvlSLOPE = [Straight], _
                             Optional ByVal drbRECESS As Boolean = False, _
                             Optional ByVal drbFRAME As Boolean = False) _
                As Long

    Dim RadianANGLE As Double
    Dim RADIUS As Single
    Dim InnerRADIUS As Single, INNERhndl As Long
    Dim OuterRADIUS As Single, OUTERhndl As Long
    Dim ORIGIN As RECT
    Dim pntX As Integer, pntY As Integer
    Dim Ocolour As HSLCol, Ncolour As HSLCol
    
    drbPIC.ScaleMode = 3 'Ensure that the Scale is PIXELS
    Select Case drbSHAPE
        Case [Rectangular]
            RADIUS = IIf(drbPIC.ScaleHeight < drbPIC.ScaleWidth, CLng(drbPIC.ScaleHeight / 2), CLng(drbPIC.ScaleWidth / 2))
            OuterRADIUS = Int(RADIUS / 2)
            If drbBvlHeight = 0 Then
                InnerRADIUS = OuterRADIUS - Int(RADIUS / 4)
                drbBvlHeight = OuterRADIUS - InnerRADIUS + 1
            Else
                InnerRADIUS = OuterRADIUS - drbBvlHeight + 1
            End If
            ORIGIN.West = OuterRADIUS
            ORIGIN.North = OuterRADIUS
            ORIGIN.East = drbPIC.ScaleWidth - ORIGIN.West - 1
            ORIGIN.South = drbPIC.ScaleHeight - ORIGIN.North - 1
            OUTERhndl = CreateRectRgn(0, 0, _
                        CLng(drbPIC.ScaleWidth + 1), _
                        CLng(drbPIC.ScaleHeight + 1))
            INNERhndl = CreateRectRgn(drbBvlHeight, drbBvlHeight, _
                        CLng(drbPIC.ScaleWidth - drbBvlHeight + 1), _
                        CLng(drbPIC.ScaleHeight - drbBvlHeight + 1))
        Case [Circular]
            OuterRADIUS = Int(drbPIC.ScaleWidth / 2)
            If drbBvlHeight = 0 Then
                InnerRADIUS = OuterRADIUS - Int(drbPIC.ScaleWidth / 8)
                drbBvlHeight = OuterRADIUS - InnerRADIUS + 1
            Else
                InnerRADIUS = OuterRADIUS - drbBvlHeight + 1
            End If
            ORIGIN.West = OuterRADIUS
            ORIGIN.North = OuterRADIUS
            ORIGIN.East = drbPIC.ScaleWidth - ORIGIN.West - 1
            ORIGIN.South = drbPIC.ScaleHeight - ORIGIN.North - 1
            OUTERhndl = CreateEllipticRgn(0, 0, _
                        CLng(drbPIC.ScaleWidth + 1), _
                        CLng(drbPIC.ScaleHeight + 1))
            INNERhndl = CreateEllipticRgn(drbBvlHeight, drbBvlHeight, _
                        CLng(drbPIC.ScaleWidth - drbBvlHeight + 1), _
                        CLng(drbPIC.ScaleHeight - drbBvlHeight + 1))
        Case [Rounded Rectangle]
            RADIUS = IIf(drbPIC.ScaleHeight < drbPIC.ScaleWidth, CLng(drbPIC.ScaleHeight / 2), CLng(drbPIC.ScaleWidth / 2))
            OuterRADIUS = Int(RADIUS / 2)
            If drbBvlHeight = 0 Then
                InnerRADIUS = OuterRADIUS - Int(RADIUS / 4)
                drbBvlHeight = OuterRADIUS - InnerRADIUS + 1
            Else
                InnerRADIUS = OuterRADIUS - drbBvlHeight + 1
            End If
            ORIGIN.West = OuterRADIUS
            ORIGIN.North = OuterRADIUS
            ORIGIN.East = drbPIC.ScaleWidth - ORIGIN.West - 1
            ORIGIN.South = drbPIC.ScaleHeight - ORIGIN.North - 1
            OUTERhndl = CreateRoundRectRgn(0, 0, _
                        CLng(drbPIC.ScaleWidth + 1), _
                        CLng(drbPIC.ScaleHeight + 1), RADIUS, RADIUS)
            INNERhndl = CreateRoundRectRgn(drbBvlHeight, drbBvlHeight, _
                        CLng(drbPIC.ScaleWidth - drbBvlHeight + 1), _
                        CLng(drbPIC.ScaleHeight - drbBvlHeight + 1), _
                        CLng(RADIUS - 2 * drbBvlHeight), _
                        CLng(RADIUS - 2 * drbBvlHeight))
        Case [Capsule]
            RADIUS = IIf(drbPIC.ScaleHeight < drbPIC.ScaleWidth, CLng(drbPIC.ScaleHeight), CLng(drbPIC.ScaleWidth))
            OuterRADIUS = Int(RADIUS / 2)
            If drbBvlHeight = 0 Then
                InnerRADIUS = OuterRADIUS - Int(RADIUS / 8)
                drbBvlHeight = OuterRADIUS - InnerRADIUS + 1
            Else
                InnerRADIUS = OuterRADIUS - drbBvlHeight + 1
            End If
            ORIGIN.West = OuterRADIUS
            ORIGIN.North = OuterRADIUS
            ORIGIN.East = drbPIC.ScaleWidth - ORIGIN.West - 1
            ORIGIN.South = drbPIC.ScaleHeight - ORIGIN.North - 1
            OUTERhndl = CreateRoundRectRgn(0, 0, _
                        CLng(drbPIC.ScaleWidth + 1), _
                        CLng(drbPIC.ScaleHeight + 1), RADIUS, RADIUS)
            INNERhndl = CreateRoundRectRgn(drbBvlHeight, drbBvlHeight, _
                        CLng(drbPIC.ScaleWidth - drbBvlHeight + 1), _
                        CLng(drbPIC.ScaleHeight - drbBvlHeight + 1), _
                        CLng(RADIUS - 2 * drbBvlHeight), _
                        CLng(RADIUS - 2 * drbBvlHeight))
    End Select
    Call CombineRgn(OUTERhndl, OUTERhndl, INNERhndl, RGN_XOR)
    Bevel_REGION = OUTERhndl ' The return value is the handle of the combined regions
    
    drbPIC.BackColor = RGBColour(drbCOLOUR)
    Ocolour = RGBtoHSL(drbPIC.BackColor)
    Ncolour = Ocolour
    
    Dim Distance As Double
    
    For pntY = 0 To drbPIC.ScaleHeight - 1: For pntX = 0 To drbPIC.ScaleWidth - 1
    If PtInRegion(Bevel_REGION, pntX, pntY) Then
    Select Case True
        'Top Left
        Case pntX < ORIGIN.West And pntY < ORIGIN.North And drbSHAPE <> [Rectangular]
            Distance = (HYPOTENUSE(ORIGIN.West - pntX, ORIGIN.North - pntY) - InnerRADIUS) / (drbBvlHeight - 1)
            RadianANGLE = HalfPi + Atn((ORIGIN.North - pntY) / (ORIGIN.West - pntX))
            Ncolour.Lum = NEWlum(Ocolour.Lum, RadianANGLE, Distance, drbSLOPE, drbRECESS, drbFRAME)
            drbPIC.PSet (pntX, pntY), HSLtoRGB(Ncolour)
        'Bottom Left
        Case pntX < ORIGIN.West And pntY > ORIGIN.South And drbSHAPE <> [Rectangular]
            Distance = (HYPOTENUSE(ORIGIN.West - pntX, pntY - ORIGIN.South) - InnerRADIUS) / (drbBvlHeight - 1)
            RadianANGLE = Pi + Atn((pntY - ORIGIN.South) / (ORIGIN.West - pntX))
            Ncolour.Lum = NEWlum(Ocolour.Lum, RadianANGLE, Distance, drbSLOPE, drbRECESS, drbFRAME)
            drbPIC.PSet (pntX, pntY), HSLtoRGB(Ncolour)
        'Bottom Right
        Case pntX > ORIGIN.East And pntY > ORIGIN.South And drbSHAPE <> [Rectangular]
            Distance = (HYPOTENUSE(pntX - ORIGIN.East, pntY - ORIGIN.South) - InnerRADIUS) / (drbBvlHeight - 1)
            RadianANGLE = Pi + HalfPi + Atn((pntY - ORIGIN.South) / (pntX - ORIGIN.East))
            Ncolour.Lum = NEWlum(Ocolour.Lum, RadianANGLE, Distance, drbSLOPE, drbRECESS, drbFRAME)
            drbPIC.PSet (pntX, pntY), HSLtoRGB(Ncolour)
        'Top Right
        Case pntX > ORIGIN.East And pntY < ORIGIN.North And drbSHAPE <> [Rectangular]
            Distance = (HYPOTENUSE(pntX - ORIGIN.East, ORIGIN.North - pntY) - InnerRADIUS) / (drbBvlHeight - 1)
            RadianANGLE = Atn((ORIGIN.North - pntY) / (pntX - ORIGIN.East))
            Ncolour.Lum = NEWlum(Ocolour.Lum, RadianANGLE, Distance, drbSLOPE, drbRECESS, drbFRAME)
            drbPIC.PSet (pntX, pntY), HSLtoRGB(Ncolour)
        'Top
        Case pntX >= ORIGIN.West And pntX <= ORIGIN.East And pntY < ORIGIN.North
            Distance = (drbBvlHeight - pntY - 1) / (drbBvlHeight - 1)
            Ncolour.Lum = NEWlum(Ocolour.Lum, HalfPi, Distance, drbSLOPE, drbRECESS, drbFRAME)
            If drbSHAPE = [Rectangular] Then
                drbPIC.Line (pntY, pntY)-(drbPIC.ScaleWidth - pntY, pntY), HSLtoRGB(Ncolour)
            Else
                drbPIC.Line (ORIGIN.West, pntY)-(ORIGIN.East + 1, pntY), HSLtoRGB(Ncolour)
            End If
            pntX = ORIGIN.East
        'Left
        Case pntX < ORIGIN.West And pntY >= ORIGIN.North And pntY <= ORIGIN.South
            Distance = (drbBvlHeight - pntX - 1) / (drbBvlHeight - 1)
            Ncolour.Lum = NEWlum(Ocolour.Lum, Pi, Distance, drbSLOPE, drbRECESS, drbFRAME)
            If drbSHAPE = [Rectangular] Then
                drbPIC.Line (pntX, pntX)-(pntX, drbPIC.ScaleHeight - pntX), HSLtoRGB(Ncolour)
            Else
                drbPIC.Line (pntX, ORIGIN.North)-(pntX, ORIGIN.South + 1), HSLtoRGB(Ncolour)
            End If
            pntY = ORIGIN.South
        'Bottom
        Case pntX >= ORIGIN.West And pntX <= ORIGIN.East And pntY > ORIGIN.South
            Distance = (pntY - ORIGIN.South - InnerRADIUS) / (drbBvlHeight - 1)
            Ncolour.Lum = NEWlum(Ocolour.Lum, Pi + HalfPi, Distance, drbSLOPE, drbRECESS, drbFRAME)
            If drbSHAPE = [Rectangular] Then
                drbPIC.Line (drbPIC.ScaleHeight - pntY, pntY)-(drbPIC.ScaleWidth - (drbPIC.ScaleHeight - pntY) + 1, pntY), HSLtoRGB(Ncolour)
            Else
                drbPIC.Line (ORIGIN.West, pntY)-(ORIGIN.East + 1, pntY), HSLtoRGB(Ncolour)
            End If
            pntX = ORIGIN.East
        'Right
        Case pntX > ORIGIN.East And pntY >= ORIGIN.North And pntY <= ORIGIN.South
            Distance = (pntX - ORIGIN.East - InnerRADIUS) / (drbBvlHeight - 1)
            Ncolour.Lum = NEWlum(Ocolour.Lum, TwoPi, Distance, drbSLOPE, drbRECESS, drbFRAME)
            If drbSHAPE = [Rectangular] Then
                drbPIC.Line (pntX, drbPIC.ScaleWidth - pntX)-(pntX, drbPIC.ScaleHeight - (drbPIC.ScaleWidth - pntX)), HSLtoRGB(Ncolour)
            Else
                drbPIC.Line (pntX, ORIGIN.North)-(pntX, ORIGIN.South + 1), HSLtoRGB(Ncolour)
            End If
            pntY = ORIGIN.South
    End Select
    End If
    Next pntX: Next pntY

End Function

'==================================================='
'                                                   '
' All credit must go to PYTHAGORAS for this routine '
'                                                   '
'==================================================='
Private Function HYPOTENUSE(ByVal Side1 As Integer, ByVal Side2 As Integer) As Double
    HYPOTENUSE = (Side1 ^ 2 + Side2 ^ 2) ^ 0.5
End Function

'============================================================='
'                                                             '
' Return a Luminosity value dependend upon passed parameters. '
'                                                             '
'============================================================='
Function NEWlum(ByRef nlSTART_LUM As Integer, _
                ByRef nlRADIAN As Double, _
                ByRef nlDISTANCE As Double, _
                ByRef nlSLOPE As bvlSLOPE, _
                ByRef nlRECESS As Boolean, _
                ByRef nlFRAME As Boolean) _
         As Integer

    Dim BvlHgt As Double, BvlAdj As Double
    Dim Lum1 As Double, Lum2 As Double
    Dim nlNEWradian As Double
    Dim nlANGLE As Double
    
    If nlDISTANCE > 1 Then nlDISTANCE = 1
    If nlDISTANCE < 0 Then nlDISTANCE = 0
    nlNEWradian = nlRADIAN
    If nlFRAME Then
        If nlDISTANCE <= 0.5 Then
            nlDISTANCE = nlDISTANCE * 2
            nlNEWradian = IIf(nlRADIAN < Pi, nlRADIAN + Pi, nlRADIAN - Pi)
            'nlRECESS = Not nlRECESS
        Else
            nlDISTANCE = 2 * (1 - nlDISTANCE)
            Select Case True
                Case [Convex]: nlSLOPE = [Concave]
                Case [Concave]: nlSLOPE = [Convex]
            End Select
        End If
    End If
    Select Case True
        Case (nlSLOPE = [Convex] And Not nlRECESS) Or _
             (nlSLOPE = [Concave] And nlRECESS)
            BvlHgt = Sin(nlDISTANCE * Pi / 2)
        Case (nlSLOPE = [Concave] And Not nlRECESS) Or _
             (nlSLOPE = [Convex] And nlRECESS)
            BvlHgt = Cos(nlDISTANCE * Pi / 2)
        Case Else
            BvlHgt = nlDISTANCE
    End Select
    
    'To give a Bevel a recessed look we offset the angle by Pi
    If nlRECESS Then
        nlANGLE = IIf(nlNEWradian < Pi, nlNEWradian + Pi, nlNEWradian - Pi)
    Else
        nlANGLE = nlNEWradian
    End If
    
    Select Case nlANGLE
        Case 0, TwoPi ' 0 or 360 degrees
            NEWlum = nlSTART_LUM - BvlHgt * (nlSTART_LUM - 40)
        Case Is < HalfPi ' 90 Degrees
            Lum1 = nlSTART_LUM + BvlHgt * (220 - nlSTART_LUM)
            Lum2 = nlSTART_LUM - BvlHgt * (nlSTART_LUM - 40)
            NEWlum = Lum2 + nlANGLE / HalfPi * (Lum1 - Lum2)
        Case HalfPi ' 90 Degrees
            NEWlum = nlSTART_LUM + BvlHgt * (220 - nlSTART_LUM)
        Case Is < HalfPi + Pi / 6 ' 120 Degrees
            BvlAdj = 20 * (nlANGLE - HalfPi) / (Pi / 6)
            NEWlum = nlSTART_LUM + BvlHgt * (BvlAdj + 220 - nlSTART_LUM)
        Case Is <= HalfPi + Pi / 3 ' 150 Degrees
            NEWlum = nlSTART_LUM + BvlHgt * (240 - nlSTART_LUM)
        Case Is < Pi ' 180 Degrees
            BvlAdj = 20 - (20 * (nlANGLE - HalfPi - Pi / 3) / (Pi / 6))
            NEWlum = nlSTART_LUM + BvlHgt * (BvlAdj + 220 - nlSTART_LUM)
        Case Pi ' 180 Degrees
            NEWlum = nlSTART_LUM + BvlHgt * (220 - nlSTART_LUM)
        Case Is < Pi + HalfPi ' 270 Degrees
            Lum1 = nlSTART_LUM - BvlHgt * (nlSTART_LUM - 40)
            Lum2 = nlSTART_LUM + BvlHgt * (220 - nlSTART_LUM)
            NEWlum = Lum2 + (nlANGLE - Pi) / HalfPi * (Lum1 - Lum2)
        Case Pi + HalfPi ' 270 Degrees
            NEWlum = nlSTART_LUM - BvlHgt * (nlSTART_LUM - 40)
        Case Is < Pi + HalfPi + Pi / 6 ' 300 Degrees
            BvlAdj = 40 * (nlANGLE - Pi - HalfPi) / (Pi / 6)
            NEWlum = nlSTART_LUM - BvlHgt * (BvlAdj + nlSTART_LUM - 40)
        Case Is <= Pi + HalfPi + Pi / 3 ' 330 Degrees
            NEWlum = nlSTART_LUM - BvlHgt * nlSTART_LUM
        Case Is < TwoPi ' 360 Degrees
            BvlAdj = 40 - (40 * (nlANGLE - Pi - HalfPi - Pi / 3) / (Pi / 6))
            NEWlum = nlSTART_LUM - BvlHgt * (BvlAdj + nlSTART_LUM - 40)
    End Select

End Function
