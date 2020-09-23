VERSION 5.00
Begin VB.UserControl bvlNUMERIC 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   EditAtDesignTime=   -1  'True
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   ToolboxBitmap   =   "bvlNUMERIC.ctx":0000
   Begin VB.PictureBox picINTERIOR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FF00&
      Height          =   1200
      Left            =   150
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   6
      Top             =   150
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox picBORDER 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   5
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label lblFUNCTION 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Std"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1875
      MouseIcon       =   "bvlNUMERIC.ctx":0312
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "STANDARD Functions"
      Top             =   1500
      Width           =   300
   End
   Begin VB.Label lblTRIGONOMETRY 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reg"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   2325
      MouseIcon       =   "bvlNUMERIC.ctx":061C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "REGULAR Triginometry"
      Top             =   1500
      Width           =   300
   End
   Begin VB.Label lblFORMAT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Float"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   825
      MouseIcon       =   "bvlNUMERIC.ctx":0926
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "FLOATING POINT Display"
      Top             =   1500
      Width           =   390
   End
   Begin VB.Label lblANGLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Deg"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1365
      MouseIcon       =   "bvlNUMERIC.ctx":0C30
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Angles in DEGREES"
      Top             =   1500
      Width           =   360
   End
   Begin VB.Label lblBASE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   375
      MouseIcon       =   "bvlNUMERIC.ctx":0F3A
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Display in DECIMAL"
      Top             =   1500
      Width           =   300
   End
End
Attribute VB_Name = "bvlNUMERIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private CTRLhndl As Long
    Private BEVLhndl As Long
    Private RTNvalue As Long
    Private LoadingProperties As Boolean
    Private Terminating As Boolean
    Private Exponent As Single
    
    Enum eNumShape
        [Rectangular] = 0
    '   [Circular] It would be silly to have a Circular display
        [Rounded Rectangle] = 2
        [Capsule] = 3
    End Enum
    
    Enum eAlignment
        [Justify LEFT]
        [Justify CENTRE]
        [Justify RIGHT]
    End Enum
    
    Enum eAngles
        [Degrees]
        [Radians]
        [Gradians]
    End Enum
    
    Enum eBase
        [Binary (Base 2)] = 2
        [Octal (Base 8)] = 8
        [Decimal (Base 10)] = 10
        [Hexadecimal (Base 16)] = 16
    End Enum
    
    Enum eFormat
        Float
        Fixed
        Scientific
        Engineering
    End Enum
    
    Enum eFunction
        Standard = False
        Inverse = True
    End Enum
    
    Enum eStatusPosition
        [Above]
        [No Status]
        [Beneath]
    End Enum
    
    Enum eTrigonometry
        Regular = False
        Hyperbolic = True
    End Enum
    
    Enum ePropertyChanged
        [Number Base]
        [Angle Type]
        [Display Format]
        [Function Shift]
        [Trigonometry Type]
    End Enum

    'Default Property Values:
    Const m_def_ActiveAngle = [Degrees]
    Const m_def_ActiveBase = [Decimal (Base 10)]
    Const m_def_ActiveFormat = [Float]
    Const m_def_ActiveFunction = [Standard]
    Const m_def_ActiveTrigonometry = [Regular]
    Const m_def_Alignment = [Justify CENTRE]
    Const m_def_Appearance = [Rounded Rectangle]
    Const m_def_BevelAutoAdjust = True
    Const m_def_BevelHeight = 0
    Const m_def_FrameColour = &H8000000F
    Const m_def_FixDecimals As Byte = 8
    Const m_def_InteriorColour = &H80000005
    Const m_def_DropShadow = True
    Const m_def_StatusPosition As Byte = [Beneath]
    Const m_def_ToolTipText = "Numeric Display"
    Const m_def_Value As Double = 0
    
    'Property Variables:
    Dim m_ActiveAngle As eAngles
    Dim m_ActiveBase As eBase
    Dim m_ActiveFormat As eFormat
    Dim m_ActiveFunction As eFunction
    Dim m_ActiveTrigonometry As eTrigonometry
    Dim m_Alignment As eAlignment
    Dim m_Appearance As eNumShape
    Dim m_BevelAutoAdjust As Boolean
    Dim m_BevelHeight As Integer
    Dim m_FrameColour As OLE_COLOR
    Dim m_FixDecimals As Byte
    Dim m_DropShadow As Boolean
    Dim m_StatusPosition As Byte
    Dim m_ToolTipText As String
    Dim m_Value As Double
    Dim m_NumericText As String
    Dim m_InteriorColour As OLE_COLOR
    
    'Event Declarations:
    Event Change(ChangedProperty As ePropertyChanged)
    Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
    Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
    Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
    Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
    Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
    Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
    Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
    Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp


'
'=====================================================================
'=                                                                   =
'=  Initialize Properties for User Control                           =
'=                                                                   =
'=====================================================================
Private Sub UserControl_InitProperties()
    m_ActiveAngle = m_def_ActiveAngle
    m_ActiveBase = m_def_ActiveBase
    m_ActiveFormat = m_def_ActiveFormat
    m_ActiveFunction = m_def_ActiveFunction
    m_ActiveTrigonometry = m_def_ActiveTrigonometry
    m_Alignment = m_def_Alignment
    m_Appearance = m_def_Appearance
    m_BevelAutoAdjust = m_def_BevelAutoAdjust
    m_BevelHeight = m_def_BevelHeight
    m_NumericText = Extender.Name
    m_FixDecimals = m_def_FixDecimals
    m_FrameColour = m_def_FrameColour
    m_InteriorColour = m_def_InteriorColour
    m_DropShadow = m_def_DropShadow
    m_StatusPosition = m_def_StatusPosition
    m_Value = m_def_Value
    Set Font = Ambient.Font
    m_ToolTipText = m_def_ToolTipText
End Sub
'
'=====================================================================
'=                                                                   =
'=  Load property values from storage                                =
'=                                                                   =
'=====================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    LoadingProperties = True
    
    With PropBag
        m_ActiveAngle = .ReadProperty("ActiveAngle", m_def_ActiveAngle)
        m_ActiveBase = .ReadProperty("ActiveBase", m_def_ActiveBase)
        m_ActiveFormat = .ReadProperty("ActiveFormat", m_def_ActiveFormat)
        m_ActiveFunction = .ReadProperty("ActiveFunction", m_def_ActiveFunction)
        m_ActiveTrigonometry = .ReadProperty("ActiveTrigonometry", m_def_ActiveTrigonometry)
        m_Alignment = .ReadProperty("Alignment", m_def_Alignment)
        m_Appearance = .ReadProperty("Appearance", m_def_Appearance)
        m_BevelAutoAdjust = .ReadProperty("BevelAutoAdjust", m_def_BevelAutoAdjust)
        m_BevelHeight = .ReadProperty("BevelHeight", m_def_BevelHeight)
        m_NumericText = .ReadProperty("NumericText", Extender.Name)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        m_FixDecimals = .ReadProperty("FixDecimals", m_def_FixDecimals)
        m_FrameColour = .ReadProperty("FrameColour", m_def_FrameColour)
        m_InteriorColour = .ReadProperty("InteriorColour", m_def_InteriorColour)
        m_DropShadow = .ReadProperty("DropShadow", m_def_DropShadow)
        m_StatusPosition = .ReadProperty("StatusPosition", m_def_StatusPosition)
        m_Value = .ReadProperty("Value", m_def_Value)
    End With
    
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblBASE.ForeColor = PropBag.ReadProperty("StatusColour", &H80000012)
    lblFORMAT.ForeColor = lblBASE.ForeColor
    lblANGLE.ForeColor = lblBASE.ForeColor
    lblFUNCTION.ForeColor = lblBASE.ForeColor
    lblTRIGONOMETRY.ForeColor = lblBASE.ForeColor
    UserControl.ForeColor = PropBag.ReadProperty("NumericColour", &H80000012)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    
    LoadingProperties = False

End Sub
'
'=====================================================================
'=                                                                   =
'=  Write property values to storage                                 =
'=                                                                   =
'=====================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("ActiveAngle", m_ActiveAngle, m_def_ActiveAngle)
        Call .WriteProperty("ActiveBase", m_ActiveBase, m_def_ActiveBase)
        Call .WriteProperty("ActiveFormat", m_ActiveFormat, m_def_ActiveFormat)
        Call .WriteProperty("ActiveFunction", m_ActiveFunction, m_def_ActiveFunction)
        Call .WriteProperty("ActiveTrigonometry", m_ActiveTrigonometry, m_def_ActiveTrigonometry)
        Call .WriteProperty("Alignment", m_Alignment, m_def_Alignment)
        Call .WriteProperty("Appearance", m_Appearance, m_def_Appearance)
        Call .WriteProperty("BevelAutoAdjust", m_BevelAutoAdjust, m_def_BevelAutoAdjust)
        Call .WriteProperty("BevelHeight", m_BevelHeight, m_def_BevelHeight)
        Call .WriteProperty("NumericText", m_NumericText, Extender.Name)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("FixDecimals", m_FixDecimals, m_def_FixDecimals)
        Call .WriteProperty("FrameColour", m_FrameColour, m_def_FrameColour)
        Call .WriteProperty("InteriorColour", m_InteriorColour, m_def_InteriorColour)
        Call .WriteProperty("DropShadow", m_DropShadow, m_def_DropShadow)
        Call .WriteProperty("StatusPosition", m_StatusPosition, m_def_StatusPosition)
        Call .WriteProperty("Value", m_Value, m_def_Value)
    End With
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("StatusColour", lblBASE.ForeColor, &H80000012)
    Call PropBag.WriteProperty("NumericColour", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)

End Sub
'
'=====================================================================
'=                                                                   =
'=  P R O P E R T I E S                                              =
'=                                                                   =
'=====================================================================


Public Property Get ActiveAngle() As eAngles
Attribute ActiveAngle.VB_Description = "Returns/sets the active angle units in use. Has no effect upon the control itself."
Attribute ActiveAngle.VB_ProcData.VB_Invoke_Property = ";Data"
    ActiveAngle = m_ActiveAngle
End Property

Public Property Let ActiveAngle(ByVal New_ActiveAngle As eAngles)
    m_ActiveAngle = New_ActiveAngle
    Select Case ActiveAngle
        Case Degrees:  lblANGLE.Caption = "Deg": lblANGLE.ToolTipText = "Angles in DEGREES"
        Case Radians:  lblANGLE.Caption = "Rad": lblANGLE.ToolTipText = "Angles in RADIANS"
        Case Gradians: lblANGLE.Caption = "Grad": lblANGLE.ToolTipText = "Angles in GRADIANS"
    End Select
    PropertyChanged "ActiveAngle"
    RaiseEvent Change([Angle Type])
End Property

Private Sub lblANGLE_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case ActiveAngle
        Case Degrees:  ActiveAngle = IIf(Button = 1, Radians, Gradians)
        Case Radians:  ActiveAngle = IIf(Button = 1, Gradians, Degrees)
        Case Gradians: ActiveAngle = IIf(Button = 1, Degrees, Radians)
    End Select
End Sub

Public Property Get ActiveBase() As eBase
Attribute ActiveBase.VB_Description = "Returns/sets the active base numbering system that the control uses to display numbers."
Attribute ActiveBase.VB_ProcData.VB_Invoke_Property = ";Data"
    ActiveBase = m_ActiveBase
End Property

Public Property Let ActiveBase(ByVal New_ActiveBase As eBase)
    m_ActiveBase = New_ActiveBase
    Select Case m_ActiveBase
        Case [Binary (Base 2)]:       lblBASE.Caption = "Bin": lblBASE.ToolTipText = "Display in BINARY"
        Case [Octal (Base 8)]:        lblBASE.Caption = "Oct": lblBASE.ToolTipText = "Display in OCTAL"
        Case [Decimal (Base 10)]:     lblBASE.Caption = "Dec": lblBASE.ToolTipText = "Display in DECIMAL"
        Case [Hexadecimal (Base 16)]: lblBASE.Caption = "Hex": lblBASE.ToolTipText = "Display in HEXADECIMAL"
    End Select
    If m_ActiveBase = [Decimal (Base 10)] Then
        lblFORMAT.Enabled = True
    Else
        lblFORMAT.Enabled = False
    End If
    PropertyChanged "ActiveBase"
    Call FormatDisplay
    RaiseEvent Change([Number Base])
End Property

Private Sub lblBASE_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case ActiveBase
        Case [Binary (Base 2)]:       ActiveBase = IIf(Button = 1, [Octal (Base 8)], [Hexadecimal (Base 16)])
        Case [Octal (Base 8)]:        ActiveBase = IIf(Button = 1, [Decimal (Base 10)], [Binary (Base 2)])
        Case [Decimal (Base 10)]:     ActiveBase = IIf(Button = 1, [Hexadecimal (Base 16)], [Octal (Base 8)])
        Case [Hexadecimal (Base 16)]: ActiveBase = IIf(Button = 1, [Binary (Base 2)], [Decimal (Base 10)])
    End Select
End Sub

Public Property Get ActiveFormat() As eFormat
Attribute ActiveFormat.VB_Description = "Returns/sets the active number display format. Only active when the ActiveBase is set to Decimal."
Attribute ActiveFormat.VB_ProcData.VB_Invoke_Property = ";Data"
    ActiveFormat = m_ActiveFormat
End Property

Public Property Let ActiveFormat(ByVal New_ActiveFormat As eFormat)
    m_ActiveFormat = New_ActiveFormat
    Select Case m_ActiveFormat
        Case Float:       lblFORMAT.Caption = "Float"
                          lblFORMAT.ToolTipText = "FLOATING POINT Display"
        Case Fixed:       lblFORMAT.Caption = "Fix" & Trim(Str(m_FixDecimals))
                          lblFORMAT.ToolTipText = "FIXED POINT (" & Trim(Str(m_FixDecimals)) & ") Display"
        Case Scientific:  lblFORMAT.Caption = "Sci" & Trim(Str(m_FixDecimals))
                          lblFORMAT.ToolTipText = "SCIENTIFIC (" & Trim(Str(m_FixDecimals)) & ") Display"
        Case Engineering: lblFORMAT.Caption = "Eng"
                          lblFORMAT.ToolTipText = "ENGINEERING Display"
    End Select
    PropertyChanged "ActiveFormat"
    Call FormatDisplay
    RaiseEvent Change([Display Format])
End Property

Private Sub lblFORMAT_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case ActiveFormat
        Case Float:       ActiveFormat = IIf(Button = 1, Fixed, Engineering)
        Case Fixed:       ActiveFormat = IIf(Button = 1, Scientific, Float)
        Case Scientific:  ActiveFormat = IIf(Button = 1, Engineering, Fixed)
        Case Engineering: ActiveFormat = IIf(Button = 1, Float, Scientific)
    End Select
End Sub

Public Property Get ActiveFunction() As eFunction
Attribute ActiveFunction.VB_Description = "Returns/sets the active function in use. Has no effect upon the control itself."
Attribute ActiveFunction.VB_ProcData.VB_Invoke_Property = ";Data"
    ActiveFunction = m_ActiveFunction
End Property

Public Property Let ActiveFunction(ByVal New_ActiveFunction As eFunction)
    m_ActiveFunction = New_ActiveFunction
    Select Case m_ActiveFunction
        Case Standard: lblFUNCTION.Caption = "Std": lblFUNCTION.ToolTipText = "STANDARD Functions"
        Case Inverse:  lblFUNCTION.Caption = "Inv": lblFUNCTION.ToolTipText = "INVERSE Functions"
    End Select
    PropertyChanged "ActiveFunction"
    RaiseEvent Change([Function Shift])
End Property

Private Sub lblFUNCTION_Click()
    Select Case ActiveFunction
        Case Standard: ActiveFunction = Inverse
        Case Inverse:  ActiveFunction = Standard
    End Select
End Sub

Public Property Get ActiveTrigonometry() As eTrigonometry
Attribute ActiveTrigonometry.VB_Description = "Returns/sets the active trigonometry system in use. Has no effect upon the control itself."
Attribute ActiveTrigonometry.VB_ProcData.VB_Invoke_Property = ";Data"
    ActiveTrigonometry = m_ActiveTrigonometry
End Property

Public Property Let ActiveTrigonometry(ByVal New_ActiveTrigonometry As eTrigonometry)
    m_ActiveTrigonometry = New_ActiveTrigonometry
    Select Case m_ActiveTrigonometry
        Case Regular:    lblTRIGONOMETRY.Caption = "Reg": lblTRIGONOMETRY.ToolTipText = "REGULAR Trigonometry"
        Case Hyperbolic: lblTRIGONOMETRY.Caption = "Hyp": lblTRIGONOMETRY.ToolTipText = "HYPERBOLIC Trigonometry"
    End Select
    PropertyChanged "ActiveTrigonometry"
    RaiseEvent Change([Trigonometry Type])
End Property

Private Sub lblTRIGONOMETRY_Click()
    Select Case ActiveTrigonometry
        Case Regular:    ActiveTrigonometry = Hyperbolic
        Case Hyperbolic: ActiveTrigonometry = Regular
    End Select
End Sub

Public Property Get Alignment() As eAlignment
Attribute Alignment.VB_Description = "Returns/sets the horizontal alignment of the numbers displayed in the control."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As eAlignment)
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
    Call DrawELEMENTS
End Property

Public Property Get Appearance() As eNumShape
Attribute Appearance.VB_Description = "Returns/sets the shape of the control."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As eNumShape)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    Call UserControl_Resize
End Property

Public Property Get BevelAutoAdjust() As Boolean
Attribute BevelAutoAdjust.VB_Description = "Returns/sets a value which determines if the Bevel Height is automatically set."
Attribute BevelAutoAdjust.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BevelAutoAdjust = m_BevelAutoAdjust
End Property

Public Property Let BevelAutoAdjust(ByVal New_BevelAutoAdjust As Boolean)
    m_BevelAutoAdjust = New_BevelAutoAdjust
    PropertyChanged "BevelAutoAdjust"
    Call DrawBEVELS
End Property

Public Property Get BevelHeight() As Integer
Attribute BevelHeight.VB_Description = "Returns/sets the height (in pixels) of the bevel. Only functional when BevelAutoAdjust is set to False."
Attribute BevelHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BevelHeight = m_BevelHeight
End Property

Public Property Let BevelHeight(ByVal New_BevelHeight As Integer)
    m_BevelHeight = New_BevelHeight
    PropertyChanged "BevelHeight"
    Call DrawBEVELS
End Property

Public Property Get DropShadow() As Boolean
Attribute DropShadow.VB_Description = "Returns/sets a value which determines if the interior shadow is displayed."
    DropShadow = m_DropShadow
End Property

Public Property Let DropShadow(ByVal New_DropShadow As Boolean)
    m_DropShadow = New_DropShadow
    PropertyChanged "DropShadow"
    Call DrawINTERIOR
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    If New_Enabled Then
        lblBASE.Enabled = True
        If m_ActiveBase = [Decimal (Base 10)] Then lblFORMAT.Enabled = True
        lblANGLE.Enabled = True
        lblFUNCTION.Enabled = True
        lblTRIGONOMETRY.Enabled = True
    Else
        lblBASE.Enabled = False
        lblFORMAT.Enabled = False
        lblANGLE.Enabled = False
        lblFUNCTION.Enabled = False
        lblTRIGONOMETRY.Enabled = False
    End If
    Call DrawELEMENTS
End Property

Public Property Get FixDecimals() As Byte
Attribute FixDecimals.VB_Description = "Returns/sets a value which determines how many decimal places will be displayed in the control when ActiveFormat is set to Fixed or Scientific display."
Attribute FixDecimals.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FixDecimals = m_FixDecimals
End Property

Public Property Let FixDecimals(ByVal New_FixDecimals As Byte)
    If New_FixDecimals < 1 Or New_FixDecimals > 14 Then
        MsgBox "Fixed Decimals must be between 1 and 14"
    Else
        m_FixDecimals = New_FixDecimals
        PropertyChanged "FixDecimals"
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Call DrawELEMENTS
End Property

Public Property Get FrameColour() As OLE_COLOR
Attribute FrameColour.VB_Description = "Returns/sets the colour of the control's frame."
Attribute FrameColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FrameColour = m_FrameColour
End Property

Public Property Let FrameColour(ByVal New_FrameColour As OLE_COLOR)
    m_FrameColour = New_FrameColour
    PropertyChanged "FrameColour"
    Call DrawBEVELS
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get InteriorColour() As OLE_COLOR
Attribute InteriorColour.VB_Description = "Returns/sets the colour of the control's interior."
Attribute InteriorColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    InteriorColour = m_InteriorColour
End Property

Public Property Let InteriorColour(ByVal New_InteriorColour As OLE_COLOR)
    m_InteriorColour = New_InteriorColour
    PropertyChanged "InteriorColour"
    Call DrawINTERIOR
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get NumericColour() As OLE_COLOR
Attribute NumericColour.VB_Description = "Returns/sets the foreground color used to display numbers in the control."
Attribute NumericColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    NumericColour = UserControl.ForeColor
End Property

Public Property Let NumericColour(ByVal New_NumericColour As OLE_COLOR)
    UserControl.ForeColor() = New_NumericColour
    PropertyChanged "NumericColour"
    Call DrawELEMENTS
End Property

Public Property Get NumericText() As String
Attribute NumericText.VB_Description = "Returns/sets a string containing the value to display in the control. Setting this also sets the contents of Value."
Attribute NumericText.VB_ProcData.VB_Invoke_Property = ";Data"
    NumericText = m_NumericText
End Property

Public Property Let NumericText(ByVal New_NumericText As String)
    m_NumericText = New_NumericText
    PropertyChanged "NumericText"
    Call GetValFromString
    Call DrawELEMENTS
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblBASE,lblBASE,-1,ForeColor
Public Property Get StatusColour() As OLE_COLOR
Attribute StatusColour.VB_Description = "Returns/sets a colour used to display the status items in the control."
Attribute StatusColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    StatusColour = lblBASE.ForeColor
End Property

Public Property Let StatusColour(ByVal New_StatusColour As OLE_COLOR)
    lblBASE.ForeColor = New_StatusColour
    lblFORMAT.ForeColor = New_StatusColour
    lblANGLE.ForeColor = New_StatusColour
    lblFUNCTION.ForeColor = New_StatusColour
    lblTRIGONOMETRY.ForeColor = New_StatusColour
    PropertyChanged "StatusColour"
End Property

Public Property Get StatusPosition() As eStatusPosition
Attribute StatusPosition.VB_Description = "Returns/sets a value which detrmines where in the control the status line will be displayed."
Attribute StatusPosition.VB_ProcData.VB_Invoke_Property = ";Appearance"
    StatusPosition = m_StatusPosition
End Property

Public Property Let StatusPosition(ByVal New_StatusPosition As eStatusPosition)
    m_StatusPosition = New_StatusPosition
    PropertyChanged "StatusPosition"
    Call DrawELEMENTS
End Property

Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Property Get Value() As Double
Attribute Value.VB_Description = "Returns/sets a double value used to determine what string of digits to display in the control. Setting this also sets the contents of NumericText."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Double)
    m_Value = New_Value
    PropertyChanged "Value"
    Call FormatDisplay
End Property

'
'=====================================================================
'=                                                                   =
'=  E V E N T S                                                      =
'=                                                                   =
'=====================================================================
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    UserControl.Refresh
End Sub

'
'===============
' CLICK EVENTS
'===============

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
'
'======================
' DOUBLE CLICK EVENTS
'======================

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub
'
'===============
' FOCUS EVENTS
'===============
Private Sub UserControl_GotFocus()
    'shpUtility.BorderStyle = vbDot
    'shpUtility.DrawMode = vbInvert
    'shpUtility.FillStyle = 1
    'shpUtility.Visible = True
End Sub

Private Sub UserControl_LostFocus()
    'shpUtility.Visible = False
End Sub
'
'==================
' KEY DOWN EVENTS
'==================

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
'
'===================
' KEY PRESS EVENTS
'===================

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
'
'================
' KEY UP EVENTS
'================

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
'
'====================
' MOUSE DOWN EVENTS
'====================

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
'
'====================
' MOUSE MOVE EVENTS
'====================

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
'
'==================
' MOUSE UP EVENTS
'==================

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Function Update()
'
End Function


Private Sub UserControl_Resize()

    Static Resizing As Boolean
    If Resizing Or Terminating Then Exit Sub
    Resizing = True
    
    Dim ucSW As Long
    Dim ucSH As Long
    Dim ucER As Long
    
    ' Ensure the Control is using PIXEL Scale Mode
    UserControl.ScaleMode = vbPixels
    
    ucSW = UserControl.Width / Screen.TwipsPerPixelX
    ucSH = UserControl.Height / Screen.TwipsPerPixelY
    
    ' Set Minimum Dimensions for the Control
    If ucSW < 60 Then ucSW = 60
    If ucSH < 30 Then ucSH = 30
    If (Int(ucSW / 2) * 2) <> ucSW Then ucSW = 2 * Int(1 + ucSW / 2)
    If (Int(ucSH / 2) * 2) <> ucSH Then ucSH = 2 * Int(1 + ucSH / 2)
    
    UserControl.Width = ucSW * Screen.TwipsPerPixelX
    UserControl.Height = ucSH * Screen.TwipsPerPixelY
    
    ' Size the border and interior picture boxes
    ' to be the same size as the user control.
    picBORDER.Move 0, 0, ucSW, ucSH
    picINTERIOR.Move 0, 0, ucSW, ucSH
    
    ' Create the shape of the Control.
    Select Case m_Appearance
        Case [Rectangular]
            ucER = IIf(ucSH < ucSW, CLng(ucSH / 2), CLng(ucSW / 2))
            CTRLhndl = CreateRectRgn(0, 0, CLng(ucSW + 1), CLng(ucSH + 1))
            RTNvalue = SetWindowRgn(UserControl.hWnd, CTRLhndl, True)
        Case [Rounded Rectangle]
            ucER = IIf(ucSH < ucSW, CLng(ucSH / 2), CLng(ucSW / 2))
            CTRLhndl = CreateRoundRectRgn(0, 0, CLng(ucSW + 1), CLng(ucSH + 1), ucER, ucER)
            RTNvalue = SetWindowRgn(UserControl.hWnd, CTRLhndl, True)
        Case [Capsule]
            ucER = IIf(ucSH < ucSW, CLng(ucSH), CLng(ucSW))
            CTRLhndl = CreateRoundRectRgn(0, 0, CLng(ucSW + 1), CLng(ucSH + 1), ucER, ucER)
            RTNvalue = SetWindowRgn(UserControl.hWnd, CTRLhndl, True)
    End Select
    
    Call DrawBEVELS
    
    Resizing = False

End Sub

Private Sub UserControl_Terminate()
    Terminating = True
    If CTRLhndl > 0 Then RTNvalue = DeleteObject(CTRLhndl)
    If BEVLhndl > 0 Then RTNvalue = DeleteObject(BEVLhndl)
End Sub

Private Sub DrawBEVELS()

    If LoadingProperties Then Exit Sub
    
    If BevelAutoAdjust Then m_BevelHeight = 0
    ' Draw the bevelled borders.
    BEVLhndl = Bevel_REGION(picBORDER, m_BevelHeight, m_FrameColour, m_Appearance, Convex, True)
    ' Cut a Hole in the Bevel Border we've just drawn.
    RTNvalue = SetWindowRgn(picBORDER.hWnd, BEVLhndl, True)
    
    Call DrawINTERIOR

End Sub

Private Sub DrawINTERIOR()

    If LoadingProperties Then Exit Sub
    
    ' Determine and set Shadow Colours
    Dim InteriorSHADOW As HSLCol
    If Not m_DropShadow Then
        picINTERIOR.BackColor = RGBColour(m_InteriorColour)
        picINTERIOR.ForeColor = RGBColour(m_InteriorColour)
        picINTERIOR.FillColor = RGBColour(m_InteriorColour)
    Else
        InteriorSHADOW = RGBtoHSL(RGBColour(m_InteriorColour))
        InteriorSHADOW.Lum = Int(InteriorSHADOW.Lum * 0.8)
        picINTERIOR.BackColor = HSLtoRGB(InteriorSHADOW)
        InteriorSHADOW.Lum = Int(InteriorSHADOW.Lum / 0.8 * 0.9)
        picINTERIOR.ForeColor = HSLtoRGB(InteriorSHADOW)
        picINTERIOR.FillColor = RGBColour(m_InteriorColour)
    End If
    
    ' Determine and draw the actual Shadow Shape
    Dim RADIUS As Long
    Select Case m_Appearance
        Case [Rectangular]
            Call Rectangle(picINTERIOR.hdc, _
                 CLng(2 * m_BevelHeight - 1), CLng(2 * m_BevelHeight - 1), _
                 CLng(UserControl.ScaleWidth), CLng(UserControl.ScaleHeight))
        Case [Rounded Rectangle]
            RADIUS = IIf(UserControl.ScaleHeight < UserControl.ScaleWidth, _
                     CLng(UserControl.ScaleHeight / 2 - 2 * m_BevelHeight), _
                     CLng(UserControl.ScaleWidth / 2 - 2 * m_BevelHeight))
            Call RoundRect(picINTERIOR.hdc, _
                 CLng(2 * m_BevelHeight - 1), CLng(2 * m_BevelHeight - 1), _
                 CLng(UserControl.ScaleWidth), CLng(UserControl.ScaleHeight), _
                 RADIUS, RADIUS)
        Case [Capsule]
            RADIUS = IIf(UserControl.ScaleHeight < UserControl.ScaleWidth, _
                     CLng(UserControl.ScaleHeight - 2 * m_BevelHeight), _
                     CLng(UserControl.ScaleWidth - 2 * m_BevelHeight))
            Call RoundRect(picINTERIOR.hdc, _
                 CLng(2 * m_BevelHeight - 1), CLng(2 * m_BevelHeight - 1), _
                 CLng(UserControl.ScaleWidth), CLng(UserControl.ScaleHeight), _
                 RADIUS, RADIUS)
    End Select
    
    Call DrawELEMENTS

End Sub

Public Sub FormatDisplay()

    Select Case m_ActiveBase
        Case [Binary (Base 2)]:       m_NumericText = OCTtoBIN(Oct(m_Value))
        Case [Octal (Base 8)]:        m_NumericText = Oct(Value)
        Case [Decimal (Base 10)]
            Select Case m_ActiveFormat
                Case Float:           m_NumericText = LTrim(Str(m_Value))
                                      If Left(m_NumericText, 2) = "-." Then m_NumericText = "-0." & Mid(m_NumericText, 3)
                Case Fixed:           m_NumericText = FixedFMT
                Case Scientific:      m_NumericText = SciFMT
                Case Engineering:     m_NumericText = EngFMT
            End Select
        Case [Hexadecimal (Base 16)]: m_NumericText = Hex(Value)
    End Select
    
    Call DrawELEMENTS

End Sub

Private Sub DrawELEMENTS()

    If LoadingProperties Then Exit Sub
    
    Dim CptnPOS As RECT
    Dim TestPOS As RECT
    Dim SavedFontSize As Integer
    Dim NewFontSize As Integer
    Dim StatusTop As Integer
    Dim wFormat As Long
    Dim ucTM As TEXTMETRIC
    Dim TestHeight As Integer
    Dim SavedColour As OLE_COLOR
    
    'Copy the image of the interior to the Background
    Call BitBlt(UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
                picINTERIOR.hdc, 0, 0, vbSrcCopy)
    
    ' Preliminary calculations of where the Text and
    ' Status Elements will be positioned in the display.
    CptnPOS.West = m_BevelHeight + 1
    CptnPOS.East = UserControl.ScaleWidth - m_BevelHeight - 1
    Select Case m_StatusPosition
        Case [No Status]
            CptnPOS.North = m_BevelHeight
            CptnPOS.South = UserControl.ScaleHeight - m_BevelHeight - 1
        Case [Above]
            CptnPOS.North = m_BevelHeight + 13
            CptnPOS.South = UserControl.ScaleHeight - m_BevelHeight - 1
            StatusTop = m_BevelHeight - 1
        Case [Beneath]
            CptnPOS.North = m_BevelHeight
            CptnPOS.South = UserControl.ScaleHeight - m_BevelHeight - 14
            StatusTop = CptnPOS.South
    End Select
    
    ' Position the Status Elements
    If m_StatusPosition = [No Status] Then
        lblBASE.Visible = False
        lblFORMAT.Visible = False
        lblANGLE.Visible = False
        lblFUNCTION.Visible = False
        lblTRIGONOMETRY.Visible = False
    Else
        lblBASE.Move CptnPOS.West + Int((UserControl.ScaleWidth - 2 * m_BevelHeight - 150) / 2 + 0.5) + 1, StatusTop, 20, 13
        lblFORMAT.Move lblBASE.Left + 30, StatusTop, 26, 13
        lblANGLE.Move lblBASE.Left + 66, StatusTop, 24, 13
        lblFUNCTION.Move lblBASE.Left + 100, StatusTop, 20, 13
        lblTRIGONOMETRY.Move lblBASE.Left + 130, StatusTop, 20, 13
        lblBASE.Visible = True
        lblFORMAT.Visible = True
        lblANGLE.Visible = True
        lblFUNCTION.Visible = True
        lblTRIGONOMETRY.Visible = True
    End If
    
    Select Case m_Alignment
        Case [Justify LEFT]:   wFormat = DT_LEFT Or DT_SINGLELINE
        Case [Justify CENTRE]: wFormat = DT_CENTER Or DT_SINGLELINE
        Case [Justify RIGHT]:  wFormat = DT_RIGHT Or DT_SINGLELINE
    End Select
    
    If Not UserControl.Enabled Then
        SavedColour = UserControl.ForeColor
        UserControl.ForeColor = RGBColour(&H80000011)
    End If
    
    ' This code scales the Font to a size where
    ' the entire Numeric Text can be displayed.
    SavedFontSize = UserControl.Font.Size
    NewFontSize = SavedFontSize
    Do
        ' We get Text Metrics because we only want to deal with the
        ' "ascent" portion of the Text Height and ignore the "descent"
        ' and "leading" portions of the text height.
        Call GetTextMetrics(UserControl.hdc, ucTM)
        TestHeight = ucTM.tmAscent - ucTM.tmInternalLeading
        ' Do a test print to determine how wide the text is.
        TestPOS = CptnPOS
        Call DrawText(UserControl.hdc, m_NumericText, -1, TestPOS, wFormat Or DT_CALCRECT)
        'If Text doesn't fit then Reduce the Font Size.
        If TestPOS.East > CptnPOS.East Or TestHeight > (CptnPOS.South - CptnPOS.North) Then
            NewFontSize = NewFontSize - 1
            If UserControl.Font.Size > 8 Then UserControl.Font.Size = NewFontSize
        End If
    ' Loop until Text fits in the display or Font becomes too small.
    Loop Until (TestPOS.East <= CptnPOS.East And _
                TestHeight <= (CptnPOS.South - CptnPOS.North)) _
                Or NewFontSize < 8
    ' Resize the rectangle to where the Text will be drawn
    CptnPOS.North = CptnPOS.North - ucTM.tmInternalLeading + 1
    If CptnPOS.South < CptnPOS.North + ucTM.tmHeight - 1 Then CptnPOS.South = CptnPOS.North + ucTM.tmHeight - 1
    Debug.Print CptnPOS.South
    ' Draw the Text
    Call DrawText(UserControl.hdc, m_NumericText, -1, CptnPOS, wFormat Or DT_VCENTER)
    ' Restore the Font Size
    UserControl.Font.Size = SavedFontSize
    
    If Not UserControl.Enabled Then
        UserControl.ForeColor = RGBColour(SavedColour)
    End If
    
    UserControl.Refresh

End Sub

'======================================'
' Extract the numerical value from the '
' input string dependent on BASE.      '
'======================================'
Private Function GetValFromString() As Integer

    On Error GoTo GetValFromString_Error

    GetValFromString = 0

    Select Case m_ActiveBase
        Case [Binary (Base 2)]:       m_Value = Val(BINtoOCT(m_NumericText))
        Case [Octal (Base 8)]:        m_Value = Val("&O" & m_NumericText)
        Case [Decimal (Base 10)]:     m_Value = Val(m_NumericText)
        Case [Hexadecimal (Base 16)]: m_Value = Val("&H" & m_NumericText)
    End Select
    
GetValFromString_Exit:

    Exit Function

GetValFromString_Error:

    Select Case Err.NUMBER
        Case 6
            GetValFromString = 3
            Resume GetValFromString_Exit
        Case 11
            GetValFromString = 2
            Resume GetValFromString_Exit
        Case Else
            Err.Raise Err.NUMBER
    End Select
    
End Function

Private Function FixedFMT() As String
    FixedFMT = LTrim(Str(Sgn(m_Value) * (Int(Abs(m_Value) * (10 ^ FixDecimals) + 0.5) / (10 ^ FixDecimals))))
    If InStr(1, UCase(FixedFMT), "E") > 0 Then
        Exponent = Val(Mid(FixedFMT, InStr(1, UCase(FixedFMT), "E") + 1))
        FixedFMT = Left(FixedFMT, InStr(1, UCase(FixedFMT), "E") - 1)
    Else
        Exponent = 0
    End If
    If FixDecimals > 0 Then
        If InStr(1, FixedFMT, ".") = 0 Then FixedFMT = FixedFMT & "."
        While Len(FixedFMT) - InStr(1, FixedFMT, ".") < FixDecimals: FixedFMT = FixedFMT & "0": Wend
        FixedFMT = Left(FixedFMT, InStr(1, FixedFMT, ".") + FixDecimals)
    End If
    If Left(FixedFMT, 2) = "-." Then FixedFMT = "-0." & Mid(FixedFMT, 3)
    Select Case Exponent
        Case Is > 0: FixedFMT = FixedFMT & "E+" & LTrim(Str(Exponent))
        Case Is < 0: FixedFMT = FixedFMT & "E-" & LTrim(Str(Abs(Exponent)))
    End Select
End Function

Private Function SciFMT() As String
    Call GetEXPONENT
    If Exponent <> 0 Then
        SciFMT = LTrim(Str(Sgn(m_Value) * (Int((Abs(m_Value) / 10 ^ Exponent) * (10 ^ FixDecimals) + 0.5) / (10 ^ FixDecimals))))
    Else
        SciFMT = LTrim(Str(m_Value))
    End If
    If FixDecimals > 0 Then
        If InStr(1, SciFMT, ".") = 0 Then SciFMT = SciFMT & "."
        While Len(SciFMT) - InStr(1, SciFMT, ".") < FixDecimals: SciFMT = SciFMT & "0": Wend
        SciFMT = Left(SciFMT, InStr(1, SciFMT, ".") + FixDecimals)
    End If
    If Left(SciFMT, 2) = "-." Then SciFMT = "-0." & Mid(SciFMT, 3)
    If Exponent >= 0 Then
        SciFMT = SciFMT & "E+" & LTrim(Str(Exponent))
    Else
        SciFMT = SciFMT & "E-" & LTrim(Str(Abs(Exponent)))
    End If
End Function

Private Function EngFMT() As String
    Call GetEXPONENT
    If Exponent <> 0 Then
        EngFMT = LTrim(Str(Sgn(m_Value) * (Int((Abs(m_Value) / 10 ^ Exponent) * (10 ^ FixDecimals) + 0.5) / (10 ^ FixDecimals))))
    Else
        EngFMT = LTrim(Str(m_Value))
    End If
    If Left(EngFMT, 2) = "-." Then EngFMT = "-0." & Mid(EngFMT, 3)
    If Exponent >= 0 Then
        EngFMT = EngFMT & "E+" & LTrim(Str(Exponent))
    Else
        EngFMT = EngFMT & "E-" & LTrim(Str(Abs(Exponent)))
    End If
End Function

Private Sub GetEXPONENT()
    If m_Value = 0 Then
        Exponent = 0
    Else
        Exponent = Log(Abs(m_Value)) / Log(10)
    End If
    Exponent = Int(Exponent)
    If m_ActiveFormat = Engineering Then Exponent = Int(Exponent / 3) * 3
End Sub

Private Function BINtoOCT(ByVal strInput As String) As String
    BINtoOCT = "&O"
    Dim intCntr As Integer
    Dim intVal As Integer
    Do While Len(strInput) <> Int(Len(strInput) / 3) * 3
        strInput = "0" & strInput
    Loop
    For intCntr = 0 To Len(strInput) / 3 - 1
        intVal = 4 * Val(Mid(strInput, intCntr * 3 + 1, 1)) + _
                 2 * Val(Mid(strInput, intCntr * 3 + 2, 1)) + _
                 1 * Val(Mid(strInput, intCntr * 3 + 3, 1))
        BINtoOCT = BINtoOCT & LTrim(Str(intVal))
    Next intCntr
End Function

Private Function OCTtoBIN(strInput As String) As String
    Dim intCntr As Integer
    For intCntr = 1 To Len(strInput)
        Select Case Mid(strInput, intCntr, 1)
            Case "0"
                OCTtoBIN = OCTtoBIN & "000"
            Case "1"
                OCTtoBIN = OCTtoBIN & "001"
            Case "2"
                OCTtoBIN = OCTtoBIN & "010"
            Case "3"
                OCTtoBIN = OCTtoBIN & "011"
            Case "4"
                OCTtoBIN = OCTtoBIN & "100"
            Case "5"
                OCTtoBIN = OCTtoBIN & "101"
            Case "6"
                OCTtoBIN = OCTtoBIN & "110"
            Case "7"
                OCTtoBIN = OCTtoBIN & "111"
        End Select
    Next intCntr
    Do While Left(OCTtoBIN, 1) = "0"
        OCTtoBIN = Mid(OCTtoBIN, 2)
    Loop
    If Len(OCTtoBIN) = 0 Then OCTtoBIN = "0"
End Function
