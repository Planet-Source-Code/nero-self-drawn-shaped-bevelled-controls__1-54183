VERSION 5.00
Begin VB.UserControl bvlWHEEL 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   155
   ToolboxBitmap   =   "bvlWHEEL.ctx":0000
   Begin VB.Timer tmrWHEEL 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   840
      Top             =   120
   End
   Begin VB.PictureBox picFRAME 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   0
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   0
      Top             =   360
      Width           =   360
   End
   Begin VB.PictureBox picWHEEL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   180
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picTOPLFT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picBTMRHT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "bvlWHEEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private CTRLhndl As Long
    Private BEVLhndl As Long
    Private TOPLFThndl As Long
    Private BTMRHThndl As Long
    Private RTNvalue As Long
    Private Focused As Boolean
    Private Grabbed As Boolean
    Private Resizing As Boolean
    Private LoadingProperties As Boolean
    Private Terminating As Boolean
    Private Exponent As Single
    Private WheelHold As Boolean
    Private WheelPos As Integer
    Private LastX As Long
    Private XChange As Long
    Private LastY As Long
    Private YChange As Long
    
    Private Enum eState
        [Raised] = &H0
        [Pressed] = &H1
        [Hover] = &H2
        [Hover Pressed] = &H3
        [Disabled] = &H4
    End Enum
    Private TOPLFTstate As eState
    Private BTMRHTstate As eState
    
    Enum eInputStates
        [Lock Input]
        [From Devices Only]
        [From Data Only]
        [Allow All Input]
    End Enum
    
    Enum eBiColourUsage
        [Never]
        [With Wheel Grab]
        [With Control Focus]
        [Always]
    End Enum
    
    Enum eBTNShape
        [No BUTTONS] = -1
        [Square Buttons] = 0
        [Circular Buttons] = 1
        [Rounded Square Buttons] = 2
    '   [Capsule] = 3
    End Enum
    
    Enum eWHLShape
        [Rectangular] = 0
    '   [Circular] It would be silly to have a Circular display
        [Rounded Rectangle] = 2
        [Capsule] = 3
    End Enum
    
    Enum eINCRdrctn
        [Left To Right]
        [Right To Left]
        [Bottom To Top]
        [Top To Bottom]
    End Enum
    
    Enum eOrientation
        [Horizontal]
        [Vertical]
    End Enum
    
    'Default Property Values:
    Const m_def_AllowedInput = [Allow All Input]
    Const m_def_ArrowColour = &H80000012
    Const m_def_ArrowHover = &H8000000E
    Const m_def_BackColour = &H8000000F
    Const m_def_BevelAutoAdjust = True
    Const m_def_BevelHeight = 0
    Const m_def_BiColourUsage = [With Control Focus]
    Const m_def_ButtonShape = [Circular Buttons]
    Const m_def_FrameShape = [Rounded Rectangle]
    Const m_def_HideNotches = False
    Const m_def_IncrementDirection = [Left To Right]
    Const m_def_Max = 100
    Const m_def_Min = 0
    Const m_def_Orientation = [Horizontal]
    Const m_def_SpinOver = True
    Const m_def_Value = 50
    Const m_def_WheelElapsed = &H8000000F
    Const m_def_WheelRemain = &H8000000E
    
    'Property Variables:
    Dim m_AllowedInput As eInputStates
    Dim m_ArrowColour As OLE_COLOR
    Dim m_ArrowHover As OLE_COLOR
    Dim m_BevelAutoAdjust As Boolean
    Dim m_BevelHeight As Integer
    Dim m_BiColourUsage As eBiColourUsage
    Dim m_ButtonShape As eBTNShape
    Dim m_BackColour As OLE_COLOR
    Dim m_FrameShape As eWHLShape
    Dim m_HideNotches As Boolean
    Dim m_IncrementDirection As eINCRdrctn
    Dim m_Max As Integer
    Dim m_Min As Integer
    Dim m_Orientation As eOrientation
    Dim m_SpinOver As Boolean
    Dim m_Value As Integer
    Dim M_ValueD As Double
    Dim m_WheelElapsed As OLE_COLOR
    Dim m_WheelRemain As OLE_COLOR
    
    'Event Declarations
    Event Change()
    Event Spin(ByVal SpinSign As Integer)

'
'=====================================================================
'=                                                                   =
'=  Initialize Properties for User Control                           =
'=                                                                   =
'=====================================================================
Private Sub UserControl_InitProperties()
    m_AllowedInput = m_def_AllowedInput
    m_ArrowColour = m_def_ArrowColour
    m_ArrowHover = m_def_ArrowHover
    m_BackColour = m_def_BackColour
    m_BevelAutoAdjust = m_def_BevelAutoAdjust
    m_BevelHeight = m_def_BevelHeight
    m_BiColourUsage = m_def_BiColourUsage
    m_ButtonShape = m_def_ButtonShape
    m_FrameShape = m_def_FrameShape
    m_HideNotches = m_def_HideNotches
    m_IncrementDirection = m_def_IncrementDirection
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_Orientation = m_def_Orientation
    m_SpinOver = m_def_SpinOver
    m_Value = m_def_Value
    m_WheelElapsed = m_def_WheelElapsed
    m_WheelRemain = m_def_WheelRemain
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
        m_AllowedInput = .ReadProperty("AllowedInput", m_def_AllowedInput)
        m_ArrowColour = .ReadProperty("ArrowColour", m_def_ArrowColour)
        m_ArrowHover = .ReadProperty("ArrowHover", m_def_ArrowHover)
        m_BackColour = .ReadProperty("BackColour", m_def_BackColour)
        m_BevelAutoAdjust = .ReadProperty("BevelAutoAdjust", m_def_BevelAutoAdjust)
        m_BevelHeight = .ReadProperty("BevelHeight", m_def_BevelHeight)
        m_BiColourUsage = .ReadProperty("BiColourUsage", m_def_BiColourUsage)
        m_ButtonShape = .ReadProperty("ButtonShape", m_def_ButtonShape)
        m_FrameShape = .ReadProperty("FrameShape", m_def_FrameShape)
        m_HideNotches = .ReadProperty("HideNotches", m_def_HideNotches)
        m_IncrementDirection = .ReadProperty("IncrementDirection", m_def_IncrementDirection)
        m_Max = .ReadProperty("Max", m_def_Max)
        m_Min = .ReadProperty("Min", m_def_Min)
        m_Orientation = .ReadProperty("Orientation", m_def_Orientation)
        m_SpinOver = .ReadProperty("SpinOver", m_def_SpinOver)
        m_Value = .ReadProperty("Value", m_def_Value)
        m_WheelElapsed = .ReadProperty("WheelElapsed", m_def_WheelElapsed)
        m_WheelRemain = .ReadProperty("WheelRemain", m_def_WheelRemain)
    End With
    
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
        Call .WriteProperty("AllowedInput", m_AllowedInput, m_def_AllowedInput)
        Call .WriteProperty("ArrowColour", m_ArrowColour, m_def_ArrowColour)
        Call .WriteProperty("ArrowHover", m_ArrowHover, m_def_ArrowHover)
        Call .WriteProperty("BackColour", m_BackColour, m_def_BackColour)
        Call .WriteProperty("BevelAutoAdjust", m_BevelAutoAdjust, m_def_BevelAutoAdjust)
        Call .WriteProperty("BevelHeight", m_BevelHeight, m_def_BevelHeight)
        Call .WriteProperty("BiColourUsage", m_BiColourUsage, m_def_BiColourUsage)
        Call .WriteProperty("ButtonShape", m_ButtonShape, m_def_ButtonShape)
        Call .WriteProperty("FrameShape", m_FrameShape, m_def_FrameShape)
        Call .WriteProperty("HideNotches", m_HideNotches, m_def_HideNotches)
        Call .WriteProperty("IncrementDirection", m_IncrementDirection, m_def_IncrementDirection)
        Call .WriteProperty("Max", m_Max, m_def_Max)
        Call .WriteProperty("Min", m_Min, m_def_Min)
        Call .WriteProperty("Orientation", m_Orientation, m_def_Orientation)
        Call .WriteProperty("SpinOver", m_SpinOver, m_def_SpinOver)
        Call .WriteProperty("Value", m_Value, m_def_Value)
        Call .WriteProperty("WheelElapsed", m_WheelElapsed, m_def_WheelElapsed)
        Call .WriteProperty("WheelRemain", m_WheelRemain, m_def_WheelRemain)
    End With

End Sub

Public Property Get AllowedInput() As eInputStates
Attribute AllowedInput.VB_Description = "Returns/sets the allowed input(s) for this control."
Attribute AllowedInput.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllowedInput = m_AllowedInput
End Property

Public Property Let AllowedInput(ByVal New_AllowedInput As eInputStates)
    m_AllowedInput = New_AllowedInput
    PropertyChanged "AllowedInput"
End Property

Public Property Get ArrowColour() As OLE_COLOR
Attribute ArrowColour.VB_Description = "Returns/sets the colour of the arrow heads on the button portion of the control."
Attribute ArrowColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ArrowColour = m_ArrowColour
End Property

Public Property Let ArrowColour(ByVal New_ArrowColour As OLE_COLOR)
    m_ArrowColour = New_ArrowColour
    PropertyChanged "ArrowColour"
    Call DrawArrows
End Property

Public Property Get ArrowHover() As OLE_COLOR
Attribute ArrowHover.VB_Description = "Returns/sets the colour of the arrow heads on the button portion of the control when the cursor is above it."
Attribute ArrowHover.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ArrowHover = m_ArrowHover
End Property

Public Property Let ArrowHover(ByVal New_ArrowHover As OLE_COLOR)
    m_ArrowHover = New_ArrowHover
    PropertyChanged "ArrowHover"
    Call DrawArrows
End Property

Public Property Get BackColour() As OLE_COLOR
Attribute BackColour.VB_Description = "Returns/sets the basic background colour of the control."
Attribute BackColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColour = m_BackColour
End Property

Public Property Let BackColour(ByVal New_BackColour As OLE_COLOR)
    m_BackColour = GetBevelColour(New_BackColour)
    PropertyChanged "BackColour"
    Call DrawBEVELS
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

Public Property Get BiColourUsage() As eBiColourUsage
Attribute BiColourUsage.VB_Description = "Returns/sets a value which governs under which conditions that dual colours will be displayed on the control."
Attribute BiColourUsage.VB_ProcData.VB_Invoke_Property = ";Behavior"
    BiColourUsage = m_BiColourUsage
End Property

Public Property Let BiColourUsage(ByVal New_BiColourUsage As eBiColourUsage)
    m_BiColourUsage = New_BiColourUsage
    PropertyChanged "BiColourUsage"
    Call DrawControlSTATE
End Property

Public Property Get ButtonShape() As eBTNShape
Attribute ButtonShape.VB_Description = "Returns/sets the shape of the button to use."
Attribute ButtonShape.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ButtonShape = m_ButtonShape
End Property

Public Property Let ButtonShape(ByVal New_ButtonShape As eBTNShape)
    Resizing = True
    If New_ButtonShape > [No BUTTONS] And m_ButtonShape = [No BUTTONS] Then
        If m_Orientation = [Horizontal] Then UserControl.Width = Screen.TwipsPerPixelX * (UserControl.ScaleWidth + 2 * UserControl.ScaleHeight)
        If m_Orientation = [Vertical] Then UserControl.Height = Screen.TwipsPerPixelY * (UserControl.ScaleHeight + 2 * UserControl.ScaleWidth)
    End If
    If New_ButtonShape = [No BUTTONS] And m_ButtonShape > [No BUTTONS] Then
        If m_Orientation = [Horizontal] Then UserControl.Width = Screen.TwipsPerPixelX * (UserControl.ScaleWidth - 2 * UserControl.ScaleHeight)
        If m_Orientation = [Vertical] Then UserControl.Height = Screen.TwipsPerPixelY * (UserControl.ScaleHeight - 2 * UserControl.ScaleWidth)
    End If
    Resizing = False
    m_ButtonShape = New_ButtonShape
    PropertyChanged "ButtonShape"
    Call UserControl_Resize
End Property

Public Property Get FrameShape() As eWHLShape
Attribute FrameShape.VB_Description = "Returns/sets the shape of the frame around the wheel to use."
Attribute FrameShape.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FrameShape = m_FrameShape
End Property

Public Property Let FrameShape(ByVal New_FrameShape As eWHLShape)
    m_FrameShape = New_FrameShape
    PropertyChanged "FrameShape"
    Call UserControl_Resize
End Property

Public Property Get HideNotches() As Boolean
Attribute HideNotches.VB_Description = "Returns/sets a value which determines if the notches in the wheel will be hidden or visible."
Attribute HideNotches.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HideNotches = m_HideNotches
End Property

Public Property Let HideNotches(ByVal New_HideNotches As Boolean)
    m_HideNotches = New_HideNotches
    PropertyChanged "HideNotches"
    Call CreateWHEEL
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get IncrementDirection() As eINCRdrctn
Attribute IncrementDirection.VB_Description = "Returns/sets a value which determines the direction in which the wheel must be moved to increment the value."
Attribute IncrementDirection.VB_ProcData.VB_Invoke_Property = ";Behavior"
    IncrementDirection = m_IncrementDirection
End Property

Public Property Let IncrementDirection(ByVal New_IncrementDirection As eINCRdrctn)
    If (m_Orientation = [Horizontal] And (New_IncrementDirection = [Left To Right] Or New_IncrementDirection = [Right To Left])) _
    Or (m_Orientation = [Vertical] And (New_IncrementDirection = [Bottom To Top] Or New_IncrementDirection = [Top To Bottom])) Then
        m_IncrementDirection = New_IncrementDirection
        PropertyChanged "IncrementDirection"
        Call DrawControlSTATE
    End If
End Property

Public Property Get Max() As Integer
Attribute Max.VB_Description = "Returns/sets the maximum value that the control can attain."
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Data"
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)

    'Exit if Data input isn't enabled
    If [From Data Only] <> (m_AllowedInput And [From Data Only]) Then Exit Property
    
    Select Case New_Max
        Case Is < m_Value
            MsgBox "Max must not be less than Value", vbCritical, "SpinWheel"
        Case Is < m_Min
            MsgBox "Max must not be greater than Min", vbCritical, "SpinWheel"
        Case Else
            m_Max = New_Max
            PropertyChanged "Max"
            Call DrawControlSTATE
    End Select

End Property

Public Property Get Min() As Integer
Attribute Min.VB_Description = "Return/sets the minimum value that the control can attain."
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Data"
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)

    'Exit if Data input isn't enabled
    If [From Data Only] <> (m_AllowedInput And [From Data Only]) Then Exit Property
    
    Select Case New_Min
        Case Is > m_Value
            MsgBox "Min must not be greater than Value", vbCritical, "SpinWheel"
        Case Is > m_Max
            MsgBox "Min must not be greater than Max", vbCritical, "SpinWheel"
        Case Else
            m_Min = New_Min
            PropertyChanged "Min"
            Call DrawControlSTATE
    End Select

End Property

Public Property Get Orientation() As eOrientation
Attribute Orientation.VB_Description = "Returns/sets the 2D orientation of the control."
Attribute Orientation.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As eOrientation)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
    Dim HeldDimension As Double
    Resizing = True
    If New_Orientation = [Horizontal] Then
        If UserControl.ScaleWidth < UserControl.ScaleHeight Then
            HeldDimension = UserControl.Height
            UserControl.Height = UserControl.Width
            UserControl.Width = HeldDimension
        End If
        If m_IncrementDirection = [Top To Bottom] Then m_IncrementDirection = [Right To Left]
        If m_IncrementDirection = [Bottom To Top] Then m_IncrementDirection = [Left To Right]
    Else
        If UserControl.ScaleHeight < UserControl.ScaleWidth Then
            HeldDimension = UserControl.Width
            UserControl.Width = UserControl.Height
            UserControl.Height = HeldDimension
        End If
        If m_IncrementDirection = [Right To Left] Then m_IncrementDirection = [Top To Bottom]
        If m_IncrementDirection = [Left To Right] Then m_IncrementDirection = [Bottom To Top]
    End If
    Resizing = False
    Call UserControl_Resize
End Property

Public Property Get SpinOver() As Boolean
Attribute SpinOver.VB_Description = "Returns/sets a value which determines if the value may cycle to Min once Max has been attained and vica versa."
Attribute SpinOver.VB_ProcData.VB_Invoke_Property = ";Behavior"
    SpinOver = m_SpinOver
End Property

Public Property Let SpinOver(ByVal New_SpinOver As Boolean)
    m_SpinOver = New_SpinOver
    PropertyChanged "SpinOver"
End Property

Public Property Get Value() As Integer
Attribute Value.VB_Description = "Returns/sets an integer which is the current value of the control."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)

    'Exit if Data input isn't enabled
    If [From Data Only] <> (m_AllowedInput And [From Data Only]) Then Exit Property
    
    Select Case New_Value
        Case Is < m_Min
            MsgBox "Value must not be less than Min", vbCritical, Extender.Name
        Case Is > m_Max
            MsgBox "Value must not be greater than Max", vbCritical, Extender.Name
        Case Else
            If m_IncrementDirection = [Bottom To Top] _
            Or m_IncrementDirection = [Right To Left] Then
                Call Wheel_Move(CInt((New_Value - m_Value) * -3))
            Else
                Call Wheel_Move(CInt((New_Value - m_Value) * 3))
            End If
            m_Value = New_Value
            PropertyChanged "Value"
    End Select

End Property

Public Property Get WheelElapsed() As OLE_COLOR
Attribute WheelElapsed.VB_Description = "Returns/sets the colour used to display the elapsed portion of the wheel when BiColourUsuage is active."
Attribute WheelElapsed.VB_ProcData.VB_Invoke_Property = ";Appearance"
    WheelElapsed = m_WheelElapsed
End Property

Public Property Let WheelElapsed(ByVal New_WheelElapsed As OLE_COLOR)
    m_WheelElapsed = GetBevelColour(New_WheelElapsed)
    PropertyChanged "WheelElapsed"
    Call CreateWHEEL
End Property

Public Property Get WheelRemain() As OLE_COLOR
Attribute WheelRemain.VB_Description = "Returns/sets the colour used to display the remaining portion of the wheel when BiColourUsuage is active."
Attribute WheelRemain.VB_ProcData.VB_Invoke_Property = ";Appearance"
    WheelRemain = m_WheelRemain
End Property

Public Property Let WheelRemain(ByVal New_WheelRemain As OLE_COLOR)
    m_WheelRemain = GetBevelColour(New_WheelRemain)
    PropertyChanged "WheelRemain"
    Call CreateWHEEL
End Property

' The picFRAME object is what actually receives
' the GotFocus and LostFocus events associated
' with the UserControl object.
Private Sub picFRAME_GotFocus()
    Focused = True
    Call DrawControlSTATE
End Sub

Private Sub picFRAME_LostFocus()
    Focused = False
    Call DrawControlSTATE
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Exit if Device input isn't enabled
    If [From Devices Only] <> (m_AllowedInput And [From Devices Only]) Then Exit Sub
    
    LastX = x: LastY = y
    Select Case True
        ' TOP or LEFT Button's region
        Case PtInRegion(TOPLFThndl, x, y)
            TOPLFTstate = [Hover Pressed]
        ' BOTTOM or RIGHT Button's region
        Case PtInRegion(BTMRHThndl, x, y)
            BTMRHTstate = [Hover Pressed]
        ' Once the Button regions have been eliminated the
        ' remaining visible control element must be the WHEEL.
        Case Else
            Grabbed = True
            Call DrawControlSTATE
    End Select

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Exit if Device input isn't enabled
    If [From Devices Only] <> (m_AllowedInput And [From Devices Only]) Then Exit Sub
    
    ' If the wheel has been "Grabbed" then we must move the Wheel.
    If Grabbed Then
        If m_Orientation = [Horizontal] Then
            Call Wheel_Move(Int(x - LastX))
        Else
            Call Wheel_Move(Int(y - LastY))
        End If
    End If
    LastX = x: LastY = y
    
    Select Case True
        ' TOP or LEFT Button's region
        Case PtInRegion(TOPLFThndl, x, y)
            TOPLFTstate = TOPLFTstate Or [Hover]
            BTMRHTstate = BTMRHTstate And Not [Hover]
            tmrWHEEL.Enabled = True
            Call DrawControlSTATE
        ' BOTTOM or RIGHT Button's region
        Case PtInRegion(BTMRHThndl, x, y)
            TOPLFTstate = TOPLFTstate And Not [Hover]
            BTMRHTstate = BTMRHTstate Or [Hover]
            tmrWHEEL.Enabled = True
            Call DrawControlSTATE
        ' Once the Button regions have been eliminated the
        ' remaining visible control element must be the WHEEL.
        Case Else
            TOPLFTstate = TOPLFTstate And Not [Hover]
            BTMRHTstate = BTMRHTstate And Not [Hover]
            Call DrawControlSTATE
    End Select
    
    'Call DrawControlSTATE

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Exit if Device input isn't enabled
    If [From Devices Only] <> (m_AllowedInput And [From Devices Only]) Then Exit Sub
    
    tmrWHEEL.Enabled = False
    Grabbed = False
    Select Case True
        ' TOP or LEFT Button's region
        Case PtInRegion(TOPLFThndl, x, y)
            TOPLFTstate = [Hover]
            BTMRHTstate = [Raised]
        ' BOTTOM or RIGHT Button's region
        Case PtInRegion(BTMRHThndl, x, y)
            TOPLFTstate = [Raised]
            BTMRHTstate = [Hover]
        ' Once the Button regions have been eliminated the
        ' remaining visible control element must be the WHEEL.
        Case Else
            TOPLFTstate = [Raised]
            BTMRHTstate = [Raised]
    End Select
    Call DrawControlSTATE

End Sub

Private Sub UserControl_Resize()

    If Resizing Or Terminating Then Exit Sub
    Resizing = True
    
    Dim ucSW As Long
    Dim ucSH As Long
    Dim ucER As Long
    
    Const MINdimension As Integer = 16
    Dim MINratio As Integer
    
    ' Ensure the Control is using PIXEL Scale Mode
    UserControl.ScaleMode = vbPixels
    
    ucSW = UserControl.Width / Screen.TwipsPerPixelX
    ucSH = UserControl.Height / Screen.TwipsPerPixelY
    
    If m_ButtonShape > [No BUTTONS] Then MINratio = 5 Else MINratio = 3
    If ucSW < ucSH Then
        m_Orientation = [Vertical]
        If ucSW < MINdimension Then ucSW = MINdimension
        If ucSH < MINdimension * MINratio Then ucSH = MINdimension * MINratio
        If ucSH < ucSW * MINratio Then ucSH = ucSW * MINratio
    Else
        m_Orientation = [Horizontal]
        If ucSW < MINdimension * MINratio Then ucSW = MINdimension * MINratio
        If ucSH < MINdimension Then ucSH = MINdimension
        If ucSW < ucSH * MINratio Then ucSW = ucSH * MINratio
    End If
    
    UserControl.Width = ucSW * Screen.TwipsPerPixelX
    UserControl.Height = ucSH * Screen.TwipsPerPixelY
    
    ' Size the Frame and Button picture boxes dependent
    ' upon whether buttons have been selected in the control.
    If m_ButtonShape = [No BUTTONS] Then
        picFRAME.Move 0, 0, ucSW, ucSH
    Else
        picFRAME.Move IIf(m_Orientation = [Horizontal], ucSH, 0), _
                      IIf(m_Orientation = [Horizontal], 0, ucSW), _
                      IIf(m_Orientation = [Horizontal], ucSW - 2 * ucSH, ucSW), _
                      IIf(m_Orientation = [Horizontal], ucSH, ucSH - 2 * ucSW)
        picTOPLFT.Move 0, 0, IIf(m_Orientation = [Horizontal], ucSH, ucSW), IIf(m_Orientation = [Horizontal], ucSH, ucSW)
        picBTMRHT.Move IIf(m_Orientation = [Horizontal], ucSW - ucSH, 0), _
                     IIf(m_Orientation = [Horizontal], 0, ucSH - ucSW), _
                     IIf(m_Orientation = [Horizontal], ucSH, ucSW), _
                     IIf(m_Orientation = [Horizontal], ucSH, ucSW)
    End If
    
    ' Create the shaped region of the Frame.
    Select Case m_FrameShape
        Case [Rectangular]
            ucER = IIf(ucSH < ucSW, CLng(ucSH / 2), CLng(ucSW / 2))
            CTRLhndl = CreateRectRgn(CLng(picFRAME.Left), _
                                     CLng(picFRAME.Top), _
                                     CLng(picFRAME.Left + picFRAME.Width), _
                                     CLng(picFRAME.Top + picFRAME.Height))
        Case [Rounded Rectangle]
            ucER = IIf(ucSH < ucSW, CLng(ucSH / 2), CLng(ucSW / 2))
            CTRLhndl = CreateRoundRectRgn(CLng(picFRAME.Left), _
                                          CLng(picFRAME.Top), _
                                          CLng(picFRAME.Left + picFRAME.Width + 1), _
                                          CLng(picFRAME.Top + picFRAME.Height + 1), _
                                          ucER, ucER)
        Case [Capsule]
            ucER = IIf(ucSH < ucSW, CLng(ucSH), CLng(ucSW))
            CTRLhndl = CreateRoundRectRgn(CLng(picFRAME.Left), _
                                          CLng(picFRAME.Top), _
                                          CLng(picFRAME.Left + picFRAME.Width + 1), _
                                          CLng(picFRAME.Top + picFRAME.Height + 1), _
                                          ucER, ucER)
    End Select
    
    If m_ButtonShape > [No BUTTONS] Then
        ' Create the shaped regions of the Buttons.
        Select Case m_ButtonShape
            Case [Square Buttons]
                ucER = IIf(ucSH < ucSW, CLng(ucSH / 2), CLng(ucSW / 2))
                TOPLFThndl = CreateRectRgn(CLng(picTOPLFT.Left), _
                                         CLng(picTOPLFT.Top), _
                                         CLng(picTOPLFT.Left + picTOPLFT.Width), _
                                         CLng(picTOPLFT.Top + picTOPLFT.Height))
                BTMRHThndl = CreateRectRgn(CLng(picBTMRHT.Left), _
                                         CLng(picBTMRHT.Top), _
                                         CLng(picBTMRHT.Left + picBTMRHT.Width), _
                                         CLng(picBTMRHT.Top + picBTMRHT.Height))
            Case [Rounded Square Buttons]
                ucER = IIf(ucSH < ucSW, CLng(ucSH / 2), CLng(ucSW / 2))
                TOPLFThndl = CreateRoundRectRgn(CLng(picTOPLFT.Left), _
                                              CLng(picTOPLFT.Top), _
                                              CLng(picTOPLFT.Left + picTOPLFT.Width + 1), _
                                              CLng(picTOPLFT.Top + picTOPLFT.Height + 1), _
                                              ucER, ucER)
                BTMRHThndl = CreateRoundRectRgn(CLng(picBTMRHT.Left), _
                                              CLng(picBTMRHT.Top), _
                                              CLng(picBTMRHT.Left + picBTMRHT.Width + 1), _
                                              CLng(picBTMRHT.Top + picBTMRHT.Height + 1), _
                                              ucER, ucER)
            Case [Circular Buttons]
                ucER = IIf(ucSH < ucSW, CLng(ucSH), CLng(ucSW))
                TOPLFThndl = CreateRoundRectRgn(CLng(picTOPLFT.Left), _
                                              CLng(picTOPLFT.Top), _
                                              CLng(picTOPLFT.Left + picTOPLFT.Width + 1), _
                                              CLng(picTOPLFT.Top + picTOPLFT.Height + 1), _
                                              ucER, ucER)
                BTMRHThndl = CreateRoundRectRgn(CLng(picBTMRHT.Left), _
                                              CLng(picBTMRHT.Top), _
                                              CLng(picBTMRHT.Left + picBTMRHT.Width + 1), _
                                              CLng(picBTMRHT.Top + picBTMRHT.Height + 1), _
                                              ucER, ucER)
        End Select
        ' The Buttons' regions are combined with the Frame region
        Call CombineRgn(CTRLhndl, CTRLhndl, TOPLFThndl, RGN_OR)
        Call CombineRgn(CTRLhndl, CTRLhndl, BTMRHThndl, RGN_OR)
    End If
    
    ' The resultant FRAME region is used to shape the control.
    RTNvalue = SetWindowRgn(UserControl.hWnd, CTRLhndl, True)
    
    Call DrawBEVELS
    
    Resizing = False

End Sub

Private Sub UserControl_Terminate()
    Terminating = True
    If CTRLhndl > 0 Then RTNvalue = DeleteObject(CTRLhndl)
    If BEVLhndl > 0 Then RTNvalue = DeleteObject(BEVLhndl)
    If TOPLFThndl > 0 Then RTNvalue = DeleteObject(TOPLFThndl)
    If BTMRHThndl > 0 Then RTNvalue = DeleteObject(BTMRHThndl)
End Sub

Private Sub DrawBEVELS()

    If LoadingProperties Then Exit Sub
    
    If BevelAutoAdjust Then m_BevelHeight = 0
    ' Draw the bevelled frame.
    BEVLhndl = Bevel_REGION(picFRAME, m_BevelHeight, m_BackColour, m_FrameShape, Convex, False, True)
    ' Cut a Hole in the Bevel Border we've just drawn.
    RTNvalue = SetWindowRgn(picFRAME.hWnd, BEVLhndl, True)
    
    If m_ButtonShape > [No BUTTONS] Then Call CreateButtons
    
    Call CreateWHEEL

End Sub

Private Sub CreateButtons()

    Call Bevel_REGION(picTOPLFT, m_BevelHeight, m_BackColour, m_ButtonShape, Convex)
    Call Bevel_REGION(picBTMRHT, m_BevelHeight, m_BackColour, m_ButtonShape, Convex, True)
    
    picTOPLFT.Height = picTOPLFT.Height * 5
    picBTMRHT.Height = picBTMRHT.Height * 5
    RTNvalue = BitBlt(picTOPLFT.hdc, 0, picTOPLFT.Width, picTOPLFT.Width, picTOPLFT.Width, _
                     picBTMRHT.hdc, 0, 0, vbSrcCopy)
    RTNvalue = BitBlt(picTOPLFT.hdc, 0, 2 * picTOPLFT.Width, picTOPLFT.Width, 2 * picTOPLFT.Width, _
                     picTOPLFT.hdc, 0, 0, vbSrcCopy)
    RTNvalue = BitBlt(picTOPLFT.hdc, 0, 4 * picTOPLFT.Width, picTOPLFT.Width, picTOPLFT.Width, _
                     picTOPLFT.hdc, 0, 0, vbSrcCopy)
    RTNvalue = BitBlt(picBTMRHT.hdc, 0, 0, picBTMRHT.Width, picBTMRHT.Height, _
                     picTOPLFT.hdc, 0, 0, vbSrcCopy)
    
    Call DrawArrows

End Sub

Private Sub DrawArrows()

    Dim ArrowHeight As Integer
    Dim Cross As Integer
    Dim Arm1 As Integer, Arm2 As Integer
    Dim ArrowRGB As OLE_COLOR, HoverRGB As OLE_COLOR
    Dim GreyRGB As OLE_COLOR, GreyHSL As HSLCol
    
    ArrowRGB = RGBColour(m_ArrowColour)
    HoverRGB = RGBColour(m_ArrowHover)
    GreyHSL = RGBtoHSL(ArrowRGB)
    If GreyHSL.Sat = 0 Then
        GreyHSL.Lum = Int(2 / 3 * GreyHSL.Lum)
    Else
        GreyHSL.Sat = 0
    End If
    GreyRGB = HSLtoRGB(GreyHSL)
    
    ArrowHeight = picTOPLFT.Width - 2 * Int(0.3 * picTOPLFT.Width)
    For Cross = 0 To ArrowHeight - 1
        If Cross < Int(ArrowHeight / 2 + 1) Then
            Arm1 = Int(ArrowHeight / 2) - Cross
            Arm2 = ArrowHeight - Arm1
        Else
            Arm1 = Int(ArrowHeight / 3 + 0.5)
            Arm2 = ArrowHeight - Arm1
        End If
        Call ArrowHead(ArrowHeight, Cross, 0, Arm1, Arm2, ArrowRGB)
        Call ArrowHead(ArrowHeight, Cross, picTOPLFT.Width, Arm1, Arm2, ArrowRGB, True)
        Call ArrowHead(ArrowHeight, Cross, 2 * picTOPLFT.Width, Arm1, Arm2, HoverRGB)
        Call ArrowHead(ArrowHeight, Cross, 3 * picTOPLFT.Width, Arm1, Arm2, HoverRGB, True)
        Call ArrowHead(ArrowHeight, Cross, 4 * picTOPLFT.Width, Arm1, Arm2, GreyRGB)
    Next Cross
    
    Call DrawControlSTATE

End Sub

Private Sub ArrowHead(ByVal ArrowHeight As Integer, _
                      ByVal Crossbar As Integer, _
                      ByVal OFFSET As Integer, _
                      ByVal Arm1 As Integer, _
                      ByVal Arm2 As Integer, _
                      ByVal AHcolour As OLE_COLOR, _
                      Optional ByVal Recess As Boolean = False)

    Dim LESSfrom As POINTAPI, LESSto As POINTAPI
    Dim MOREfrom As POINTAPI, MOREto As POINTAPI
    Dim Inset As Integer
    
    If m_Orientation = [Horizontal] Then
        LESSfrom.x = Crossbar
        LESSto.x = Crossbar
        LESSfrom.y = Arm1 + OFFSET
        LESSto.y = Arm2 + OFFSET
        MOREfrom.x = ArrowHeight - 1 - Crossbar
        MOREto.x = ArrowHeight - 1 - Crossbar
        MOREfrom.y = Arm1 + OFFSET
        MOREto.y = Arm2 + OFFSET
    Else
        LESSfrom.x = Arm1
        LESSto.x = Arm2
        LESSfrom.y = Crossbar + OFFSET
        LESSto.y = Crossbar + OFFSET
        MOREfrom.x = Arm1
        MOREto.x = Arm2
        MOREfrom.y = ArrowHeight - 1 - Crossbar + OFFSET
        MOREto.y = ArrowHeight - 1 - Crossbar + OFFSET
    End If
    Inset = Int(0.3 * picTOPLFT.Width)
    If Recess Then Inset = Inset + 1
    
    picTOPLFT.Line (Inset + LESSfrom.x, Inset + LESSfrom.y)-(Inset + LESSto.x, Inset + LESSto.y), RGBColour(AHcolour)
    picBTMRHT.Line (Inset + MOREfrom.x, Inset + MOREfrom.y)-(Inset + MOREto.x, Inset + MOREto.y), RGBColour(AHcolour)

End Sub

Private Sub CreateWHEEL()

    Dim WorkHSL() As HSLCol
    
    Dim WHLOffset As Integer
    Dim LineOffset As Integer
    Dim LineLength As Integer
    Dim LineFinish As Integer
    
    Dim CurveRadius As Double
    Dim CurveOffset As Double
    
    Dim NotchCount As Integer
    Dim CrrntNotch As Integer
    Dim CrrntAngle As Double
    Dim NotchPosition As Integer
    Dim UnlitPosition As Integer
    Dim LightPosition As Integer
    
    ' Determine and Set the dimensions needed to plot the Wheel.
    If m_Orientation = [Horizontal] Then
        picWHEEL.Move picFRAME.Left + m_BevelHeight, picFRAME.Top + m_BevelHeight, _
                      picFRAME.Width - 2 * m_BevelHeight, 27
        LineFinish = picWHEEL.Width - 1
    Else
        picWHEEL.Move picFRAME.Left + m_BevelHeight, picFRAME.Top + m_BevelHeight, _
                      27, picFRAME.Height - 2 * m_BevelHeight
        LineFinish = picWHEEL.Height - 1
    End If
    CurveRadius = 2 / 3 * Pi
    CurveOffset = (Pi - CurveRadius) / 2
    NotchCount = IIf(LineFinish < 49, 7, 1 + Int(LineFinish / 7))
    
    ReDim WorkHSL(LineFinish)
    
    'Loop through each of the three Wheel Colours
    For WHLOffset = 0 To 18 Step 9
        'Loop through the 9 positions the wheel can be in.
        For LineOffset = 0 To 8
            'Create each colour point on the Wheel Line.
            For LineLength = 0 To LineFinish
                'Get the right starting colour dependent upon
                'which of the Wheels we are actually drawing
                Select Case WHLOffset
                    Case 0:  WorkHSL(LineLength) = RGBtoHSL(RGBColour(m_BackColour))
                    Case 9:  WorkHSL(LineLength) = RGBtoHSL(RGBColour(m_WheelElapsed))
                    Case 18: WorkHSL(LineLength) = RGBtoHSL(RGBColour(m_WheelRemain))
                End Select
                'Determine the actual colour to use.
                If LineLength <= Int(LineFinish / 2) Then
                    WorkHSL(LineLength).Lum = NEWlum(WorkHSL(LineLength).Lum, _
                                              CDbl(HalfPi), _
                                              CDbl((Int(LineFinish / 2) - LineLength) / (Int(LineFinish / 2))), _
                                              Convex, False, False)
                Else
                    WorkHSL(LineLength).Lum = NEWlum(WorkHSL(LineLength).Lum, _
                                              CDbl(Pi + HalfPi), _
                                              CDbl((LineLength - Int(LineFinish / 2)) / (Int(LineFinish / 2))), _
                                              Convex, False, False)
                End If
            Next LineLength
            'Add the Notches to the Wheel if they are not hidden.
            If Not HideNotches Then
                For CrrntNotch = 1 To NotchCount
                    CrrntAngle = CurveOffset _
                               + LineOffset * CurveRadius / NotchCount / 9 _
                               + CurveRadius * (CrrntNotch - 1) / NotchCount
                    NotchPosition = Int(LineFinish * ((Cos(CrrntAngle) + Cos(CurveOffset)) / (2 * Cos(CurveOffset))))
                    Select Case NotchPosition
                        Case 0
                            UnlitPosition = LineFinish
                            LightPosition = NotchPosition + 1
                        Case LineFinish
                            UnlitPosition = NotchPosition - 1
                            LightPosition = 0
                        Case Else
                            UnlitPosition = NotchPosition - 1
                            LightPosition = NotchPosition + 1
                    End Select
                    WorkHSL(UnlitPosition).Lum = WorkHSL(UnlitPosition).Lum - 0.4 * WorkHSL(UnlitPosition).Lum
                    WorkHSL(NotchPosition).Lum = WorkHSL(NotchPosition).Lum - 0.9 * WorkHSL(NotchPosition).Lum
                    WorkHSL(LightPosition).Lum = WorkHSL(LightPosition).Lum + 0.8 * (240 - WorkHSL(LightPosition).Lum)
                Next CrrntNotch
            End If
            'Plot the Wheel
            For LineLength = 0 To LineFinish
                If m_Orientation = [Horizontal] Then
                    picWHEEL.PSet (LineLength, WHLOffset + 8 - LineOffset), HSLtoRGB(WorkHSL(LineLength))
                Else
                    picWHEEL.PSet (WHLOffset + 8 - LineOffset, LineLength), HSLtoRGB(WorkHSL(LineLength))
                End If
            Next LineLength
        Next LineOffset
    Next WHLOffset
    
    Call DrawControlSTATE

End Sub

Private Sub Wheel_Move(Increment As Integer)

    Dim intCntr As Integer
    Dim Inverter As Integer  ' Used to change Max and Min positions
    Dim Magnitude As Integer ' The actual size of the increment
    Dim ValChange As Integer ' The sign of the magnitude of the change
    
    If m_IncrementDirection = [Bottom To Top] _
    Or m_IncrementDirection = [Right To Left] Then
        Inverter = -1
    Else
        Inverter = 1
    End If
    
    ValChange = Sgn(Increment * Inverter)
    Magnitude = Abs(Increment * Inverter)
    If Not m_SpinOver Then
        Select Case ValChange
            Case Is < 0: If m_Value - Magnitude < m_Min Then Magnitude = m_Value - m_Min
            Case Is > 0: If m_Value + Magnitude > m_Max Then Magnitude = m_Max - m_Value
        End Select
    End If
    Magnitude = Magnitude * ValChange
    
    If Magnitude = 0 Then Exit Sub
    
    'Loop once each time for the magnitude of the increment
    For intCntr = ValChange To Magnitude Step ValChange
        WheelPos = WheelPos + Sgn(Increment)
        'Upper and Lower Bounds of wheel position are 8 & 0
        If WheelPos > 8 Then WheelPos = 0
        If WheelPos < 0 Then WheelPos = 8
        Select Case WheelPos
            'VALUE is only changed on every third movement
            Case 1, 4, 7
                'Increment the value by an amount of 1
                m_Value = m_Value + ValChange
                If ValChange < 0 Then
                    If m_Value < m_Min And m_SpinOver Then m_Value = m_Max
                Else
                    If m_Value > m_Max And m_SpinOver Then m_Value = m_Min
                End If
                RaiseEvent Change
        End Select
        RaiseEvent Spin(Sgn(Increment))
        Call DrawControlSTATE
    Next intCntr

End Sub

Private Sub DrawControlSTATE()
    
    ' Draw the buttons
    If m_ButtonShape > [No BUTTONS] Then
        RTNvalue = BitBlt(UserControl.hdc, 0, 0, picTOPLFT.Width, picTOPLFT.Width, _
                          picTOPLFT.hdc, 0, TOPLFTstate * picTOPLFT.Width, vbSrcCopy)
        RTNvalue = BitBlt(UserControl.hdc, _
                          IIf(m_Orientation = [Horizontal], UserControl.ScaleWidth - picBTMRHT.Width, 0), _
                          IIf(m_Orientation = [Horizontal], 0, UserControl.ScaleHeight - picBTMRHT.Width), _
                          picBTMRHT.Width, picBTMRHT.Width, _
                          picBTMRHT.hdc, 0, BTMRHTstate * picBTMRHT.Width, vbSrcCopy)
    End If
    
    ' Draw the wheel
    If WheelPos < 0 Or WheelPos > 8 Then Exit Sub
    
    Dim AmountComplete As Double
    Dim CompletePixels As Integer
    Dim RemainPixels As Integer
    
    Dim WhlLEFT As Integer, WhlTOP As Integer
    Dim WhlWIDTH As Integer, WhlHEIGHT As Integer
    Dim Clr1OFST As Integer, Clr2OFST As Integer
    
    WhlLEFT = picFRAME.Left + m_BevelHeight
    WhlTOP = picFRAME.Top + m_BevelHeight
    WhlWIDTH = picFRAME.Width - 2 * m_BevelHeight
    WhlHEIGHT = picFRAME.Height - 2 * m_BevelHeight
    
    AmountComplete = (m_Value - m_Min) / (m_Max - m_Min)
    If (m_BiColourUsage = [With Wheel Grab] And Grabbed) _
    Or (m_BiColourUsage = [With Control Focus] And Focused) _
    Or (m_BiColourUsage = [Always]) Then
        If m_IncrementDirection = [Left To Right] Or m_IncrementDirection = [Top To Bottom] Then
            Clr1OFST = 9: Clr2OFST = 18
        Else
            AmountComplete = 1 - AmountComplete
            Clr1OFST = 18: Clr2OFST = 9
        End If
    Else
        Clr1OFST = 0: Clr2OFST = 0
    End If
    
    CompletePixels = Int(AmountComplete * IIf(m_Orientation = [Horizontal], WhlWIDTH, WhlHEIGHT))
    RemainPixels = IIf(m_Orientation = [Horizontal], WhlWIDTH, WhlHEIGHT) - CompletePixels
    If CompletePixels > 0 Then
        Call StretchBlt(UserControl.hdc, _
                        CLng(WhlLEFT), CLng(WhlTOP), _
                        CLng(IIf(m_Orientation = [Horizontal], CompletePixels, WhlWIDTH)), _
                        CLng(IIf(m_Orientation = [Horizontal], WhlHEIGHT, CompletePixels)), _
                        picWHEEL.hdc, _
                        CLng(IIf(m_Orientation = [Horizontal], 0, Clr1OFST + WheelPos)), _
                        CLng(IIf(m_Orientation = [Horizontal], Clr1OFST + WheelPos, 0)), _
                        CLng(IIf(m_Orientation = [Horizontal], CompletePixels, 1)), _
                        CLng(IIf(m_Orientation = [Horizontal], 1, CompletePixels)), _
                        vbSrcCopy)
    End If
    If RemainPixels > 0 Then
        Call StretchBlt(UserControl.hdc, _
                        CLng(IIf(m_Orientation = [Horizontal], CompletePixels + WhlLEFT, WhlLEFT)), _
                        CLng(IIf(m_Orientation = [Horizontal], WhlTOP, CompletePixels + WhlTOP)), _
                        CLng(IIf(m_Orientation = [Horizontal], RemainPixels, WhlWIDTH)), _
                        CLng(IIf(m_Orientation = [Horizontal], WhlHEIGHT, RemainPixels)), _
                        picWHEEL.hdc, _
                        CLng(IIf(m_Orientation = [Horizontal], CompletePixels, Clr2OFST + WheelPos)), _
                        CLng(IIf(m_Orientation = [Horizontal], Clr2OFST + WheelPos, CompletePixels)), _
                        CLng(IIf(m_Orientation = [Horizontal], RemainPixels, 1)), _
                        CLng(IIf(m_Orientation = [Horizontal], 1, RemainPixels)), _
                        vbSrcCopy)
    End If
    
    UserControl.Refresh

End Sub

Private Sub tmrWHEEL_Timer()
    If MouseIsOver(UserControl.hWnd) Then
        ' TOP or LEFT Button's region
        If PtInRegion(TOPLFThndl, LastX, LastY) Then
            If (TOPLFTstate And [Pressed]) = [Pressed] Then Call Wheel_Move(-3)
        Else
            TOPLFTstate = TOPLFTstate And Not [Hover]
        End If
        ' BOTTOM or RIGHT Button's region
        If PtInRegion(BTMRHThndl, LastX, LastY) Then
            If (BTMRHTstate And [Pressed]) = [Pressed] Then Call Wheel_Move(3)
        Else
            BTMRHTstate = BTMRHTstate And Not [Hover]
        End If
    Else
        TOPLFTstate = TOPLFTstate And Not [Hover]
        BTMRHTstate = BTMRHTstate And Not [Hover]
        Call DrawControlSTATE
        tmrWHEEL.Enabled = False
    End If
End Sub

Private Function GetBevelColour(ByVal NewColour As OLE_COLOR) As OLE_COLOR
    Dim NewHSL As HSLCol
    NewHSL = RGBtoHSL(RGBColour(NewColour))
    'Minimum Saturation for a Bevel Colour is 60
    If NewHSL.Sat < 60 Then NewHSL.Sat = 60
    'Minimum Luminosity is dependent upon Saturation
    If NewHSL.Lum < (180 - (NewHSL.Sat / 2)) Then NewHSL.Lum = 180 - (NewHSL.Sat / 2)
    'Maximum Luminosity is set at 200
    If NewHSL.Lum > 200 Then NewHSL.Lum = 200
    'If colour is unaltered we use the original value as this
    'could be a system value and not an actual RGB colour.
    If HSLtoRGB(NewHSL) = RGBColour(NewColour) Then
        GetBevelColour = NewColour
    Else
        GetBevelColour = HSLtoRGB(NewHSL)
    End If
End Function
