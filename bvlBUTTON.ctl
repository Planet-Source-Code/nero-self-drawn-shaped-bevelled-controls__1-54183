VERSION 5.00
Begin VB.UserControl bvlBUTTON 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   DefaultCancel   =   -1  'True
   FillColor       =   &H8000000F&
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   ToolboxBitmap   =   "bvlBUTTON.ctx":0000
   Begin VB.Timer tmrRCB 
      Enabled         =   0   'False
      Interval        =   125
      Left            =   3600
      Top             =   120
   End
   Begin VB.PictureBox picRAISED 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   240
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox picSUNKEN 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   480
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox picBASIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   720
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox picHOVER 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   960
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox picWORK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   1200
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   2700
   End
End
Attribute VB_Name = "bvlBUTTON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Dim RgnHndl As Long
    Dim RtrnVal As Long
    Dim CPTN As RECT
    Dim PICT As RECT
    Dim LoadingProperties As Boolean
    Dim Terminating As Boolean
    Dim MouseIsDown As Boolean
    Dim NewX As Integer, NewY As Integer
    
    Private Type tPosition
        Left As Long
        Top As Long
        Width As Integer
        Height As Integer
    End Type
    Dim BASICpstn As tPosition
    Dim HOVERpstn As tPosition
    
    Enum eJustify
        [Top Left]
        [Top Centred]
        [Top Right]
        [Middle Left]
        [Middle Centred]
        [Middle Right]
        [Bottom Left]
        [Bottom Centred]
        [Bottom Right]
    End Enum
    
    Enum eShape
        [Rectangular]
        [Circular]
        [Rounded Rectangle]
        [Capsule]
    End Enum
    
    Enum eMouseKeys
        [Left Button] = &H1
        [Right Button] = &H2
        [Left and Right] = &H3
        [Middle Button] = &H4
        [Left and Middle] = &H5
        [Middle and Right] = &H6
        [All Buttons] = &H7
    End Enum
    
    Private Enum eState
        [Raised] = &H0
        [Pressed] = &H1
        [Hover] = &H2
        [Hover Pressed] = &H3
        [Disabled] = &H4
        [Disabled Pressed] = &H5
        [Button Flash] = &HB
    End Enum
    Private ButtonState As eState
    
    Enum eStyle
        [Command Button]
        [Check Button]
        [Option Button]
    End Enum
    
    'Default Property Values:
    Const m_def_Appearance = [Rounded Rectangle]
    Const m_def_BackColour = &H8000000F
    Const m_def_Behaviour = [Command Button]
    Const m_def_BevelAutoAdjust = True
    Const m_def_BevelHeight = 0
    Const m_def_CaptionJustify = [Middle Centred]
    Const m_def_CaptionColour = &H80000008
    Const m_def_CaptionOffset = "(0,0)"
    Const m_def_Enabled = True
    Const m_def_CaptionHoverColour = &H8000000E
    Const m_def_MouseKeys = [Left Button]
    Const m_def_PictureJustify = [Middle Centred]
    Const m_def_PictureOffset = "(0,0)"
    Const m_def_Transparency = &HC0C0C0
    Const m_def_UseTransparency = True
    Const m_def_Value = False
    'Property Variables:
    Dim m_Appearance As eShape
    Dim m_BackColour As OLE_COLOR
    Dim m_Behaviour As Byte
    Dim m_BevelAutoAdjust As Boolean
    Dim m_BevelHeight As Integer
    Dim m_Caption As String
    Dim m_CaptionJustify As eJustify
    Dim m_CaptionColour As OLE_COLOR
    Dim m_CaptionOffset As String
    Dim m_CaptionOffsetX As Integer
    Dim m_CaptionOffsetY As Integer
    Dim m_Enabled As Boolean
    Dim m_CaptionHoverColour As OLE_COLOR
    Dim m_MouseKeys As eMouseKeys
    Dim m_PictureJustify As eJustify
    Dim m_PictureOffset As String
    Dim m_PictureOffsetX As Integer
    Dim m_PictureOffsetY As Integer
    Dim m_Transparency As OLE_COLOR
    Dim m_UseTransparency As Boolean
    Dim m_Value As Boolean
    'Event Declarations:
    Event Click()
    Event DblClick()
    Event KeyDown(KeyCode As Integer, Shift As Integer)
    Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
    Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
    Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Event MouseExit()
    Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
    Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Appearance = m_def_Appearance
    m_BackColour = m_def_BackColour
    m_Behaviour = m_def_Behaviour
    m_BevelAutoAdjust = m_def_BevelAutoAdjust
    m_BevelHeight = m_def_BevelHeight
    m_Caption = Extender.Name
    m_CaptionJustify = m_def_CaptionJustify
    m_CaptionColour = m_def_CaptionColour
    m_CaptionHoverColour = m_def_CaptionHoverColour
    m_CaptionOffset = m_def_CaptionOffset
    m_Enabled = m_def_Enabled
    m_MouseKeys = m_def_MouseKeys
    m_PictureJustify = m_def_PictureJustify
    m_PictureOffset = m_def_PictureOffset
    m_Transparency = m_def_Transparency
    m_UseTransparency = m_def_UseTransparency
    m_Value = m_def_Value
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    LoadingProperties = True
    
    With PropBag
        UserControl.AccessKeys = .ReadProperty("AccessKeys", "")
        m_Appearance = .ReadProperty("Appearance", m_def_Appearance)
        m_BackColour = .ReadProperty("BackColour", m_def_BackColour)
        m_Behaviour = .ReadProperty("Behaviour", m_def_Behaviour)
        m_BevelAutoAdjust = .ReadProperty("BevelAutoAdjust", m_def_BevelAutoAdjust)
        m_BevelHeight = .ReadProperty("BevelHeight", m_def_BevelHeight)
        m_Caption = .ReadProperty("Caption", Extender.Name)
        m_CaptionJustify = .ReadProperty("CaptionJustify", m_def_CaptionJustify)
        m_CaptionColour = .ReadProperty("CaptionColour", m_def_CaptionColour)
        m_CaptionHoverColour = .ReadProperty("CaptionHoverColour", m_def_CaptionHoverColour)
        m_CaptionOffset = .ReadProperty("CaptionOffset", m_def_CaptionOffset)
        m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
        Set Font = .ReadProperty("Font", Ambient.Font)
        m_MouseKeys = .ReadProperty("MouseKeys", m_def_MouseKeys)
        Set Picture = .ReadProperty("Picture", Nothing)
        m_PictureJustify = .ReadProperty("PictureJustify", m_def_PictureJustify)
        Set PictureHover = .ReadProperty("PictureHover", Nothing)
        m_PictureOffset = .ReadProperty("PictureOffset", m_def_PictureOffset)
        m_Transparency = .ReadProperty("Transparency", m_def_Transparency)
        m_UseTransparency = .ReadProperty("UseTransparency", m_def_UseTransparency)
        m_Value = .ReadProperty("Value", m_def_Value)
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)
    End With
    
    Call Cartesian(m_CaptionOffset, m_CaptionOffsetX, m_CaptionOffsetY)
    Call Cartesian(m_PictureOffset, m_PictureOffsetX, m_PictureOffsetY)
    
    LoadingProperties = False
    
    If (m_Behaviour = [Check Button] Or m_Behaviour = [Option Button]) _
    And m_Value = True Then
        ButtonState = ButtonState Or [Pressed]
        Call DrawBUTTON(False)
    End If

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("AccessKeys", UserControl.AccessKeys, "")
        Call .WriteProperty("Appearance", m_Appearance, m_def_Appearance)
        Call .WriteProperty("BackColour", m_BackColour, m_def_BackColour)
        Call .WriteProperty("Behaviour", m_Behaviour, m_def_Behaviour)
        Call .WriteProperty("BevelAutoAdjust", m_BevelAutoAdjust, m_def_BevelAutoAdjust)
        Call .WriteProperty("BevelHeight", m_BevelHeight, m_def_BevelHeight)
        Call .WriteProperty("Caption", m_Caption, Extender.Name)
        Call .WriteProperty("CaptionJustify", m_CaptionJustify, m_def_CaptionJustify)
        Call .WriteProperty("CaptionColour", m_CaptionColour, m_def_CaptionColour)
        Call .WriteProperty("CaptionHoverColour", m_CaptionHoverColour, m_def_CaptionHoverColour)
        Call .WriteProperty("CaptionOffset", m_CaptionOffset, m_def_CaptionOffset)
        Call .WriteProperty("Enabled", m_Enabled, m_def_Enabled)
        Call .WriteProperty("Font", Font, Ambient.Font)
        Call .WriteProperty("MouseKeys", m_MouseKeys, m_def_MouseKeys)
        Call .WriteProperty("Picture", Picture, Nothing)
        Call .WriteProperty("PictureJustify", m_PictureJustify, m_def_PictureJustify)
        Call .WriteProperty("PictureHover", PictureHover, Nothing)
        Call .WriteProperty("PictureOffset", m_PictureOffset, m_def_PictureOffset)
        Call .WriteProperty("Transparency", m_Transparency, m_def_Transparency)
        Call .WriteProperty("UseTransparency", m_UseTransparency, m_def_UseTransparency)
        Call .WriteProperty("Value", m_Value, m_def_Value)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)
    End With

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AccessKeys
Public Property Get AccessKeys() As String
Attribute AccessKeys.VB_Description = "Returns/sets a Key that when pressed with the Alt key will trigger the button's Click event."
Attribute AccessKeys.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AccessKeys = UserControl.AccessKeys
End Property

Public Property Let AccessKeys(ByVal New_AccessKeys As String)
    UserControl.AccessKeys() = New_AccessKeys
    PropertyChanged "AccessKeys"
End Property

Public Property Get Appearance() As eShape
Attribute Appearance.VB_Description = "Returns/sets the shape of the button to use."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As eShape)
    m_Appearance = New_Appearance
    Call UserControl_Resize
    PropertyChanged "Appearance"
End Property

Public Property Get BackColour() As OLE_COLOR
Attribute BackColour.VB_Description = "Returns/sets the basic colour of the button."
Attribute BackColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColour = m_BackColour
End Property

Public Property Let BackColour(ByVal New_BackColour As OLE_COLOR)
    m_BackColour = New_BackColour
    Call CreateBevelFX
    PropertyChanged "BackColour"
End Property

Public Property Get Behaviour() As eStyle
Attribute Behaviour.VB_Description = "Returns/sets a value which determines whether the button will behave as a Command Button, Option Button or Check Button."
Attribute Behaviour.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Behaviour = m_Behaviour
End Property

Public Property Let Behaviour(ByVal New_Behaviour As eStyle)
    m_Behaviour = New_Behaviour
    PropertyChanged "Behaviour"
End Property

Public Property Get BevelAutoAdjust() As Boolean
Attribute BevelAutoAdjust.VB_Description = "Returns/sets a value which determines if the Bevel Height is automatically set."
Attribute BevelAutoAdjust.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BevelAutoAdjust = m_BevelAutoAdjust
End Property

Public Property Let BevelAutoAdjust(ByVal New_BevelAutoAdjust As Boolean)
    m_BevelAutoAdjust = New_BevelAutoAdjust
    Call CreateBevelFX
    PropertyChanged "BevelAutoAdjust"
End Property

Public Property Get BevelHeight() As Integer
Attribute BevelHeight.VB_Description = "Returns/sets the height (in pixels) of the bevel. Only functional when BevelAutoAdjust is set to False."
Attribute BevelHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BevelHeight = m_BevelHeight
End Property

Public Property Let BevelHeight(ByVal New_BevelHeight As Integer)
    m_BevelHeight = New_BevelHeight
    Call CreateBevelFX
    PropertyChanged "BevelHeight"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the Text to appear on the button."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    Call DrawBUTTON
    PropertyChanged "Caption"
End Property

Public Property Get CaptionColour() As OLE_COLOR
Attribute CaptionColour.VB_Description = "Return/sets the colour of the Text on the button when the cursor is NOT above the button."
Attribute CaptionColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CaptionColour = m_CaptionColour
End Property

Public Property Let CaptionColour(ByVal New_CaptionColour As OLE_COLOR)
    m_CaptionColour = New_CaptionColour
    Call DrawBUTTON
    PropertyChanged "CaptionColour"
End Property

Public Property Get CaptionHoverColour() As OLE_COLOR
Attribute CaptionHoverColour.VB_Description = "Return/sets the colour of the Text on the button when the cursor is Above the button."
Attribute CaptionHoverColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CaptionHoverColour = m_CaptionHoverColour
End Property

Public Property Let CaptionHoverColour(ByVal New_CaptionHoverColour As OLE_COLOR)
    m_CaptionHoverColour = New_CaptionHoverColour
    PropertyChanged "CaptionHoverColour"
End Property

Public Property Get CaptionJustify() As eJustify
Attribute CaptionJustify.VB_Description = "Returns/sets where Text is to primarily be located on the button."
Attribute CaptionJustify.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CaptionJustify = m_CaptionJustify
End Property

Public Property Let CaptionJustify(ByVal New_CaptionJustify As eJustify)
    m_CaptionJustify = New_CaptionJustify
    Call AlignObjects
    Call DrawBUTTON
    PropertyChanged "CaptionJustify"
End Property

Public Property Get CaptionOffset() As String
Attribute CaptionOffset.VB_Description = "Returns/sets the amount of Pixels in cartesian co-ordinates (x,y) to offset the Text on the button by."
Attribute CaptionOffset.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CaptionOffset = m_CaptionOffset
End Property

Public Property Let CaptionOffset(ByVal New_CaptionOffset As String)
    Dim NewX As Integer, NewY As Integer
    If Cartesian(New_CaptionOffset, NewX, NewY) Then
        m_CaptionOffset = New_CaptionOffset
        m_CaptionOffsetX = NewX
        m_CaptionOffsetY = NewY
        Call AlignObjects
        Call DrawBUTTON
        PropertyChanged "CaptionOffset"
    End If
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    If m_Enabled Then
        ButtonState = ButtonState And Not [Disabled]
    Else
        ButtonState = ButtonState Or [Disabled]
    End If
    Call DrawBUTTON
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBUTTON,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Call DrawBUTTON
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a Custom Mouse Icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MouseKeys() As eMouseKeys
Attribute MouseKeys.VB_Description = "Returns/sets which Mouse Buttons the control will respond to."
Attribute MouseKeys.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MouseKeys = m_MouseKeys
End Property

Public Property Let MouseKeys(ByVal New_MouseKeys As eMouseKeys)
    m_MouseKeys = New_MouseKeys
    PropertyChanged "MouseKeys"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPICTURE,picBASIC,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets the graphic to be displayed on the button when the cursor is NOT ABOVE the button."
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Picture = picBASIC.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set picBASIC.Picture = New_Picture
    Call SetUpPictures
    Call DrawBUTTON
    PropertyChanged "Picture"
End Property

Public Property Get PictureHover() As Picture
Attribute PictureHover.VB_Description = "Returns/sets the graphic to be displayed on the button when the cursor is ABOVE the button."
Attribute PictureHover.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set PictureHover = picHOVER.Picture
End Property

Public Property Set PictureHover(ByVal New_PictureHover As Picture)
    Set picHOVER.Picture = New_PictureHover
    Call SetUpPictures
    PropertyChanged "PictureHover"
End Property

Public Property Get PictureJustify() As eJustify
Attribute PictureJustify.VB_Description = "Returns/sets where the active Picture is to primarily be located on the button."
Attribute PictureJustify.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PictureJustify = m_PictureJustify
End Property

Public Property Let PictureJustify(ByVal New_PictureJustify As eJustify)
    m_PictureJustify = New_PictureJustify
    Call AlignObjects
    Call DrawBUTTON
    PropertyChanged "PictureJustify"
End Property

Public Property Get PictureOffset() As String
Attribute PictureOffset.VB_Description = "Returns/sets the amount of Pixels in cartesian co-ordinates (x,y) to offset the active Picture by."
Attribute PictureOffset.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PictureOffset = m_PictureOffset
End Property

Public Property Let PictureOffset(ByVal New_PictureOffset As String)
    Dim NewX As Integer, NewY As Integer
    If Cartesian(New_PictureOffset, NewX, NewY) Then
        m_PictureOffset = New_PictureOffset
        m_PictureOffsetX = NewX
        m_PictureOffsetY = NewY
        Call AlignObjects
        Call DrawBUTTON
        PropertyChanged "PictureOffset"
    End If
End Property

Public Property Get Transparency() As OLE_COLOR
Attribute Transparency.VB_Description = "Returns/sets the colour that specifies transparent areas in all Button pictures."
Attribute Transparency.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Transparency = m_Transparency
End Property

Public Property Let Transparency(ByVal New_Transparency As OLE_COLOR)
    m_Transparency = New_Transparency
    Call SetUpPictures
    Call DrawBUTTON
    PropertyChanged "Transparency"
End Property

Public Property Get UseTransparency() As Boolean
Attribute UseTransparency.VB_Description = "Returs/sets a value that determines whether the Transparency colour is to be applied when drawing button pictures."
Attribute UseTransparency.VB_ProcData.VB_Invoke_Property = ";Appearance"
    UseTransparency = m_UseTransparency
End Property

Public Property Let UseTransparency(ByVal New_UseTransparency As Boolean)
    m_UseTransparency = New_UseTransparency
    Call SetUpPictures
    Call DrawBUTTON
    PropertyChanged "UseTransparency"
End Property

Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Returns/sets the value of the button when it is set to behave as a Check or Option button."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    If m_Behaviour <> [Command Button] _
    And m_Value <> New_Value Then
        m_Value = New_Value
        If m_Value Then
            If m_Behaviour = [Option Button] Then
                Call SetOptionsFalse
                m_Value = True
            End If
            ButtonState = (ButtonState Or [Pressed])
            If Not MouseIsDown Then RaiseEvent Click
        Else
            ButtonState = (ButtonState And Not [Pressed])
            If m_Behaviour = [Check Button] And Not MouseIsDown Then _
                RaiseEvent Click
        End If
        Call DrawBUTTON
        PropertyChanged "Value"
    End If
End Property


Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Call Flash
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    If m_Enabled Then RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If m_Enabled Then RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    ButtonState = [Raised]
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_Enabled And ((MouseKeys And Button) = Button) Then
        MouseIsDown = True
        ButtonState = [Hover Pressed]
        Call DrawBUTTON(False)
        RaiseEvent MouseDown(Button, Shift, x, y)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_Enabled Then
        If MouseIsOver(UserControl.hWnd) Then
            ButtonState = ButtonState Or [Hover]
        Else
            ButtonState = ButtonState And Not [Hover]
        End If
        Call DrawBUTTON(False)
        If tmrRCB.Enabled = False Then tmrRCB.Enabled = True
        If (ButtonState And [Pressed]) = Pressed Then
            RaiseEvent MouseMove(Button, Shift, x - 2, y - 2)
        Else
            RaiseEvent MouseMove(Button, Shift, x, y)
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_Enabled Then
        If MouseIsOver(UserControl.hWnd) Then
            Select Case m_Behaviour
                Case [Command Button]: ButtonState = [Hover]
                Case [Check Button]
                    If m_Value Then
                        Value = False
                        ButtonState = [Hover]
                    Else
                        Value = True
                    End If
                Case [Option Button]
                    If Not m_Value Then
                        Value = True
                    End If
            End Select
            RaiseEvent MouseUp(Button, Shift, x, y)
        Else
            Select Case m_Behaviour
                Case [Command Button]: ButtonState = [Raised]
                Case [Check Button], [Option Button]
                    If m_Value Then
                        ButtonState = [Pressed]
                    Else
                        ButtonState = [Raised]
                    End If
            End Select
        End If
        Call DrawBUTTON(False)
        MouseIsDown = False
    End If
End Sub

Private Sub SetOptionsFalse()
    Dim iCntr As Integer
    Dim OptFound As Boolean
    On Error GoTo SetOptionsFalse_ERROR
    For iCntr = 0 To UserControl.ParentControls.Count - 1
        If UserControl.ParentControls.Item(iCntr).Container.Name = _
           UserControl.Extender.Container.Name Then
            OptFound = True
            If UserControl.ParentControls.Item(iCntr).Behaviour = [Option Button] _
            And UserControl.ParentControls.Item(iCntr).Value = True Then
                If OptFound Then UserControl.ParentControls.Item(iCntr).Value = False
            End If
        End If
    Next iCntr
SetOptionsFalse_ERROR:
    If Err.NUMBER = 438 Then OptFound = False: Resume Next
End Sub

Public Sub Refresh()
     Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()

    Static Resizing As Boolean
    If Resizing Or Terminating Then Exit Sub
    Resizing = True
    
    Dim ucSW As Long
    Dim ucSH As Long
    Dim ucER As Long
    
    UserControl.ScaleMode = 3
    
    ucSW = UserControl.Width / Screen.TwipsPerPixelX
    ucSH = UserControl.Height / Screen.TwipsPerPixelY
    
    Select Case m_Appearance
        Case [Rectangular]
            ucER = IIf(ucSH < ucSW, CLng(ucSH / 2), CLng(ucSW / 2))
            RgnHndl = CreateRectRgn(0, 0, CLng(ucSW + 1), CLng(ucSH + 1))
            RtrnVal = SetWindowRgn(UserControl.hWnd, RgnHndl, True)
        Case [Circular]
            If ucSW < ucSH Then ucSW = ucSH
            If ucSH < ucSW Then ucSH = ucSW
            UserControl.Width = ucSW * Screen.TwipsPerPixelX
            UserControl.Height = ucSH * Screen.TwipsPerPixelY
            RgnHndl = CreateEllipticRgn(0, 0, CLng(ucSW + 1), CLng(ucSH + 1))
            RtrnVal = SetWindowRgn(UserControl.hWnd, RgnHndl, True)
        Case [Rounded Rectangle]
            ucER = IIf(ucSH < ucSW, CLng(ucSH / 2), CLng(ucSW / 2))
            RgnHndl = CreateRoundRectRgn(0, 0, CLng(ucSW + 1), CLng(ucSH + 1), ucER, ucER)
            RtrnVal = SetWindowRgn(UserControl.hWnd, RgnHndl, True)
        Case [Capsule]
            ucER = IIf(ucSH < ucSW, CLng(ucSH), CLng(ucSW))
            RgnHndl = CreateRoundRectRgn(0, 0, CLng(ucSW + 1), CLng(ucSH + 1), ucER, ucER)
            RtrnVal = SetWindowRgn(UserControl.hWnd, RgnHndl, True)
    End Select
    
    picRAISED.Move 32, 32, ucSW, ucSH
    picSUNKEN.Move 48, 48, ucSW, ucSH
    
    Call CreateBevelFX
    
    Resizing = False

End Sub

Private Sub UserControl_Terminate()
    If RgnHndl > 0 Then RtrnVal = DeleteObject(RgnHndl)
    Terminating = True
End Sub

Private Sub CreateBevelFX()

    If LoadingProperties Then Exit Sub
    
    Dim RAISEDhndl
    Dim SUNKENhndl
    
    If BevelAutoAdjust Then m_BevelHeight = 0
    RAISEDhndl = Bevel_REGION(picRAISED, m_BevelHeight, m_BackColour, m_Appearance, Convex)
    SUNKENhndl = Bevel_REGION(picSUNKEN, m_BevelHeight, m_BackColour, m_Appearance, Concave, True)
    
    Call DeleteObject(RAISEDhndl)
    Call DeleteObject(SUNKENhndl)
    
    Call SetUpPictures
    Call DrawBUTTON

End Sub

Private Sub SetUpPictures()

    If LoadingProperties Then Exit Sub
    
    Dim CntrX As Integer, CntrY As Integer
    Dim TrnsRGB As Long, CrntRGB As Long, NewHSL As HSLCol, NewRGB As Long
    
    Call AlignObjects
    
    If picBASIC.Picture.Type = vbPicTypeNone And _
       picHOVER.Picture.Type = vbPicTypeNone Then Exit Sub
    
    picWORK.Cls
    picWORK.Width = IIf(4 * BASICpstn.Width > 2 * HOVERpstn.Width, 4 * BASICpstn.Width, 2 * HOVERpstn.Width)
    picWORK.Height = BASICpstn.Height + HOVERpstn.Height
    
    TrnsRGB = RGBColour(m_Transparency)
    
    If Not picBASIC.Picture.Type = vbPicTypeNone Then
        ' 4 copies of the BASIC picture are copied into the work area, the first
        ' is used to create the cutout, the second has the transparent colour
        ' converted to Black, the third is greyed and the fourth is greyed with
        ' a Black Transparency
        For CntrX = 0 To 3 * BASICpstn.Width Step BASICpstn.Width
            BitBlt picWORK.hdc, CntrX, 0, BASICpstn.Width, BASICpstn.Height, _
                   picBASIC.hdc, 0, 0, vbSrcCopy
        Next CntrX
        For CntrY = 0 To BASICpstn.Height - 1: For CntrX = 0 To BASICpstn.Width - 1
            CrntRGB = picBASIC.Point(CntrX, CntrY)
            NewHSL = RGBtoHSL(CrntRGB)
            NewHSL.Sat = 0: NewHSL.Lum = IIf(NewHSL.Lum < 220, NewHSL.Lum + 20, 240)
            NewRGB = HSLtoRGB(NewHSL)
            If CrntRGB = TrnsRGB Then
                picWORK.PSet (CntrX, CntrY), vbWhite
                picWORK.PSet (BASICpstn.Width + CntrX, CntrY), vbBlack
                picWORK.PSet (3 * BASICpstn.Width + CntrX, CntrY), vbBlack
            Else
                picWORK.PSet (CntrX, CntrY), vbBlack
                picWORK.PSet (3 * BASICpstn.Width + CntrX, CntrY), NewRGB
            End If
            picWORK.PSet (2 * BASICpstn.Width + CntrX, CntrY), NewRGB
        Next CntrX: Next CntrY
    End If
    
    If Not picHOVER.Picture.Type = vbPicTypeNone Then
        ' Only 2 copies of the HOVER picture are required as there is no
        ' necessity for a disabled Hover picture
        For CntrX = 0 To HOVERpstn.Width Step HOVERpstn.Width
            BitBlt picWORK.hdc, CntrX, BASICpstn.Height, HOVERpstn.Width, HOVERpstn.Height, _
                   picHOVER.hdc, 0, 0, vbSrcCopy
        Next CntrX
        For CntrY = 0 To HOVERpstn.Height - 1: For CntrX = 0 To HOVERpstn.Width - 1
            CrntRGB = picHOVER.Point(CntrX, CntrY)
            If CrntRGB = TrnsRGB Then
                picWORK.PSet (CntrX, BASICpstn.Height + CntrY), vbWhite
                picWORK.PSet (HOVERpstn.Width + CntrX, BASICpstn.Height + CntrY), vbBlack
            Else
                picWORK.PSet (CntrX, BASICpstn.Height + CntrY), vbBlack
            End If
        Next CntrX: Next CntrY
    End If
    
    picWORK.Refresh

End Sub

Private Sub AlignObjects()

    CPTN.West = m_BevelHeight + m_CaptionOffsetX
    CPTN.North = m_BevelHeight + m_CaptionOffsetY
    CPTN.East = UserControl.ScaleWidth - m_BevelHeight + m_CaptionOffsetX - 1
    CPTN.South = UserControl.ScaleHeight - m_BevelHeight + m_CaptionOffsetY - 1
    
    PICT.West = m_BevelHeight + m_PictureOffsetX
    PICT.North = m_BevelHeight + m_PictureOffsetY
    PICT.East = UserControl.ScaleWidth - m_BevelHeight + m_PictureOffsetX - 1
    PICT.South = UserControl.ScaleHeight - m_BevelHeight + m_PictureOffsetY - 1
    
    Call AlignPicture(picBASIC, BASICpstn)
    Call AlignPicture(picHOVER, HOVERpstn)

End Sub

Private Sub AlignPicture(ByRef PB As PictureBox, ByRef PP As tPosition)

    If LoadingProperties Then Exit Sub
    
    If PB.Picture.Type = vbPicTypeNone Then
        PP.Left = 0: PP.Top = 0: PP.Width = 0: PP.Height = 0
        Exit Sub
    End If
    
    PP.Width = PB.Width: PP.Height = PB.Height
    
    Select Case m_PictureJustify
        Case [Top Left], [Middle Left], [Bottom Left]
            PP.Left = PICT.West
        Case [Top Centred], [Middle Centred], [Bottom Centred]
            PP.Left = PICT.West + Int((PICT.East - PICT.West - PB.Width) / 2)
        Case [Top Right], [Middle Right], [Bottom Right]
            PP.Left = PICT.East - PB.Width
    End Select
    
    Select Case m_PictureJustify
        Case [Top Left], [Top Centred], [Top Right]
            PP.Top = PICT.North
        Case [Middle Left], [Middle Centred], [Middle Right]
            PP.Top = PICT.North + Int((PICT.South - PICT.North - PB.Height) / 2)
        Case [Bottom Left], [Bottom Centred], [Bottom Right]
            PP.Top = PICT.South - PB.Height
    End Select

End Sub

Private Sub DrawBUTTON(Optional ByVal PropertyChange As Boolean = True)

    Static LastButtonState As eState
    
    If LoadingProperties Then LastButtonState = -1: Exit Sub
    If ButtonState = LastButtonState And Not PropertyChange Then Exit Sub
    
    If (ButtonState And [Pressed]) = [Pressed] Then
        BitBlt UserControl.hdc, 0, 0, picSUNKEN.Width, picSUNKEN.Height, _
               picSUNKEN.hdc, 0, 0, vbSrcCopy
    Else
        BitBlt UserControl.hdc, 0, 0, picRAISED.Width, picRAISED.Height, _
               picRAISED.hdc, 0, 0, vbSrcCopy
        'UserControl.Move 0, 0
    End If
    
    Call AddPICTURE
    Call DrawCAPTION
    
    LastButtonState = ButtonState

End Sub

Private Sub AddPICTURE()

    Dim UseHOVERpic As Boolean
    
    If (ButtonState And [Pressed]) = [Pressed] Then
        BASICpstn.Left = BASICpstn.Left + 1: BASICpstn.Top = BASICpstn.Top + 1
        HOVERpstn.Left = HOVERpstn.Left + 1: HOVERpstn.Top = HOVERpstn.Top + 1
    End If
        
    If (ButtonState And [Hover]) <> [Hover] _
    Or (ButtonState And [Disabled]) = [Disabled] Then
        If picBASIC.Picture.Type = vbPicTypeNone Then Exit Sub
    Else
        If picHOVER.Picture.Type = vbPicTypeNone Then
            If picBASIC.Picture.Type = vbPicTypeNone Then Exit Sub
        Else
            If (ButtonState And [Hover]) = [Hover] Then UseHOVERpic = True
        End If
    End If
    
    If Not m_UseTransparency Then
        If Not UseHOVERpic Then
            If (ButtonState And [Disabled]) = [Disabled] Then
                BitBlt UserControl.hdc, BASICpstn.Left, BASICpstn.Top, _
                       BASICpstn.Width, BASICpstn.Height, _
                       picWORK.hdc, 2 * BASICpstn.Width, 0, vbSrcCopy
            Else
                BitBlt UserControl.hdc, BASICpstn.Left, BASICpstn.Top, _
                       BASICpstn.Width, BASICpstn.Height, _
                       picBASIC.hdc, 0, 0, vbSrcCopy
            End If
        Else
            BitBlt UserControl.hdc, HOVERpstn.Left, HOVERpstn.Top, _
                   HOVERpstn.Width, HOVERpstn.Height, _
                   picHOVER.hdc, 0, 0, vbSrcCopy
        End If
    Else
        If Not UseHOVERpic Then
            BitBlt UserControl.hdc, BASICpstn.Left, BASICpstn.Top, _
                   BASICpstn.Width, BASICpstn.Height, _
                   picWORK.hdc, 0, 0, vbSrcAnd
            If (ButtonState And [Disabled]) = [Disabled] Then
                BitBlt UserControl.hdc, BASICpstn.Left, BASICpstn.Top, _
                       BASICpstn.Width, BASICpstn.Height, _
                       picWORK.hdc, 3 * BASICpstn.Width, 0, vbSrcPaint
            Else
                BitBlt UserControl.hdc, BASICpstn.Left, BASICpstn.Top, _
                       BASICpstn.Width, BASICpstn.Height, _
                       picWORK.hdc, BASICpstn.Width, 0, vbSrcPaint
            End If
        Else
            BitBlt UserControl.hdc, HOVERpstn.Left, HOVERpstn.Top, _
                   HOVERpstn.Width, HOVERpstn.Height, _
                   picWORK.hdc, 0, BASICpstn.Height, vbSrcAnd
            BitBlt UserControl.hdc, HOVERpstn.Left, HOVERpstn.Top, _
                   HOVERpstn.Width, HOVERpstn.Height, _
                   picWORK.hdc, HOVERpstn.Width, BASICpstn.Height, vbSrcPaint
        End If
    End If
    
    If (ButtonState And [Pressed]) = [Pressed] Then
        BASICpstn.Left = BASICpstn.Left - 1: BASICpstn.Top = BASICpstn.Top - 1
        HOVERpstn.Left = HOVERpstn.Left - 1: HOVERpstn.Top = HOVERpstn.Top - 1
    End If

End Sub

Private Sub DrawCAPTION()

    Dim wFormat As Long
    Dim DisableColour As HSLCol
    Dim PSTN As RECT
    
    Select Case m_CaptionJustify
        Case [Top Left]:       wFormat = DT_LEFT Or DT_TOP Or DT_SINGLELINE
        Case [Top Centred]:    wFormat = DT_CENTER Or DT_TOP Or DT_SINGLELINE
        Case [Top Right]:      wFormat = DT_RIGHT Or DT_TOP Or DT_SINGLELINE
        Case [Middle Left]:    wFormat = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
        Case [Middle Centred]: wFormat = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        Case [Middle Right]:   wFormat = DT_RIGHT Or DT_VCENTER Or DT_SINGLELINE
        Case [Bottom Left]:    wFormat = DT_LEFT Or DT_BOTTOM Or DT_SINGLELINE
        Case [Bottom Centred]: wFormat = DT_CENTER Or DT_BOTTOM Or DT_SINGLELINE
        Case [Bottom Right]:   wFormat = DT_RIGHT Or DT_BOTTOM Or DT_SINGLELINE
    End Select
    
    If (ButtonState And [Disabled]) = [Disabled] Then
        DisableColour = RGBtoHSL(picRAISED.BackColor)
        DisableColour.Lum = IIf(DisableColour.Lum * 1.2 < 240, DisableColour.Lum * 1.2, 240)
        UserControl.ForeColor = HSLtoRGB(DisableColour)
    Else
        If (ButtonState And [Hover]) = [Hover] Then
            UserControl.ForeColor = m_CaptionHoverColour
        Else
            UserControl.ForeColor = m_CaptionColour
        End If
    End If
    
    PSTN = CPTN
    If (ButtonState And [Pressed]) = [Pressed] Then
        PSTN.West = PSTN.West + 1: PSTN.East = PSTN.East + 1
        PSTN.North = PSTN.North + 1: PSTN.South = PSTN.South + 1
    End If
    
    DrawText UserControl.hdc, m_Caption, -1, PSTN, wFormat
    UserControl.Refresh

End Sub

Public Sub Flash()
    If m_Enabled Then
        ButtonState = [Button Flash]
        Call DrawBUTTON(False)
        tmrRCB.Enabled = True
    End If
End Sub

Private Sub tmrRCB_Timer()
    If MouseIsOver(UserControl.hWnd) Then
        If ButtonState = [Button Flash] Then
            ButtonState = [Hover]
            tmrRCB.Enabled = False
        End If
    Else
        tmrRCB.Enabled = False
        If ButtonState = [Button Flash] Then
            ButtonState = [Raised]
        Else
            ButtonState = ButtonState And Not [Hover]
            RaiseEvent MouseExit
        End If
    End If
    Call DrawBUTTON(False)
End Sub

