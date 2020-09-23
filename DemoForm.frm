VERSION 5.00
Object = "*\ABVLcontrols.vbp"
Begin VB.Form DemoForm 
   Caption         =   "Demo Form"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   367
   StartUpPosition =   3  'Windows Default
   Begin BVLcontrols.bvlBUTTON OB 
      Height          =   390
      Index           =   0
      Left            =   750
      TabIndex        =   3
      Top             =   2175
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Appearance      =   1
      Behaviour       =   2
      BevelHeight     =   4
      Caption         =   "O1"
      CaptionColour   =   16711680
      CaptionHoverColour=   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BVLcontrols.bvlBUTTON CB 
      Height          =   450
      Index           =   0
      Left            =   720
      TabIndex        =   7
      Top             =   3120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      Appearance      =   0
      Behaviour       =   1
      BevelHeight     =   4
      Caption         =   "Check1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin BVLcontrols.bvlBUTTON bvlBUTTON2 
      Height          =   375
      Left            =   1500
      TabIndex        =   2
      Top             =   1080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Appearance      =   3
      BevelHeight     =   4
      Caption         =   "Command 2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BVLcontrols.bvlBUTTON bvlBUTTON1 
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Top             =   1080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BevelHeight     =   4
      Caption         =   "Command 1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BVLcontrols.bvlNUMERIC bvlNUMERIC1 
      Height          =   750
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   1323
      BevelHeight     =   7
      NumericText     =   "3.1415927"
      InteriorColour  =   16776960
      Value           =   3.1415927
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StatusColour    =   255
      NumericColour   =   16711680
   End
   Begin BVLcontrols.bvlWHEEL bvlWHEEL5 
      Height          =   375
      Left            =   150
      TabIndex        =   12
      Top             =   1575
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BevelHeight     =   4
      BiColourUsage   =   3
      ButtonShape     =   0
      FrameShape      =   0
      WheelElapsed    =   33023
      WheelRemain     =   65280
   End
   Begin BVLcontrols.bvlWHEEL bvlWHEEL4 
      Height          =   2475
      Left            =   2925
      TabIndex        =   11
      Top             =   2025
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   4366
      BevelAutoAdjust =   0   'False
      BevelHeight     =   5
      ButtonShape     =   -1
      FrameShape      =   0
      IncrementDirection=   2
      Orientation     =   1
      Value           =   66
      WheelElapsed    =   16776960
      WheelRemain     =   16744448
   End
   Begin BVLcontrols.bvlWHEEL bvlWHEEL2 
      Height          =   4770
      Left            =   4920
      TabIndex        =   9
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   8414
      ArrowColour     =   255
      ArrowHover      =   16776960
      BevelAutoAdjust =   0   'False
      BevelHeight     =   5
      BiColourUsage   =   3
      ButtonShape     =   -1
      FrameShape      =   3
      HideNotches     =   -1  'True
      IncrementDirection=   2
      Max             =   255
      Orientation     =   1
      SpinOver        =   0   'False
      Value           =   128
      WheelElapsed    =   16776960
      WheelRemain     =   16711935
   End
   Begin BVLcontrols.bvlWHEEL bvlWHEEL1 
      Height          =   2475
      Left            =   150
      TabIndex        =   8
      Top             =   2025
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   4366
      ArrowColour     =   16711680
      ArrowHover      =   255
      BevelAutoAdjust =   0   'False
      BevelHeight     =   5
      ButtonShape     =   2
      FrameShape      =   3
      IncrementDirection=   3
      Orientation     =   1
      SpinOver        =   0   'False
      Value           =   25
      WheelElapsed    =   16776960
      WheelRemain     =   16711935
   End
   Begin BVLcontrols.bvlWHEEL bvlWHEEL3 
      Height          =   375
      Left            =   150
      TabIndex        =   10
      Top             =   4575
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BevelAutoAdjust =   0   'False
      BevelHeight     =   5
      BiColourUsage   =   3
      HideNotches     =   -1  'True
      SpinOver        =   0   'False
      WheelElapsed    =   65535
      WheelRemain     =   16711680
   End
   Begin BVLcontrols.bvlBUTTON OB 
      Height          =   390
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   2175
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Appearance      =   1
      Behaviour       =   2
      BevelHeight     =   4
      Caption         =   "O2"
      CaptionColour   =   16711680
      CaptionHoverColour=   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin BVLcontrols.bvlBUTTON OB 
      Height          =   390
      Index           =   2
      Left            =   750
      TabIndex        =   5
      Top             =   2625
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Appearance      =   1
      Behaviour       =   2
      BevelHeight     =   4
      Caption         =   "O3"
      CaptionColour   =   16711680
      CaptionHoverColour=   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BVLcontrols.bvlBUTTON OB 
      Height          =   390
      Index           =   3
      Left            =   1200
      TabIndex        =   6
      Top             =   2625
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      Appearance      =   1
      Behaviour       =   2
      BevelHeight     =   4
      Caption         =   "O4"
      CaptionColour   =   16711680
      CaptionHoverColour=   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BVLcontrols.bvlBUTTON CB 
      Height          =   450
      Index           =   1
      Left            =   720
      TabIndex        =   13
      Top             =   3720
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      Appearance      =   0
      Behaviour       =   1
      BevelHeight     =   4
      Caption         =   "Check2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
End
Attribute VB_Name = "DemoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bvlBUTTON1_Click()
    bvlWHEEL1.SetFocus
End Sub
