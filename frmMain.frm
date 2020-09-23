VERSION 5.00
Object = "{58F5C475-6154-4EAD-AAA1-60BEBAEEDA36}#1.0#0"; "GradPanel.ocx"
Begin VB.Form frmMain 
   Caption         =   "Gradient Panel Example"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboCaptionStyle 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1800
      Width           =   2085
   End
   Begin VB.TextBox txtToolTip 
      Height          =   315
      Left            =   1380
      TabIndex        =   16
      Top             =   2520
      Width           =   2085
   End
   Begin VB.ComboBox cboStyle 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2160
      Width           =   2085
   End
   Begin VB.ComboBox cboCaptionAlignment 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1440
      Width           =   2085
   End
   Begin VB.TextBox txtCaption 
      Height          =   315
      Left            =   1380
      TabIndex        =   13
      Text            =   "gpWorking"
      Top             =   1080
      Width           =   2085
   End
   Begin VB.TextBox txtBevelWidth 
      Height          =   315
      Left            =   1380
      TabIndex        =   12
      Text            =   "3"
      Top             =   720
      Width           =   2085
   End
   Begin VB.TextBox txtBevelIntensity 
      Height          =   315
      Left            =   1380
      TabIndex        =   11
      Text            =   "20"
      Top             =   360
      Width           =   2085
   End
   Begin VB.ComboBox cboAppearance 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   1380
      List            =   "frmMain.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   2085
   End
   Begin GradPanel.GradientPanel gpPrimary 
      Height          =   3435
      Left            =   3510
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6059
      Alignment       =   0
      Caption         =   "Primary Gradient Panel"
      CaptionStyle    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientAngle   =   135
      GradientColor1  =   255
      GradientColor2  =   65535
      Style           =   1
      Begin GradPanel.GradientPanel gpWorking 
         Height          =   2805
         Left            =   338
         TabIndex        =   2
         Top             =   308
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   4948
         Caption         =   "gpWorking"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   45
         GradientColor1  =   16744703
         GradientColor2  =   16744576
         Picture         =   "frmMain.frx":0004
      End
   End
   Begin VB.Label lblCreditLabel 
      AutoSize        =   -1  'True
      Caption         =   "Credits:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   19
      Top             =   3180
      Width           =   660
   End
   Begin VB.Label lblCredits 
      Height          =   795
      Left            =   150
      TabIndex        =   18
      Top             =   3450
      Width           =   7515
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "ToolTipText:"
      Height          =   195
      Index           =   7
      Left            =   30
      TabIndex        =   10
      Top             =   2580
      Width           =   900
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Style:"
      Height          =   195
      Index           =   6
      Left            =   30
      TabIndex        =   9
      Top             =   2220
      Width           =   390
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Caption Style:"
      Height          =   195
      Index           =   5
      Left            =   30
      TabIndex        =   8
      Top             =   1830
      Width           =   975
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Caption:"
      Height          =   195
      Index           =   4
      Left            =   30
      TabIndex        =   7
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Bevel Width:"
      Height          =   195
      Index           =   3
      Left            =   30
      TabIndex        =   6
      Top             =   750
      Width           =   915
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Bevel Intensity:"
      Height          =   195
      Index           =   2
      Left            =   30
      TabIndex        =   5
      Top             =   420
      Width           =   1080
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Appearance:"
      Height          =   195
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Caption Alignment:"
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   3
      Top             =   1470
      Width           =   1320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CREDITS = "Kath-Rock Software - Gradient Class (Modified for improved speed)" & vbCrLf & _
                        "Microsoft - Transparent Paint Routines (Modified to work outside of a class)" & vbCrLf & _
                        "Stuart Pennington - Bevel Coding (Modified for variable intensity, default offsets, added Depressed state, and adapted to other styles)"

Private Sub cboAppearance_Click()
    gpWorking.Appearance = cboAppearance.ItemData(cboAppearance.ListIndex)
End Sub

Private Sub cboCaptionAlignment_Click()
    gpWorking.Alignment = cboCaptionAlignment.ItemData(cboCaptionAlignment.ListIndex)
End Sub

Private Sub cboCaptionStyle_Click()
    gpWorking.CaptionStyle = cboCaptionStyle.ItemData(cboCaptionStyle.ListIndex)
End Sub

Private Sub cboStyle_Click()
    gpWorking.Style = cboStyle.ItemData(cboStyle.ListIndex)
End Sub

Private Sub txtBevelIntensity_Change()
    gpWorking.BevelIntensity = CLng(Val(txtBevelIntensity.Text))
End Sub

Private Sub txtBevelWidth_Change()
    gpWorking.BevelWidth = CLng(Val(txtBevelWidth.Text))
End Sub

Private Sub txtCaption_Change()
    gpWorking.Caption = txtCaption.Text
End Sub

Private Sub txtToolTip_Change()
    gpWorking.ToolTipText = txtToolTip.Text
End Sub

'***************************************************
'This code to fill comboboxes
'***************************************************

Private Sub Form_Load()
    mFillAppearance cboAppearance
    mFillCaptionAlignment cboCaptionAlignment
    mFillCaptionStyle cboCaptionStyle
    mFillStyle cboStyle
    lblCredits.Caption = CREDITS
End Sub

Private Sub mFillAppearance(Box As ComboBox)
    Box.AddItem "None"
    Box.ItemData(Box.NewIndex) = gpaNone
    Box.AddItem "Flat Raised"
    Box.ItemData(Box.NewIndex) = gpaFlatRaised
    Box.AddItem "Flat Inset"
    Box.ItemData(Box.NewIndex) = gpaFlatInset
    Box.AddItem "3D Raised"
    Box.ItemData(Box.NewIndex) = gpa3DRaised
    Box.AddItem "3D Inset"
    Box.ItemData(Box.NewIndex) = gpa3DInset
    Box.AddItem "Etched"
    Box.ItemData(Box.NewIndex) = gpaEtched
    Box.ListIndex = Box.NewIndex    'Set default
    Box.AddItem "Bevel Raised"
    Box.ItemData(Box.NewIndex) = gpaBevelRaised
    Box.AddItem "Bevel Inset"
    Box.ItemData(Box.NewIndex) = gpaBevelInset
End Sub

Private Sub mFillCaptionAlignment(Box As ComboBox)
    Box.AddItem "Left Top"
    Box.ItemData(Box.NewIndex) = gpaLeftTop
    Box.AddItem "Left Middle"
    Box.ItemData(Box.NewIndex) = gpaLeftMiddle
    Box.AddItem "Left Bottom"
    Box.ItemData(Box.NewIndex) = gpaLeftBottom
    Box.AddItem "Right Top"
    Box.ItemData(Box.NewIndex) = gpaRightTop
    Box.AddItem "Right Middle"
    Box.ItemData(Box.NewIndex) = gpaRightMiddle
    Box.AddItem "Right Bottom"
    Box.ItemData(Box.NewIndex) = gpaRightBottom
    Box.AddItem "Center Top"
    Box.ItemData(Box.NewIndex) = gpaCenterTop
    Box.AddItem "Center Middle"
    Box.ItemData(Box.NewIndex) = gpaCenterMiddle
    Box.ListIndex = Box.NewIndex    'Set default
    Box.AddItem "Center Bottom"
    Box.ItemData(Box.NewIndex) = gpaCenterBottom
End Sub

Private Sub mFillCaptionStyle(Box As ComboBox)
    Box.AddItem "Standard"
    Box.ItemData(Box.NewIndex) = gpcStandard
    Box.ListIndex = Box.NewIndex    'Set default
    Box.AddItem "Light Inset"
    Box.ItemData(Box.NewIndex) = gpcInsetLight
    Box.AddItem "Heavy Inset"
    Box.ItemData(Box.NewIndex) = gpcInsetHeavy
    Box.AddItem "Light Raised"
    Box.ItemData(Box.NewIndex) = gpcRaisedLight
    Box.AddItem "Heavy Raised"
    Box.ItemData(Box.NewIndex) = gpcRaisedHeavy
    Box.AddItem "Drop Shadow"
    Box.ItemData(Box.NewIndex) = gpcDropShadow
End Sub

Private Sub mFillStyle(Box As ComboBox)
    Box.AddItem "Standard"
    Box.ItemData(Box.NewIndex) = gpsStandard
    Box.ListIndex = Box.NewIndex    'Set default
    Box.AddItem "Gradient"
    Box.ItemData(Box.NewIndex) = gpsGradient
    Box.AddItem "Picture"
    Box.ItemData(Box.NewIndex) = gpsPicture
    Box.AddItem "Transparent"
    Box.ItemData(Box.NewIndex) = gpsTransparent
End Sub
