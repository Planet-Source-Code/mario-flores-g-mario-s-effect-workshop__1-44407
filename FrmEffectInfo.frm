VERSION 5.00
Begin VB.Form FrmEffectInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Effect Checker"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8385
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox Files 
      Height          =   480
      Left            =   7440
      Pattern         =   "*.aft;*.gif;*.class"
      TabIndex        =   117
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check"
      Height          =   375
      Left            =   5760
      TabIndex        =   116
      Top             =   8880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6840
      TabIndex        =   115
      Top             =   8880
      Width           =   855
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   24
      Left            =   2520
      TabIndex        =   120
      Top             =   3120
      Width           =   705
   End
   Begin VB.Label ll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Watermark"
      Height          =   195
      Left            =   720
      TabIndex        =   119
      Top             =   3120
      Width           =   780
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "alpha"
      Height          =   195
      Index           =   23
      Left            =   7560
      TabIndex        =   118
      Top             =   1320
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   23
      Left            =   6120
      TabIndex        =   114
      Top             =   6840
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   22
      Left            =   6120
      TabIndex        =   113
      Top             =   6600
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   21
      Left            =   6120
      TabIndex        =   112
      Top             =   6360
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   20
      Left            =   6120
      TabIndex        =   111
      Top             =   6120
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wheel"
      Height          =   195
      Index           =   22
      Left            =   3960
      TabIndex        =   110
      Top             =   6840
      Width           =   465
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Strips"
      Height          =   195
      Index           =   21
      Left            =   3960
      TabIndex        =   109
      Top             =   6600
      Width           =   390
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spiral"
      Height          =   195
      Index           =   20
      Left            =   3960
      TabIndex        =   108
      Top             =   6360
      Width           =   390
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transitions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3960
      TabIndex        =   107
      Top             =   3600
      Width           =   1365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Applets"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3960
      TabIndex        =   106
      Top             =   1800
      Width           =   930
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3DFX Effects"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3960
      TabIndex        =   105
      Top             =   120
      Width           =   1530
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3D Effects"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   720
      TabIndex        =   104
      Top             =   3600
      Width           =   1230
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2D Effects"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   720
      TabIndex        =   103
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   10
      Left            =   2520
      TabIndex        =   102
      Top             =   2880
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wave"
      Height          =   195
      Index           =   10
      Left            =   720
      TabIndex        =   101
      Top             =   2880
      Width           =   435
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   8
      Left            =   2520
      TabIndex        =   100
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engrave"
      Height          =   195
      Index           =   8
      Left            =   720
      TabIndex        =   99
      Top             =   2400
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rotate"
      Height          =   195
      Index           =   66
      Left            =   720
      TabIndex        =   98
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   7
      Left            =   2520
      TabIndex        =   97
      Top             =   2160
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emboss"
      Height          =   195
      Index           =   7
      Left            =   720
      TabIndex        =   96
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   6
      Left            =   2520
      TabIndex        =   95
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BasicImage"
      Height          =   195
      Index           =   6
      Left            =   7560
      TabIndex        =   94
      Top             =   1080
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   5
      Left            =   2520
      TabIndex        =   93
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blur"
      Height          =   195
      Index           =   5
      Left            =   720
      TabIndex        =   92
      Top             =   1680
      Width           =   270
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mirror"
      Height          =   195
      Left            =   720
      TabIndex        =   91
      Top             =   1440
      Width           =   390
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   2520
      TabIndex        =   90
      Top             =   1440
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BasicImage"
      Height          =   195
      Index           =   4
      Left            =   7560
      TabIndex        =   89
      Top             =   840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Xray"
      Height          =   195
      Left            =   720
      TabIndex        =   88
      Top             =   1200
      Width           =   315
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   87
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BasicImage"
      Height          =   195
      Index           =   3
      Left            =   7560
      TabIndex        =   86
      Top             =   600
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invert"
      Height          =   195
      Left            =   720
      TabIndex        =   85
      Top             =   960
      Width           =   405
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2520
      TabIndex        =   84
      Top             =   960
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BasicImage"
      Height          =   195
      Index           =   2
      Left            =   7560
      TabIndex        =   83
      Top             =   360
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gray"
      Height          =   195
      Left            =   720
      TabIndex        =   82
      Top             =   720
      Width           =   330
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   81
      Top             =   720
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BasicImage"
      Height          =   195
      Index           =   1
      Left            =   7560
      TabIndex        =   80
      Top             =   120
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Light"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   79
      Top             =   480
      Width           =   345
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   2520
      TabIndex        =   78
      Top             =   480
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alpha"
      Height          =   195
      Index           =   38
      Left            =   -600
      TabIndex        =   77
      Top             =   480
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BurnFilm"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   76
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checkerboard"
      Height          =   195
      Index           =   13
      Left            =   3960
      TabIndex        =   75
      Top             =   4440
      Width           =   1005
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fade"
      Height          =   195
      Index           =   14
      Left            =   3960
      TabIndex        =   74
      Top             =   4680
      Width           =   360
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blinds"
      Height          =   195
      Index           =   12
      Left            =   3960
      TabIndex        =   73
      Top             =   4200
      Width           =   420
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CenterCurls"
      Height          =   195
      Index           =   6
      Left            =   720
      TabIndex        =   72
      Top             =   5400
      Width           =   810
   End
   Begin VB.Label LabelApplet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Snow"
      Height          =   195
      Index           =   0
      Left            =   3960
      TabIndex        =   71
      Top             =   2160
      Width           =   405
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ColorFade"
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   70
      Top             =   4440
      Width           =   720
   End
   Begin VB.Label LabelApplet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stretch"
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   69
      Top             =   2400
      Width           =   510
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curls"
      Height          =   195
      Index           =   21
      Left            =   720
      TabIndex        =   68
      Top             =   9000
      Width           =   345
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curtains"
      Height          =   195
      Index           =   16
      Left            =   720
      TabIndex        =   67
      Top             =   7800
      Width           =   570
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inset"
      Height          =   195
      Index           =   16
      Left            =   3960
      TabIndex        =   66
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FlowMotion"
      Height          =   195
      Index           =   14
      Left            =   720
      TabIndex        =   65
      Top             =   7320
      Width           =   810
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FadeWhite"
      Height          =   195
      Index           =   12
      Left            =   720
      TabIndex        =   64
      Top             =   6840
      Width           =   780
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RadialWipe"
      Height          =   195
      Index           =   17
      Left            =   3960
      TabIndex        =   63
      Top             =   5640
      Width           =   825
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GlassBlock"
      Height          =   195
      Index           =   3
      Left            =   720
      TabIndex        =   62
      Top             =   4680
      Width           =   795
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RandomBars"
      Height          =   195
      Index           =   18
      Left            =   3960
      TabIndex        =   61
      Top             =   5880
      Width           =   915
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RandomDissolve"
      Height          =   195
      Index           =   19
      Left            =   3960
      TabIndex        =   60
      Top             =   6120
      Width           =   1200
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jaws"
      Height          =   195
      Index           =   13
      Left            =   720
      TabIndex        =   59
      Top             =   7080
      Width           =   360
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lens"
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   58
      Top             =   4200
      Width           =   345
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LightWipe"
      Height          =   195
      Index           =   9
      Left            =   720
      TabIndex        =   57
      Top             =   6120
      Width           =   720
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liquid"
      Height          =   195
      Index           =   4
      Left            =   720
      TabIndex        =   56
      Top             =   4920
      Width           =   420
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PageCurl"
      Height          =   195
      Index           =   7
      Left            =   720
      TabIndex        =   55
      Top             =   5640
      Width           =   645
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PeelABCD"
      Height          =   195
      Index           =   15
      Left            =   720
      TabIndex        =   54
      Top             =   7560
      Width           =   750
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pixelate"
      Height          =   195
      Index           =   9
      Left            =   720
      TabIndex        =   53
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grid"
      Height          =   195
      Index           =   18
      Left            =   720
      TabIndex        =   52
      Top             =   8280
      Width           =   285
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ripple"
      Height          =   195
      Index           =   20
      Left            =   720
      TabIndex        =   51
      Top             =   8760
      Width           =   450
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RollDown"
      Height          =   195
      Index           =   10
      Left            =   720
      TabIndex        =   50
      Top             =   6360
      Width           =   690
   End
   Begin VB.Label LabelMicrosoft2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ripple"
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   49
      Top             =   720
      Width           =   450
   End
   Begin VB.Label LabelMicrosoft2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HeightField"
      Height          =   195
      Index           =   2
      Left            =   3960
      TabIndex        =   48
      Top             =   960
      Width           =   795
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WormHole"
      Height          =   195
      Index           =   11
      Left            =   720
      TabIndex        =   47
      Top             =   6600
      Width           =   750
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Iris"
      Height          =   195
      Index           =   36
      Left            =   3960
      TabIndex        =   46
      Top             =   5400
      Width           =   195
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GradientWipe"
      Height          =   195
      Index           =   15
      Left            =   3960
      TabIndex        =   45
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Water"
      Height          =   195
      Index           =   8
      Left            =   720
      TabIndex        =   44
      Top             =   5880
      Width           =   435
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vacuum"
      Height          =   195
      Index           =   17
      Left            =   720
      TabIndex        =   43
      Top             =   8040
      Width           =   585
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Twister"
      Height          =   195
      Index           =   5
      Left            =   720
      TabIndex        =   42
      Top             =   5160
      Width           =   510
   End
   Begin VB.Label MetaCreations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Threshold"
      Height          =   195
      Index           =   19
      Left            =   720
      TabIndex        =   41
      Top             =   8520
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barn"
      Height          =   195
      Index           =   11
      Left            =   3960
      TabIndex        =   40
      Top             =   3960
      Width           =   330
   End
   Begin VB.Label LabelApplet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom"
      Height          =   195
      Index           =   2
      Left            =   3960
      TabIndex        =   39
      Top             =   2640
      Width           =   405
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   2520
      TabIndex        =   38
      Top             =   3960
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   18
      Left            =   6120
      TabIndex        =   37
      Top             =   5640
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   19
      Left            =   6120
      TabIndex        =   36
      Top             =   5880
      Width           =   705
   End
   Begin VB.Label LabelR3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   6120
      TabIndex        =   35
      Top             =   480
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   6
      Left            =   2520
      TabIndex        =   34
      Top             =   5400
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   11
      Left            =   6120
      TabIndex        =   33
      Top             =   3960
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2520
      TabIndex        =   32
      Top             =   4440
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   12
      Left            =   6120
      TabIndex        =   31
      Top             =   4200
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   13
      Left            =   6120
      TabIndex        =   30
      Top             =   4440
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   14
      Left            =   6120
      TabIndex        =   29
      Top             =   4680
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   15
      Left            =   6120
      TabIndex        =   28
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   16
      Left            =   6120
      TabIndex        =   27
      Top             =   5160
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   17
      Left            =   6120
      TabIndex        =   26
      Top             =   5400
      Width           =   705
   End
   Begin VB.Label LabelR3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   6120
      TabIndex        =   25
      Top             =   720
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   24
      Top             =   4680
      Width           =   705
   End
   Begin VB.Label LabelR3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   6120
      TabIndex        =   23
      Top             =   960
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   21
      Left            =   2520
      TabIndex        =   22
      Top             =   9000
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   20
      Left            =   2520
      TabIndex        =   21
      Top             =   8760
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   20
      Top             =   4200
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   9
      Left            =   2520
      TabIndex        =   19
      Top             =   6120
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   2520
      TabIndex        =   18
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   7
      Left            =   2520
      TabIndex        =   17
      Top             =   5640
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   14
      Left            =   2520
      TabIndex        =   16
      Top             =   7320
      Width           =   705
   End
   Begin VB.Label LabelR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   9
      Left            =   2520
      TabIndex        =   15
      Top             =   2640
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   15
      Left            =   2520
      TabIndex        =   14
      Top             =   7560
      Width           =   705
   End
   Begin VB.Label LabelRA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   6120
      TabIndex        =   13
      Top             =   2160
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   10
      Left            =   2520
      TabIndex        =   12
      Top             =   6360
      Width           =   705
   End
   Begin VB.Label LabelRA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   6120
      TabIndex        =   11
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label LabelRA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   6120
      TabIndex        =   10
      Top             =   2640
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   19
      Left            =   2520
      TabIndex        =   9
      Top             =   8520
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   18
      Left            =   2520
      TabIndex        =   8
      Top             =   8280
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   17
      Left            =   2520
      TabIndex        =   7
      Top             =   8040
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   5
      Left            =   2520
      TabIndex        =   6
      Top             =   5160
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   16
      Left            =   2520
      TabIndex        =   5
      Top             =   7800
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   8
      Left            =   2520
      TabIndex        =   4
      Top             =   5880
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   12
      Left            =   2520
      TabIndex        =   3
      Top             =   6840
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   13
      Left            =   2520
      TabIndex        =   2
      Top             =   7080
      Width           =   705
   End
   Begin VB.Label LabelR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnTested"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   11
      Left            =   2520
      TabIndex        =   1
      Top             =   6600
      Width           =   705
   End
   Begin VB.Label LabelMicrosoft2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CrShatter"
      Height          =   195
      Index           =   0
      Left            =   3960
      TabIndex        =   0
      Top             =   480
      Width           =   660
   End
End
Attribute VB_Name = "FrmEffectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
 '
 'THIS FORM IS USED TO CHECK ALL EFFECTS WORKING IN MARIO'S EFFECT WORKSHOP.
 '                           sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------

Dim LocalizarTransitions As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next

'*****CHECK MICROSOFT EFFECTS """

For i = 0 To 24
    
    LocalizarTransitions = ViewStringRegistryValue(HKEY_CLASSES_ROOT, "DXImageTransform.Microsoft." & LabelMicrosoft(i).Caption)
    If LocalizarTransitions = "YES" Then
        LabelR(i).ForeColor = vbBlue
        LabelR(i).Caption = "WORKING"
      End If
    If LocalizarTransitions = "NO" Then
        LabelR(i).Caption = "INSTALL"
    End If
  
Next i

'*****CHECK METACREATIONS EFFECTS """

For i = 0 To 21

    LocalizarTransitions = ViewStringRegistryValue(HKEY_CLASSES_ROOT, "DXImageTransform.Metacreations." & MetaCreations(i).Caption)
    If LocalizarTransitions = "YES" Then
        LabelR2(i).ForeColor = vbBlue
        LabelR2(i).Caption = "WORKING"
      End If
    If LocalizarTransitions = "NO" Then
        LabelR2(i).Caption = "INSTALL"
    End If

Next i

'*****CHECK MICROSOFT 3D EFFECTS """

For i = 0 To 2
    
    LocalizarTransitions = ViewStringRegistryValue(HKEY_CLASSES_ROOT, "DX3DTransform.Microsoft." & LabelMicrosoft2(i).Caption)
    If LocalizarTransitions = "YES" Then
        LabelR3(i).ForeColor = vbBlue
        LabelR3(i).Caption = "WORKING"
      End If
    If LocalizarTransitions = "NO" Then
        LabelR3(i).Caption = "INSTALL"
    End If

Next i

Dim TempVer, X
files.Path = App.Path & "\3rd"

'*****CHECK ANFY SNOW EFFECT """

For i = 0 To 2
     TempVer = 0
     If LabelApplet(i).Caption = "Snow" Then
        For X = 1 To files.ListCount
            files.ListIndex = X - 1
            If files.FileName = "ansnow.class" Then TempVer = TempVer + 1
        Next X
     End If

'*****CHECK ANFY STRETCH EFFECT """
     
     If LabelApplet(i).Caption = "Stretch" Then
        For X = 1 To files.ListCount
            files.ListIndex = X - 1
            If files.FileName = "anstretch.class" Then TempVer = TempVer + 1
        Next X
     End If

'*****CHECK ANFY ZOOM & PAN EFFECT """
     
     If LabelApplet(i).Caption = "Zoom" Then
        For X = 1 To files.ListCount
            files.ListIndex = X - 1
            If files.FileName = "zoompan.class" Then TempVer = TempVer + 1
        Next X
     End If
     
     If TempVer = 1 Then
            LabelRA(i).ForeColor = vbBlue
            LabelRA(i).Caption = "WORKING"
     End If
     If TempVer = 0 Then LabelRA(i).Caption = "INSTALL"
Next i

'*****CHECK ANFY HEADER CLASSES """


For i = 0 To 2
    TempVer = 0
    For X = 1 To files.ListCount
            files.ListIndex = X - 1
            If files.FileName = "Lware.class" Then TempVer = TempVer + 1
            If files.FileName = "anfy.class" Then TempVer = TempVer + 1
    Next X
        
    If TempVer <> 2 Then LabelRA(i).Caption = "MISSING CLASS"
Next i


End Sub
Private Sub Form_LostFocus()
Unload Me
End Sub

