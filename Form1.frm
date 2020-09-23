VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mario's effect Workshop"
   ClientHeight    =   11145
   ClientLeft      =   105
   ClientTop       =   675
   ClientWidth     =   15270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   743
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox FilesX 
      Height          =   480
      Left            =   5880
      Pattern         =   "*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf"
      TabIndex        =   454
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SlideTab 
      Height          =   900
      Left            =   0
      TabIndex        =   451
      Top             =   360
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   1588
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      ForeColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "."
      TabPicture(0)   =   "Form1.frx":062A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ImageBrowse"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   453
         Top             =   600
         Width           =   525
      End
      Begin VB.Image ImageBrowse 
         Height          =   480
         Left            =   480
         MouseIcon       =   "Form1.frx":0646
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":0950
         ToolTipText     =   "Browse For Photos "
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slide"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1275
         TabIndex        =   452
         Top             =   600
         Width           =   345
      End
      Begin VB.Image Command1 
         Height          =   480
         Left            =   1200
         MouseIcon       =   "Form1.frx":121A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1524
         ToolTipText     =   "Start Slide Show"
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame FrameAll 
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   2520
      TabIndex        =   61
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Frame FrameEffectOption 
         Caption         =   "SpotLight"
         Height          =   7095
         Index           =   37
         Left            =   240
         TabIndex        =   489
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   66
            Left            =   360
            TabIndex        =   490
            Top             =   6360
            Width           =   1215
         End
         Begin EffectWorkshop.cpvSlider SpotOpacity 
            Height          =   1020
            Left            =   1320
            Top             =   1080
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":20E6
            RailPicture     =   "Form1.frx":25E0
            Max             =   100
         End
         Begin EffectWorkshop.cpvSlider SpotSize 
            Height          =   1020
            Left            =   480
            Top             =   1080
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   1799
            SliderIcon      =   "Form1.frx":3B32
            RailPicture     =   "Form1.frx":402C
            Min             =   1
            Max             =   300
            Value           =   50
         End
         Begin VB.Label lblspeed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Opacity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   1200
            TabIndex        =   492
            Top             =   2160
            Width           =   570
         End
         Begin VB.Label lblspeed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   450
            TabIndex        =   491
            Top             =   2160
            Width           =   300
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "WaterMark"
         Height          =   7095
         Index           =   11
         Left            =   240
         TabIndex        =   472
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.TextBox TextWM 
            Height          =   315
            Left            =   240
            TabIndex        =   483
            Text            =   "Mario's Effect Workshop"
            Top             =   5640
            Width           =   1695
         End
         Begin VB.CheckBox CheckWMB 
            Caption         =   "Border"
            Height          =   255
            Left            =   720
            TabIndex        =   482
            Top             =   1800
            Width           =   855
         End
         Begin VB.Frame WaterMarkMark 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1455
            Left            =   25
            TabIndex        =   478
            Top             =   2280
            Visible         =   0   'False
            Width           =   2000
            Begin VB.PictureBox ColorFilterColor2 
               Height          =   615
               Left            =   480
               MouseIcon       =   "Form1.frx":557E
               MousePointer    =   99  'Custom
               ScaleHeight     =   555
               ScaleWidth      =   1035
               TabIndex        =   479
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label LabelColorName2 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Please Select One"
               ForeColor       =   &H00BE6B47&
               Height          =   180
               Left            =   0
               TabIndex        =   481
               Top             =   1200
               Width           =   1905
            End
            Begin VB.Label lblEffectTitle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Actual Border Color"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   45
               Left            =   330
               TabIndex        =   480
               Top             =   120
               Width           =   1410
            End
         End
         Begin VB.PictureBox ColorFilterColor 
            Height          =   615
            Left            =   480
            MouseIcon       =   "Form1.frx":5888
            MousePointer    =   99  'Custom
            ScaleHeight     =   555
            ScaleWidth      =   1035
            TabIndex        =   475
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   65
            Left            =   360
            TabIndex        =   473
            Top             =   6360
            Width           =   1215
         End
         Begin EffectWorkshop.cpvSlider WMOpacity 
            Height          =   1020
            Left            =   1560
            Top             =   3840
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":5B92
            RailPicture     =   "Form1.frx":608C
            Max             =   100
            Value           =   50
         End
         Begin EffectWorkshop.cpvSlider WMPOSY 
            Height          =   1020
            Left            =   600
            Top             =   3840
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":75DE
            RailPicture     =   "Form1.frx":7AD8
            Max             =   100
         End
         Begin EffectWorkshop.cpvSlider WMPOSX 
            Height          =   1020
            Left            =   120
            Top             =   3840
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":902A
            RailPicture     =   "Form1.frx":9524
            Max             =   100
         End
         Begin EffectWorkshop.cpvSlider WMSIZE 
            Height          =   1020
            Left            =   1080
            Top             =   3840
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   1799
            SliderIcon      =   "Form1.frx":AA76
            RailPicture     =   "Form1.frx":AF70
            Min             =   8
            Max             =   48
            Value           =   14
         End
         Begin VB.Label lblspeed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   1050
            TabIndex        =   487
            Top             =   4920
            Width           =   300
         End
         Begin VB.Label lblspeed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PosX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   486
            Top             =   4920
            Width           =   360
         End
         Begin VB.Label lblspeed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PosY"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   585
            TabIndex        =   485
            Top             =   4920
            Width           =   360
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Text"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   46
            Left            =   900
            TabIndex        =   484
            Top             =   5400
            Width           =   330
         End
         Begin VB.Label LabelColorName 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Please Select One"
            ForeColor       =   &H00BE6B47&
            Height          =   180
            Left            =   120
            TabIndex        =   476
            Top             =   1440
            Width           =   1905
         End
         Begin VB.Label lblspeed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Opacity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   477
            Top             =   4920
            Width           =   570
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Actual Font Color"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   44
            Left            =   405
            TabIndex        =   474
            Top             =   360
            Width           =   1260
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Lighting"
         Height          =   7095
         Index           =   29
         Left            =   240
         TabIndex        =   203
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.OptionButton OptionDLight 
            Caption         =   "4 Light's"
            Height          =   495
            Index           =   3
            Left            =   480
            TabIndex        =   209
            Top             =   3840
            Width           =   1095
         End
         Begin VB.OptionButton OptionDLight 
            Caption         =   "3 Light's"
            Height          =   495
            Index           =   2
            Left            =   480
            TabIndex        =   208
            Top             =   3480
            Width           =   1095
         End
         Begin VB.OptionButton OptionDLight 
            Caption         =   "2 Light's"
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   207
            Top             =   3120
            Width           =   1095
         End
         Begin VB.OptionButton OptionDLight 
            Caption         =   "1 Light's"
            Height          =   495
            Index           =   0
            Left            =   480
            TabIndex        =   206
            Top             =   2760
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   21
            Left            =   360
            TabIndex        =   204
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label Label55 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dinamic Lighting"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   360
            TabIndex        =   205
            Top             =   1560
            Width           =   1140
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Burn"
         Height          =   7095
         Index           =   12
         Left            =   240
         TabIndex        =   130
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   6
            Left            =   360
            TabIndex        =   131
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Burn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   795
            TabIndex        =   132
            Top             =   2880
            Width           =   330
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Light"
         Height          =   7095
         Index           =   0
         Left            =   240
         TabIndex        =   63
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.ComboBox AlphaComboStyle 
            Height          =   315
            ItemData        =   "Form1.frx":C4C2
            Left            =   200
            List            =   "Form1.frx":C4D2
            TabIndex        =   64
            Text            =   "0 - Uniform"
            Top             =   1080
            Width           =   1695
         End
         Begin EffectWorkshop.cpvSlider OpacitySliderS 
            Height          =   1020
            Left            =   480
            Top             =   2400
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":C50C
            RailPicture     =   "Form1.frx":CA06
            Max             =   100
            Value           =   100
         End
         Begin EffectWorkshop.cpvSlider OpacitySliderE 
            Height          =   1020
            Left            =   1245
            Top             =   2400
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":CA22
            RailPicture     =   "Form1.frx":CF1C
            Max             =   100
            Value           =   50
         End
         Begin EffectWorkshop.cpvSlider OpacitySliderX1 
            Height          =   1020
            Left            =   120
            Top             =   4680
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":E46E
            RailPicture     =   "Form1.frx":E968
            Max             =   100
         End
         Begin EffectWorkshop.cpvSlider OpacitySliderX2 
            Height          =   1020
            Left            =   720
            Top             =   4680
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":E984
            RailPicture     =   "Form1.frx":EE7E
            Max             =   100
         End
         Begin EffectWorkshop.cpvSlider OpacitySliderY1 
            Height          =   1020
            Left            =   1200
            Top             =   4680
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":103D0
            RailPicture     =   "Form1.frx":108CA
            Max             =   100
         End
         Begin EffectWorkshop.cpvSlider OpacitySliderY2 
            Height          =   1020
            Left            =   1680
            Top             =   4680
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":108E6
            RailPicture     =   "Form1.frx":10DE0
            Max             =   100
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "y2"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   1680
            TabIndex        =   113
            Top             =   4440
            Width           =   165
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "y1"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   1200
            TabIndex        =   112
            Top             =   4440
            Width           =   165
         End
         Begin VB.Label LabelLightY1 
            AutoSize        =   -1  'True
            Caption         =   "0 %"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   1080
            TabIndex        =   111
            Top             =   5760
            Width           =   255
         End
         Begin VB.Label LabelLightY2 
            AutoSize        =   -1  'True
            Caption         =   "0 %"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   1560
            TabIndex        =   110
            Top             =   5760
            Width           =   255
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x1 "
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   120
            TabIndex        =   109
            Top             =   4440
            Width           =   210
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x2"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   720
            TabIndex        =   108
            Top             =   4440
            Width           =   165
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Linear Opacity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   465
            TabIndex        =   107
            Top             =   4080
            Width           =   1050
         End
         Begin VB.Label LabelLightX1 
            AutoSize        =   -1  'True
            Caption         =   "0 %"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   75
            TabIndex        =   106
            Top             =   5760
            Width           =   255
         End
         Begin VB.Label LabelLightX2 
            AutoSize        =   -1  'True
            Caption         =   "0 %"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   600
            TabIndex        =   105
            Top             =   5760
            Width           =   255
         End
         Begin VB.Label PercentOE 
            AutoSize        =   -1  'True
            Caption         =   "50 %"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   1200
            TabIndex        =   70
            Top             =   3480
            Width           =   345
         End
         Begin VB.Label PercentOS 
            AutoSize        =   -1  'True
            Caption         =   "100 %"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   480
            TabIndex        =   69
            Top             =   3480
            Width           =   435
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Opacity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   675
            TabIndex        =   68
            Top             =   1800
            Width           =   570
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   195
            Left            =   1200
            TabIndex        =   67
            Top             =   2160
            Width           =   285
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   195
            Left            =   450
            TabIndex        =   66
            Top             =   2160
            Width           =   330
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Style"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   200
            TabIndex        =   65
            Top             =   720
            Width           =   345
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Animated Text"
         Height          =   7095
         Index           =   36
         Left            =   240
         TabIndex        =   456
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.TextBox AniTSymbol 
            Alignment       =   2  'Center
            Height          =   295
            Left            =   1080
            TabIndex        =   468
            Text            =   "&"
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox AniTDirection 
            Height          =   315
            ItemData        =   "Form1.frx":12332
            Left            =   480
            List            =   "Form1.frx":1233F
            TabIndex        =   465
            Text            =   "ellipse"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.ComboBox AniTStyle 
            Height          =   315
            ItemData        =   "Form1.frx":12362
            Left            =   240
            List            =   "Form1.frx":12372
            TabIndex        =   463
            Text            =   "A"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox AniTText 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   458
            Text            =   "Form1.frx":12382
            Top             =   3960
            Width           =   1815
         End
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   64
            Left            =   360
            TabIndex        =   457
            Top             =   6000
            Width           =   1215
         End
         Begin EffectWorkshop.cpvSlider AniTStars 
            Height          =   1020
            Left            =   360
            Top             =   2520
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   1799
            SliderIcon      =   "Form1.frx":12393
            RailPicture     =   "Form1.frx":1288D
            Max             =   200
            Value           =   16
         End
         Begin EffectWorkshop.cpvSlider AniTBlur 
            Height          =   1020
            Left            =   960
            Top             =   2520
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":128A9
            RailPicture     =   "Form1.frx":12DA3
            Min             =   1
            Max             =   255
            Value           =   25
         End
         Begin EffectWorkshop.cpvSlider AniTSpeed 
            Height          =   1020
            Left            =   1560
            Top             =   2520
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":12DBF
            RailPicture     =   "Form1.frx":132B9
            Max             =   200
            Value           =   40
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H00BE6B47&
            BackStyle       =   0  'Transparent
            Caption         =   "Star Symbol"
            ForeColor       =   &H00BE6B47&
            Height          =   195
            Left            =   1080
            TabIndex        =   467
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Speed"
            Height          =   195
            Left            =   1380
            TabIndex        =   466
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00BE6B47&
            BackStyle       =   0  'Transparent
            Caption         =   "Text Direction"
            ForeColor       =   &H00BE6B47&
            Height          =   195
            Left            =   480
            TabIndex        =   464
            Top             =   1200
            Width           =   990
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Blur"
            Height          =   195
            Left            =   960
            TabIndex        =   462
            Top             =   2160
            Width           =   270
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stars"
            Height          =   195
            Left            =   360
            TabIndex        =   461
            Top             =   2160
            Width           =   360
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00BE6B47&
            BackStyle       =   0  'Transparent
            Caption         =   "Style"
            ForeColor       =   &H00BE6B47&
            Height          =   195
            Left            =   360
            TabIndex        =   460
            Top             =   360
            Width           =   345
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00BE6B47&
            BackStyle       =   0  'Transparent
            Caption         =   "Text"
            ForeColor       =   &H00BE6B47&
            Height          =   195
            Left            =   840
            TabIndex        =   459
            Top             =   3600
            Width           =   315
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Threshold"
         Height          =   7095
         Index           =   33
         Left            =   240
         TabIndex        =   441
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   63
            Left            =   360
            TabIndex        =   442
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Threshold"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   43
            Left            =   600
            TabIndex        =   443
            Top             =   2880
            Width           =   720
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Grid"
         Height          =   7095
         Index           =   32
         Left            =   240
         TabIndex        =   438
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   62
            Left            =   360
            TabIndex        =   439
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Grid"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   42
            Left            =   810
            TabIndex        =   440
            Top             =   2880
            Width           =   300
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Jaws"
         Height          =   7095
         Index           =   31
         Left            =   240
         TabIndex        =   435
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   61
            Left            =   360
            TabIndex        =   436
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Jaws"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   41
            Left            =   780
            TabIndex        =   437
            Top             =   2880
            Width           =   360
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Color Fade"
         Height          =   7095
         Index           =   30
         Left            =   240
         TabIndex        =   432
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   60
            Left            =   360
            TabIndex        =   433
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ColorFade"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   40
            Left            =   585
            TabIndex        =   434
            Top             =   2880
            Width           =   750
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Curtains"
         Height          =   7095
         Index           =   29
         Left            =   240
         TabIndex        =   429
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   59
            Left            =   360
            TabIndex        =   430
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Curtains"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   39
            Left            =   660
            TabIndex        =   431
            Top             =   2880
            Width           =   600
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "PeelABCD"
         Height          =   7095
         Index           =   28
         Left            =   240
         TabIndex        =   426
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   58
            Left            =   360
            TabIndex        =   427
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PeelABCD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   38
            Left            =   600
            TabIndex        =   428
            Top             =   2880
            Width           =   720
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Curls"
         Height          =   7095
         Index           =   27
         Left            =   240
         TabIndex        =   423
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   57
            Left            =   360
            TabIndex        =   424
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Curls"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   37
            Left            =   780
            TabIndex        =   425
            Top             =   2880
            Width           =   360
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Ripple"
         Height          =   7095
         Index           =   26
         Left            =   240
         TabIndex        =   420
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   56
            Left            =   360
            TabIndex        =   421
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ripple"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   36
            Left            =   735
            TabIndex        =   422
            Top             =   2880
            Width           =   450
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Vacuum"
         Height          =   7095
         Index           =   25
         Left            =   240
         TabIndex        =   417
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   55
            Left            =   360
            TabIndex        =   418
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vaccum"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   35
            Left            =   690
            TabIndex        =   419
            Top             =   2880
            Width           =   540
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "FlowMotion"
         Height          =   7095
         Index           =   24
         Left            =   240
         TabIndex        =   414
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   54
            Left            =   360
            TabIndex        =   415
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "FlowMotion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   34
            Left            =   555
            TabIndex        =   416
            Top             =   2880
            Width           =   810
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Lens"
         Height          =   7095
         Index           =   23
         Left            =   240
         TabIndex        =   411
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   53
            Left            =   360
            TabIndex        =   412
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Lens"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   33
            Left            =   795
            TabIndex        =   413
            Top             =   2880
            Width           =   330
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "WormHole"
         Height          =   7095
         Index           =   22
         Left            =   240
         TabIndex        =   408
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   52
            Left            =   360
            TabIndex        =   409
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WormHole"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   32
            Left            =   585
            TabIndex        =   410
            Top             =   2880
            Width           =   750
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "RollDown"
         Height          =   7095
         Index           =   21
         Left            =   240
         TabIndex        =   405
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   51
            Left            =   360
            TabIndex        =   406
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "RollDown"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   31
            Left            =   630
            TabIndex        =   407
            Top             =   2880
            Width           =   660
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "LightWipe"
         Height          =   7095
         Index           =   20
         Left            =   240
         TabIndex        =   402
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   50
            Left            =   360
            TabIndex        =   403
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "LightWipe"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   30
            Left            =   600
            TabIndex        =   404
            Top             =   2880
            Width           =   720
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Water"
         Height          =   7095
         Index           =   19
         Left            =   240
         TabIndex        =   399
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   49
            Left            =   360
            TabIndex        =   400
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Water"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   29
            Left            =   735
            TabIndex        =   401
            Top             =   2880
            Width           =   450
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "PageCurl"
         Height          =   7095
         Index           =   18
         Left            =   240
         TabIndex        =   396
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   48
            Left            =   360
            TabIndex        =   397
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PageCurl"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   28
            Left            =   630
            TabIndex        =   398
            Top             =   2880
            Width           =   660
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Twister"
         Height          =   7095
         Index           =   17
         Left            =   240
         TabIndex        =   393
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   47
            Left            =   360
            TabIndex        =   394
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Twister"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   27
            Left            =   690
            TabIndex        =   395
            Top             =   2880
            Width           =   540
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Liquid"
         Height          =   7095
         Index           =   16
         Left            =   240
         TabIndex        =   390
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   46
            Left            =   360
            TabIndex        =   391
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Liquid"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   26
            Left            =   750
            TabIndex        =   392
            Top             =   2880
            Width           =   420
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "GlassBlock"
         Height          =   7095
         Index           =   15
         Left            =   240
         TabIndex        =   387
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   45
            Left            =   360
            TabIndex        =   388
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "GlassBlock"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   25
            Left            =   585
            TabIndex        =   389
            Top             =   2880
            Width           =   750
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Center Curls"
         Height          =   7095
         Index           =   14
         Left            =   240
         TabIndex        =   384
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   44
            Left            =   360
            TabIndex        =   385
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CenterCurls"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   525
            TabIndex        =   386
            Top             =   2880
            Width           =   870
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Burn"
         Height          =   7095
         Index           =   13
         Left            =   240
         TabIndex        =   381
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   43
            Left            =   360
            TabIndex        =   382
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Burn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   795
            TabIndex        =   383
            Top             =   2880
            Width           =   330
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Wheel"
         Height          =   7095
         Index           =   12
         Left            =   240
         TabIndex        =   373
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   42
            Left            =   360
            TabIndex        =   375
            Top             =   6000
            Width           =   1215
         End
         Begin VB.ComboBox WheelSpikes 
            Height          =   315
            ItemData        =   "Form1.frx":132D5
            Left            =   600
            List            =   "Form1.frx":132E8
            TabIndex        =   374
            Text            =   "2"
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label95 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Spikes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   735
            TabIndex        =   376
            Top             =   1080
            Width           =   450
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Strips"
         Height          =   7095
         Index           =   11
         Left            =   240
         TabIndex        =   369
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.ComboBox StripsMotion 
            Height          =   315
            ItemData        =   "Form1.frx":132FC
            Left            =   360
            List            =   "Form1.frx":1330C
            TabIndex        =   371
            Text            =   "rightup"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   41
            Left            =   360
            TabIndex        =   370
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label Label96 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Motion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   720
            TabIndex        =   372
            Top             =   1080
            Width           =   480
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Spiral"
         Height          =   7095
         Index           =   10
         Left            =   240
         TabIndex        =   363
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   40
            Left            =   360
            TabIndex        =   366
            Top             =   6000
            Width           =   1215
         End
         Begin VB.ComboBox SpiralX 
            Height          =   315
            ItemData        =   "Form1.frx":13336
            Left            =   600
            List            =   "Form1.frx":13346
            TabIndex        =   365
            Text            =   "8"
            Top             =   1800
            Width           =   735
         End
         Begin VB.ComboBox SpiralY 
            Height          =   315
            ItemData        =   "Form1.frx":13359
            Left            =   600
            List            =   "Form1.frx":13369
            TabIndex        =   364
            Text            =   "8"
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label92 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "GridSizeX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   630
            TabIndex        =   368
            Top             =   1440
            Width           =   660
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "GridSizeY"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   630
            TabIndex        =   367
            Top             =   2400
            Width           =   660
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "RandomDislove"
         Height          =   7095
         Index           =   9
         Left            =   240
         TabIndex        =   361
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   39
            Left            =   360
            TabIndex        =   362
            Top             =   6000
            Width           =   1215
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "RandomBars"
         Height          =   7095
         Index           =   8
         Left            =   240
         TabIndex        =   357
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.ComboBox RBOrientation 
            Height          =   315
            ItemData        =   "Form1.frx":1337C
            Left            =   360
            List            =   "Form1.frx":13386
            TabIndex        =   359
            Text            =   "horizontal"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   38
            Left            =   360
            TabIndex        =   358
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label Label89 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Orientation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   555
            TabIndex        =   360
            Top             =   720
            Width           =   810
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "GradientWipe"
         Height          =   7095
         Index           =   7
         Left            =   240
         TabIndex        =   353
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   37
            Left            =   360
            TabIndex        =   355
            Top             =   6000
            Width           =   1215
         End
         Begin VB.ComboBox wipeStyle 
            Height          =   315
            ItemData        =   "Form1.frx":133A0
            Left            =   360
            List            =   "Form1.frx":133AD
            TabIndex        =   354
            Text            =   "Radial"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label87 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WipeStyle"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   600
            TabIndex        =   356
            Top             =   720
            Width           =   720
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Iris"
         Height          =   7095
         Index           =   6
         Left            =   240
         TabIndex        =   347
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.ComboBox IrisMotion 
            Height          =   315
            ItemData        =   "Form1.frx":133C7
            Left            =   600
            List            =   "Form1.frx":133D1
            TabIndex        =   350
            Text            =   "in"
            Top             =   2760
            Width           =   735
         End
         Begin VB.ComboBox IrisStyle 
            Height          =   315
            ItemData        =   "Form1.frx":133DE
            Left            =   360
            List            =   "Form1.frx":133F4
            TabIndex        =   349
            Text            =   "Diamond"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   36
            Left            =   360
            TabIndex        =   348
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label Label86 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Motion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   720
            TabIndex        =   352
            Top             =   2400
            Width           =   480
         End
         Begin VB.Label Label85 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "IrisStyle"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   660
            TabIndex        =   351
            Top             =   720
            Width           =   600
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Inset"
         Height          =   7095
         Index           =   5
         Left            =   240
         TabIndex        =   345
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   35
            Left            =   360
            TabIndex        =   346
            Top             =   6000
            Width           =   1215
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "GradientWipe"
         Height          =   7095
         Index           =   4
         Left            =   240
         TabIndex        =   337
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   34
            Left            =   360
            TabIndex        =   341
            Top             =   6000
            Width           =   1215
         End
         Begin VB.ComboBox GWSize 
            Height          =   315
            ItemData        =   "Form1.frx":13424
            Left            =   360
            List            =   "Form1.frx":13437
            TabIndex        =   340
            Text            =   "0.50"
            Top             =   840
            Width           =   1215
         End
         Begin VB.ComboBox GWStyle 
            Height          =   315
            ItemData        =   "Form1.frx":13459
            Left            =   360
            List            =   "Form1.frx":13463
            TabIndex        =   339
            Text            =   "Left-to-Right"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.ComboBox GWMotion 
            Height          =   315
            ItemData        =   "Form1.frx":13485
            Left            =   360
            List            =   "Form1.frx":1348F
            TabIndex        =   338
            Text            =   "reverse"
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label81 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "GradientSize"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   510
            TabIndex        =   344
            Top             =   480
            Width           =   900
         End
         Begin VB.Label Label78 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WhipeStyle"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   555
            TabIndex        =   343
            Top             =   1680
            Width           =   810
         End
         Begin VB.Label Label75 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Motion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   720
            TabIndex        =   342
            Top             =   2760
            Width           =   480
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Fade"
         Height          =   7095
         Index           =   3
         Left            =   240
         TabIndex        =   335
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   33
            Left            =   360
            TabIndex        =   336
            Top             =   6000
            Width           =   1215
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "CheckerBoard"
         Height          =   7095
         Index           =   2
         Left            =   240
         TabIndex        =   327
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.ComboBox ChBoardy 
            Height          =   315
            ItemData        =   "Form1.frx":134A5
            Left            =   360
            List            =   "Form1.frx":134BE
            TabIndex        =   333
            Text            =   "2"
            Top             =   3120
            Width           =   1215
         End
         Begin VB.ComboBox ChBoardx 
            Height          =   315
            ItemData        =   "Form1.frx":134DA
            Left            =   360
            List            =   "Form1.frx":134F3
            TabIndex        =   331
            Text            =   "2"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox ChBoardDir 
            Height          =   315
            ItemData        =   "Form1.frx":1350F
            Left            =   360
            List            =   "Form1.frx":1351F
            TabIndex        =   329
            Text            =   "Up"
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   32
            Left            =   360
            TabIndex        =   328
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label Label74 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Squares Y"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   600
            TabIndex        =   334
            Top             =   2760
            Width           =   720
         End
         Begin VB.Label Label71 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Squares X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   600
            TabIndex        =   332
            Top             =   1680
            Width           =   720
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Direction"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   645
            TabIndex        =   330
            Top             =   480
            Width           =   630
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Blinds"
         Height          =   7095
         Index           =   1
         Left            =   240
         TabIndex        =   321
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   31
            Left            =   360
            TabIndex        =   324
            Top             =   6000
            Width           =   1215
         End
         Begin VB.ComboBox BlindsDirection 
            Height          =   315
            ItemData        =   "Form1.frx":1353A
            Left            =   360
            List            =   "Form1.frx":1354A
            TabIndex        =   323
            Text            =   "Up"
            Top             =   840
            Width           =   1215
         End
         Begin VB.ComboBox BlindsBands 
            Height          =   315
            ItemData        =   "Form1.frx":13565
            Left            =   600
            List            =   "Form1.frx":13578
            TabIndex        =   322
            Text            =   "2"
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Direction"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   645
            TabIndex        =   326
            Top             =   480
            Width           =   630
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Bands"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   735
            TabIndex        =   325
            Top             =   1440
            Width           =   450
         End
      End
      Begin VB.Frame FrameTransitionOption 
         Caption         =   "Barn"
         Height          =   7095
         Index           =   0
         Left            =   240
         TabIndex        =   312
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.ComboBox BarnMotion 
            Height          =   315
            ItemData        =   "Form1.frx":1358C
            Left            =   360
            List            =   "Form1.frx":13596
            TabIndex        =   320
            Text            =   "in"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.ComboBox BarnOrientation 
            Height          =   315
            ItemData        =   "Form1.frx":135A3
            Left            =   360
            List            =   "Form1.frx":135AD
            TabIndex        =   318
            Text            =   "horizontal"
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   30
            Left            =   360
            TabIndex        =   313
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Motion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   720
            TabIndex        =   319
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Orientation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   555
            TabIndex        =   314
            Top             =   480
            Width           =   810
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Zoom"
         Height          =   7095
         Index           =   35
         Left            =   240
         TabIndex        =   245
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   28
            Left            =   360
            TabIndex        =   246
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblAppletTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Zoom"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   765
            TabIndex        =   247
            Top             =   2880
            Width           =   390
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Snow"
         Height          =   7095
         Index           =   34
         Left            =   240
         TabIndex        =   236
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   27
            Left            =   360
            TabIndex        =   237
            Top             =   6000
            Width           =   1215
         End
         Begin EffectWorkshop.cpvSlider SnowFlake1 
            Height          =   1020
            Left            =   240
            Top             =   1080
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":135C7
            RailPicture     =   "Form1.frx":13AC1
            Max             =   500
            Value           =   80
         End
         Begin EffectWorkshop.cpvSlider SnowFlake2 
            Height          =   1020
            Left            =   600
            Top             =   1080
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":13ADD
            RailPicture     =   "Form1.frx":13FD7
            Max             =   500
            Value           =   100
         End
         Begin EffectWorkshop.cpvSlider SnowFlake3 
            Height          =   1020
            Left            =   960
            Top             =   1080
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":15529
            RailPicture     =   "Form1.frx":15A23
            Max             =   500
            Value           =   150
         End
         Begin EffectWorkshop.cpvSlider SnowFlake4 
            Height          =   1020
            Left            =   1395
            Top             =   1080
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":15A3F
            RailPicture     =   "Form1.frx":15F39
            Max             =   500
            Value           =   50
         End
         Begin EffectWorkshop.cpvSlider SnowSpeed 
            Height          =   1020
            Left            =   840
            Top             =   3120
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":1748B
            RailPicture     =   "Form1.frx":17985
            Max             =   100
            Value           =   10
         End
         Begin VB.Label Label69 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Speed"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   720
            TabIndex        =   243
            Top             =   2880
            Width           =   450
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   195
            Left            =   1050
            TabIndex        =   242
            Top             =   840
            Width           =   90
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            Height          =   195
            Left            =   1440
            TabIndex        =   241
            Top             =   840
            Width           =   90
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   195
            Left            =   360
            TabIndex        =   240
            Top             =   840
            Width           =   90
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   195
            Left            =   680
            TabIndex        =   239
            Top             =   840
            Width           =   90
         End
         Begin VB.Label Label68 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SnowFlakes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   495
            TabIndex        =   238
            Top             =   360
            Width           =   840
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Stretch"
         Height          =   7095
         Index           =   33
         Left            =   240
         TabIndex        =   232
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   26
            Left            =   360
            TabIndex        =   233
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblAppletTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Stretch"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   660
            TabIndex        =   234
            Top             =   2880
            Width           =   540
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Ripple"
         Height          =   7095
         Index           =   32
         Left            =   240
         TabIndex        =   225
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   25
            Left            =   360
            TabIndex        =   226
            Top             =   6000
            Width           =   1215
         End
         Begin EffectWorkshop.cpvSlider RippleTime 
            Height          =   1020
            Left            =   360
            Top             =   3600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":18ED7
            RailPicture     =   "Form1.frx":193D1
            Max             =   100
            Value           =   15
         End
         Begin EffectWorkshop.cpvSlider RippleRotate 
            Height          =   1020
            Left            =   720
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":193ED
            RailPicture     =   "Form1.frx":198E7
            Max             =   365
         End
         Begin EffectWorkshop.cpvSlider RippleSize 
            Height          =   1020
            Left            =   1200
            Top             =   3600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":19903
            RailPicture     =   "Form1.frx":19DFD
            Max             =   100
            Value           =   57
         End
         Begin VB.Label Label66 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   345
            TabIndex        =   230
            Top             =   3360
            Width           =   330
         End
         Begin VB.Label Label65 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1185
            TabIndex        =   229
            Top             =   3360
            Width           =   300
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "0 Grades"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   480
            TabIndex        =   228
            Top             =   2640
            Width           =   645
         End
         Begin VB.Label Label62 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rotate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   600
            TabIndex        =   227
            Top             =   1080
            Width           =   510
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "HeightField"
         Height          =   7095
         Index           =   31
         Left            =   240
         TabIndex        =   218
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   24
            Left            =   360
            TabIndex        =   219
            Top             =   6000
            Width           =   1215
         End
         Begin EffectWorkshop.cpvSlider HeightFieldTime 
            Height          =   1020
            Left            =   360
            Top             =   3600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":19E19
            RailPicture     =   "Form1.frx":1A313
            Max             =   100
            Value           =   15
         End
         Begin EffectWorkshop.cpvSlider HeightFieldRotate 
            Height          =   1020
            Left            =   750
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":1A32F
            RailPicture     =   "Form1.frx":1A829
            Max             =   365
         End
         Begin EffectWorkshop.cpvSlider HeightFieldSize 
            Height          =   1020
            Left            =   1200
            Top             =   3600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":1A845
            RailPicture     =   "Form1.frx":1AD3F
            Max             =   100
            Value           =   57
         End
         Begin VB.Label Label63 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rotate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   600
            TabIndex        =   223
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label LabelHeightFieldRotate 
            AutoSize        =   -1  'True
            Caption         =   "0 Grades"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   480
            TabIndex        =   222
            Top             =   2640
            Width           =   645
         End
         Begin VB.Label Label61 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1185
            TabIndex        =   221
            Top             =   3360
            Width           =   300
         End
         Begin VB.Label Label60 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   345
            TabIndex        =   220
            Top             =   3360
            Width           =   330
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Shatter"
         Height          =   7095
         Index           =   30
         Left            =   240
         TabIndex        =   210
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin EffectWorkshop.cpvSlider ShatterTime 
            Height          =   1020
            Left            =   360
            Top             =   3600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":1AD5B
            RailPicture     =   "Form1.frx":1B255
            Max             =   100
            Value           =   15
         End
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   23
            Left            =   360
            TabIndex        =   211
            Top             =   6000
            Width           =   1215
         End
         Begin EffectWorkshop.cpvSlider ShatterRotate 
            Height          =   1020
            Left            =   750
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":1B271
            RailPicture     =   "Form1.frx":1B76B
            Max             =   365
         End
         Begin EffectWorkshop.cpvSlider ShatterSize 
            Height          =   1020
            Left            =   1200
            Top             =   3600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":1B787
            RailPicture     =   "Form1.frx":1BC81
            Max             =   100
            Value           =   57
         End
         Begin VB.Label Label58 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   345
            TabIndex        =   215
            Top             =   3360
            Width           =   330
         End
         Begin VB.Label Label59 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1185
            TabIndex        =   214
            Top             =   3360
            Width           =   300
         End
         Begin VB.Label LabelShatterrotate 
            AutoSize        =   -1  'True
            Caption         =   "0 Grades"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   480
            TabIndex        =   213
            Top             =   2640
            Width           =   645
         End
         Begin VB.Label Label57 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rotate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   600
            TabIndex        =   212
            Top             =   1080
            Width           =   510
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Curtains"
         Height          =   7095
         Index           =   28
         Left            =   240
         TabIndex        =   199
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   20
            Left            =   360
            TabIndex        =   200
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Curtains"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   22
            Left            =   720
            TabIndex        =   201
            Top             =   2880
            Width           =   600
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "PeelABCD"
         Height          =   7095
         Index           =   27
         Left            =   240
         TabIndex        =   195
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   18
            Left            =   360
            TabIndex        =   196
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PeelABCD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   720
            TabIndex        =   197
            Top             =   2880
            Width           =   720
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Curls"
         Height          =   7095
         Index           =   26
         Left            =   240
         TabIndex        =   191
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   17
            Left            =   360
            TabIndex        =   192
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Curls"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   840
            TabIndex        =   193
            Top             =   2880
            Width           =   360
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Ripple"
         Height          =   7095
         Index           =   25
         Left            =   240
         TabIndex        =   184
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   22
            Left            =   360
            TabIndex        =   185
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ripple"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   705
            TabIndex        =   186
            Top             =   2880
            Width           =   450
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Vacuum"
         Height          =   7095
         Index           =   24
         Left            =   240
         TabIndex        =   181
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   19
            Left            =   360
            TabIndex        =   182
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vacuum"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   645
            TabIndex        =   183
            Top             =   2880
            Width           =   570
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "FlowMotion"
         Height          =   7095
         Index           =   23
         Left            =   240
         TabIndex        =   187
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   16
            Left            =   360
            TabIndex        =   188
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "FlowMotion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   600
            TabIndex        =   189
            Top             =   3120
            Width           =   810
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Lens"
         Height          =   7095
         Index           =   22
         Left            =   240
         TabIndex        =   178
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   15
            Left            =   360
            TabIndex        =   179
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Lens"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   765
            TabIndex        =   180
            Top             =   2880
            Width           =   330
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Wormhole"
         Height          =   7095
         Index           =   21
         Left            =   240
         TabIndex        =   175
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   14
            Left            =   360
            TabIndex        =   176
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Wormhole"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   570
            TabIndex        =   177
            Top             =   2880
            Width           =   720
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "RollDown"
         Height          =   7095
         Index           =   20
         Left            =   240
         TabIndex        =   154
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   13
            Left            =   360
            TabIndex        =   155
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "RollDown"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   600
            TabIndex        =   156
            Top             =   2880
            Width           =   660
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "LightWipe"
         Height          =   7095
         Index           =   19
         Left            =   240
         TabIndex        =   151
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   12
            Left            =   360
            TabIndex        =   152
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "LightWipe"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   570
            TabIndex        =   153
            Top             =   2880
            Width           =   720
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Water"
         Height          =   7095
         Index           =   18
         Left            =   240
         TabIndex        =   148
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   11
            Left            =   360
            TabIndex        =   149
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Water"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   705
            TabIndex        =   150
            Top             =   2880
            Width           =   450
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Twister"
         Height          =   7095
         Index           =   16
         Left            =   240
         TabIndex        =   142
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   10
            Left            =   360
            TabIndex        =   143
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Twister"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   660
            TabIndex        =   144
            Top             =   2880
            Width           =   540
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "PageCurl"
         Height          =   7095
         Index           =   17
         Left            =   240
         TabIndex        =   145
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   29
            Left            =   360
            TabIndex        =   146
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PageCurl"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   600
            TabIndex        =   147
            Top             =   2880
            Width           =   660
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Liquid"
         Height          =   7095
         Index           =   15
         Left            =   240
         TabIndex        =   139
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   9
            Left            =   360
            TabIndex        =   140
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Liquid"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   750
            TabIndex        =   141
            Top             =   2880
            Width           =   420
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "GlassBlock"
         Height          =   7095
         Index           =   14
         Left            =   240
         TabIndex        =   136
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   8
            Left            =   360
            TabIndex        =   137
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "GlassBlock"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   585
            TabIndex        =   138
            Top             =   2880
            Width           =   750
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "CenterCurls"
         Height          =   7095
         Index           =   13
         Left            =   240
         TabIndex        =   133
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   7
            Left            =   360
            TabIndex        =   134
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CenterCurls"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   525
            TabIndex        =   135
            Top             =   2880
            Width           =   870
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Wave"
         Height          =   7095
         Index           =   10
         Left            =   240
         TabIndex        =   96
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin EffectWorkshop.cpvSlider SliderGlowFreq 
            Height          =   1020
            Left            =   360
            Top             =   1440
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":1BC9D
            RailPicture     =   "Form1.frx":1C197
            Max             =   50
         End
         Begin EffectWorkshop.cpvSlider SliderGlowLight 
            Height          =   1020
            Left            =   1320
            Top             =   1440
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":1C1B3
            RailPicture     =   "Form1.frx":1C6AD
            Max             =   100
         End
         Begin EffectWorkshop.cpvSlider SliderGlowPhase 
            Height          =   1020
            Left            =   345
            Top             =   3480
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":1C6C9
            RailPicture     =   "Form1.frx":1CBC3
            Max             =   100
         End
         Begin EffectWorkshop.cpvSlider SliderGlowStrength 
            Height          =   1020
            Left            =   1275
            Top             =   3480
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":1CBDF
            RailPicture     =   "Form1.frx":1D0D9
            Max             =   100
         End
         Begin VB.Label lblGlowStrength 
            AutoSize        =   -1  'True
            Caption         =   "0 %"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   1275
            TabIndex        =   104
            Top             =   4560
            Width           =   255
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Strength"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1110
            TabIndex        =   103
            Top             =   3120
            Width           =   630
         End
         Begin VB.Label lblGlowPhase 
            AutoSize        =   -1  'True
            Caption         =   "0 %"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   345
            TabIndex        =   102
            Top             =   4560
            Width           =   255
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Phase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   270
            TabIndex        =   101
            Top             =   3120
            Width           =   450
         End
         Begin VB.Label lblGlowLight 
            AutoSize        =   -1  'True
            Caption         =   "0 %"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   1320
            TabIndex        =   100
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Light"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1290
            TabIndex        =   99
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lblGlowFreq 
            AutoSize        =   -1  'True
            Caption         =   "0 %"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   360
            TabIndex        =   98
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Frequency"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   1080
            Width           =   780
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Pixelate"
         Height          =   7095
         Index           =   9
         Left            =   240
         TabIndex        =   93
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin EffectWorkshop.cpvSlider SliderPixelate 
            Height          =   1020
            Left            =   840
            Top             =   2280
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":1D0F5
            RailPicture     =   "Form1.frx":1D5EF
            Min             =   2
            Max             =   50
            Value           =   2
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pixelate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   675
            TabIndex        =   95
            Top             =   1920
            Width           =   570
         End
         Begin VB.Label LabelPixel 
            AutoSize        =   -1  'True
            Caption         =   "0 %"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   840
            TabIndex        =   94
            Top             =   3360
            Width           =   255
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Engrave"
         Height          =   7095
         Index           =   8
         Left            =   240
         TabIndex        =   90
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   5
            Left            =   360
            TabIndex        =   91
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Engrave"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   660
            TabIndex        =   92
            Top             =   2880
            Width           =   600
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Emboss"
         Height          =   7095
         Index           =   7
         Left            =   240
         TabIndex        =   87
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   4
            Left            =   360
            TabIndex        =   88
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Emboss"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   690
            TabIndex        =   89
            Top             =   2880
            Width           =   540
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Rotate"
         Height          =   7095
         Index           =   6
         Left            =   240
         TabIndex        =   84
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.ComboBox ComboRotate 
            Height          =   315
            ItemData        =   "Form1.frx":1D60B
            Left            =   240
            List            =   "Form1.frx":1D61B
            TabIndex        =   85
            Text            =   "0     Degrees"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rotation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   600
            TabIndex        =   86
            Top             =   1080
            Width           =   630
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Mirror"
         Height          =   7095
         Index           =   5
         Left            =   240
         TabIndex        =   81
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   3
            Left            =   360
            TabIndex        =   82
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Mirror"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   750
            TabIndex        =   83
            Top             =   2880
            Width           =   420
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Xray"
         Height          =   7095
         Index           =   4
         Left            =   240
         TabIndex        =   78
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   2
            Left            =   360
            TabIndex        =   79
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Xray"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   795
            TabIndex        =   80
            Top             =   2880
            Width           =   330
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Invert"
         Height          =   7095
         Index           =   3
         Left            =   240
         TabIndex        =   75
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   1
            Left            =   360
            TabIndex        =   76
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Invert"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   735
            TabIndex        =   77
            Top             =   2880
            Width           =   450
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "GrayScale"
         Height          =   7095
         Index           =   2
         Left            =   240
         TabIndex        =   72
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton ButtonActivate 
            Caption         =   "Activate"
            Height          =   495
            Index           =   0
            Left            =   360
            TabIndex        =   74
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblEffectTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "GrayScale"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   73
            Top             =   2880
            Width           =   720
         End
      End
      Begin VB.Frame FrameEffectOption 
         Caption         =   "Blur"
         Height          =   7095
         Index           =   1
         Left            =   240
         TabIndex        =   71
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         Begin VB.Frame FrameMotionBlur 
            BorderStyle     =   0  'None
            Height          =   4335
            Left            =   120
            TabIndex        =   115
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
            Begin VB.OptionButton OptionBlurDir 
               Alignment       =   1  'Right Justify
               Caption         =   "315"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   129
               Top             =   1200
               Width           =   615
            End
            Begin VB.OptionButton OptionBlurDir 
               Alignment       =   1  'Right Justify
               Caption         =   "270"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   128
               Top             =   1560
               Width           =   615
            End
            Begin VB.OptionButton OptionBlurDir 
               Alignment       =   1  'Right Justify
               Caption         =   "225"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   127
               Top             =   1920
               Width           =   615
            End
            Begin VB.OptionButton OptionBlurDir 
               Caption         =   "180"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   4
               Left            =   840
               TabIndex        =   126
               Top             =   2160
               Width           =   615
            End
            Begin VB.OptionButton OptionBlurDir 
               Caption         =   "135"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   3
               Left            =   1200
               TabIndex        =   125
               Top             =   1920
               Width           =   615
            End
            Begin VB.OptionButton OptionBlurDir 
               Caption         =   "90"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   1320
               TabIndex        =   124
               Top             =   1560
               Width           =   495
            End
            Begin VB.OptionButton OptionBlurDir 
               Caption         =   "45"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   123
               Top             =   1200
               Width           =   495
            End
            Begin VB.OptionButton OptionBlurDir 
               Caption         =   "0"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   0
               Left            =   840
               TabIndex        =   122
               Top             =   960
               Value           =   -1  'True
               Width           =   375
            End
            Begin EffectWorkshop.cpvSlider MotionBlurStr 
               Height          =   240
               Left            =   360
               Top             =   3720
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   423
               SliderIcon      =   "Form1.frx":1D656
               Orientation     =   0
               RailPicture     =   "Form1.frx":1DB70
               Max             =   100
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Strength"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   480
               TabIndex        =   118
               Top             =   3360
               Width           =   630
            End
            Begin VB.Label LabelBlurStr 
               AutoSize        =   -1  'True
               Caption         =   "0 %"
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   720
               TabIndex        =   117
               Top             =   3960
               Width           =   255
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Direction"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   480
               TabIndex        =   116
               Top             =   240
               Width           =   630
            End
         End
         Begin VB.Frame FrameBlur1 
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   600
            TabIndex        =   119
            Top             =   360
            Width           =   735
            Begin EffectWorkshop.cpvSlider BlurX 
               Height          =   1020
               Left            =   240
               Top             =   480
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   423
               SliderIcon      =   "Form1.frx":1DB8C
               RailPicture     =   "Form1.frx":1E086
               Max             =   25
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Blur"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   240
               TabIndex        =   121
               Top             =   240
               Width           =   270
            End
            Begin VB.Label PercentBX 
               AutoSize        =   -1  'True
               Caption         =   "0 %"
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   240
               TabIndex        =   120
               Top             =   1560
               Width           =   255
            End
         End
         Begin VB.CheckBox CheckMotionBlur 
            Caption         =   "Motion Blur"
            Height          =   495
            Left            =   480
            TabIndex        =   114
            Top             =   6240
            Width           =   1215
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   7935
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   13996
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Options"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab EffectsTab 
      Height          =   915
      Left            =   2400
      TabIndex        =   57
      Top             =   10200
      Visible         =   0   'False
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   1614
      _Version        =   393216
      TabHeight       =   529
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "2-D"
      TabPicture(0)   =   "Form1.frx":1E0A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Option1(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Option1(10)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option1(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option1(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option1(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option1(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option1(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option1(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Option1(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Option1(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Option1(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "3-D"
      TabPicture(1)   =   "Form1.frx":1E0BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ButtonLeftEffect(0)"
      Tab(1).Control(1)=   "ButtonRightEffect(0)"
      Tab(1).Control(2)=   "FrameEffects(1)"
      Tab(1).Control(3)=   "FrameEffects(2)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Dynamic"
      TabPicture(2)   =   "Form1.frx":1E0DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Option1(29)"
      Tab(2).Control(1)=   "Option1(30)"
      Tab(2).Control(2)=   "Option1(31)"
      Tab(2).Control(3)=   "Option1(32)"
      Tab(2).Control(4)=   "Option1(33)"
      Tab(2).Control(5)=   "Option1(34)"
      Tab(2).Control(6)=   "Option1(35)"
      Tab(2).Control(7)=   "Option1(36)"
      Tab(2).Control(8)=   "Option1(37)"
      Tab(2).ControlCount=   9
      Begin VB.OptionButton Option1 
         Caption         =   "SpotLight"
         Height          =   495
         Index           =   37
         Left            =   -73320
         TabIndex        =   488
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Watermark"
         Height          =   495
         Index           =   11
         Left            =   10680
         TabIndex        =   471
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "AnimText"
         Height          =   495
         Index           =   36
         Left            =   -64680
         TabIndex        =   455
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Light"
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   310
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "GrayScale"
         Height          =   495
         Index           =   2
         Left            =   2040
         TabIndex        =   309
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Invert"
         Height          =   495
         Index           =   3
         Left            =   3240
         TabIndex        =   308
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Blur"
         Height          =   495
         Index           =   1
         Left            =   1320
         TabIndex        =   307
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Xray"
         Height          =   495
         Index           =   4
         Left            =   4080
         TabIndex        =   306
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mirror"
         Height          =   495
         Index           =   5
         Left            =   4920
         TabIndex        =   305
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rotate"
         Height          =   495
         Index           =   6
         Left            =   5760
         TabIndex        =   304
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Emboss"
         Height          =   495
         Index           =   7
         Left            =   6720
         TabIndex        =   303
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Engrave"
         Height          =   495
         Index           =   8
         Left            =   7680
         TabIndex        =   302
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Wave"
         Height          =   495
         Index           =   10
         Left            =   9840
         TabIndex        =   301
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pixelate"
         Height          =   495
         Index           =   9
         Left            =   8760
         TabIndex        =   300
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame FrameEffects 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   2
         Left            =   -73440
         TabIndex        =   157
         Top             =   360
         Visible         =   0   'False
         Width           =   9975
         Begin VB.OptionButton Option1 
            Caption         =   "BurnFilm"
            Height          =   495
            Index           =   12
            Left            =   0
            TabIndex        =   166
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "CenterCurls"
            Height          =   495
            Index           =   13
            Left            =   1080
            TabIndex        =   165
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "GlassBlock"
            Height          =   495
            Index           =   14
            Left            =   2400
            TabIndex        =   164
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Liquid"
            Height          =   495
            Index           =   15
            Left            =   3600
            TabIndex        =   163
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Twister"
            Height          =   495
            Index           =   16
            Left            =   4440
            TabIndex        =   162
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "PageCurl"
            Height          =   495
            Index           =   17
            Left            =   5400
            TabIndex        =   161
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Water"
            Height          =   495
            Index           =   18
            Left            =   6600
            TabIndex        =   160
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "LightWipe"
            Height          =   495
            Index           =   19
            Left            =   7560
            TabIndex        =   159
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "RollDown"
            Height          =   495
            Index           =   20
            Left            =   8760
            TabIndex        =   158
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Zoom"
         Height          =   495
         Index           =   35
         Left            =   -65640
         TabIndex        =   244
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Snow"
         Height          =   495
         Index           =   34
         Left            =   -66720
         TabIndex        =   235
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Stretch"
         Height          =   495
         Index           =   33
         Left            =   -67680
         TabIndex        =   231
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ripple3D"
         Height          =   495
         Index           =   32
         Left            =   -68880
         TabIndex        =   224
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HeightField"
         Height          =   495
         Index           =   31
         Left            =   -70080
         TabIndex        =   217
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CrShatter"
         Height          =   495
         Index           =   30
         Left            =   -71160
         TabIndex        =   216
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame FrameEffects 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   -72960
         TabIndex        =   169
         Top             =   360
         Width           =   9015
         Begin VB.OptionButton Option1 
            Caption         =   "Curtains"
            Height          =   495
            Index           =   28
            Left            =   8040
            TabIndex        =   198
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "PeelABCD"
            Height          =   495
            Index           =   27
            Left            =   6840
            TabIndex        =   194
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Curls"
            Height          =   495
            Index           =   26
            Left            =   6000
            TabIndex        =   190
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ripple"
            Height          =   495
            Index           =   25
            Left            =   5040
            TabIndex        =   174
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Vacuum"
            Height          =   495
            Index           =   24
            Left            =   3840
            TabIndex        =   173
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "FlowMotion"
            Height          =   495
            Index           =   23
            Left            =   2400
            TabIndex        =   172
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Lens"
            Height          =   495
            Index           =   22
            Left            =   1440
            TabIndex        =   171
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Wormhole"
            Height          =   495
            Index           =   21
            Left            =   120
            TabIndex        =   170
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Lighting"
         Height          =   495
         Index           =   29
         Left            =   -72240
         TabIndex        =   202
         Top             =   360
         Width           =   1095
      End
      Begin VB.PictureBox ButtonRightEffect 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -63000
         Picture         =   "Form1.frx":1E0F6
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   168
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox ButtonLeftEffect 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -74520
         Picture         =   "Form1.frx":1E4A7
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   167
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.PictureBox TempPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   7320
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin EffectWorkshop.cpvSlider SliderH 
      Height          =   3390
      Left            =   14640
      Top             =   3720
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   5980
      SliderIcon      =   "Form1.frx":1E85A
      RailPicture     =   "Form1.frx":1E96C
      RailStyle       =   99
   End
   Begin VB.PictureBox PictureMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   5280
      Picture         =   "Form1.frx":1FE8E
      ScaleHeight     =   543
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   607
      TabIndex        =   0
      Top             =   1200
      Width           =   9135
      Begin VB.PictureBox PictureShow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7335
         Left            =   360
         ScaleHeight     =   487
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   551
         TabIndex        =   1
         Top             =   360
         Width           =   8295
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            CausesValidation=   0   'False
            Height          =   615
            Left            =   480
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   720
            Width           =   3855
            ExtentX         =   6800
            ExtentY         =   1085
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   1
            RegisterAsDropTarget=   0
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   5280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Photo"
      Filter          =   "*JPG|*.JPG|*.BMP|*.Bmp|*.GIF|*.GIF"
   End
   Begin EffectWorkshop.cpvSlider SliderV 
      Height          =   240
      Left            =   8400
      Top             =   9720
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   423
      SliderIcon      =   "Form1.frx":25D71
      Orientation     =   0
      RailPicture     =   "Form1.frx":25ECB
      RailStyle       =   99
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9315
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   16431
      _Version        =   393216
      TabOrientation  =   3
      TabHeight       =   529
      ShowFocusRect   =   0   'False
      ForeColor       =   8474663
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Effects"
      TabPicture(0)   =   "Form1.frx":273FD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ImagePrv"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FrameTSpeed"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Transitions"
      TabPicture(1)   =   "Form1.frx":27419
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameT2Speed"
      Tab(1).Control(1)=   "OptionSizeTRansition(1)"
      Tab(1).Control(2)=   "OptionSizeTRansition(0)"
      Tab(1).Control(3)=   "Label5(9)"
      Tab(1).Control(4)=   "Label5(10)"
      Tab(1).Control(5)=   "Label26"
      Tab(1).Control(6)=   "Label5(8)"
      Tab(1).Control(7)=   "ImagePrv2"
      Tab(1).Control(8)=   "ImagePrv1"
      Tab(1).Control(9)=   "Label5(7)"
      Tab(1).Control(10)=   "Label5(6)"
      Tab(1).Control(11)=   "Label1"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "SlideShow"
      TabPicture(2)   =   "Form1.frx":27435
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(0)"
      Tab(2).Control(1)=   "Label2(1)"
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(3)=   "FrameTransitions(2)"
      Tab(2).Control(4)=   "SaverTime"
      Tab(2).Control(5)=   "RandomSaverCheck"
      Tab(2).Control(6)=   "AllSaverCheck"
      Tab(2).Control(7)=   "ControlBoxSaverCheck"
      Tab(2).Control(8)=   "ButtonRight"
      Tab(2).Control(9)=   "ButtonLeft"
      Tab(2).Control(10)=   "FrameTransitions(1)"
      Tab(2).Control(11)=   "FrameTransitions(0)"
      Tab(2).ControlCount=   12
      Begin VB.Frame FrameTransitions 
         Caption         =   "Transitions"
         Height          =   6015
         Index           =   0
         Left            =   -74880
         TabIndex        =   38
         Top             =   120
         Width           =   1815
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Split-out Vertical"
            Height          =   495
            Index           =   14
            Left            =   120
            TabIndex        =   53
            Top             =   5400
            Width           =   1455
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Split-in Vertical"
            Height          =   495
            Index           =   13
            Left            =   120
            TabIndex        =   52
            Top             =   5040
            Width           =   1455
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Disolve"
            Height          =   495
            Index           =   12
            Left            =   120
            TabIndex        =   51
            Top             =   4680
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "ChecherBoard 2"
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   50
            Top             =   4320
            Width           =   1575
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "CheckerBoard 1"
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   49
            Top             =   3960
            Width           =   1455
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Horizontal Blinds"
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   48
            Top             =   3600
            Width           =   1575
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Vertical Blinds"
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   47
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Wipe Left"
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   46
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Wipe Right"
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   45
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Wipe Down"
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   44
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Wipe Up"
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   43
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Circle In"
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Circle Out"
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   41
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Box Out"
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Box In"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame FrameTransitions 
         Caption         =   "Transitions"
         Height          =   6015
         Index           =   1
         Left            =   -74880
         TabIndex        =   14
         Top             =   120
         Width           =   1815
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "BurnFilm"
            Height          =   495
            Index           =   24
            Left            =   120
            TabIndex        =   450
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Split-out Vertical"
            Height          =   495
            Index           =   16
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Strips-down Left"
            Height          =   495
            Index           =   17
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Strips-up Left"
            Height          =   495
            Index           =   18
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Strips-down Right"
            Height          =   495
            Index           =   19
            Left            =   120
            TabIndex        =   24
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Horizontal Bars"
            Height          =   495
            Index           =   21
            Left            =   120
            TabIndex        =   23
            Top             =   2520
            Width           =   1455
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Vertical Bars"
            Height          =   495
            Index           =   22
            Left            =   120
            TabIndex        =   22
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Blend"
            Height          =   495
            Index           =   23
            Left            =   120
            TabIndex        =   21
            Top             =   3240
            Width           =   1095
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Strips-up Right"
            Height          =   495
            Index           =   20
            Left            =   120
            TabIndex        =   20
            Top             =   2160
            Width           =   1455
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "GlassBlock"
            Height          =   495
            Index           =   25
            Left            =   120
            TabIndex        =   19
            Top             =   3960
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Liquid"
            Height          =   495
            Index           =   26
            Left            =   120
            TabIndex        =   18
            Top             =   4320
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Twister"
            Height          =   495
            Index           =   27
            Left            =   120
            TabIndex        =   17
            Top             =   4680
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "CenterCurls"
            Height          =   495
            Index           =   28
            Left            =   120
            TabIndex        =   16
            Top             =   5040
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "PageCurl"
            Height          =   495
            Index           =   29
            Left            =   120
            TabIndex        =   15
            Top             =   5400
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Split-in Horizontal"
            Height          =   495
            Index           =   15
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame FrameT2Speed 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   975
         Left            =   -74520
         TabIndex        =   444
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         Begin EffectWorkshop.cpvSlider Transition2Speed 
            Height          =   240
            Left            =   120
            Top             =   600
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":27451
            Orientation     =   0
            RailPicture     =   "Form1.frx":2796B
            Min             =   1
            Max             =   50
            Value           =   10
         End
         Begin VB.Label lblspeed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Speed"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   445
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Frame FrameTSpeed 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   975
         Left            =   480
         TabIndex        =   379
         Top             =   3480
         Visible         =   0   'False
         Width           =   1215
         Begin EffectWorkshop.cpvSlider TransitionSpeed 
            Height          =   240
            Left            =   120
            Top             =   600
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   423
            SliderIcon      =   "Form1.frx":27987
            Orientation     =   0
            RailPicture     =   "Form1.frx":27EA1
            Min             =   1
            Max             =   50
            Value           =   5
         End
         Begin VB.Label lblspeed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Speed"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   380
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.OptionButton OptionSizeTRansition 
         Caption         =   "Keep Photo 2"
         ForeColor       =   &H00815027&
         Height          =   375
         Index           =   1
         Left            =   -74640
         TabIndex        =   316
         Top             =   4920
         Width           =   1455
      End
      Begin VB.OptionButton OptionSizeTRansition 
         Caption         =   "Keep Photo 1  "
         ForeColor       =   &H00815027&
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   315
         Top             =   4560
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.PictureBox ButtonLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74280
         Picture         =   "Form1.frx":27EBD
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Top             =   6240
         Width           =   255
      End
      Begin VB.PictureBox ButtonRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73920
         Picture         =   "Form1.frx":28270
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   12
         Top             =   6240
         Width           =   255
      End
      Begin VB.CheckBox ControlBoxSaverCheck 
         Caption         =   "Interpolate Pictures"
         Height          =   495
         Left            =   -74760
         TabIndex        =   8
         Top             =   7440
         Width           =   1815
      End
      Begin VB.CheckBox AllSaverCheck 
         Caption         =   "Select All"
         Height          =   495
         Left            =   -74760
         TabIndex        =   7
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CheckBox RandomSaverCheck 
         Caption         =   "Random Transitions"
         Height          =   495
         Left            =   -74760
         TabIndex        =   6
         Top             =   6480
         Width           =   1815
      End
      Begin EffectWorkshop.cpvSlider SaverTime 
         Height          =   240
         Left            =   -74760
         Top             =   8520
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   423
         SliderIcon      =   "Form1.frx":28621
         Orientation     =   0
         RailPicture     =   "Form1.frx":2877B
         RailStyle       =   1
         Min             =   1
         Max             =   20
         Value           =   10
      End
      Begin VB.Frame FrameTransitions 
         Caption         =   "Transitions"
         Height          =   6015
         Index           =   2
         Left            =   -74880
         TabIndex        =   29
         Top             =   120
         Width           =   1815
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Curtains"
            Height          =   495
            Index           =   41
            Left            =   120
            TabIndex        =   449
            Top             =   4320
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "PeelABCD"
            Height          =   495
            Index           =   40
            Left            =   120
            TabIndex        =   448
            Top             =   3960
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Curls"
            Height          =   495
            Index           =   39
            Left            =   120
            TabIndex        =   447
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Ripple"
            Height          =   495
            Index           =   38
            Left            =   120
            TabIndex        =   446
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Water"
            Height          =   495
            Index           =   30
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "RollDown"
            Height          =   495
            Index           =   32
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "LightWipe"
            Height          =   495
            Index           =   31
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Wormhole"
            Height          =   495
            Index           =   33
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Lens"
            Height          =   495
            Index           =   34
            Left            =   120
            TabIndex        =   33
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "FadeWhite"
            Height          =   495
            Index           =   35
            Left            =   120
            TabIndex        =   32
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "FlowMotion"
            Height          =   495
            Index           =   36
            Left            =   120
            TabIndex        =   31
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CheckBox SaverTransitions 
            Caption         =   "Vacuum"
            Height          =   495
            Index           =   37
            Left            =   120
            TabIndex        =   30
            Top             =   2880
            Width           =   1215
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Script"
         ForeColor       =   &H00815027&
         Height          =   195
         Index           =   9
         Left            =   -74400
         TabIndex        =   378
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save Preview"
         ForeColor       =   &H00815027&
         Height          =   195
         Index           =   10
         Left            =   -74400
         TabIndex        =   377
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size Radio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   -74400
         TabIndex        =   317
         Top             =   4200
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Transition"
         ForeColor       =   &H00815027&
         Height          =   195
         Index           =   8
         Left            =   -74400
         TabIndex        =   311
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Image ImagePrv2 
         Height          =   855
         Left            =   -74520
         MouseIcon       =   "Form1.frx":2964D
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         ToolTipText     =   "Reload From Memory"
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Image ImagePrv1 
         Height          =   855
         Left            =   -74520
         MouseIcon       =   "Form1.frx":2979F
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         ToolTipText     =   "Reload From Memory"
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Image ImagePrv 
         Height          =   855
         Left            =   480
         MouseIcon       =   "Form1.frx":298F1
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         ToolTipText     =   "Reload From Memory"
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get Picture 2"
         ForeColor       =   &H00815027&
         Height          =   195
         Index           =   7
         Left            =   -74400
         TabIndex        =   250
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get Picture 1"
         ForeColor       =   &H00815027&
         Height          =   195
         Index           =   6
         Left            =   -74400
         TabIndex        =   249
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00815027&
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   248
         Top             =   8520
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Script"
         ForeColor       =   &H00815027&
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   60
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save Preview"
         ForeColor       =   &H00815027&
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   59
         Top             =   2400
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Info"
         ForeColor       =   &H00815027&
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   58
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00CE7D48&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Welcome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   135
         TabIndex        =   56
         Top             =   240
         Width           =   2070
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Effect"
         ForeColor       =   &H00815027&
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   55
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get Picture"
         ForeColor       =   &H00815027&
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   54
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time between Pictures"
         Height          =   195
         Left            =   -74680
         TabIndex        =   11
         Top             =   8160
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Less"
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   10
         Top             =   8760
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "More"
         Height          =   195
         Index           =   0
         Left            =   -73200
         TabIndex        =   9
         Top             =   8760
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transitions"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   -74400
         TabIndex        =   5
         Top             =   120
         Width           =   765
      End
   End
   Begin TabDlg.SSTab TransitionsTab 
      Height          =   915
      Left            =   2400
      TabIndex        =   251
      Top             =   10200
      Visible         =   0   'False
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   1614
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "2D"
      TabPicture(0)   =   "Form1.frx":29A43
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ButtonLeftEffect(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ButtonRightEffect(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameEffects(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameEffects(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "3D"
      TabPicture(1)   =   "Form1.frx":29A5F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameEffects(6)"
      Tab(1).Control(1)=   "FrameEffects(7)"
      Tab(1).Control(2)=   "ButtonLeftEffect(2)"
      Tab(1).Control(3)=   "ButtonRightEffect(2)"
      Tab(1).Control(4)=   "FrameEffects(5)"
      Tab(1).ControlCount=   5
      Begin VB.Frame FrameEffects 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   5
         Left            =   -73680
         TabIndex        =   285
         Top             =   360
         Width           =   9975
         Begin VB.OptionButton Option2 
            Caption         =   "RollDown"
            Height          =   495
            Index           =   21
            Left            =   8760
            TabIndex        =   294
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "LightWipe"
            Height          =   495
            Index           =   20
            Left            =   7560
            TabIndex        =   293
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Water"
            Height          =   495
            Index           =   19
            Left            =   6600
            TabIndex        =   292
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "PageCurl"
            Height          =   495
            Index           =   18
            Left            =   5400
            TabIndex        =   291
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Twister"
            Height          =   495
            Index           =   17
            Left            =   4440
            TabIndex        =   290
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Liquid"
            Height          =   495
            Index           =   16
            Left            =   3600
            TabIndex        =   289
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "GlassBlock"
            Height          =   495
            Index           =   15
            Left            =   2400
            TabIndex        =   288
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "CenterCurls"
            Height          =   495
            Index           =   14
            Left            =   1200
            TabIndex        =   287
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "BurnFilm"
            Height          =   495
            Index           =   13
            Left            =   240
            TabIndex        =   286
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.Frame FrameEffects 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   3
         Left            =   1680
         TabIndex        =   259
         Top             =   360
         Width           =   9615
         Begin VB.OptionButton Option2 
            Caption         =   "Barn"
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   268
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Blinds"
            Height          =   495
            Index           =   1
            Left            =   720
            TabIndex        =   267
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Checkerboard"
            Height          =   495
            Index           =   2
            Left            =   1560
            TabIndex        =   266
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Fade"
            Height          =   495
            Index           =   3
            Left            =   3000
            TabIndex        =   265
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "GradientWipe"
            Height          =   495
            Index           =   4
            Left            =   3840
            TabIndex        =   264
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Inset"
            Height          =   495
            Index           =   5
            Left            =   5280
            TabIndex        =   263
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Iris"
            Height          =   495
            Index           =   6
            Left            =   6120
            TabIndex        =   262
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "RadialWipe"
            Height          =   495
            Index           =   7
            Left            =   6840
            TabIndex        =   261
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "RandomBars"
            Height          =   495
            Index           =   8
            Left            =   8160
            TabIndex        =   260
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.PictureBox ButtonRightEffect 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   -63120
         Picture         =   "Form1.frx":29A7B
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   284
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox ButtonLeftEffect 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   -74640
         Picture         =   "Form1.frx":29E2C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   283
         Top             =   480
         Width           =   255
      End
      Begin VB.Frame FrameEffects 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   4
         Left            =   1680
         TabIndex        =   269
         Top             =   360
         Visible         =   0   'False
         Width           =   5415
         Begin VB.OptionButton Option2 
            Caption         =   "RandomDissolve"
            Height          =   495
            Index           =   9
            Left            =   0
            TabIndex        =   273
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Spiral"
            Height          =   495
            Index           =   10
            Left            =   1680
            TabIndex        =   272
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Strips"
            Height          =   495
            Index           =   11
            Left            =   2520
            TabIndex        =   271
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Wheel"
            Height          =   495
            Index           =   12
            Left            =   3360
            TabIndex        =   270
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox ButtonRightEffect 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   12000
         Picture         =   "Form1.frx":2A1DF
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   258
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox ButtonLeftEffect 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   480
         Picture         =   "Form1.frx":2A590
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   257
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74520
         Picture         =   "Form1.frx":2A943
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   256
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -63000
         Picture         =   "Form1.frx":2ACF6
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   255
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Lighting "
         Height          =   495
         Index           =   58
         Left            =   -72360
         TabIndex        =   254
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CrShatter"
         Height          =   495
         Index           =   40
         Left            =   -71280
         TabIndex        =   253
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HeightField"
         Height          =   495
         Index           =   39
         Left            =   -70200
         TabIndex        =   252
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame FrameEffects 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   7
         Left            =   -73680
         TabIndex        =   295
         Top             =   360
         Visible         =   0   'False
         Width           =   4815
         Begin VB.OptionButton Option2 
            Caption         =   "ColorFade"
            Height          =   495
            Index           =   30
            Left            =   0
            TabIndex        =   299
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Jaws"
            Height          =   495
            Index           =   31
            Left            =   1320
            TabIndex        =   298
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Grid"
            Height          =   495
            Index           =   32
            Left            =   2400
            TabIndex        =   297
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Threshold"
            Height          =   495
            Index           =   33
            Left            =   3600
            TabIndex        =   296
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame FrameEffects 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   6
         Left            =   -73200
         TabIndex        =   274
         Top             =   360
         Visible         =   0   'False
         Width           =   9015
         Begin VB.OptionButton Option2 
            Caption         =   "Wormhole"
            Height          =   495
            Index           =   22
            Left            =   120
            TabIndex        =   282
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Lens"
            Height          =   495
            Index           =   23
            Left            =   1440
            TabIndex        =   281
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "FlowMotion"
            Height          =   495
            Index           =   24
            Left            =   2400
            TabIndex        =   280
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Vacuum"
            Height          =   495
            Index           =   25
            Left            =   3840
            TabIndex        =   279
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Ripple"
            Height          =   495
            Index           =   26
            Left            =   5040
            TabIndex        =   278
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Curls"
            Height          =   495
            Index           =   27
            Left            =   6120
            TabIndex        =   277
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "PeelABCD"
            Height          =   495
            Index           =   28
            Left            =   6840
            TabIndex        =   276
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Curtains"
            Height          =   495
            Index           =   29
            Left            =   8040
            TabIndex        =   275
            Top             =   0
            Width           =   1095
         End
      End
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "By MArio Flores G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   13680
      TabIndex        =   470
      Top             =   75
      Width           =   1455
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "EFFECT WORKSHOP BETA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   720
      TabIndex        =   469
      Top             =   75
      Width           =   2325
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   375
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '------------------------------------------------------------------------------
 '
 '                            THIS FORM IS THE MAIN FORM
 '                            IN MARIO'S EFFECT WORKSHOP.
 '
 '                          sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------

Dim Flag As Boolean
Dim FlickFlag As Boolean
Dim FTNum0, FTNum1 As Integer
Public MotionBlurDir As Integer

Private Sub AllSaverCheck_Click()

If AllSaverCheck.Value = 1 Then
    RandomSaverCheck.Value = 0
    For i = SaverTransitions.LBound To SaverTransitions.UBound
            SaverTransitions(i).Value = 1
    Next i
End If

If AllSaverCheck.Value = 0 Then
    For i = SaverTransitions.LBound To SaverTransitions.UBound
            SaverTransitions(i).Value = 0
    Next i
End If

End Sub

Private Sub AlphaComboStyle_Change()
AlphaComboStyle_Click
End Sub

Private Sub AlphaComboStyle_Click()
CreateHTML
WebBrowser1.Navigate (App.Path & "\Temp.html")
End Sub

Private Sub BarnMotion_Change()
AlphaComboStyle_Change
End Sub

Private Sub BarnMotion_Click()
AlphaComboStyle_Change
End Sub

Private Sub BlurX_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub BlurX_ValueChanged()
PercentBX.Caption = BlurX.Value * 4 & " %"
End Sub

Private Sub ButtonActivate_Click(Index As Integer)
If TransitionName = "SpotLight" Then MsgBox "Move your Mouse Cursor Over The Image to Activate Effect", vbInformation, "HINT"
If TransitionName = "Zoom" Then MsgBox "Use Mouse Left and Right Button Over Image to Activate Effect", vbInformation, "HINT"
If TransitionName = "Stretch" Then MsgBox "Left Click Over Image and Drag to Activate Effect", vbInformation, "HINT"
If TransitionName = "Lighting" Then MsgBox "Move your Mouse Cursor Over The Image to Activate Effect", vbInformation, "HINT"
CallEffectRefresh
End Sub

Private Sub ButtonLeft_Click()
FTNum0 = FTNum0 - 1

If FTNum0 < 0 Then
   FTNum0 = 0
   Exit Sub
End If

FrameTransitions(0).Visible = False
FrameTransitions(1).Visible = False
FrameTransitions(2).Visible = False
FrameTransitions(FTNum0).Visible = True

End Sub

Private Sub ButtonLeftEffect_Click(Index As Integer)
FTNum1 = FTNum1 - 1

If Index = 0 Then If FTNum1 < 1 Then FTNum1 = 2
If Index = 1 Then If FTNum1 < 3 Then FTNum1 = 4
If Index = 2 Then If FTNum1 < 5 Then FTNum1 = 7

For i = 1 To 7
    FrameEffects(i).Visible = False
Next i

FrameEffects(FTNum1).Visible = True
End Sub

Private Sub ButtonMirror_Click()
CallEffectRefresh
End Sub

Private Sub ButtonRight_Click()
FTNum0 = FTNum0 + 1

If FTNum0 > 2 Then
   FTNum0 = 2
   Exit Sub
End If

FrameTransitions(0).Visible = False
FrameTransitions(1).Visible = False
FrameTransitions(2).Visible = False
FrameTransitions(FTNum0).Visible = True
End Sub

Private Sub ButtonRightEffect_Click(Index As Integer)
FTNum1 = FTNum1 + 1

If Index = 0 Then If FTNum1 > 2 Or FTNum1 < 1 Then FTNum1 = 1
If Index = 1 Then If FTNum1 > 4 Or FTNum1 < 3 Then FTNum1 = 3
If Index = 2 Then If FTNum1 > 7 Or FTNum1 < 5 Then FTNum1 = 5


For i = 1 To 7
    FrameEffects(i).Visible = False
Next i

FrameEffects(FTNum1).Visible = True
End Sub

Private Sub ButtonUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureShow.Top = PictureShow.Top + 25
End Sub

Private Sub CheckWaterMark_Click()
If CheckWaterMark.Value = 1 Then frmFont.Show 1
End Sub

Private Sub CheckWMB_Click()
WaterMarkMark.Visible = Not WaterMarkMark.Visible
End Sub

Private Sub CheckMotionBlur_Click()
FrameMotionBlur.Visible = CheckMotionBlur.Value
FrameBlur1.Visible = Not (FrameMotionBlur.Visible)
End Sub

Private Sub ColorFilterColor_Click()
ColorPicker.Show 1
End Sub

Private Sub ColorFilterColor2_Click()
ColorPicker.BorderColor
ColorPicker.Show
End Sub

Private Sub ComboRotate_Change()
ComboRotate_Click
End Sub

Private Sub ComboRotate_Click()
CallEffectRefresh
End Sub

Public Sub CreateIn()
 
 PictureShow.Width = PictureX
 PictureShow.Height = PictureY
 WebBrowser1.Resizable = True
 
 'Important *******DO NOT ALTERATE*********
 WebBrowser1.Left = -2 '-12
 WebBrowser1.Top = -2 '-17
 WebBrowser1.Width = PictureShow.Width + 18
 WebBrowser1.Height = PictureShow.Height + 18
'********************************************
 
 'Center the Picture**************************
 PictureShow.Move Round(PictureMain.Width / 2) - Round(PictureShow.Width / 2), Round(PictureMain.Height / 2) - Round(PictureShow.Height / 2)
 '********************************************

 'Set & Center The Slider**********************
 SliderH.Min = (-PictureShow.Height)
 SliderH.Max = (PictureMain.Height)
 SliderH.Value = (SliderH.Max + SliderH.Min) / 2
 SliderV.Min = (-PictureShow.Width)
 SliderV.Max = (PictureMain.Width)
 SliderV.Value = (SliderV.Max + SliderV.Min) / 2
 '*********************************************
 
  WebBrowser1.Navigate (App.Path & "\Temp.html")
  
End Sub

Private Sub Command2_Click()
CreateHTML
End Sub

Private Sub Command1_Click()
Dim TempBoolean

If PicsLoaded = False Then
    MsgBox "Please Browse For Photos First!", , ""
    Exit Sub
End If

TempBoolean = 0
For i = SaverTransitions.LBound To SaverTransitions.UBound
 If SaverTransitions(i).Value = 1 Then TempBoolean = 1
Next i

If TempBoolean = 0 Then
    MsgBox "Please Select at Least one Transition First!", , ""
    Exit Sub
End If
WebBrowser1.Stop
Unload FrmSlideShow
FrmSlideShow.Show

End Sub

Private Sub ControlBoxSaverCheck_Click()
MsgBox "Sorry Working on it to Perfectionate!", , ""
End Sub

Private Sub EffectsTab_Click(PreviousTab As Integer)
FTNum1 = 0: ButtonRightEffect_Click (0)
End Sub

Private Sub Form_Load()
 EFFECTSTYLE = "2D"
 Call Intro_Logo
 'Set & Center The Slider**********************
 SliderH.Min = (-PictureShow.Height)
 SliderH.Max = (PictureMain.Height)
 SliderH.Value = (SliderH.Max + SliderH.Min) / 2
 SliderV.Min = (-PictureShow.Width)
 SliderV.Max = (PictureMain.Width)
 SliderV.Value = (SliderV.Max + SliderV.Min) / 2
 '*********************************************
 FTNum0 = 0: FTNum1 = 1
 SSTab1.Tab = 0
 EffectsTab.Tab = 0
 TransitionsTab.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub HeightFieldRotate_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub HeightFieldRotate_ValueChanged()
LabelHeightFieldRotate = HeightFieldRotate.Value & " %"
End Sub

Private Sub HeightFieldSize_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub HeightFieldTime_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub ImageBrowse_Click()
VBA.ChDir App.Path
ShowTree
End Sub

Private Sub ImagePrv_Click()
Call ChargePhoto(PHOTOT0, ImagePrv, True)
End Sub

Private Sub ImagePrv1_Click()
Call ChargePhoto(PHOTOT1, ImagePrv1, True)
End Sub

Private Sub ImagePrv2_Click()
Dim X, xx
X = MsgBox("Swap Photos?", vbQuestion + vbYesNo, "Cant'Be Primary Photo")
If X = vbYes Then
     xx = PHOTOT2: PHOTOT2 = PHOTOT1: PHOTOT1 = xx
     Call ChargePhoto(PHOTOT1, ImagePrv1, True)
     Call ChargePhoto(PHOTOT2, ImagePrv2, True)
End If
End Sub

Private Sub Label5_Click(Index As Integer)
Dim Temp, Temp2

If Index = 1 And Len(PHOTOT0) > 0 Then
    HideTabs
    EffectsTab.Visible = True
End If

If Index = 0 Then Call ChargePhoto(PHOTOT0, ImagePrv)
If Index = 1 And Len(PHOTOT0) = 0 Then MsgBox "Load Picture", vbInformation, "Can't Edit Nothing"
If Index = 2 Then FrmEffectInfo.Show
If Index = 3 And TransitionName <> vbNullString Then FrmPreview.Show
If Index = 3 And TransitionName = vbNullString Then MsgBox "No Effect has Been Activated!", vbInformation, ""
If Index = 4 Or Index = 9 Then FrmScript.Show
If Index = 5 Then About_Logo
If Index = 6 Then Call ChargePhoto(PHOTOT1, ImagePrv1)
If Index = 7 Then Call ChargePhoto(PHOTOT2, ImagePrv2)
If Index = 8 And (Len(PHOTOT1) = 0 Or Len(PHOTOT2) = 0) Then MsgBox "Load 2 Pictures", vbInformation, "Can't Edit"

If Index = 8 And Len(PHOTOT1) > 0 And Len(PHOTOT2) > 0 Then
     HideTabs
     TransitionsTab.Visible = True
     TransitionsTab_Click (0)
End If

If Index = 10 And TransitionName <> vbNullString Then FrmPreview.Show

End Sub
Private Sub HideTabs()
 TransitionsTab.Visible = False
 EffectsTab.Visible = False
 FrameAll.Visible = False
   
   For i = 0 To FrameEffectOption.UBound
     FrameEffectOption(i).Visible = False
     Option1(i).Value = False
   Next
        
    For i = 0 To FrameTransitionOption.UBound
     FrameTransitionOption(i).Visible = False
     Option2(i).Value = False
    Next
   
End Sub

Private Sub Label5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Flag = True Then Exit Sub
    Flag = True
    Label5(Index).FontUnderline = True
    Label5(Index).ForeColor = vbWhite
End Sub

Private Sub MotionBlurDir_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub MotionBlurStr_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub MotionBlurStr_ValueChanged()
LabelBlurStr = MotionBlurStr.Value & " %"
End Sub

Private Sub OpacitySliderE_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub OpacitySliderE_ValueChanged()
PercentOE.Caption = OpacitySliderE.Value & " %"
End Sub

Private Sub OpacitySliderS_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub OpacitySliderS_ValueChanged()
PercentOS.Caption = OpacitySliderS.Value & " %"
End Sub

Private Sub OpacitySliderX1_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub OpacitySliderX1_ValueChanged()
LabelLightX1 = OpacitySliderX1.Value & " %"
End Sub

Private Sub OpacitySliderX2_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub OpacitySliderX2_ValueChanged()
LabelLightX2 = OpacitySliderX2.Value & " %"
End Sub

Private Sub OpacitySliderY1_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub OpacitySliderY1_ValueChanged()
LabelLightY1 = OpacitySliderY1.Value & " %"
End Sub

Private Sub OpacitySliderY2_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub OpacitySliderY2_ValueChanged()
LabelLightY2 = OpacitySliderY2.Value & " %"
End Sub

Private Sub Option1_Click(Index As Integer)
EFFECTSTYLE = "2D"
If Index >= 12 Then EFFECTSTYLE = "3D"
If Index >= 28 Then EFFECTSTYLE = "3DFX"
If Index >= 33 Then EFFECTSTYLE = "SPECIALFX"
If Index = 28 Or Index = 37 Then EFFECTSTYLE = "2D"

If Index = 11 Then
WMPOSX.Max = PictureX
WMPOSY.Max = PictureY
End If

For i = 0 To FrameEffectOption.UBound
    If Index <> i Then FrameEffectOption(i).Visible = False
    If Index <> i Then Option1(i).Value = False
Next
FrameAll.Visible = True: FrameTSpeed.Visible = False
FrameEffectOption(Index).Visible = True: Option1(Index).Value = True
If EFFECTSTYLE = "3D" Then FrameTSpeed.Visible = True
TransitionName = Option1(Index).Caption
 
End Sub

Private Sub Option2_Click(Index As Integer)
EFFECTSTYLE = "TRANSITION"
If Index > 12 Then EFFECTSTYLE = "TRANSITION3D"

For i = 0 To FrameTransitionOption.UBound
    If Index <> i Then FrameTransitionOption(i).Visible = False
    If Index <> i Then Option2(i).Value = False
Next
FrameAll.Visible = True
FrameTransitionOption(Index).Visible = True: Option2(Index).Value = True
FrameT2Speed.Visible = True
TransitionName = Option2(Index).Caption
End Sub

Private Sub OptionBlurDir_Click(Index As Integer)
MotionBlurDir = Index * 45
CallEffectRefresh
End Sub

Private Sub OptionDLight_Click(Index As Integer)
NumDTLight = Index
CallEffectRefresh
End Sub

Private Sub Picturepropertys_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim s
If Button.Index = 2 Then
    Dim TT
    TT = TransitionName
    TransitionName = "Display"
    CreateHTML
    CreateIn
    TransitionName = TT
End If

End Sub

Private Sub RandomSaverCheck_Click()

If RandomSaverCheck.Value = 1 Then
    AllSaverCheck.Value = 0
    For i = SaverTransitions.LBound To SaverTransitions.UBound
            SaverTransitions(i).Value = 0
            SaverTransitions(i).Value = Rnd(2)
    Next i
End If

End Sub

Private Sub RippleRotate_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub RippleSize_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub RippleTime_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub


Private Sub ShatterRotate_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub ShatterRotate_ValueChanged()
LabelShatterrotate = ShatterRotate.Value & " Grades"
End Sub

Private Sub ShatterSize_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub ShatterTime_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SliderGlowFreq_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SliderGlowFreq_ValueChanged()
lblGlowFreq.Caption = SliderGlowFreq.Value * 2 & " %"
End Sub

Private Sub SliderGlowLight_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SliderGlowLight_ValueChanged()
lblGlowLight.Caption = SliderGlowLight.Value & " %"
End Sub

Private Sub SliderGlowPhase_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SliderGlowPhase_ValueChanged()
lblGlowPhase.Caption = SliderGlowPhase.Value & " %"
End Sub

Private Sub SliderGlowStrength_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SliderGlowStrength_ValueChanged()
lblGlowStrength.Caption = SliderGlowStrength.Value & " %"
End Sub

Private Sub SliderH_ValueChanged()
PictureShow.Top = SliderH.Value
End Sub

Private Sub SliderPixelate_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SliderPixelate_ValueChanged()
LabelPixel.Caption = SliderPixelate.Value * 2 & " %"
If LabelPixel.Caption = "4 %" Then LabelPixel.Caption = "0 %"
End Sub

Private Sub SliderV_ValueChanged()
PictureShow.Left = SliderV.Value
End Sub

Private Sub SnowFlake1_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SnowFlake2_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SnowFlake3_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SnowFlake4_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SnowSpeed_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SpotOpacity_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SpotSize_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
TransitionsTab.Visible = False: EffectsTab.Visible = False: FrameAll.Visible = False: FrameTSpeed.Visible = False: FrameT2Speed.Visible = False: SlideTab.Visible = False
If SSTab1.Tab = 2 Then SlideTab.Visible = True
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Flag = False Then Exit Sub
    Flag = False
For i = Label5.LBound To Label5.UBound
    Label5(i).FontUnderline = False
    Label5(i).ForeColor = &H815027
Next
End Sub

Private Sub TxtWML_Change()
CallEffectRefresh
End Sub

Private Sub TxtWMX_Change()
CallEffectRefresh
End Sub

Private Sub ZoomAuto_Click()
CallEffectRefresh
End Sub

Public Sub ChargePhoto(Name As String, Pic As IMAGE, Optional RELOAD As Boolean)
    
    If RELOAD = False Then
        Dialog.FileName = vbNullString
        Dialog.ShowOpen
        If Dialog.FileName <> vbNullString Then Name = Dialog.FileName
    End If
    Temp = TransitionName
    Temp2 = EFFECTSTYLE
    TransitionName = "Display"
    EFFECTSTYLE = "2D"
    If Len(Name) > 0 Then
        WebBrowser1.Stop
        CreateHTML
        CreateIn
    Pic.Picture = LoadPicture(Name)
    TransitionName = Temp
    EFFECTSTYLE = Temp2
    End If
    
End Sub

Private Sub Transition2Speed_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub TransitionSpeed_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub TransitionsTab_Click(PreviousTab As Integer)
FTNum1 = 0
If TransitionsTab.Tab = 0 Then ButtonRightEffect_Click (1)
If TransitionsTab.Tab = 1 Then ButtonRightEffect_Click (2)
End Sub

Private Sub WMOpacity_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub WMPOSX_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub WMPOSY_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub

Private Sub WMSIZE_MouseUp(Shift As Integer)
CallEffectRefresh
End Sub
