VERSION 5.00
Begin VB.Form ColorPicker 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color Picker"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   3840
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   3960
      ItemData        =   "ColorPicker.frx":0000
      Left            =   240
      List            =   "ColorPicker.frx":01A8
      TabIndex        =   5
      Top             =   240
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color"
      Height          =   3135
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1755
         ScaleWidth      =   1395
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "B="
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "G="
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   210
      End
      Begin VB.Label Label1 
         Caption         =   "R="
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label LR 
         AutoSize        =   -1  'True
         Caption         =   "R="
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   210
      End
      Begin VB.Label LG 
         AutoSize        =   -1  'True
         Caption         =   "G="
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   210
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "B="
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   195
      End
   End
End
Attribute VB_Name = "ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
 '
 '                 THIS FORM IS THE COLOR PICKER FORM USED IN WATERMARK EFFECT
 '                                  MARIO'S EFFECT WORKSHOP.
 '
 '                                sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------
Private Style As Boolean
Private R, G, B As Integer

Private Sub Command1_Click()
If Style = False Then
    Main.ColorFilterColor.BackColor = Picture2.BackColor
    Main.LabelColorName = List1.List(List1.ListIndex)
     
End If
If Style = True Then
    Main.ColorFilterColor2.BackColor = Picture2.BackColor
    Main.LabelColorName2 = List1.List(List1.ListIndex)
    Style = False
End If

End Sub

Private Sub Command2_Click()
Unload Me
CallEffectRefresh
Main.Refresh
End Sub

Private Sub Form_Load()
List1.ListIndex = 0
End Sub

Private Sub List1_Click()
ColorHtml List1.ListIndex
End Sub

Private Sub List1_Scroll()
List1_Click
End Sub

Private Sub ColorHtml(Color As Integer)
Me.Caption = List1.List(Color)
If Color = 0 Then LR = 240: LG = 248: LB = 255
If Color = 1 Then LR = 250: LG = 235: LB = 215
If Color = 2 Then LR = 0: LG = 255: LB = 255
If Color = 3 Then LR = 127: LG = 255: LB = 212
If Color = 4 Then LR = 240: LG = 255: LB = 255
If Color = 5 Then LR = 245: LG = 245: LB = 220
If Color = 6 Then LR = 255: LG = 228: LB = 196
If Color = 7 Then LR = 0: LG = 0: LB = 0
If Color = 8 Then LR = 255: LG = 235: LB = 205
If Color = 9 Then LR = 0: LG = 0: LB = 255
If Color = 10 Then LR = 138: LG = 43: LB = 226
If Color = 11 Then LR = 165: LG = 42: LB = 42
If Color = 12 Then LR = 222: LG = 184: LB = 135
If Color = 13 Then LR = 95: LG = 158: LB = 160
If Color = 14 Then LR = 127: LG = 255: LB = 0
If Color = 15 Then LR = 210: LG = 105: LB = 30
If Color = 16 Then LR = 255: LG = 127: LB = 80
If Color = 17 Then LR = 100: LG = 149: LB = 237
If Color = 18 Then LR = 255: LG = 248: LB = 220
If Color = 19 Then LR = 220: LG = 20: LB = 60
If Color = 20 Then LR = 0: LG = 255: LB = 255
If Color = 21 Then LR = 0: LG = 0: LB = 139
If Color = 22 Then LR = 0: LG = 139: LB = 139
If Color = 23 Then LR = 184: LG = 134: LB = 11
If Color = 24 Then LR = 169: LG = 169: LB = 169
If Color = 25 Then LR = 0: LG = 100: LB = 0
If Color = 26 Then LR = 189: LG = 183: LB = 107
If Color = 27 Then LR = 139: LG = 0: LB = 139
If Color = 28 Then LR = 85: LG = 107: LB = 47
If Color = 29 Then LR = 255: LG = 140: LB = 0
If Color = 30 Then LR = 153: LG = 50: LB = 204
If Color = 31 Then LR = 139: LG = 0: LB = 0
If Color = 32 Then LR = 233: LG = 150: LB = 122
If Color = 33 Then LR = 143: LG = 188: LB = 143
If Color = 34 Then LR = 72: LG = 61: LB = 139
If Color = 35 Then LR = 47: LG = 79: LB = 79
If Color = 36 Then LR = 0: LG = 206: LB = 209
If Color = 37 Then LR = 148: LG = 0: LB = 211
If Color = 38 Then LR = 255: LG = 20: LB = 147
If Color = 39 Then LR = 0: LG = 191: LB = 255
If Color = 40 Then LR = 105: LG = 105: LB = 105
If Color = 41 Then LR = 30: LG = 144: LB = 255
If Color = 42 Then LR = 178: LG = 34: LB = 34
If Color = 43 Then LR = 255: LG = 250: LB = 240
If Color = 44 Then LR = 34: LG = 139: LB = 34
If Color = 45 Then LR = 255: LG = 0: LB = 255
If Color = 46 Then LR = 220: LG = 220: LB = 220
If Color = 47 Then LR = 248: LG = 248: LB = 255
If Color = 48 Then LR = 255: LG = 215: LB = 0
If Color = 49 Then LR = 218: LG = 165: LB = 32
If Color = 50 Then LR = 128: LG = 128: LB = 128
If Color = 51 Then LR = 0: LG = 128: LB = 0
If Color = 52 Then LR = 173: LG = 255: LB = 47
If Color = 53 Then LR = 240: LG = 255: LB = 240
If Color = 54 Then LR = 255: LG = 105: LB = 180
If Color = 55 Then LR = 205: LG = 92: LB = 92
If Color = 56 Then LR = 75: LG = 0: LB = 130
If Color = 57 Then LR = 255: LG = 255: LB = 240
If Color = 58 Then LR = 240: LG = 230: LB = 140
If Color = 59 Then LR = 230: LG = 230: LB = 250
If Color = 60 Then LR = 255: LG = 240: LB = 245
If Color = 61 Then LR = 124: LG = 252: LB = 0
If Color = 62 Then LR = 255: LG = 250: LB = 205
If Color = 63 Then LR = 173: LG = 216: LB = 230
If Color = 64 Then LR = 240: LG = 128: LB = 128
If Color = 65 Then LR = 224: LG = 255: LB = 255
If Color = 66 Then LR = 250: LG = 250: LB = 210
If Color = 67 Then LR = 144: LG = 238: LB = 144
If Color = 68 Then LR = 211: LG = 211: LB = 211
If Color = 69 Then LR = 255: LG = 182: LB = 193
If Color = 70 Then LR = 255: LG = 160: LB = 122
If Color = 71 Then LR = 32: LG = 178: LB = 170
If Color = 72 Then LR = 135: LG = 206: LB = 250
If Color = 73 Then LR = 119: LG = 136: LB = 153
If Color = 74 Then LR = 176: LG = 196: LB = 222
If Color = 75 Then LR = 255: LG = 255: LB = 224
If Color = 76 Then LR = 0: LG = 255: LB = 0
If Color = 77 Then LR = 50: LG = 205: LB = 50
If Color = 78 Then LR = 250: LG = 240: LB = 230
If Color = 79 Then LR = 255: LG = 0: LB = 255
If Color = 80 Then LR = 128: LG = 0: LB = 0
If Color = 81 Then LR = 102: LG = 205: LB = 170
If Color = 82 Then LR = 0: LG = 0: LB = 205
If Color = 83 Then LR = 186: LG = 85: LB = 211
If Color = 84 Then LR = 147: LG = 112: LB = 219
If Color = 85 Then LR = 60: LG = 179: LB = 113
If Color = 86 Then LR = 123: LG = 104: LB = 238
If Color = 87 Then LR = 0: LG = 250: LB = 154
If Color = 88 Then LR = 72: LG = 209: LB = 204
If Color = 89 Then LR = 199: LG = 21: LB = 133
If Color = 90 Then LR = 25: LG = 25: LB = 112
If Color = 91 Then LR = 245: LG = 255: LB = 250
If Color = 92 Then LR = 255: LG = 228: LB = 225
If Color = 93 Then LR = 255: LG = 228: LB = 181
If Color = 94 Then LR = 255: LG = 222: LB = 173
If Color = 95 Then LR = 0: LG = 0: LB = 128
If Color = 96 Then LR = 253: LG = 245: LB = 230
If Color = 97 Then LR = 128: LG = 128: LB = 0
If Color = 98 Then LR = 107: LG = 142: LB = 35
If Color = 99 Then LR = 255: LG = 135: LB = 0
If Color = 100 Then LR = 255: LG = 69: LB = 0
If Color = 101 Then LR = 218: LG = 112: LB = 214
If Color = 102 Then LR = 238: LG = 232: LB = 170
If Color = 103 Then LR = 152: LG = 251: LB = 152
If Color = 104 Then LR = 175: LG = 238: LB = 238
If Color = 105 Then LR = 219: LG = 112: LB = 147
If Color = 106 Then LR = 255: LG = 239: LB = 213
If Color = 107 Then LR = 255: LG = 218: LB = 185
If Color = 108 Then LR = 205: LG = 133: LB = 63
If Color = 109 Then LR = 255: LG = 192: LB = 203
If Color = 110 Then LR = 221: LG = 160: LB = 221
If Color = 111 Then LR = 176: LG = 224: LB = 230
If Color = 112 Then LR = 128: LG = 0: LB = 128
If Color = 113 Then LR = 255: LG = 0: LB = 0
If Color = 114 Then LR = 188: LG = 143: LB = 143
If Color = 115 Then LR = 65: LG = 105: LB = 225
If Color = 116 Then LR = 139: LG = 69: LB = 19
If Color = 117 Then LR = 250: LG = 128: LB = 114
If Color = 118 Then LR = 244: LG = 164: LB = 96
If Color = 119 Then LR = 46: LG = 139: LB = 87
If Color = 120 Then LR = 255: LG = 245: LB = 238
If Color = 121 Then LR = 160: LG = 82: LB = 45
If Color = 122 Then LR = 192: LG = 192: LB = 192
If Color = 123 Then LR = 135: LG = 206: LB = 235
If Color = 124 Then LR = 106: LG = 90: LB = 205
If Color = 125 Then LR = 112: LG = 128: LB = 144
If Color = 126 Then LR = 255: LG = 250: LB = 250
If Color = 127 Then LR = 0: LG = 255: LB = 127
If Color = 128 Then LR = 70: LG = 130: LB = 180
If Color = 129 Then LR = 210: LG = 180: LB = 140
If Color = 130 Then LR = 0: LG = 128: LB = 128
If Color = 131 Then LR = 216: LG = 191: LB = 216
If Color = 132 Then LR = 255: LG = 99: LB = 71
If Color = 133 Then LR = 64: LG = 224: LB = 208
If Color = 134 Then LR = 238: LG = 130: LB = 238
If Color = 135 Then LR = 245: LG = 222: LB = 179
If Color = 136 Then LR = 255: LG = 255: LB = 255
If Color = 137 Then LR = 245: LG = 245: LB = 245
If Color = 138 Then LR = 255: LG = 255: LB = 0
If Color = 139 Then LR = 154: LG = 205: LB = 50

Picture2.BackColor = RGB(LR, LG, LB)

End Sub

Public Sub BorderColor()
Style = True
End Sub
