VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmPreview 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   5040
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   367
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your Photo is to Big To Fit!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.PictureBox PictureMenu 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
      Begin VB.Image Image4 
         Height          =   240
         Left            =   120
         Picture         =   "FrmPreview.frx":0000
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
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
         Index           =   3
         Left            =   480
         TabIndex        =   31
         Top             =   1320
         Width           =   330
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         FillColor       =   &H00FF0000&
         Height          =   1680
         Left            =   0
         Top             =   0
         Width           =   1320
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   120
         Picture         =   "FrmPreview.frx":058A
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Picture         =   "FrmPreview.frx":0B14
         Top             =   480
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "FrmPreview.frx":109E
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Freeze"
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
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh"
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
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   555
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
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
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   120
         Width           =   450
      End
   End
   Begin VB.PictureBox PictureShow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   5280
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   0
      Top             =   1800
      Width           =   4695
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         CausesValidation=   0   'False
         Height          =   615
         Left            =   480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
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
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   13320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Photo"
      Filter          =   "*.BMP|*.Bmp"
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   23
      Left            =   7920
      TabIndex        =   33
      Top             =   11400
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   22
      Left            =   10440
      TabIndex        =   32
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   21
      Left            =   12840
      TabIndex        =   30
      Top             =   9240
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   20
      Left            =   8760
      TabIndex        =   29
      Top             =   10080
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   19
      Left            =   2400
      TabIndex        =   28
      Top             =   11040
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   18
      Left            =   13440
      TabIndex        =   27
      Top             =   10920
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   17
      Left            =   14040
      TabIndex        =   26
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   16
      Left            =   0
      TabIndex        =   25
      Top             =   9120
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   15
      Left            =   6840
      TabIndex        =   24
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   14
      Left            =   11280
      TabIndex        =   23
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   3840
      TabIndex        =   22
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   1560
      TabIndex        =   21
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   5760
      TabIndex        =   20
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   2520
      TabIndex        =   19
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   8880
      TabIndex        =   18
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   12960
      TabIndex        =   17
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   7560
      TabIndex        =   16
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   15
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   720
      TabIndex        =   14
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   12600
      TabIndex        =   13
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   12
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   9600
      TabIndex        =   11
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect WorkShop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FrmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '------------------------------------------------------------------------------
 '
 'THIS FORM IS USED TO PREVIEW IN REAL SIZE ONE PHOTO AND SAVE IT IN DISK
 '                           MARIO'S EFFECT WORKSHOP.
 '
 '                          sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------

Dim Flag As Boolean

Private Sub Command1_Click()
Picture1.Visible = False
End Sub

Private Sub Form_Activate()
CallEffectRefresh
WebBrowser1.Visible = True
Me.Visible = True
If TransitionName = "Rotate" Then CallEffectRefresh
End Sub

Private Sub Form_Load()
CallEffectRefresh
PictureShow.Width = Main.PictureShow.Width
PictureShow.Height = Main.PictureShow.Height
WebBrowser1.Resizable = True
  
 
 'Important *******DO NOT ALTERATE*********
 WebBrowser1.Left = Main.WebBrowser1.Left
 WebBrowser1.Top = Main.WebBrowser1.Top
 WebBrowser1.Width = Main.WebBrowser1.Width
 WebBrowser1.Height = Main.WebBrowser1.Height
'********************************************
 'Center the Picture**************************
 PictureShow.Move Round((Screen.Width / Screen.TwipsPerPixelX) / 2) - Round(PictureShow.Width / 2), Round((Screen.Height / Screen.TwipsPerPixelY) / 2) - Round(PictureShow.Height / 2)
 Picture1.Move Round((Screen.Width / Screen.TwipsPerPixelX) / 2) - Round(Picture1.Width / 2), Round((Screen.Height / Screen.TwipsPerPixelY) / 2) - Round(Picture1.Height / 2)
 '********************************************
If PictureShow.Width > (Screen.Width / Screen.TwipsPerPixelX) Or PictureShow.Height > (Screen.Height / Screen.TwipsPerPixelY) Then Picture1.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
CallEffectRefresh
End Sub

Private Sub Label1_Click(Index As Integer)
If Index = 0 Then
    PictureMenu.Visible = False
    Dialog.ShowSave
    Me.Refresh
    PictureShow.Picture = CaptureClient(Me)
    If Dialog.FileName <> "" Then SavePicture PictureShow.Picture, Dialog.FileName
    PictureMenu.Visible = True
End If

If Index = 1 Then
    CallEffectRefresh
    WebBrowser1.Visible = True
End If

If Index = 2 Then
    PictureMenu.Visible = False
    Me.Refresh
    PictureShow.Picture = CaptureClient(Me)
    WebBrowser1.Visible = False: WebBrowser1.Stop
    PictureMenu.Visible = True
End If

If Index = 3 Then Unload Me

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Flag = True Then Exit Sub
    Flag = True
    Label1(Index).FontUnderline = True
    Label1(Index).ForeColor = vbWhite
End Sub

Private Sub PictureMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Flag = False Then Exit Sub
    Flag = False
For i = Label1.LBound To Label1.UBound
    Label1(i).FontUnderline = False
    Label1(i).ForeColor = &H815027
Next
End Sub
