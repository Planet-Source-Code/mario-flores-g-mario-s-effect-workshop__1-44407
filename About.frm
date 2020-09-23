VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form FrmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ABOUT !"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8805
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   587
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   7800
      Width           =   1215
   End
   Begin VB.PictureBox PictureShow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   240
      ScaleHeight     =   487
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   551
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         CausesValidation=   0   'False
         Height          =   1095
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   2415
         ExtentX         =   4260
         ExtentY         =   1931
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
         Location        =   ""
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By MArio FLores G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   8040
      Width           =   1590
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect Workshop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   480
      TabIndex        =   4
      Top             =   7800
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sistec_de_juarez@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3840
      TabIndex        =   2
      Top             =   7920
      Width           =   2625
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '------------------------------------------------------------------------------
 '
 '  THIS FORM IS USED TO DISPLAY THE CREDITS IN MARIO'S EFFECT WORKSHOP.
 '                      sistec_de_juarez@ hotmail.com
 '------------------------------------------------------------------------------


Private Sub Command1_Click()
MsgBox "IF YOU LIKED MY APP PLEASE VOTE FOR ME AT PSC THANKS!", , ""
Unload Me
End Sub

Private Sub Form_Activate()
WebBrowser1.Resizable = True
WebBrowser1.Width = Main.PictureMain.Width + 29
WebBrowser1.Left = -12
WebBrowser1.Top = -17
WebBrowser1.Height = Main.PictureMain.Height + 35
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate (App.Path & "\3rd\ini.html")
End Sub

Private Sub Form_Unload(Cancel As Integer)
WebBrowser1.Stop
End Sub
