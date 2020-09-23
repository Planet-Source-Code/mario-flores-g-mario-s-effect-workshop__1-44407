VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form FrmSlideShow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "SlideShow"
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   91
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictureFrameY 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleMode       =   0  'User
      ScaleWidth      =   4935
      TabIndex        =   4
      Top             =   1095
      Width           =   4935
   End
   Begin VB.PictureBox PictureFramex 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   4665
      ScaleHeight     =   1095
      ScaleMode       =   0  'User
      ScaleWidth      =   270
      TabIndex        =   0
      Top             =   0
      Width           =   270
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   1560
   End
   Begin SHDocVwCtl.WebBrowser SaverScreen 
      CausesValidation=   0   'False
      Height          =   135
      Left            =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3120
      Width           =   135
      ExtentX         =   238
      ExtentY         =   238
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mario's Effect Workshop ScreenSaver"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   2715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "FrmSlideShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SaverPhoto1
Dim SaverPhoto2
Dim Time, LT, UT As Integer

Private Sub Form_Activate()
Time = Main.SaverTime.Value
End Sub


Private Sub Form_Load()
Dim SX, SY As Integer
SX = (Screen.Width / Screen.TwipsPerPixelX)
SY = (Screen.Height / Screen.TwipsPerPixelY)
SaverScreen.Visible = False
SaverScreen.Resizable = True
SaverScreen.FullScreen = True
SaverScreen.Top = 0
SaverScreen.Left = 0
SaverScreen.Width = SX
SaverScreen.Height = SY
LT = Main.SaverTransitions.LBound - 1
Timer1.Enabled = True
PictureFramex.Height = SY
PictureFrameY.Width = SX
PictureFramex.Align = 4
PictureFrameY.Align = 2
End Sub


Private Sub CreateSSaver(Index As Integer)

 On Error Resume Next
 
 Main.TempPicture.Picture = LoadPicture(PHOTO1)
  
 PictureX = Round(Main.TempPicture.Picture.Width / 26.4583)
 PictureY = Round(Main.TempPicture.Picture.Height / 26.4583)

    
 Open App.Path & "\ScreenSaver.html" For Output As #1
     
 'If EFFECTSTYLE = "TRANSITION3D" Then Print #1, Transition3D(PHOTO1, PictureX, PictureY, Main.SaverTime.Value, TransitionName, PHOTOT2)
 If Main.SaverTransitions(Index).Index < 23 Then Print #1, Basics(PHOTO1, PHOTOT2, PictureX, PictureY, Main.SaverTime.Value, Main.SaverTransitions(Index).Index)
  
 Close #1

End Sub



Private Sub Timer1_Timer()
If Timer1.Enabled = False Then Exit Sub
Time = Time + 1

If Time >= Main.SaverTime.Value Then

Check:

LT = LT + 1
If LT > Main.SaverTransitions.UBound Then
       Unload Me
        Exit Sub
End If
If Main.SaverTransitions(LT).Value = 1 Then
   PHOTO1 = "C:\My Documents\Morritas\4.jpg"
   PHOTOT2 = "C:\My Documents\Morritas\t2.jpg"
   CreateSSaver (LT)
   SaverScreen.Stop
   SaverScreen.Navigate (App.Path & "\ScreenSaver.html")
   SaverScreen.Visible = True
   Time = 0
End If

If Main.SaverTransitions(LT).Value = 0 Then GoTo Check

          
End If


End Sub

