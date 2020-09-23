VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form FrmSlideShow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "SlideShow"
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
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
      Top             =   1080
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
      Left            =   480
      Top             =   2400
   End
   Begin SHDocVwCtl.WebBrowser SaverScreen 
      CausesValidation=   0   'False
      Height          =   1740
      Left            =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3000
      Width           =   2340
      ExtentX         =   4128
      ExtentY         =   3069
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
      Caption         =   "Mario's Effect Workshop SlideShow"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblLoading 
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
 '------------------------------------------------------------------------------
 '
 '             THIS FORM IS USED TO CREATE ONE SLIDE SHOW PRESENTATION
 '                            IN MARIO'S EFFECT WORKSHOP.
 '
 '                          sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------

Dim Z As Variant
Dim Running As Boolean
Dim N, TempN
Dim Flip As Boolean
Dim Time, LT As Integer

Private Sub Form_Activate()
If Me.Visible = True Then Timer1.Enabled = True
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
PictureFramex.Height = SY
PictureFrameY.Width = SX
PictureFramex.Align = 4
PictureFrameY.Align = 2
N = 0
Call GetFiles(Main.FilesX)
Time = Main.SaverTime.Value
End Sub

'******************************************************************************************
'**                                 Sub CreateSlide                                      **
'**     Creates a Web Page and Displays it in Full Screen to get the Slide Appearance    **
'******************************************************************************************

Private Sub CreateSlide(Index As Integer)

 On Error Resume Next
 
 lblLoading.Visible = False
 
 Running = True
 
  Main.TempPicture.Picture = LoadPicture(PHOTO1)
  
 PictureX = Round(Main.TempPicture.Picture.Width / 26.4583)
 PictureY = Round(Main.TempPicture.Picture.Height / 26.4583)

    
 Open App.Path & "\ScreenSaver.html" For Output As #1
     
 
 Flip = Not Flip
 If Main.SaverTransitions(Index).Index < 24 Then Print #1, Basics(PHOTO1, PictureX, PictureY, Main.SaverTime.Value, Main.SaverTransitions(Index).Index, Flip)
 If Main.SaverTransitions(Index).Index > 23 Then Print #1, Transition3D(PHOTO1, PictureX, PictureY, Main.SaverTime.Value, Main.SaverTransitions(Index).Caption)
 Close #1

End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
Running = False
End Sub

Private Sub Timer1_Timer()
If Running = True And SaverScreen.Visible = False Then Unload Me

If Timer1.Enabled = False Then Exit Sub
 Time = Time + 1
 
If N = Pfiles.Count And TempN > 0 And Time > Main.SaverTime.Value Then Unload Me

TempN = 0

For Each Z In Pfiles
    If N = TempN And Time >= Main.SaverTime.Value Then Call DisplayFiles(Z(0))
    TempN = TempN + 1
Next Z

    
End Sub
'*********************************************************************************
'*******      Sub GetFiles (Get Total Files to Display in Slide Show     *********
'*********************************************************************************
Private Sub GetFiles(files As FileListBox)
On Error Resume Next
 
 Set Pfiles = New Collection
    
    For i = 1 To files.ListCount
        files.ListIndex = i - 1
        Pfiles.add Array(SourcePath & "\" & files.FileName)
    Next i

End Sub

'*********************************************************************************
'*******      Sub DisplayFiles (Displays The Files in GetFiles)          *********
'*********************************************************************************
Private Sub DisplayFiles(File)
    
    N = N + 1
Check:

    LT = LT + 1
    If LT > Main.SaverTransitions.UBound Then LT = 0

    If Main.SaverTransitions(LT).Value = 1 Then
        PHOTO1 = File
        Time = 0
        SaverScreen.Visible = False
        CreateSlide (LT)
        SaverScreen.Navigate (App.Path & "\ScreenSaver.html")
        SaverScreen.Visible = True
    End If

    If Main.SaverTransitions(LT).Value = 0 Then GoTo Check
    
End Sub
