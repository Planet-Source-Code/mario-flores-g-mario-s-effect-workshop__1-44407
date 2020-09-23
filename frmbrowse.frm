VERSION 5.00
Begin VB.Form FrmBrowse 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Browse for Pictures"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
   End
   Begin VB.FileListBox files 
      Height          =   4770
      Left            =   4080
      Pattern         =   "*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf"
      TabIndex        =   1
      Top             =   360
      Width           =   3195
   End
   Begin VB.PictureBox PictureBrowse 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   120
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Browse:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   690
   End
   Begin VB.Label Labelnumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   5640
      TabIndex        =   3
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pictures Found"
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
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "FrmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '------------------------------------------------------------------------------
 '
 'THIS FORM IS USED TO BROWSE FOR PHOTOS IN FOLDER OF USER CHOICE
 '                           MARIO'S EFFECT WORKSHOP.
 '
 '                          sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------
Public Sub PathChange()
On Error Resume Next
files.Path = m_CurrentDirectory
Labelnumber = files.ListCount

If files.ListCount > 0 Then Command1.Enabled = True
If files.ListCount = 0 Then Command1.Enabled = False

End Sub

Private Sub Command1_Click()
PicsLoaded = True
SourcePath = m_CurrentDirectory
Main.FilesX.Path = SourcePath
Unload Me
End Sub

Private Sub Command2_Click()
files.Path = vbNullString
Unload Me
End Sub

Private Sub Form_Load()
m_CurrentDirectory = ""
VBA.ChDir (App.Path)
PathChange
End Sub

Private Sub Form_LostFocus()
CloseUP
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CloseUP
End Sub

Private Sub Form_Unload(Cancel As Integer)
CloseUP
End Sub

