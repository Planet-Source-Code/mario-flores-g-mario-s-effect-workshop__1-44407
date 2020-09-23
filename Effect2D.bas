Attribute VB_Name = "Effect2D"
 '------------------------------------------------------------------------------
 '
 'THIS MODULE IS USED TO CALL THE 2D EFFECTS THAT ARE MANIPULATED IN
 '                           MARIO'S EFFECT WORKSHOP.
 '
 '                          sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------
 
 
 
'************************************************************************
'******                        LIGHT EFFECT                       *******
'******                                                           *******
'************************************************************************
Public Function Light(PhotoDir As String, Width As Integer, Height As Integer, OpacityS As Integer, OpacityF As Integer, Style As Integer, x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.Alpha(style=" & Style & " ,opacity=" & OpacityS & ",finishOpacity=" & OpacityF & ",startX=" & x1 & ",finishX=" & x2 & ",startY=" & y1 & ",finishY=" & y2 & "); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"
Light = STYLEX & vbCrLf & IMAGE & vbCrLf & DisableRightClick

End Function

'************************************************************************
'******                        GRAY EFFECT                        *******
'******                                                           *******
'************************************************************************
Public Function Gray(PhotoDir As String, Width As Integer, Height As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.BasicImage(GrayScale=1); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"
Gray = STYLEX & vbCrLf & IMAGE & vbCrLf & DisableRightClick

End Function
 
'************************************************************************
'******                        INVERT EFFECT                      *******
'******                                                           *******
'************************************************************************
Public Function Invert(PhotoDir As String, Width As Integer, Height As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.BasicImage(Invert=1); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"
Invert = STYLEX & vbCrLf & IMAGE & DisableRightClick

End Function

'************************************************************************
'******                        XRAY EFFECT                        *******
'******                                                           *******
'************************************************************************
Public Function XRay(PhotoDir As String, Width As Integer, Height As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.BasicImage(Xray=1); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"
XRay = STYLEX & vbCrLf & IMAGE & DisableRightClick

End Function

'************************************************************************
'******                        MIRROR EFFECT                      *******
'******                                                           *******
'************************************************************************
Public Function Mirror(PhotoDir As String, Width As Integer, Height As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.BasicImage(Mirror=1); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"
Mirror = STYLEX & vbCrLf & IMAGE & DisableRightClick

End Function

'************************************************************************
'******                        BLUR EFFECT                        *******
'******                                                           *******
'************************************************************************
Public Function Blur(PhotoDir As String, Width As Integer, Height As Integer, BlurX As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.Blur(pixelradius=" & BlurX & "); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"
Blur = STYLEX & vbCrLf & IMAGE & DisableRightClick

End Function

'************************************************************************
'******                        MOTIONBLUR EFFECT                  *******
'******                                                           *******
'************************************************************************
Public Function MotionBlur(PhotoDir As String, Width As Integer, Height As Integer, Direction As Integer, Strength As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.MotionBlur(direction=" & Direction & ",strength=" & Strength & "); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"
MotionBlur = STYLEX & vbCrLf & IMAGE & DisableRightClick

End Function

 
'************************************************************************
'******                        ROTATE EFFECT                      *******
'******                                                           *******
'************************************************************************
Public Function Rotate(PhotoDir As String, Width As Integer, Height As Integer, Angle As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.BasicImage(Rotation=" & Angle & "); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"
Rotate = STYLEX & vbCrLf & IMAGE & DisableRightClick

End Function

'************************************************************************
'******                        EMBOSS EFFECT                      *******
'******                                                           *******
'************************************************************************
Public Function Emboss(PhotoDir As String, Width As Integer, Height As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.Emboss(); LEFT: -3px; POSITION: absolute; TOP: -3px" & """"
Emboss = STYLEX & vbCrLf & IMAGE & DisableRightClick

End Function

'************************************************************************
'******                        ENGRAVE EFFECT                     *******
'******                                                           *******
'************************************************************************
Public Function Engrave(PhotoDir As String, Width As Integer, Height As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.Engrave(); LEFT: -3px; POSITION: absolute; TOP: -3px" & """"
Engrave = STYLEX & vbCrLf & IMAGE & DisableRightClick

End Function

'************************************************************************
'******                        PIXELATE EFFECT                    *******
'******                                                           *******
'************************************************************************
Public Function Pixelate(PhotoDir As String, Width As Integer, Height As Integer, Pixel As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.Pixelate(maxsquare=" & Pixel & "); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"
Pixelate = STYLEX & vbCrLf & IMAGE & DisableRightClick

End Function

'************************************************************************
'******                        WAVE EFFECT                        *******
'******                                                           *******
'************************************************************************
Public Function Wave(PhotoDir As String, Width As Integer, Height As Integer, Freq As Integer, Light As Integer, Phase As Integer, Strength) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.Wave(freq=" & Freq & ",LightStrength=" & Light & ",Phase=" & Phase & ",Strength=" & Strength & " ); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"
Wave = STYLEX & vbCrLf & IMAGE & DisableRightClick

End Function
'************************************************************************
'******                          WATERMARK FILTER                 *******
'******                                                           *******
'************************************************************************
Public Function Watermark(PhotoDir As String, Width As Integer, Height As Integer, Opacity As Integer, Color As String, Color2 As String, Text As String, Left As Integer, Top As Integer, Size As Integer) As String

STYLEX = "<BODY leftMargin=0 topMargin=0><CENTER style=""FILTER: progid:DXImageTransform.Microsoft.alpha(opacity=" & Opacity & "); LEFT:" & Left & "px; WIDTH:" & Width & "px; POSITION: absolute; TOP:" & Top & "px; HEIGHT: 0px"">" & vbCrLf & _
"<P style=""PADDING-RIGHT:0px; FONT-WEIGHT: bold; FONT-SIZE:" & Size & "pt;WIDTH: 100%; COLOR:" & Color & "; PADDING-TOP:0px; BACKGROUND-COLOR:" & Color2 & """>" & Text & "</P></CENTER>"
IMAGE = "<IMG height=" & Height & " width=" & Width & " src=""" & HtmlFormatPath(PhotoDir) & """ TOP: 0px ></BODY>"
Watermark = STYLEX & vbCrLf & IMAGE & DisableRightClick
  
End Function

'************************************************************************
'******                        DISPLAY NORMAL PICTURE             *******
'******                                                           *******
'************************************************************************
Public Function DPHOTO(PhotoDir As String, Width As Integer, Height As Integer) As String

IMAGE = " src=" & """" & PhotoDir & """" & "width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.BasicImage(Mirror=0); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"
DPHOTO = STYLEX & vbCrLf & IMAGE & DisableRightClick

End Function

Public Function DisableRightClick() As String
DisableRightClick = "<SCRIPT language=JavaScript>" & vbCrLf & _
"document.oncontextmenu=new Function(""return false"")" & vbCrLf & _
"</SCRIPT>"
End Function

