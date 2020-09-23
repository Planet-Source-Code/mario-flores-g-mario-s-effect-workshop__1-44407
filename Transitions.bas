Attribute VB_Name = "Transitions"
  '------------------------------------------------------------------------------
 '
 'THIS MODULE IS USED TO CALL THE DIFERENT TRANSITIONS THAT ARE MANIPULATED IN
 '                           MARIO'S EFFECT WORKSHOP.
 '
 '                          sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------
 
 
 
'************************************************************************
'******                          BARN TRANSITION                  *******
'******                                                           *******
'************************************************************************
Public Function Barn(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer, Orientation As String, Motion As String) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.Barn(duration=" & Speed & " , orientation='" & Orientation & "' , motion='" & Motion & "')"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
Barn = SCRIPTJ & DisableRightClick

End Function
'************************************************************************
'******                        BLINDS TRANSITION                  *******
'******                                                           *******
'************************************************************************
Public Function Blinds(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer, Direction As String, Bands As Integer) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.Blinds(duration=" & Speed & " , Bands='" & Bands & "' ,Direction='" & Direction & "')"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
Blinds = SCRIPTJ & DisableRightClick

End Function

'************************************************************************
'******                    CHECKERBOARD TRANSITION                *******
'******                                                           *******
'************************************************************************
Public Function CheckerBoard(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer, Direction As String, SQx As Integer, SQy As Integer) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.Checkerboard(duration=" & Speed & " ,Direction='" & Direction & "' ,SquaresX='" & SQx & "' ,SquaresY='" & SQy & "')"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
CheckerBoard = SCRIPTJ & DisableRightClick

End Function
'************************************************************************
'******                       FADE TRANSITION                     *******
'******                                                           *******
'************************************************************************
Public Function Fade(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.Fade(duration=" & Speed & ")"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
Fade = SCRIPTJ & DisableRightClick

End Function
'************************************************************************
'******                    GRADIENTWIPE TRANSITION                *******
'******                                                           *******
'************************************************************************
Public Function GradientWipe(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer, Size As String, Style As Integer, Motion As String) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.GradientWipe(duration=" & Speed & " , GradientSize='" & Size & "' , wipestyle='" & Style & "' ,motion='" & Motion & "')"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
GradientWipe = SCRIPTJ & DisableRightClick

End Function
'************************************************************************
'******                       INSET TRANSITION                     *******
'******                                                           *******
'************************************************************************
Public Function Inset(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.Inset(duration=" & Speed & ")"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
Inset = SCRIPTJ & DisableRightClick

End Function

'************************************************************************
'******                       IRIS TRANSITION                     *******
'******                                                           *******
'************************************************************************
Public Function Iris(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer, Style As String, Motion As String) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.Iris(duration=" & Speed & " , irisstyle='" & Style & "' , motion='" & Motion & "')"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
Iris = SCRIPTJ & DisableRightClick

End Function

'************************************************************************
'******                    RADIALWIPE TRANSITION                  *******
'******                                                           *******
'************************************************************************
Public Function RadialWipe(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer, Style As String) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.RadialWipe(duration=" & Speed & " , wipestyle='" & Style & "')"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
RadialWipe = SCRIPTJ & DisableRightClick

End Function

'************************************************************************
'******                    RANDOMBARS TRANSITION                  *******
'******                                                           *******
'************************************************************************
Public Function RandomBars(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer, Style As String) As String

SCRIPTJ = Head(PhotoDir2, Style) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.RandomBars(duration=" & Speed & ")"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
RandomBars = SCRIPTJ & DisableRightClick

End Function
'************************************************************************
'******                  RANDOMDISOLVE TRANSITION                 *******
'******                                                           *******
'************************************************************************
Public Function RandomDisolve(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.RandomDissolve(duration=" & Speed & ")"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
RandomDisolve = SCRIPTJ & DisableRightClick

End Function
'************************************************************************
'******                         SPIRAL TRANSITION                 *******
'******                                                           *******
'************************************************************************
Public Function Spiral(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer, X As Integer, Y As Integer) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.Spiral(duration=" & Speed & " GridSizeX='" & X & " , GridSizeY='" & Y & "')"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
Spiral = SCRIPTJ & DisableRightClick

End Function
'************************************************************************
'******                         STRIPS TRANSITION                 *******
'******                                                           *******
'************************************************************************
Public Function Strips(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer, Motion As String) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.Strips(duration=" & Speed & ", motion ='" & Motion & "')"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
Strips = SCRIPTJ & DisableRightClick

End Function
'************************************************************************
'******                         WHEEL TRANSITION                  *******
'******                                                           *******
'************************************************************************
Public Function Wheel(PhotoDir As String, PhotoDir2 As String, Width As Integer, Height As Integer, Speed As Integer, Spokes As Integer) As String

SCRIPTJ = Head(PhotoDir2) & vbCrLf & _
"<IMG id=SampleID  style=""FILTER: progid:DXImageTransform.Microsoft.Wheel(duration=" & Speed & ", spokes ='" & Spokes & "')"";" & vbCrLf & _
"onload = startTrans() src=""" & PhotoDir & """ width=" & Width & " height=" & Height & ">"
Wheel = SCRIPTJ & DisableRightClick

End Function
'************************************************************************
'******                         BASICS TRANSITION                 *******
'******                                                           *******
'************************************************************************
Public Function Basics(PhotoDir As String, Width As Integer, Height As Integer, Speed As Integer, Style As String, Selection As Boolean) As String
Dim X, Y
X = ((Screen.Width / Screen.TwipsPerPixelX) - Width) / 2: Y = ((Screen.Height / Screen.TwipsPerPixelY) - Height) / 2

SCRIPTJ = "<BODY bgColor=0><SCRIPT FOR=window EVENT=onload LANGUAGE=JavaScript>" & vbCrLf & _
"flttgt.filters[0].Apply();"
If Selection = False Then SCRIPTJ = SCRIPTJ & "flttgt.innerHTML =""<img src='" & HtmlFormatPath(PhotoDir) & "'>"";" & vbCrLf
If Selection = True Then SCRIPTJ = SCRIPTJ & "flttgt.innerHTML ="""";" & vbCrLf
SCRIPTJ = SCRIPTJ & "flttgt.filters[0].Play(); </script>" & vbCrLf
If Style <> "23" Then SCRIPTJ = SCRIPTJ & "<div id=""flttgt"" style=""position: relative; width:" & Width & "; height: " & Height & " ; top:" & Y & "; left:" & X & ";background-color:Black;filter:revealTrans(transition=" & Style & ",duration=" & Speed & ")"">" & vbCrLf
If Style = "23" Then SCRIPTJ = SCRIPTJ & "<div id=""flttgt"" style=""position: relative; width:" & Width & "; height: " & Height & " ; top:" & Y & "; left:" & X & ";background-color:Black;filter:blendTrans(duration=" & Speed & ")"">" & vbCrLf
If Selection = True Then SCRIPTJ = SCRIPTJ & "<img src='" & PhotoDir & "'>" & vbCrLf
SCRIPTJ = SCRIPTJ & "</div>"

Basics = SCRIPTJ & vbCrLf & EnableRightClickExit
End Function

Public Function HtmlFormatPath(Path As String) As String
Dim Temp, Temp2
For i = 0 To Len(Path)
    Temp = Right(Left(Path, i), 1): If Temp = "\" Then Temp = "/"
    Temp2 = Temp2 & Temp
Next
HtmlFormatPath = Temp2
End Function

Private Function Head(PhotoDir2 As String, Optional Style As String) As String

PhotoDir2 = HtmlFormatPath(PhotoDir2)
Head = "<SCRIPT language=JavaScript>" & vbCrLf & _
"var fRunning = 0 " & vbCrLf & _
"function startTrans()" & vbCrLf & _
"{ if (fRunning == 0){ fRunning = 1 " & vbCrLf & _
"SampleID.filters[0].Apply();"
If Style <> vbNullString Then Head = Head & "SampleID.filters[0].orientation='" & Style & "'" & vbCrLf
                       Head = Head & "SampleID.src = """ & PhotoDir2 & """;" & vbCrLf & _
"SampleID.filters[0].Play() }} </SCRIPT>"

End Function

Public Function EnableRightClickExit() As String
EnableRightClickExit = "<SCRIPT language=JavaScript>" & vbCrLf & _
"document.oncontextmenu=new Function(""close();return false"")" & vbCrLf & _
"</SCRIPT>"
End Function


