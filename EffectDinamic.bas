Attribute VB_Name = "EffectDinamic"
 '------------------------------------------------------------------------------
 '
 'THIS MODULE IS USED TO CALL MOST OF THE DINAMIC EFFECTS THAT ARE MANIPULATED IN
 '                           MARIO'S EFFECT WORKSHOP.
 '
 '                          sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------
 

'************************************************************************
'******                        DINAMIC_LIGHT EFFECT               *******
'******                                                           *******
'************************************************************************
Public Function DLight(PhotoDir As String, Width As Integer, Height As Integer, NLights As Integer) As String

IMAGE = "src=" & """" & PhotoDir & """" & " width=" & Width & " height=" & Height & ">"
STYLEX = "<IMG id=myimage Style = " & """" & "FILTER: progid:DXImageTransform.Microsoft.Light(enabled=1); LEFT: 0px; POSITION: absolute; TOP: 0px" & """"

OBJID = STYLEX & vbCrLf & IMAGE

SCRIPTJ = "<SCRIPT language=JavaScript>" & vbCrLf & _
"var g_numlights = " & NLights & ";" & vbCrLf & _
"var Click = 0;" & vbCrLf & _
"window.onload = setlights;" & vbCrLf & _
"document.onclick=keyhandler;" & vbCrLf & _
"myimage.onmousemove=mousehandler;" & vbCrLf & _
"function setlights(){" & vbCrLf & _
    "myimage.filters[0].clear();" & vbCrLf & _
    "myimage.filters[0].addcone(0,0,5,100,100,255,255,0,60,15);" & vbCrLf & _
    "if(g_numlights>0){ myimage.filters[0].addcone(0," & Height & ",5,100,100,255,0,0,60,15);" & vbCrLf & _
    "if(g_numlights>1){myimage.filters[0].addcone(" & Width & "," & Height & ",5,100,100,0,255,255,60,15);" & vbCrLf & _
    "if(g_numlights>2){myimage.filters[0].addcone(" & Width & ",0,5,100,100,255,0,255,60,15);" & vbCrLf & _
            "}}}}"

SCRIPTJ = SCRIPTJ & vbCrLf & "function keyhandler(){if(Click == 0){Click=1;}else Click=0; }" & vbCrLf & _
"function mousehandler(){" & vbCrLf & _
    "if( Click == 0  )  { x = (window.event.x - 80); y = (window.event.y - 80);" & vbCrLf & _
                       "myimage.filters[0].movelight(0,x,y,5,1);" & vbCrLf & _
                       "if( g_numlights > 0 ){ myimage.filters[0].movelight(1,x,y,5,1);" & vbCrLf & _
                       "if( g_numlights > 1 ){ myimage.filters[0].movelight(2,x,y,5,1);" & vbCrLf & _
                       "if( g_numlights > 2 ){ myimage.filters[0].movelight(3,x,y,5,1);" & vbCrLf & _
                                             "}}}}}" & vbCrLf & _
                       "</SCRIPT>"

DLight = IMAGE & vbCrLf & OBJID & vbCrLf & SCRIPTJ & vbCrLf & DisableRightClick
End Function

'************************************************************************
'******                        SPOT_LIGHT EFFECT                  *******
'******                                                           *******
'************************************************************************

Public Function SpotLight(PhotoDir As String, Width As Integer, Height As Integer, Size As Integer, Opacity As Integer) As String

IMAGE = "<IMG id=myimage src=""" & HtmlFormatPath(PhotoDir) & """ width=""" & Width & """ height=""" & Height & """>"

SCRIPTJ = "<BODY leftMargin=0 topMargin=0><STYLE> #myimage {FILTER: light} </STYLE>" & vbCrLf & _
"<SCRIPT> var Click = 0;" & vbCrLf & _
"document.onclick=keyhandler;" & vbCrLf & _
"document.all.myimage.onmousemove = getMouseXY;" & vbCrLf & _
"function getMouseXY(e){ if( Click == 0  ){" & vbCrLf & _
"myimage.filters.light.MoveLight(1,event.offsetX,event.offsetY," & Size & ",true);}}" & vbCrLf & _
"myimage.filters.light.addAmbient(255,255,255," & Opacity & ")" & vbCrLf & _
"myimage.filters.light.addPoint(0,0," & Size & ",255,255,255,255)" & vbCrLf & _
"function keyhandler(){if(Click == 0){Click=1;}else Click=0; }" & vbCrLf & _
"</SCRIPT>"

SpotLight = IMAGE & vbCrLf & SCRIPTJ & vbCrLf & DisableRightClick
End Function

'************************************************************************
'******                        RIPPLE EFFECT                      *******
'******                      HEIGHTFIELD EFFECT                   *******
'******                        SHATTER EFFECT                     *******
'******                                                           *******
'************************************************************************

Public Function Transform3D(PhotoDir As String, Width As Integer, Height As Integer, Name As String, Rotate As Integer, Size As Integer, Time As Integer) As String

OBJID = "<OBJECT ID=""DAControl"" STYLE=""width:" & Width & ";height:" & Height & "; background-color:black; LEFT: 0px; POSITION: absolute; TOP: 0px"" "
CLSID = "CLASSID=""CLSID:B6FFC24C-7E13-11D0-9B47-00C04FC2F51D""></OBJECT>"
IMAGE = "<IMG id=imag1 src=" & """" & PhotoDir & """" & " style=""DISPLAY: none"" width=" & """" & Width & """" & "height=" & """" & Height & """" & ">"
If Name = "Ripple3D" Then Name = "Ripple"
SCRIPTJ = "<SCRIPT LANGUAGE=""JScript"">" & vbCrLf & _
"m = DAControl.PixelLibrary;" & vbCrLf & _
"swGeo = m.ModifiableBehavior(m.EmptyGeometry);" & vbCrLf & _
"swImg = m.ModifiableBehavior(m.EmptyImage);" & vbCrLf & _
"rawImg1 = m.ImportImage(imag1.src);" & vbCrLf & _
"xf =new ActiveXObject(""DX3DTransform.Microsoft." & Name & """);"
If Name = "HeightField" Then SCRIPTJ = SCRIPTJ & "xf.Depth = 3.0;"
SCRIPTJ = SCRIPTJ & vbCrLf & _
"holdTime = m.DANumber(0).Duration(1);" & vbCrLf & _
"forward = m.Interpolate(0, 1," & Time & ");" & vbCrLf & _
"back = m.Interpolate(1, 0, " & Time & ");" & vbCrLf & _
"evaluator = m.Sequence(holdTime, m.Sequence(forward, back)).RepeatForever();" & vbCrLf & _
"result = m.ApplyDXTransform(xf, new Array(rawImg1), evaluator);" & vbCrLf & _
"realGeo = result.OutputBvr;" & vbCrLf & _
"realTransScale = m.Scale3Uniform(" & (Size) / 1000 & ");" & vbCrLf & _
"realTransRotY = m.Rotate3Degrees( m.YVector3 ," & Rotate & ");" & vbCrLf & _
"realTransF = m.Compose3Array(new Array( realTransScale, realTransRotY));" & vbCrLf & _
"realGeo = realGeo.Transform(realTransF);" & vbCrLf & _
"swGeo.SwitchTo(realGeo);" & vbCrLf & _
"camera = m.PerspectiveCamera(0.06, 0.033);" & vbCrLf & _
"light = m.UnionGeometry(m.AmbientLight, m.DirectionalLight.Transform(m.Rotate3Degrees(m.XVector3,-90)));" & vbCrLf & _
"lightAndGeo = m.UnionGeometry(swGeo, light);" & vbCrLf & _
"swImg.SwitchTo(lightAndGeo.render(camera));" & vbCrLf & _
"DAControl.Image = swImg;DAControl.Start();" & vbCrLf & _
"</SCRIPT>"
If Name = "Ripple" Then Name = "Ripple3D"
Transform3D = IMAGE & vbCrLf & OBJID & vbCrLf & CLSID & vbCrLf & SCRIPTJ

End Function

'************************************************************************
'******                        DINAMIC_STRETCH EFFECT             *******
'******                                                           *******
'************************************************************************
Public Function Stretch(PhotoDir As String, Width As Integer, Height As Integer) As String
VBA.FileCopy PhotoDir, App.Path & "\3rd\DemoTemp.jpg"


SCRIPTJ = "<applet code=""anstretch.class"" width=" & """" & Width & """" & "height=" & """" & Height & """" & ">" & vbCrLf & _
"<param name=""credits"" value=""Applet by Fabio Ciucci (www.anfyteam.com)"">" & vbCrLf & _
"<param name=""regnewframe"" value=""YES"">" & vbCrLf & _
"<param name=""image"" value=""DemoTemp.jpg"">" & vbCrLf & _
"<param name=""memdelay"" value=""1000"">" & vbCrLf & _
"<param name=""priority"" value=""3"">" & vbCrLf & _
"</applet>"
Stretch = SCRIPTJ

End Function
'************************************************************************
'******                         DINAMIC_SNOW EFFECT               *******
'******                                                           *******
'************************************************************************
Public Function Snow(PhotoDir As String, Width As Integer, Height As Integer, Speed As Integer, S1 As Integer, S2 As Integer, S3 As Integer, S4 As Integer) As String
VBA.FileCopy PhotoDir, App.Path & "\3rd\DemoTemp.jpg"


SCRIPTJ = "<applet code=""ansnow.class"" width=" & """" & Width & """" & "height=" & """" & Height & """" & ">" & vbCrLf & _
"<param name=""credits"" value=""Applet by Fabio Ciucci (www.anfyteam.com)"">" & vbCrLf & _
"<param name=""regnewframe"" value=""YES"">" & vbCrLf & _
"<param name=flakes1 value=""" & S1 & """>" & vbCrLf & _
"<param name=flakes2 value=""" & S2 & """>" & vbCrLf & _
"<param name=flakes3 value=""" & S3 & """>" & vbCrLf & _
"<param name=flakes4 value=""" & S4 & """>" & vbCrLf & _
"<param name=windmax value=""1"">" & vbCrLf & _
"<param name=windvariation value=""7"">" & vbCrLf & _
"<param name=speed value=""" & Speed & """>" & vbCrLf & _
"<param name=""backimage"" value=""DemoTemp.jpg"">" & vbCrLf & _
"<param name=bgcolor value=""000133"">" & vbCrLf & _
"<param name=overtext value=""NO"">" & vbCrLf & _
"<param name=""memdelay"" value=""1000"">" & vbCrLf & _
"<param name=""priority"" value=""3"">" & vbCrLf & _
"<param name=MinSYNC value=""10"">Sorry, your browser doesn't support This Effect." & vbCrLf & _
"</applet>"
Snow = SCRIPTJ

End Function
'************************************************************************
'******                         DINAMIC_ZOOM EFFECT               *******
'******                                                           *******
'************************************************************************
Public Function Zoom(PhotoDir As String, Width As Integer, Height As Integer) As String
VBA.FileCopy PhotoDir, App.Path & "\3rd\DemoTemp.jpg"

SCRIPTJ = "<applet code=""zoompan.class"" width=" & """" & Width & """" & "height=" & """" & Height & """" & ">" & vbCrLf & _
"<param name=""credits"" value=""Applet by Anfy Team (www.anfyteam.com)"">" & vbCrLf & _
"<param name=""regnewframe"" value=""YES"">" & vbCrLf & _
"<param name=""image"" value=""DemoTemp.jpg"">" & vbCrLf & _
"<param name=""autodesign"" value=""NO"">" & vbCrLf & _
"<param name=""memdelay"" value=""1000"">" & vbCrLf & _
"<param name=""priority"" value=""3"">" & vbCrLf & _
"<param name=""MinSYNC"" value=""10"">" & vbCrLf & _
"<param name=""zoomspeed"" value=""60"">" & vbCrLf & _
"<param name=""xmovespeed"" value=""2"">" & vbCrLf & _
"<param name=""ymovespeed"" value=""2"">" & vbCrLf & _
"<param name=""xborder"" value=""60"">" & vbCrLf & _
"<param name=""yborder"" value=""60"">" & vbCrLf & _
"<param name=""maxzoom"" value=""30"">" & vbCrLf & _
"<param name=""movex"" value=""2"">" & vbCrLf & _
"<param name=""movey"" value=""2"">" & vbCrLf & _
"<param name=""rightclick"" value=""yes"">" & vbCrLf & _
"<param name=""xaccelerate"" value=""40"">" & vbCrLf & _
"<param name=""yaccelerate"" value=""40"">" & vbCrLf & _
"<param name=""stretch"" value=""no"">Sorry, your browser doesn't support This Effect." & vbCrLf & _
"</applet>"
Zoom = SCRIPTJ

End Function
'************************************************************************
'******                         DINAMIC_BLUR EFFECT               *******
'******                                                           *******
'************************************************************************
Public Function DBlur(Style As Integer, Width As Integer, Height As Integer, Texto As TextBox, TSpeed As Integer, Blur As Integer, Direction As String, Stars As Integer, Star As String) As String
Texto = UCase(Texto)
Open App.Path & "\3rd\Temp.txt" For Output As #2: Print #2, Texto: Close #2
Dim Estilo As String
If Style = 0 Then Estilo = "bennyfnt.aft"
If Style = 1 Then Estilo = "1morefnt.aft"
If Style = 2 Then Estilo = "escali_f.aft"
If Style = 3 Then Estilo = "hmd_fontr.aft"

SCRIPTJ = "<applet code=""zoomblur.class"" width=" & """" & Width & """" & " height=" & """" & Height & """" & ">" & vbCrLf & _
"<param name=""credits"" value=""Applet by Fabio Ciucci (www.anfyteam.com)"">" & vbCrLf & _
"<param name=""regnewframe"" value=""YES"">" & vbCrLf & _
"<param name=""font"" value=""" & Estilo & """>" & vbCrLf & _
"<param name=""text"" value=""Temp.txt"">" & vbCrLf & _
"<param name=""bgcolor"" value=""0"">" & vbCrLf & _
"<param name=""blur"" value=""" & Blur & """>" & vbCrLf & _
"<param name=""displaymode"" value=""over"">" & vbCrLf & _
"<param name=""texthspace"" value=""0"">" & vbCrLf & _
"<param name=""delay"" value=""100"">" & vbCrLf & _
"<param name=""delay1"" value=""100"">" & vbCrLf & _
"<param name=""zbmode"" value=""normal"">" & vbCrLf & _
"<param name=""zbcoef"" value=""80"">" & vbCrLf & _
"<param name=""zbmove"" value=""no"">" & vbCrLf & _
"<param name=""dx"" value=""0"">" & vbCrLf & _
"<param name=""dy"" value=""0"">"
SCRIPTJ = SCRIPTJ & vbCrLf & _
"<param name=""txtmove"" value=""" & Direction & """>" & vbCrLf & _
"<param name=""speed"" value=""" & TSpeed & """>" & vbCrLf & _
"<param name=""rx"" value=""20"">" & vbCrLf & _
"<param name=""ry"" value=""30"">" & vbCrLf & _
"<param name=""nbstars"" value=""" & Stars & """>" & vbCrLf & _
"<param name=""star"" value=""" & Star & """>" & vbCrLf & _
"<param name=""stardisplaymode"" value=""over"">" & vbCrLf & _
"<param name=""memdelay"" value=""1000"">" & vbCrLf & _
"<param name=""priority"" value=""3"">" & vbCrLf & _
"<param name=""MinSYNC"" value=""10"">Sorry, your browser doesn't support This Effect." & vbCrLf & _
"</applet>"
DBlur = SCRIPTJ

End Function


