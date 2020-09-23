Attribute VB_Name = "Effect3D"
 '------------------------------------------------------------------------------
 '
 'THIS MODULE IS USED TO CALL THE DIFERENT 3D EFFECTS THAT ARE MANIPULATED IN
 '                           MARIO'S EFFECT WORKSHOP.
 '
 '                          sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------
 
 
 
'************************************************************************
'******                          3D EFFECTs                       *******
'******                                                           *******
'************************************************************************

Public Function Transition3D(PhotoDir As String, Width As Integer, Height As Integer, Speed As Integer, Name As String, Optional PhotoDir2 As String) As String
IMAGE = "<IMG id=imag1 src=" & """" & PhotoDir & """" & " style=""DISPLAY: none"" width=" & """" & Width & """" & "height=" & """" & Height & """" & ">"
Image2 = "<IMG id=imag2 src=" & """" & PhotoDir2 & """" & " style=""DISPLAY: none"" width=" & """" & Width & """" & "height=" & """" & Height & """" & ">"

CLSID = "CLASSID=""CLSID:B6FFC24C-7E13-11D0-9B47-00C04FC2F51D""> "
                  
If Name = "BurnFilm" Then LICENSE = "107045D0-06E0-11D2-8D6D-00C04F8EF8E0"
If Name = "Lens" Then LICENSE = "107045CA-06E0-11D2-8D6D-00C04F8EF8E0"
If Name = "ColorFade" Then LICENSE = "2A54C908-07AA-11D2-8D6D-00C04F8EF8E0"
If Name = "GlassBlock" Then LICENSE = "2A54C913-07AA-11D2-8D6D-00C04F8EF8E0"
If Name = "Liquid" Then LICENSE = "AA0D4D0A-06A3-11D2-8F98-00C04FB92EB7"
If Name = "Twister" Then LICENSE = "107045CF-06E0-11D2-8D6D-00C04F8EF8E0"
If Name = "CenterCurls" Then LICENSE = "AA0D4D0C-06A3-11D2-8F98-00C04FB92EB7"
If Name = "PageCurl" Then LICENSE = "AA0D4D08-06A3-11D2-8F98-00C04FB92EB7"
If Name = "Water" Then LICENSE = "107045C5-06E0-11D2-8D6D-00C04F8EF8E0"
If Name = "LightWipe" Then LICENSE = "107045C8-06E0-11D2-8D6D-00C04F8EF8E0"
If Name = "RollDown" Then LICENSE = "9C61F46E-0530-11D2-8F98-00C04FB92EB7"
If Name = "Wormhole" Then LICENSE = "0E6AE022-0C83-11D2-8CD4-00104BC75D9A"
If Name = "FadeWhite" Then LICENSE = "107045CC-06E0-11D2-8D6D-00C04F8EF8E0"
If Name = "Jaws" Then LICENSE = "2A54C904-07AA-11D2-8D6D-00C04F8EF8E0"
If Name = "FlowMotion" Then LICENSE = "2A54C90B-07AA-11D2-8D6D-00C04F8EF8E0"
If Name = "Vacuum" Then LICENSE = "2A54C90D-07AA-11D2-8D6D-00C04F8EF8E0"
If Name = "Grid" Then LICENSE = "2A54C911-07AA-11D2-8D6D-00C04F8EF8E0"
If Name = "Threshold" Then LICENSE = "2A54C915-07AA-11D2-8D6D-00C04F8EF8E0"
If Name = "Ripple" Then LICENSE = "AA0D4D03-06A3-11D2-8F98-00C04FB92EB7"
If Name = "Curls" Then LICENSE = "AA0D4D0E-06A3-11D2-8F98-00C04FB92EB7"
If Name = "PeelABCD" Then LICENSE = "AA0D4D10-06A3-11D2-8F98-00C04FB92EB7"
If Name = "Curtains" Then LICENSE = "AA0D4D12-06A3-11D2-8F98-00C04FB92EB7"

Style = "STYLE = ""Black; width:" & Width & " px; height:" & Height & "px """""""
OBJID = "<OBJECT ID=""DAControl"" " & vbCrLf & Style & vbCrLf & CLSID & vbCrLf & "</OBJECT>"

SCRIPTJ = "<SCRIPT LANGUAGE=""JScript""> " & vbCrLf & _
"<!-- " & vbCrLf & _
       "m = DAControl.PixelLibrary; "

If Len(PhotoDir2) > 0 Then SCRIPTJ = SCRIPTJ & EffectOrTransition(PhotoDir, PhotoDir2, "Transition")
If Len(PhotoDir2) = 0 Then SCRIPTJ = SCRIPTJ & EffectOrTransition(PhotoDir, PhotoDir2, "Effect")

SCRIPTJ = SCRIPTJ & "transf = new ActiveXObject(""DXImageTransform.MetaCreations." & Name & """); " & vbCrLf & _
       "transf.Copyright = ""Copyright MetaCreations Corp. 1998.  Unauthorized duplication of this string is illegal. {" & LICENSE & "}""; " & vbCrLf & _
       "function myInterp() " & vbCrLf & _
           "{" & vbCrLf & _
                "forward = m.Interpolate(0, 1," & Speed * 1.4 & ");" & vbCrLf & _
                "back = m.Interpolate(1, 0," & Speed * 1.4 & " ); " & vbCrLf
If FrmSlideShow.Visible = False Then SCRIPTJ = SCRIPTJ & "return m.Sequence(forward, back).RepeatForever();" & vbCrLf
SCRIPTJ = SCRIPTJ & "}" & vbCrLf & _
        "result = m.ApplyDXTransform(transf, inputImgs, myInterp());" & vbCrLf & _
        "finalImage = result.OutputBvr;" & vbCrLf & _
        "DAControl.Image = finalImage; " & vbCrLf & "DAControl.Start(); " & vbCrLf & _
        "--> " & vbCrLf & _
"</SCRIPT>"


Transition3D = "<BODY bgColor=0><CENTER> " & vbCrLf & IMAGE & Image2 & vbCrLf & OBJID & vbCrLf & SCRIPTVB & vbCrLf & SCRIPTJ & vbCrLf & " </CENTER>"
End Function



Private Function EffectOrTransition(PhotoDir As String, PhotoDir2 As String, Style As String) As String


If Style = "Effect" Then
EffectOrTransition = "img1 = m.ImportImage(imag1.src); " & vbCrLf & _
                     "inputImgs = new Array(img1, img1);"
End If

If Style = "Transition" Then
EffectOrTransition = "img1 = m.ImportImage(imag1.src); " & vbCrLf & _
                     "img2 = m.ImportImage(imag2.src); " & vbCrLf & _
                     "inputImgs = new Array(img1, img2);"
End If


End Function
