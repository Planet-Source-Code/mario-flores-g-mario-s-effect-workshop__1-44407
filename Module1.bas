Attribute VB_Name = "CALLS"
  '------------------------------------------------------------------------------
 '
 'THIS MODULE IS USED TO CREATE(HTML) THE DIFERENT EFFECTS THAT ARE MANIPULATED IN
 '                           MARIO'S EFFECT WORKSHOP.
 '
 '                          sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------
 
 
 
 Public Sub CreateHTML()
 On Error Resume Next
 If Main.SSTab1.Tab = 0 Then PHOTO1 = PHOTOT0
 If Main.SSTab1.Tab = 1 Then PHOTO1 = PHOTOT1
 
 Main.TempPicture.Picture = LoadPicture(PHOTO1)
  
 PictureX = Round(Main.TempPicture.Picture.Width / 26.4583)
 PictureY = Round(Main.TempPicture.Picture.Height / 26.4583)

 If EFFECTSTYLE = "TRANSITION" And Main.OptionSizeTRansition(1).Value = True Then
 Main.TempPicture.Picture = LoadPicture(PHOTOT2)
 PictureX = Round(Main.TempPicture.Picture.Width / 26.4583)
 PictureY = Round(Main.TempPicture.Picture.Height / 26.4583)
 End If
 
 
 If EFFECTSTYLE = "SPECIALFX" Then GoTo SFX
 
 Open App.Path & "\Temp.html" For Output As #1
 
 If TransitionName = "Light" Then
    If Main.AlphaComboStyle.ListIndex <= 0 Then Main.AlphaComboStyle.ListIndex = 0
    Print #1, Light(PHOTO1, PictureX, PictureY, Main.OpacitySliderS.Value, Main.OpacitySliderE, Main.AlphaComboStyle.ListIndex, Main.OpacitySliderX1.Value, Main.OpacitySliderX2.Value, Main.OpacitySliderY1.Value, Main.OpacitySliderY2.Value)
 End If
 
 If TransitionName = "Blur" And Main.CheckMotionBlur.Value = 0 Then Print #1, Blur(PHOTO1, PictureX, PictureY, Main.BlurX.Value)
      
 If TransitionName = "Blur" And Main.CheckMotionBlur.Value = 1 Then Print #1, MotionBlur(PHOTO1, PictureX, PictureY, Main.MotionBlurDir, Main.MotionBlurStr.Value)
 
 
 If TransitionName = "GrayScale" Then Print #1, Gray(PHOTO1, PictureX, PictureY)
 
      
 If TransitionName = "Invert" Then Print #1, Invert(PHOTO1, PictureX, PictureY)
 
     
 If TransitionName = "Xray" Then Print #1, XRay(PHOTO1, PictureX, PictureY)
 
 
 If TransitionName = "Mirror" Then Print #1, Mirror(PHOTO1, PictureX, PictureY)
 
 
 If TransitionName = "Emboss" Then Print #1, Emboss(PHOTO1, PictureX, PictureY)
 
 
 If TransitionName = "Engrave" Then Print #1, Engrave(PHOTO1, PictureX, PictureY)
 
 
 If TransitionName = "Rotate" Then Print #1, Rotate(PHOTO1, PictureX, PictureY, Main.ComboRotate.ListIndex)
 
 
 If TransitionName = "Pixelate" Then Print #1, Pixelate(PHOTO1, PictureX, PictureY, Main.SliderPixelate.Value)
 
      
 If TransitionName = "Wave" Then Print #1, Wave(PHOTO1, PictureX, PictureY, Main.SliderGlowFreq.Value, Main.SliderGlowLight.Value, Main.SliderGlowPhase.Value, Main.SliderGlowStrength.Value)
 
 
 If TransitionName = "Watermark" Then
    If Main.CheckWMB.Value = 1 Then Print #1, Watermark(PHOTO1, PictureX, PictureY, Main.WMOpacity.Value, Main.LabelColorName.Caption, Main.LabelColorName2.Caption, Main.TextWM, Main.WMPOSX.Value, Main.WMPOSY.Value, Main.WMSIZE.Value)
    If Main.CheckWMB.Value = 0 Then Print #1, Watermark(PHOTO1, PictureX, PictureY, Main.WMOpacity.Value, Main.LabelColorName.Caption, "", Main.TextWM, Main.WMPOSX.Value, Main.WMPOSY.Value, Main.WMSIZE.Value)
 End If
 
 If EFFECTSTYLE = "3D" Then Print #1, Transition3D(PHOTO1, PictureX, PictureY, Main.TransitionSpeed.Value, TransitionName)
 
 If EFFECTSTYLE = "TRANSITION3D" Then Print #1, Transition3D(PHOTO1, PictureX, PictureY, Main.Transition2Speed.Value, TransitionName, PHOTOT2)
     
     
 If TransitionName = "CrShatter" Then Print #1, Transform3D(PHOTO1, PictureX, PictureY, TransitionName, Main.ShatterRotate.Value, Main.ShatterSize.Value, Main.ShatterTime.Value)
 
    
 If TransitionName = "HeightField" Then Print #1, Transform3D(PHOTO1, PictureX, PictureY, TransitionName, Main.HeightFieldRotate.Value, Main.HeightFieldSize.Value, Main.HeightFieldTime.Value)
 
    
 If TransitionName = "Ripple3D" Then Print #1, Transform3D(PHOTO1, PictureX, PictureY, TransitionName, Main.RippleRotate.Value, Main.RippleSize.Value, Main.RippleTime.Value)
 
 
 If TransitionName = "Lighting" Then Print #1, DLight(PHOTO1, PictureX, PictureY, NumDTLight)
 
 
 If TransitionName = "SpotLight" Then Print #1, SpotLight(PHOTO1, PictureX, PictureY, Main.SpotSize.Value, Main.SpotOpacity.Value)
    
 If TransitionName = "Display" Then Print #1, DPHOTO(PHOTO1, PictureX, PictureY)
 
 
 If TransitionName = "Barn" Then Print #1, Barn(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value, Main.BarnOrientation.List(Main.BarnOrientation.ListIndex), Main.BarnMotion.List(Main.BarnMotion.ListIndex))
 
 If TransitionName = "Blinds" Then
 If Main.BlindsBands.ListIndex < 0 Then Main.BlindsBands.ListIndex = 0
 If Main.BlindsDirection.ListIndex < 0 Then Main.BlindsDirection.ListIndex = 0
 Print #1, Blinds(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value, Main.BlindsDirection.List(Main.BlindsDirection.ListIndex), Main.BlindsBands.List(Main.BlindsBands.ListIndex))
 End If
 
 If TransitionName = "Checkerboard" Then
 If Main.ChBoardDir.ListIndex < 0 Then Main.ChBoardDir.ListIndex = 0
 If Main.ChBoardx.ListIndex < 0 Then Main.ChBoardx.ListIndex = 0
 If Main.ChBoardy.ListIndex < 0 Then Main.ChBoardy.ListIndex = 0
 Print #1, CheckerBoard(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value, Main.ChBoardDir.List(Main.ChBoardDir.ListIndex), Main.ChBoardx.List(Main.ChBoardx.ListIndex), Main.ChBoardy.List(Main.ChBoardy.ListIndex))
 End If

 If TransitionName = "Fade" Then Print #1, Fade(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value)
 
 If TransitionName = "GradientWipe" Then
 If Main.GWMotion.ListIndex < 0 Then Main.GWMotion.ListIndex = 0
 If Main.GWSize.ListIndex < 0 Then Main.GWSize.ListIndex = 0
 If Main.GWStyle.ListIndex < 0 Then Main.GWStyle.ListIndex = 0
 Print #1, GradientWipe(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value, Main.GWSize.List(Main.GWSize.ListIndex), Main.GWStyle.ListIndex, Main.GWMotion.List(Main.GWMotion.ListIndex))
 End If
 
 If TransitionName = "Inset" Then Print #1, Inset(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value)
 
 If TransitionName = "Iris" Then
 If Main.IrisMotion.ListIndex < 0 Then Main.IrisMotion.ListIndex = 0
 If Main.IrisStyle.ListIndex < 0 Then Main.IrisStyle.ListIndex = 0
 Print #1, Iris(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value, Main.IrisStyle.List(Main.IrisStyle.ListIndex), Main.IrisMotion.List(Main.IrisMotion.ListIndex))
 End If
 
 If TransitionName = "RadialWipe" Then
  If Main.wipeStyle.ListIndex < 0 Then Main.wipeStyle.ListIndex = 0
 Print #1, RadialWipe(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value, Main.wipeStyle.List(Main.wipeStyle.ListIndex))
 End If
 
 If TransitionName = "RandomBars" Then
  If Main.RBOrientation.ListIndex < 0 Then Main.RBOrientation.ListIndex = 0
 Print #1, RandomBars(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value, Main.RBOrientation.List(Main.RBOrientation.ListIndex))
 End If
 
 If TransitionName = "RandomDissolve" Then Print #1, RandomDisolve(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value)
 
 If TransitionName = "Spiral" Then
 If Main.SpiralX.ListIndex < 0 Then Main.SpiralX.ListIndex = 0
 If Main.SpiralY.ListIndex < 0 Then Main.SpiralY.ListIndex = 0
 Print #1, Spiral(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value, Main.SpiralX.List(Main.SpiralX.ListIndex), Main.SpiralY.List(Main.SpiralY.ListIndex))
 End If
 
 If TransitionName = "Strips" Then
 If Main.StripsMotion.ListIndex < 0 Then Main.StripsMotion.ListIndex = 0
  Print #1, Strips(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value, Main.StripsMotion.List(Main.StripsMotion.ListIndex))
 End If
 
 If TransitionName = "Wheel" Then
 If Main.WheelSpikes.ListIndex < 0 Then Main.WheelSpikes.ListIndex = 0
 Print #1, Wheel(PHOTO1, PHOTOT2, PictureX, PictureY, Main.TransitionSpeed.Value, Main.WheelSpikes.List(Main.WheelSpikes.ListIndex))
 End If
 
 Close #1

Exit Sub

SFX:

Open App.Path & "\3rd\Temp.html" For Output As #1
    
 If TransitionName = "Stretch" Then Print #1, Stretch(PHOTO1, PictureX, PictureY)
 
    
 If TransitionName = "Snow" Then Print #1, Snow(PHOTO1, PictureX, PictureY, Main.SnowSpeed.Value, Main.SnowFlake1.Value, Main.SnowFlake2.Value, Main.SnowFlake3.Value, Main.SnowFlake4.Value)
 
 If TransitionName = "Zoom" Then Print #1, Zoom(PHOTO1, PictureX, PictureY)
 
 
 If TransitionName = "AnimText" Then
 If Main.AniTStyle.ListIndex < 0 Then Main.AniTStyle.ListIndex = 0
 If Main.AniTDirection.ListIndex < 0 Then Main.AniTDirection.ListIndex = 0
 Print #1, DBlur(Main.AniTStyle.ListIndex, PictureX, PictureY, Main.AniTText, Main.AniTSpeed.Value, Main.AniTBlur.Value, Main.AniTDirection.List(Main.AniTDirection.ListIndex), Main.AniTStars.Value, Trim(UCase(Main.AniTSymbol.Text)))
 End If
 
 Close #1
     
End Sub

'-----------------------------ABOUT EFFECT DISPLAY  ------------------
Public Sub About_Logo()
  
FrmAbout.Show 1
 
   
End Sub
'-----------------------------INITIALIZE EFFECT DISPLAY  ------------------
Public Sub Intro_Logo()
   Main.WebBrowser1.Left = -2
   Main.WebBrowser1.Top = -2
   Main.WebBrowser1.Width = Main.PictureShow.Width + 18
   Main.WebBrowser1.Height = Main.PictureShow.Height + 18
   Main.WebBrowser1.Navigate (App.Path & "\ReadMe.html")
End Sub

'--------------------------TO REFRESH PAN DISPLAY  -------------------------
Public Sub CallEffectRefresh()

Dim WL As Integer
Dim WT As Integer
Dim WW As Integer
Dim WH As Integer

Call CreateHTML

    WL = -12
    WT = -17
    WW = 29
    WH = 35


If EFFECTSTYLE = "2D" Or EFFECTSTYLE = "3DFX" Then
    WL = -2
    WT = -2
    WW = 18
    WH = 18
End If

Main.PictureShow.Width = PictureX
Main.PictureShow.Height = PictureY


If TransitionName = "Rotate" Then
    If Main.ComboRotate.ListIndex = 1 Or Main.ComboRotate.ListIndex = 3 Then
      Main.PictureShow.Width = PictureY
      Main.PictureShow.Height = PictureX
    End If
End If

Main.WebBrowser1.Left = WL
Main.WebBrowser1.Top = WT
Main.WebBrowser1.Width = Main.PictureShow.Width + WW
Main.WebBrowser1.Height = Main.PictureShow.Height + WH
    
Main.WebBrowser1.Stop

If FrmPreview.Visible = True Then FrmPreview.WebBrowser1.Stop

If EFFECTSTYLE = "SPECIALFX" Then
If FrmPreview.Visible = True Then
        FrmPreview.WebBrowser1.Navigate (App.Path & "\3rd\Temp.html")
        Exit Sub
    End If
    Main.WebBrowser1.Navigate (App.Path & "\3rd\Temp.html")
    Exit Sub
End If

If FrmPreview.Visible = True Then
    FrmPreview.WebBrowser1.Navigate (App.Path & "\Temp.html")
    Exit Sub
End If

Main.WebBrowser1.Navigate (App.Path & "\Temp.html")


End Sub
