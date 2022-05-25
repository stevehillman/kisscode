' *************************************
'    KISS Music V1.02
'
Const KissMusicV=1.02
' *************************************
Dim Track(10)

Track(1) =  "bgout_song-0x1E.mp3"  
Track(2) =  "bgout_song-0x20.mp3"  
Track(3) =  "bgout_song-0x1A.mp3"  
Track(4) =  "bgout_song-0x22.mp3"  
Track(5) =  "bgout_song-0x24.mp3"  
Track(6) =  "bgout_song-0x28.mp3"  
Track(7) =  "bgout_song-0x2A.mp3"  
Track(8) =  "bgout_song-0x2D.mp3"   

Track(9) =  "bgout_song-0x1C.mp3"  
Track(10)=  "bgout_song-0x26.mp3" 

Dim Title(10)
Title(1)  = "Detroit Rock City"
Title(2)  = "Deuce"
Title(3)  = "Black Diamond"
Title(4)  = "Hotter Than Hell"
Title(5)  = "Lick It Up"
Title(6)  = "Love It Loud"
Title(7)  = "Rock + Roll"
Title(8) = "Shout It Out"

Title(9)  = "Calling Dr. Love"
Title(10) = "Love Gun"

Dim City(16)
City(1) = "Detroit" ' 0x288
City(2) = "Chicago"
City(3) = "Pittsburgh"
City(4) = "Seattle"
City(5) = "Portland"
City(6) = "Los Angeles"
City(7) = "Houston"
City(8) = "New Orleans"
City(9) = "Atlanta"
City(10) = "Orlando"
City(11) = "Tokyo"
City(12) = "London"
City(13) = "New York"
City(14) = "San Francisco"
City(15) = "Mexico City"

Dim slen(100) ' Song Length
slen(01)=1120
slen(02)=4750
slen(03)=680
slen(04)=1360
slen(05)=1120
slen(06)=8575  ' Match Vid
slen(07)=9480
slen(08)=1360 'missing 
slen(09)=76960  ' high score entry
slen(10)=560
slen(11)=1160
slen(12)=10880
slen(13)=4300
slen(14)=3080
slen(15)=4280
slen(16)=1400  ' missing
slen(17)=2950   'missing
slen(18)=440
slen(19)=920
slen(20)=1400
slen(21)=3780
slen(22)=6360
slen(23)=3000
slen(24)=1560
slen(25)=1600
slen(26)=2760
slen(27)=920
slen(28)=520
slen(29)=1640
slen(30)=2360
slen(31)=2800
slen(32)=1520
slen(33)=1560
slen(34)=3700
slen(35)=1320
slen(36)=18600
slen(37)=900
slen(38)=500
slen(39)=920
slen(40)=2460
slen(41)=2480
slen(42)=2600
slen(43)=2500
slen(44)=470
slen(45)=4280
slen(46)=2360
slen(47)=3480
slen(48)=3480
slen(49)=3480
slen(50)=3480
slen(51)=1000 ' 10120
slen(52)=130
slen(53)=5160
slen(54)=240
slen(55)=3360
slen(56)=3320
slen(57)=240
slen(58)=2080
slen(59)=240
slen(60)=9880
slen(61)=1400
slen(62)=2800
slen(63)=4000
slen(64)=6600
slen(65)=1040
slen(66)=1960
slen(67)=7320
slen(68)=4800
slen(69)=2320
slen(70)=2840
slen(71)=4600
slen(72)=1240
slen(73)=1520
slen(74)=3160
slen(75)=00
slen(76)=00
slen(77)=00
slen(78)=00
slen(79)=00
slen(80)=700
slen(81)=680
slen(82)=600
slen(83)=13850
slen(84)=00
slen(85)=00
slen(86)=00
slen(87)=00
slen(88)=00
slen(89)=00
slen(100)=4490




' *************************************
'    KISS Core Functions 
'
Const KissCoreV = 1.02
' *************************************
' 20161031 - changes to reduce UltraDMD DMD flicker
' 20161101 - audio additions
Dim DesktopMode:DesktopMode = Table1.ShowDT

'

' *************************************
'    KISS UltraDMDCode 
'
Const KissDMDV=1.08
' *************************************
'
' 20161031 - use ModifyScene to reduce flicker
' 20161101 - End of game audio
' 20161102 - New Match Gifs - be sure to download them and put in ultradmd directory.
' 20161113 - New UDMD Location code from Seraph74
' 20161120 - Updated UDMD Location code from Seraph74

Dim UltraDMD
Dim BonusLights1, BonusLights2, BonusLights3, BonusLights4

Const UltraDMD_VideoMode_Stretch = 0
Const UltraDMD_VideoMode_Top = 1
Const UltraDMD_VideoMode_Middle = 2
Const UltraDMD_VideoMode_Bottom = 3


Const UltraDMD_Animation_FadeIn = 0
Const UltraDMD_Animation_FadeOut = 1
Const UltraDMD_Animation_ZoomIn = 2
Const UltraDMD_Animation_ZoomOut = 3
Const UltraDMD_Animation_ScrollOffLeft = 4
Const UltraDMD_Animation_ScrollOffRight = 5
Const UltraDMD_Animation_ScrollOnLeft = 6
Const UltraDMD_Animation_ScrollOnRight = 7
Const UltraDMD_Animation_ScrollOffUp = 8
Const UltraDMD_Animation_ScrollOffDown = 9
Const UltraDMD_Animation_ScrollOnUp = 10
Const UltraDMD_Animation_ScrollOnDown = 11
Const UltraDMD_Animation_None = 14

Sub LoadUltraDMD
    ' Set UltraDMD = CreateObject("UltraDMD.DMDObject")
	Dim FlexDMD
    Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
	Set UltraDMD = FlexDMD.NewUltraDMD()
    UltraDMD.Init

    Dim fso, curDir
    Set fso = CreateObject("Scripting.FileSystemObject")
    curDir = fso.GetAbsolutePathName(".")
    Set fso = nothing

    ' A Major version change indicates the version is no longer backward compatible
    If Not UltraDMD.GetMajorVersion = 1 Then
        MsgBox "Incompatible Version of UltraDMD found."
        Exit Sub
    End If

    'A Minor version change indicates new features that are all backward compatible
    If UltraDMD.GetMinorVersion < 3 Then
        MsgBox "Incompatible Version of UltraDMD found.  Please update to version 1.4 or newer."
        Exit Sub
    End If

    UltraDMD.SetProjectFolder curDir & "\KISS.UltraDMD"
    UltraDMD.SetVideoStretchMode UltraDMD_VideoMode_Middle
    UltraDMD.SetScoreboardBackgroundImage "1.png",15,13

    ImgList = "kiss.png,kiss.png,kiss.png,kiss.png,kiss1.png,kiss2.png,kiss3.png,kiss4.png,kiss.png,kiss.png,kiss.png,kiss.png"
    BonusLights1 = UltraDMD.CreateAnimationFromImages(4, false, imgList)
    ImgList = "Army.png,Army.png,Army.png,Army.png,Army1.png,Army2.png,Army3.png,Army4.png,Army.png,Army.png,Army.png,Army.png"
    BonusLights2 = UltraDMD.CreateAnimationFromImages(4, false, imgList)
    ImgList = "faces.png,faces.png,faces.png,faces.png,faces1.png,faces2.png,faces3.png,faces4.png,faces.png,faces.png,faces.png,faces.png"
    BonusLights3 = UltraDMD.CreateAnimationFromImages(4, false, imgList)
    ImgList = "inst.png,inst.png,inst.png,inst.png,inst1.png,inst2.png,inst3.png,inst4.png,inst.png,inst.png,inst.png,inst.png"
    BonusLights4 = UltraDMD.CreateAnimationFromImages(4, false, imgList)

    OnScoreboardChanged()
End Sub

'---------- UltraDMD Unique Table Color preference -------------
' http://www.vpforums.org/index.php?showtopic=26602&page=21#entry362581
'
Dim DMDColor, DMDColorSelect, UseFullColor
Dim DMDPosition, DMDPosX, DMDPosY

Sub GetDMDColor
    Dim WshShell,filecheck,directory
    Set WshShell = CreateObject("WScript.Shell")
    If DMDPosition then
        WshShell.RegWrite "HKCU\Software\UltraDMD\x",DMDPosX,"REG_DWORD"
        WshShell.RegWrite "HKCU\Software\UltraDMD\y",DMDPosY,"REG_DWORD"
    End if
    WshShell.RegWrite "HKCU\Software\UltraDMD\fullcolor",UseFullColor,"REG_SZ"
    WshShell.RegWrite "HKCU\Software\UltraDMD\color",DMDColorSelect,"REG_SZ"
End Sub
'---------------------------------------------------
'---------------------------------------------------

' *****************************************
'    DMD Interactions
' *****************************************
Sub DMDTextI(txt1,txt2,img)  ' Pass in Image
  debug.print "Text: " & txt1 & " " & txt2 
  DMDTextPauseI txt1,txt2,500,img
End Sub

Sub DMDText(txt1,txt2)
  debug.print "Text: " & txt1 & " " & txt2 
  DMDTextPause txt1,txt2,500
End Sub

Sub DMDTextPauseI(txt1,txt2,pause,img)
dim PriorState
  debug.print "Text: " & txt1 & " " & txt2 
  PriorState=UDMDTimer.Enabled
  UDMDTimer.Enabled=False
  if UseUDMD then
    UltraDMD.DisplayScene00Ex img, txt1, 15, 2, txt2, 15, 2, UltraDMD_Animation_None, pause, UltraDMD_Animation_None
  End if
  UDMDTimer.interval=100:UDMDTimer.Enabled=PriorState
End Sub

Sub DMDTextPause(txt1,txt2,pause)
  debug.print "Text: " & txt1 & " " & txt2 
  UDMDTimer.Enabled=False
  if UseUDMD then
    UltraDMD.DisplayScene00Ex "scene01.gif", txt1, 15, 2, txt2, 15, 2, UltraDMD_Animation_None, pause, UltraDMD_Animation_None
    UDMDTimer.interval=100:UDMDTimer.Enabled = True
  End if
End Sub

Sub DMDGif(img1,txt1,txt2,p)
  debug.print "Txt: >" & txt1 & "< Gif:" & img1 & " Pause =" & p
  if NOT UseUDMD then Exit Sub

  UDMDTimer.Enabled=False
  if txt2 = "" then
    UltraDMD.DisplayScene00Ex img1, txt1, 0, 15, "", -1, -1, UltraDMD_Animation_None, p, UltraDMD_Animation_None
  else
    UltraDMD.DisplayScene00Ex img1, txt1, 15, 2, txt2, 13, 2, UltraDMD_Animation_None, p, UltraDMD_Animation_None
  end if
  UDMDTimer.interval=p:UDMDTimer.Enabled = True
End Sub

Sub DisplayI(id)
  debug.print "Showing ID" & id
  if NOT useUDMD Then Exit Sub

  DMDFlush()
  Select Case id
    Case 2: DMDTextI "DANGER","", bgi
    Case 1: DMDTextI "TILT","", bgi
    Case 3: DMDTextI "BASS INSTRUMENT","LIT", bgi
    Case 4: DMDTextI "DRUM INSTRUMENT","LIT", bgi
    Case 5: DMDTextI "KISS COMBO","LIT", bgi                          ' right Lane
    Case 6: DMDTextI "SPELL DEMON TO","RAISE JACKPOT!", bgi
    Case 7: DMDTextI "LOCK","LIT", bgi
'when both green lights hit
    Case 8: DMDTextI "DEUCE","", bgi
' ARMY COMPLETED SPINNER VALUE 24,000 (spinning army target) 4 right targets
    Case 9: DMDTextI "SPINNER VALUE 5,000","ARMY COMPLETED", bgi
' FRONT ROW IS LIT .. when all left kiss targets hit
    Case 10: DMDTextI "FRONT ROW IS","LIT", bgi
    Case 11: DMDTextI "ARMY COMBO IS","LIT", bgi
' Left out lane - shows crows cheer the "FRONT ROW AWARDED | ROCK AGAIN!", dont wait to drain just pop ball and autoplunge
    Case 13: DMDTextI "FRONT ROW","ROCK AGAIN!", bgi
    Case 14: DMDTextI "GUITAR INSTRUMENT","LIT", bgi
    Case 15: DMDTextI "EXTRA BALL","", bgi
    Case 16: DMDGif "scene10.gif","REPLAY","",slen(10)
             ' PlaySound "audio"  'x175

    Case 17: DMDTextI "JACKPOT","", bgi:PlaySound "audio663"
    Case 18: DMDTextI "DEMON","JACKPOT", bgi:PlaySound "audio663"
    Case 19: DMDTextI "DOUBLE","JACKPOT", bgi:PlaySound "audio666"
    Case 20: DMDTextI "SUPER","JACKPOT", bgi
    Case 21: DMDTextI "DOUBLE SUPER","JACKPOT", bgi
    Case 22: DMDTextI "BONUS","2X", bgi:PlaySound "audio681"
    Case 23: DMDTextI "BONUS","3X", bgi:PlaySound "audio682"
    Case 24: DMDTextI "COLOSSAL","BONUS", bgi

    Case 25: DMDGif  "scene11.gif","EXTRA BALL","LIT",slen(11)
    Case 26: DMDGif  "scene11.gif","ROCK AGAIN!","",slen(11)
    Case 27: DMDGif  "scene45.gif","SUPER RAMPS","COMPLETED",slen(45)
    Case 28: DMDGif  "scene45.gif","SUPER BUMPERS","COMPLETED",slen(45)
    Case 29: DMDGif  "scene45.gif","SUPER SPINNER","COMPLETED",slen(45)
    Case 30: DMDGif  "scene45.gif","SUPER TARGETS","COMPLETED",slen(45)
    Case 31: DMDGif  "scene45.gif","SUPER RAMPS", (10-RampCnt(CurPlayer)) & " REMAINING",300
    Case 32: DMDGif  "scene45.gif","SUPER BUMPERS", (50-BumperCnt(CurPlayer)) & " REMAINING",300
    Case 33: DMDGif  "scene45.gif","SUPER SPINNER", (100-SpinCnt(CurPlayer)) & " REMAINING",300
    Case 34: DMDGif  "scene45.gif","SUPER TARGETS", (100-TargetCnt(CurPlayer)) & " REMAINING",300
  End Select
End Sub

Sub UDMDTimer_Timer
    If Not UltraDMD.IsRendering and NOT hsbModeActive Then
        'When the scene finishes rendering, then immediately display the scoreboard
        UDMDTimer.Enabled = False:UDMDTimer.interval=100 
        OnScoreboardChanged()
    End If
End Sub

Sub BV(val)
' only show dmd if timer is active so as not to overwhelm the display
DIM tstr,bstr
    if NOT useUDMD Then Exit Sub
    if bvtimer.enabled=False and svtimer.enabled=False and UltraDMD.IsRendering then Exit Sub  ' Some other animation is going on

    if RND*10 < 5 then
      tstr=string(INT(5*RND)," ") & ((bumpercolor(val)+1)*5) & "K"
    else
      tstr=((bumpercolor(val)+1)*5) & "K" & string(INT(5*RND)," ") 
    end if

    bstr=" "

    UDMDTimer.Enabled=False

    bvtimer.enabled=False
    bvtimer.interval=100
    bvtimer.enabled=True   
    debug.print "BV()"
    if CurScene <> "bv" then
      CurScene="bv"
      debug.print "BV Display scene TEST"
      UltraDMD.CancelRendering:UltraDMD.Clear
      if RND*10 < 5 then
      	  UltraDMD.DisplayScene00ExWithId "bvid", FALSE, "scene18.gif", tstr, 14, INT(RND*4)-1, bstr, -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
      else
    	  UltraDMD.DisplayScene00ExWithId "bvid", FALSE, "scene19.gif", bstr, -1, -1, tstr, 14, INT(RND*4)-1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
      end if
    Else
      debug.print "ModifyScene"
      if NOT UltraDMD.IsRendering then debug.print "bv is no longer rendering??"
      if RND*10 < 5 then
      	  UltraDMD.ModifyScene00 "bvid",  tstr, bstr
      else
    	  UltraDMD.ModifyScene00 "bvid",  bstr, tstr
      end if
    End if
    UDMDTimer.interval=500:UDMDTimer.Enabled = True
End Sub

Sub bvtimer_timer()
  if Not UltraDMD.IsRendering Then
    bvtimer.enabled=False
    debug.print "BV timer expired"
    If CurScene="bv" then 
      CurScene=""
    End If
    OnScoreboardChanged()
    UDMDTimer.interval=100:UDMDTimer.Enabled = True
  End if
End Sub

Sub SpinV(txt)  ' spinner video
Dim bstr,tstr
    if NOT useUDMD Then Exit Sub
    if svtimer.enabled=False and bvtimer.enabled=False and UltraDMD.IsRendering then Exit Sub  ' Some other animation is going on

    if RND*10 < 7 then
      tstr=string(INT(5*RND)," ") & txt
    else
      tstr=txt & string(INT(5*RND)," ") 
    end if

    bstr=" "

    if left(txt,7)="Spinner" then
      tstr="Spinner Count"
      bstr=SpinCnt(CurPlayer)
    End if
    UDMDTimer.Enabled=False

    svtimer.enabled=False
    svtimer.interval=100
    svtimer.enabled=True   
    debug.print "SV() >" & tstr & "<" & "[" & bstr & "]"
    if CurScene <> "sv" then
      UltraDMD.CancelRendering:UltraDMD.Clear
      CurScene="sv"
      debug.print "New Scene"
      if RND*10 < 5 then
    	  UltraDMD.DisplayScene00ExWithId "sv", FALSE, "scene45.gif", tstr, 14, INT(RND*4)-1, bstr, -1, -1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
      else
    	  UltraDMD.DisplayScene00ExWithId "sv", FALSE, "scene45.gif", bstr, -1, -1, tstr, 14, INT(RND*4)-1, UltraDMD_Animation_None, 500, UltraDMD_Animation_None
      end if
    Else
      debug.print "ModifyScene"
      if NOT UltraDMD.IsRendering then debug.print "sv is no longer rendering??"
      if RND*10 < 5 then
      	  UltraDMD.ModifyScene00 "sv",  tstr, bstr
      else
    	  UltraDMD.ModifyScene00 "sv",  bstr, tstr
      end if
    End if
    UDMDTimer.interval=500:UDMDTimer.Enabled = True
End Sub

Sub svtimer_timer()
  if Not UltraDMD.IsRendering Then
    svtimer.enabled=False
    debug.print "SV timer expired"
    If CurScene="sv" then 
      CurScene=""
    End If
    OnScoreboardChanged()
    UDMDTimer.interval=100:UDMDTimer.Enabled = True
  End if
End Sub

Sub OnScoreboardChanged()
  if NOT useUDMD Then Exit Sub
  If UltraDMD.IsRendering Then Exit Sub
  if BVTimer.Enabled=True then Exit Sub ' Lets show the Bumper Animations
  if SVTimer.Enabled=True then Exit Sub ' Lets show the Spinner Animations

  debug.print "OnScoreboardChanged()"
  if UltraDMD.GetMinorVersion > 3 then
    if CurPlayer = 0 or BallsRemaining(CurPlayer)=0 then
      UltraDMD.DisplayScoreboard00 PlayersPlayingGame, 0, Score(1),  Score(2),  Score(3),  Score(4), "credits " & Credits, ""
    else
      UltraDMD.DisplayScoreboard00 PlayersPlayingGame, CurPlayer, Score(1),  Score(2),  Score(3),  Score(4), "credits " & Credits, "ball " & BallsPerGame-BallsRemaining(CurPlayer)+1
    end if
  else
    if CurPlayer = 0 or BallsRemaining(CurPlayer)=0 then
      UltraDMD.DisplayScoreboard PlayersPlayingGame, 0, Score(1),  Score(2),  Score(3),  Score(4), "credits " & Credits, ""
    else
      UltraDMD.DisplayScoreboard PlayersPlayingGame, CurPlayer, Score(1),  Score(2),  Score(3),  Score(4), "credits " & Credits, "ball " & BallsPerGame-BallsRemaining(CurPlayer)+1
    end if
  End if
End Sub

Sub RandomScene()
  RScene(INT(RND*20)+1)
End Sub

Sub RScene(x)
  if NOT useUDMD Then Exit Sub
  If UltraDMD.IsRendering Then Exit Sub
  UDMDTimer.Enabled=False
  debug.print "RandomScene " & x
  Select Case x
    Case 1: DMDGif "scene14.gif","","",slen(14)
    Case 2: DMDGif "scene24.gif","","",slen(24)
    Case 3: DMDGif "scene25.gif","","",slen(25)
    Case 4: DMDGif "scene27.gif","","",slen(27)
    Case 5: DMDGif "scene28.gif","","",slen(28)
    Case 6: DMDGif "scene29.gif","","",slen(29)
    Case 7: DMDGif "scene30.gif","","",slen(30)
    Case 8: DMDGif "scene31.gif","","",slen(31)
    Case 9: DMDGif "scene32.gif","","",slen(32)
    Case 10: DMDGif "scene33.gif","","",slen(33)
    Case 11: DMDGif "scene35.gif","","",slen(35)
    Case 12: DMDGif "scene36.gif","","",slen(36)
    Case 13: DMDGif "scene45.gif","","",slen(45)
    Case 14: DMDGif "scene51.gif","","",slen(51)
    Case 15: DMDGif "scene69.gif","","",slen(69)
    Case 16: DMDGif "scene70.gif","","",slen(70)
    Case 17: DMDGif "scene71.gif","","",slen(71)
    Case 18: DMDGif "scene72.gif","","",slen(72)
    Case 19: DMDGif "scene73.gif","","",slen(73)
    Case 20: DMDGif "scene74.gif","","",slen(74)
  End Select
End Sub



Dim MatchFN,MatchFlag
Sub CheckMatch()
   Dim X, XX, Y, Z, ZZ, abortLoop, tempSort, tmpScore
   Dim Match, divider	
   Dim fname

   Debug.print "CheckMatch()"
   UDMDTimer.Enabled=False
   match = INT(RND*10)
   matchFlag = 0
   for x=1 To (Players)	
        tmpScore=scores(x)							'Break player's scores down into 2 digit numbers for match
	divider = 1000000000							'Divider starts at 1 billion
  	for xx=0 To 7								'Seven places will get us the last 2 digits of a 10 digit score		
		if (tmpScore >= divider) Then
			tmpScore = tmpScore MOD divider
		End If
		divider = Divider / 10						
	Next
	if (tmpScore = (Match * 10)) Then	'Did we match?	
		matchFlag  = matchFlag  +  1						'Count it up!
	End If
   Next
   MatchFN = CStr(Match) & "0"
  ' fname="match" & MatchFN & ".gif"
   debug.print fname
   if UseUDMD then
     UltraDMD.CancelRendering:UltraDMD.Clear
     DMDGif "scene06.gif","","",12000
   end if
   UDMDTimer.Enabled=False
   MatchTimer.Interval=12000
   MatchTimer.Enabled=True
End Sub

Sub MatchTimer_Timer
    debug.print "MatchTimer()"
    MatchTimer.Enabled=False
    Debug.print  ">" & MatchFN & "<"
   if UseUDMD then
     UltraDMD.CancelRendering:UltraDMD.Clear
     DMDGif "black.bmp","MATCH",MatchFN,3000
   end if
    if (matchFlag) Then										'Does one of the player's scores match?
        PlaySound "audio746",0,4  ' Match      
        PlaySound SoundFXDOF("fx_kicker",141,DOFPulse,DOFContactors), 0, 1, 0.1, 0.1
	credits  = credits  +  matchFlag								'Award a credit for each match!
    End If
    ' set the machine into game over mode
   MatchTimer2.Interval=3000
   MatchTimer2.Enabled=True
End Sub

Sub MatchTimer2_Timer
    MatchTimer2.Enabled=False
    debug.print "MatchTimer2"
    UDMDTimer.interval=200:UDMDTimer.Enabled=True 
    If INT(RND*10) > 5 then
       PlaySound "audio415"
    Else  
      If Int(RND*10) > 5 Then  
        PlaySound "audio416"
      Else 
        if Int(RND*10) > 5 Then
          PlaySound "audio427"
        End If
      End If
    End If 
    EndOfGame()
End Sub


' *************************************
'    KISS Code 
'
Const KissCodeV=1.02
' *************************************
' Code to handle the various modes
' *************************************
' 20161031 track shots made by player by song as some songs require more/less shots to complete

' LO-sw14/58, sc-68, MR-sw64/73, D-sw42/79, RR-sw56/92, ro-sw24/98


Dim Curcolor, Curcolorfull

Sub SaveRGB(ll)
  CurColor=ll.color
  Curcolorfull=ll.colorfull
End Sub

Sub RestoreRGB(ll)
  ll.color=Curcolor
  ll.colorfull=Curcolorfull
End Sub

Sub ProcessLO
  Debug.print "ProcessLO Cursong is " & CurSong(CurPlayer)
  SaveRGB(i58):i58.state=LightStateOff
  if DemonMBMode then
    if LastShot(CurPlayer) = -1 then
      DisplayI(18)
      AddScore(1000000)
    else
     DisplayI(19)
      AddScore(2000000)
    end if
    LastShot(CurPlayer)=1
    Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
    exit sub
  else
    if LoveGunMode then
      if LastShot(CurPlayer) = -1 then
        DisplayI(17)
        AddScore(1000000)
      else
        DisplayI(19)
        AddScore(2000000)
      end if
      LastShot(CurPlayer)=1
      Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
      exit sub
    end if
  end if
  LastShot(CurPlayer)=1
  Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1

  Select Case CurSong(CurPlayer)
    case 5: 
      ShowShot(Shots(CurPlayer,CurSong(CurPlayer))*200000)
      if i58.state=LightStateOff and i68.state=LightStateOff and i73.state=LightStateOff and i79.state=LightStateOff and i92.state=LightStateOff and i98.state=LightStateOff  then ' reset lights
        InitMode(CurSong(CurPlayer))
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 10 = 0 then SongComplete()
    case 2:
      ShowShot(100000)
      if i68.state=LightStateBlinking then
        RestoreRGB(i68):i68.state=LightStateBlinking:RestoreRGB(i73):i73.state=LightStateBlinking:
        SetLightColor i58, "white", 0
      else
        RestoreRGB(i58):i58.state=LightStateBlinking:RestoreRGB(i68):i68.state=LightStateBlinking
        SetLightColor i98, "white", 0
      end if
        DisplayI(8) ' Deuce
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 3:
      ShowShot(100000):i98.state=LightStateBlinking:RestoreRGB(i98)
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 4:
      if Shots(CurPlayer,CurSong(CurPlayer)) > 17 then
        ShowShot(2000000)
      else
        ShowShot(250000+(100000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if sw38.isdropped=False then i35.state=LightStateBlinking end if
      if sw39.isdropped=False then i36.state=LightStateBlinking end if
      if sw40.isdropped=False then i37.state=LightStateBlinking end if
      if sw41.isdropped=False then i38.state=LightStateBlinking end if
      if Shots(CurPlayer,CurSong(CurPlayer)) > 10 then SongComplete()
    case 1:
      if Shots(CurPlayer,CurSong(CurPlayer)) > 20 then
        ShowShot(2000000)
      else
        ShowShot(50000+ (50000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if NOT i68.state=LightStateOff then  ' Move Left
        i98.state=LightStateBlinking:RestoreRGB(i98):i58.state=LightStateBlinking:RestoreRGB(i58)
        SetLightColor i68, "white", 0
      else
        SetLightColor i98, "white", 0
        RestoreRGB(i58):i58.state=LightStateBlinking:RestoreRGB(i68):i68.state=LightStateBlinking
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 6:
      ShowShot(300000)
      if i58.state=LightStateOff and i98.state=LightStateoff then ' Both Orbits complete
        RestoreRGB(i68):RestoreRGB(i73):RestoreRGB(i79):RestoreRGB(i92)
        i68.state=LightStateBlinking:i73.state=LightStateBlinking:i79.state=LightStateBlinking:i92.state=LightStateBlinking
      end if
    case 7:
      ShowShot(100000):RestoreRGB(i73):i73.state=LightStateBlinking
    case 8:
       ShowShot(100000)
       if i58.state=LightStateOff and i98.state=LightStateOff then 
          RestoreRGB(i68):i68.state=LightStateBlinking
          RestoreRGB(i79):i79.state=LightStateBlinking
          RestoreRGB(i92):i92.state=LightStateBlinking
       end if
    case 9:
    case 10:
  End Select
End Sub

Sub ProcessSC
  debug.print "ProcessSC"
  SaveRGB(i68)
  SetLightColor i68, "white", 0
  LastShot(CurPlayer)=2
  Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
  debug.print "Shots=" & Shots(CurPlayer,CurSong(CurPlayer))
  Select Case CurSong(CurPlayer)
    case 5: 
      ShowShot(Shots(CurPlayer,CurSong(CurPlayer))*200000)
      if i58.state=LightStateOff and i68.state=LightStateOff and i73.state=LightStateOff and i79.state=LightStateOff and i92.state=LightStateOff and i98.state=LightStateOff  then ' reset lights
        InitMode(CurSong(CurPlayer))
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) = 10 then SongComplete()
    case 2:
      ShowShot(100000)
      if i73.state=LightStateBlinking then
        RestoreRGB(i73):i73.state=LightStateBlinking:RestoreRGB(i79):i79.state=LightStateBlinking:SetLightColor i68, "white", 0
      else
        RestoreRGB(i68):i68.state=LightStateBlinking:RestoreRGB(i73):i73.state=LightStateBlinking:SetLightColor i58, "white", 0
      end if
   '   DisplayI(8) ' Deuce
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 3:
      ShowShot(100000):i98.state=LightStateBlinking:RestoreRGB(i98)
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 4:
      if Shots(CurPlayer,CurSong(CurPlayer)) > 17 then
        ShowShot(2000000)
      else
        ShowShot(250000+(100000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if sw38.isdropped=False then i35.state=LightStateBlinking end if
      if sw39.isdropped=False then i36.state=LightStateBlinking end if
      if sw40.isdropped=False then i37.state=LightStateBlinking end if
      if sw41.isdropped=False then i38.state=LightStateBlinking end if
      if Shots(CurPlayer,CurSong(CurPlayer)) > 10 then SongComplete()
    case 1:
      if Shots(CurPlayer,CurSong(CurPlayer)) > 20 then
        ShowShot(2000000)
      else
        ShowShot(50000+ (50000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if NOT i73.state=LightStateOff then  ' Move Left
        RestoreRGB(i58):i58.state=LightStateBlinking:RestoreRGB(i68):i68.state=LightStateBlinking:SetLightColor i73, "white", 0
      else
        SetLightColor i58, "white", 0:RestoreRGB(i68):i68.state=LightStateBlinking:RestoreRGB(i73):i73.state=LightStateBlinking
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 6:
      ShowShot(200000*(Shots(CurPlayer,CurSong(CurPlayer))-1))
      if Shots(CurPlayer,CurSong(CurPlayer)) > 10 then
        SongComplete()
        SetLightColor i68, "white", 0
        SetLightColor i73, "white", 0
        SetLightColor i79, "white", 0
        SetLightColor i92, "white", 0
      else
        RestoreRGB(i73):RestoreRGB(i79):RestoreRGB(i92)
        SetLightColor i68, "white", 0
        i73.state=LightStateBlinking:i79.state=LightStateBlinking:i92.state=LightStateBlinking
      end if
    case 7:
      ShowShot(100000):RestoreRGB(i73):i73.state=LightStateBlinking
    case 8:
      ShowShot(100000)
      if i68.state=LightStateOff and i79.state=LightStateOff and i92.state=LightStateOff then
         RestoreRGB(i73):i73.state=LightStateBlinking
         SongComplete()
      End If
    case 9:
    case 10:
  End Select
End Sub

Sub RandomBD ' When Bumper is hit & Black Diamond then target moves if its not the RO
  if i98.state=LightStateOff and CurSong(CurPlayer) = 3 then
    debug.print "RandomBD"
    SetLightColor i58,"white", 0
    SetLightColor i68,"white", 0
    SetLightColor i73,"white", 0
    SetLightColor i79,"white", 0
    SetLightColor i92,"white", 0
      Select Case INT(RND*5)+1
        case 1: SetLightColor i58,RGBColors(cRGB), 2
        case 2: SetLightColor i68,RGBColors(cRGB), 2
        case 3: SetLightColor i73,RGBColors(cRGB), 2
        case 4: SetLightColor i79,RGBColors(cRGB), 2
        case 5: SetLightColor i92,RGBColors(cRGB), 2
      End Select
  End If
End Sub

Sub ProcessTarget
   LastShot(CurPlayer)=7 
   ' If Random Shot not already lit then light one
   if i58.state=LightStateOff and i68.state=LightStateOff and i73.state=LightStateOff and i79.state=LightStateOff and i92.state=LightStateOff and i98.state=LightStateOff then
     Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
     if Shots(CurPlayer,CurSong(CurPlayer)) > 5 then
       ShowShot(2000000)
     else
       ShowShot(250000+(100000*Shots(CurPlayer,CurSong(CurPlayer))))
     end if
    ' light random shot
     debug.print "Light Random shot "
   
     Select Case INT(RND*6)+1
      case 1: SetLightColor i58, "red", 2
      case 2: SetLightColor i68, "red", 2
      case 3: SetLightColor i73, "red", 2
      case 4: SetLightColor i79, "red", 2
      case 5: SetLightColor i92, "red", 2
      case 6: SetLightColor i98, "red", 2
     End Select
  End If
  if Shots(CurPlayer,CurSong(CurPlayer)) > 10 then SongComplete()
End Sub

Sub ProcessMR
  SaveRGB(i73):SetLightColor i73, "white", 0
  if DemonMBMode then
    if LastShot(CurPlayer) = -1 then
      DisplayI(18)
      AddScore(1000000)
    else
      DisplayI(19)
      AddScore(2000000)
    end if
    LastShot(CurPlayer)=3
    Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
    exit sub
  else
    if LoveGunMode then
      if LastShot(CurPlayer) = -1 then
        DisplayI(17)
        AddScore(1000000)
      else
        DisplayI(19)
        AddScore(2000000)
      end if
      LastShot(CurPlayer)=3
      Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
      exit sub
    end if
  end if
  LastShot(CurPlayer)=3
  Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
  debug.print "Shots=" & Shots(CurPlayer,CurSong(CurPlayer))
  Select Case CurSong(CurPlayer)
    case 5: 
      ShowShot(Shots(CurPlayer,CurSong(CurPlayer))*200000)
      if i58.state=LightStateOff and i68.state=LightStateOff and i73.state=LightStateOff and i79.state=LightStateOff and i92.state=LightStateOff and i98.state=LightStateOff  then ' reset lights
        InitMode(CurSong(CurPlayer))
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) = 10 then SongComplete()
    case 2:
   '   DisplayI(8) ' Deuce
      ShowShot(100000)
       if i79.state=LightStateBlinking then
        RestoreRGB(i79):i79.state=LightStateBlinking:RestoreRGB(i92):i92.state=LightStateBlinking:SetLightColor i73, "white", 0
      else
        RestoreRGB(i73):i73.state=LightStateBlinking:RestoreRGB(i79):i79.state=LightStateBlinking:SetLightColor i68, "white", 0
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 3:
      ShowShot(100000):i98.state=LightStateBlinking:RestoreRGB(i98)
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 4:
      if Shots(CurPlayer,CurSong(CurPlayer)) > 17 then
        ShowShot(2000000)
      else
        ShowShot(250000+(100000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if sw38.isdropped=False then i35.state=LightStateBlinking end if
      if sw39.isdropped=False then i36.state=LightStateBlinking end if
      if sw40.isdropped=False then i37.state=LightStateBlinking end if
      if sw41.isdropped=False then i38.state=LightStateBlinking end if
      if Shots(CurPlayer,CurSong(CurPlayer)) > 10 then SongComplete()
    case 1:
      if Shots(CurPlayer,CurSong(CurPlayer)) > 20 then
        ShowShot(2000000)
      else
        ShowShot(50000+ (50000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if Not i79.state=LightStateOff then  ' Move Left
        RestoreRGB(i68):i68.state=LightStateBlinking:RestoreRGB(i73):i73.state=LightStateBlinking:SetLightColor i79, "white", 0
      else
        SetLightColor i68, "white", 0:RestoreRGB(i73):i73.state=LightStateBlinking:RestoreRGB(i79):i79.state=LightStateBlinking
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 6:
      ShowShot(200000*(Shots(CurPlayer,CurSong(CurPlayer))-1))
      if Shots(CurPlayer,CurSong(CurPlayer)) > 10 then
        SongComplete()
        SetLightColor i68, "white", 0
        SetLightColor i73, "white", 0
        SetLightColor i79, "white", 0
        SetLightColor i92, "white", 0
      else
        RestoreRGB(i68):RestoreRGB(i79):RestoreRGB(i92)
        SetLightColor i73, "white", 0
        i68.state=LightStateBlinking:i79.state=LightStateBlinking:i92.state=LightStateBlinking
      end if
    case 7:
      ShowShot(100000)
      ' Light Next Light In Sequence
      Select Case Shots(CurPlayer,CurSong(CurPlayer))
        Case 1: RestoreRGB(i68):i68.state=LightStateBlinking
        Case 3: RestoreRGB(i79):i79.state=LightStateBlinking
        Case 5: RestoreRGB(i92):i92.state=LightStateBlinking
        Case 7: RestoreRGB(i98):i98.state=LightStateBlinking
        Case 9: SongComplete()
      End Select  
    case 8:
      debug.print "R&R All night light Orbits"
      ShowShot(100000)
      SaveRGB(i58):i58.state=LightStateBlinking
      RestoreRGB(i98):i98.state=LightStateBlinking  ' Two orbits
    case 9:
    case 10:
  End Select
End Sub

Sub ProcessD
  SaveRGB(i79):SetLightColor i79, "white", 0
  if DemonMBMode then
    if LastShot(CurPlayer) = -1 then
      DisplayI(18)
      AddScore(1000000)
    else
      DisplayI(19)
      AddScore(2000000)
    end if
    LastShot(CurPlayer)=4
    Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
    Exit sub
  else
    if LoveGunMode then
      if LastShot(CurPlayer) = -1 then
        DisplayI(17)
        AddScore(1000000)
      else
        DisplayI(19)
        AddScore(2000000)
      end if
      LastShot(CurPlayer)=4
      Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
      exit sub
    end if
  end if
  LastShot(CurPlayer)=4
  Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
  debug.print "Shots=" & Shots(CurPlayer,CurSong(CurPlayer))
  Select Case CurSong(CurPlayer)
    case 5: 
      ShowShot(Shots(CurPlayer,CurSong(CurPlayer))*200000)
      if i58.state=LightStateOff and i68.state=LightStateOff and i73.state=LightStateOff and i79.state=LightStateOff and i92.state=LightStateOff and i98.state=LightStateOff  then ' reset lights
        InitMode(CurSong(CurPlayer))
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) = 10 then SongComplete()
    case 2:
      ShowShot(100000)
       if i92.state=LightStateBlinking then
        RestoreRGB(i92):i92.state=LightStateBlinking:SaveRGB(i98):i98.state=LightStateBlinking:SetLightColor i79, "white", 0
      else
        RestoreRGB(i79):i79.state=LightStateBlinking:RestoreRGB(i92):i92.state=LightStateBlinking:SetLightColor i73, "white", 0
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 3:
      ShowShot(100000):i98.state=LightStateBlinking:RestoreRGB(i98)
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 4:
      if Shots(CurPlayer,CurSong(CurPlayer)) > 17 then
        ShowShot(2000000)
      else
        ShowShot(250000+(100000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if sw38.isdropped=False then i35.state=LightStateBlinking end if
      if sw39.isdropped=False then i36.state=LightStateBlinking end if
      if sw40.isdropped=False then i37.state=LightStateBlinking end if
      if sw41.isdropped=False then i38.state=LightStateBlinking end if
      if Shots(CurPlayer,CurSong(CurPlayer)) > 10 then SongComplete()
    case 1:
      if Shots(CurPlayer,CurSong(CurPlayer)) > 20 then
        ShowShot(2000000)
      else
        ShowShot(50000+ (50000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if NOT i92.state=LightStateOff then  ' Move Left
        RestoreRGB(i73):i73.state=LightStateBlinking:RestoreRGB(i79):i79.state=LightStateBlinking:SetLightColor i92, "white", 0
      else
        SetLightColor i73, "white", 0:RestoreRGB(i79):i79.state=LightStateBlinking:RestoreRGB(i92):i92.state=LightStateBlinking
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 6:
      ShowShot(200000*(Shots(CurPlayer,CurSong(CurPlayer))-1))
      if Shots(CurPlayer,CurSong(CurPlayer)) > 10 then 
        SongComplete()
        SetLightColor i68, "white", 0
        SetLightColor i73, "white", 0
        SetLightColor i79, "white", 0
        SetLightColor i92, "white", 0
      else
        RestoreRGB(i68):RestoreRGB(i73):RestoreRGB(i92)
        SetLightColor i79, "white", 0
        i68.state=LightStateBlinking:i73.state=LightStateBlinking:i92.state=LightStateBlinking
      end if
    case 7:
      ShowShot(100000):RestoreRGB(i73):i73.state=LightStateBlinking
    case 8:
      ShowShot(100000)
      if i68.state=LightStateOff and i79.state=LightStateOff and i92.state=LightStateOff then
        RestoreRGB(i73):i73.state=LightStateBlinking
        SongComplete()
      End If
    case 9:
    case 10:
  End Select
End Sub

Sub ProcessRR
  SaveRGB(i92):SetLightColor i92, "white", 0
  if DemonMBMode then
    if LastShot(CurPlayer) = -1 then
      DisplayI(17)
      AddScore(1000000)
    else
      DisplayI(28)
      AddScore(2000000)
    end if
    LastShot(CurPlayer)=5
    Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
    exit sub
  else
    if LoveGunMode then
      if LastShot(CurPlayer) = -1 then
        DisplayI(19)
        AddScore(1000000)
      else
        DisplayI(19)
        AddScore(2000000)
      end if
      LastShot(CurPlayer)=5
      Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
      exit sub
    end if
  end if
  LastShot(CurPlayer)=5
  Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
  debug.print "Shots=" & Shots(CurPlayer,CurSong(CurPlayer))
  Select Case CurSong(CurPlayer)
    case 5: 
      ShowShot(Shots(CurPlayer,CurSong(CurPlayer))*200000)
      if i58.state=LightStateOff and i68.state=LightStateOff and i73.state=LightStateOff and i79.state=LightStateOff and i92.state=LightStateOff and i98.state=LightStateOff  then ' reset lights
        InitMode(CurSong(CurPlayer))
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) = 10 then SongComplete()
    case 2:
   '   DisplayI(8) ' Deuce
      ShowShot(100000)
       if i98.state=LightStateBlinking then
        RestoreRGB(i98):i98.state=LightStateBlinking:RestoreRGB(i58):i58.state=LightStateBlinking:SetLightColor i92, "white", 0
      else
        SetLightColor i79, "white", 0:RestoreRGB(i98):i98.state=LightStateBlinking:RestoreRGB(i92):i92.state=LightStateBlinking
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 3:
      ShowShot(100000):i98.state=LightStateBlinking:RestoreRGB(i98)
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 4:
      if Shots(CurPlayer,CurSong(CurPlayer)) > 17 then
        ShowShot(2000000)
      else
        ShowShot(250000+(100000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if sw38.isdropped=False then i35.state=LightStateBlinking end if
      if sw39.isdropped=False then i36.state=LightStateBlinking end if
      if sw40.isdropped=False then i37.state=LightStateBlinking end if
      if sw41.isdropped=False then i38.state=LightStateBlinking end if
      if Shots(CurPlayer,CurSong(CurPlayer)) > 10 then SongComplete()
    case 1:
      if Shots(CurPlayer,CurSong(CurPlayer)) > 20 then
        ShowShot(2000000)
      else
        ShowShot(50000+ (50000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if i98.state=LightStateBlinking then  ' Move Left
        RestoreRGB(i79):i79.state=LightStateBlinking:RestoreRGB(i92):i92.state=LightStateBlinking:SetLightColor i98, "white", 0
      else
        SetLightColor i79, "white", 0:RestoreRGB(i92):i92.state=LightStateBlinking:RestoreRGB(i98):i98.state=LightStateBlinking
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 6:
      ShowShot(200000*(Shots(CurPlayer,CurSong(CurPlayer))-1))
      if Shots(CurPlayer,CurSong(CurPlayer)) > 10 then
        SongComplete()
        SetLightColor i68, "white", 0
        SetLightColor i73, "white", 0
        SetLightColor i79, "white", 0
        SetLightColor i92, "white", 0
      else
        RestoreRGB(i68):RestoreRGB(i73):RestoreRGB(i79):SetLightColor i92, "white", 0
        i68.state=LightStateBlinking:i73.state=LightStateBlinking:i79.state=LightStateBlinking
      end if
    case 7:
      ShowShot(100000):RestoreRGB(i73):i73.state=LightStateBlinking
    case 8:
      ShowShot(100000)
      if i68.state=LightStateOff and i79.state=LightStateOff and i92.state=LightStateOff then
         RestoreRGB(i73):i73.state=LightStateBlinking
         SongComplete()
     End if
    case 9:
    case 10:
  End Select
End Sub

Sub ProcessRO
  debug.print "ProcessRO"
  SaveRGB(i98):SetLightColor i98, "white", 0
  if DemonMBMode then
    if LastShot(CurPlayer) = -1 then
      DisplayI(18)
      AddScore(1000000)
    else
      DisplayI(19)
      AddScore(2000000)
    end if
    LastShot(CurPlayer)=6
    Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
    exit sub
  else
    if LoveGunMode then
      if LastShot(CurPlayer) = -1 then
        DisplayI(17)
        AddScore(1000000)
      else
        DisplayI(19)
        AddScore(2000000)
      end if
      LastShot(CurPlayer)=6
      Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
      exit sub
    end if
  end if
  LastShot(CurPlayer)=6
  Shots(CurPlayer,CurSong(CurPlayer))=Shots(CurPlayer,CurSong(CurPlayer))+1
  debug.print "Shots=" & Shots(CurPlayer,CurSong(CurPlayer))
  Select Case CurSong(CurPlayer)
    case 5: 
      ShowShot(Shots(CurPlayer,CurSong(CurPlayer))*200000)
      if i58.state=LightStateOff and i68.state=LightStateOff and i73.state=LightStateOff and i79.state=LightStateOff and i92.state=LightStateOff and i98.state=LightStateOff  then ' reset lights
        InitMode(CurSong(CurPlayer))
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) = 10 then SongComplete()
    case 2:
   '   DisplayI(8) ' Deuce
      ShowShot(100000)
      if i58.state=LightStateBlinking then
        RestoreRGB(i58):i58.state=LightStateBlinking:RestoreRGB(i68):i68.state=LightStateBlinking:SetLightColor i98, "white", 0
      else
        RestoreRGB(i98):i98.state=LightStateBlinking:RestoreRGB(i58):i58.state=LightStateBlinking:SetLightColor i98, "white", 0
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 3: ' Goes to a random light
      ShowShot(100000)
      Select Case INT(RND*5)+1
        case 1: i58.state=LightStateBlinking:RestoreRGB(i58)
        case 2: i68.state=LightStateBlinking:RestoreRGB(i68)
        case 3: i73.state=LightStateBlinking:RestoreRGB(i73)
        case 4: i79.state=LightStateBlinking:RestoreRGB(i79)
        case 5: i92.state=LightStateBlinking:RestoreRGB(i92)
      End Select
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then SongComplete()
    case 4: 
      if Shots(CurPlayer,CurSong(CurPlayer)) > 17 then
        ShowShot(2000000)
      else
        ShowShot(250000+(100000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if sw38.isdropped=False then i35.state=LightStateBlinking end if
      if sw39.isdropped=False then i36.state=LightStateBlinking end if
      if sw40.isdropped=False then i37.state=LightStateBlinking end if
      if sw41.isdropped=False then i38.state=LightStateBlinking end if
      if Shots(CurPlayer,CurSong(CurPlayer)) > 10 then SongComplete()
    case 1:
      if Shots(CurPlayer,CurSong(CurPlayer)) > 20 then
        ShowShot(2000000)
      else
        ShowShot(50000+ (50000*Shots(CurPlayer,CurSong(CurPlayer))))
      end if
      if i58.state=LightStateBlinking then  ' Move Left
        RestoreRGB(i92):i92.state=LightStateBlinking:RestoreRGB(i98):i98.state=LightStateBlinking:SetLightColor i58, "white", 0
      else
        SetLightColor i92, "white", 0:RestoreRGB(i98):i98.state=LightStateBlinking:RestoreRGB(i58):i58.state=LightStateBlinking
      end if
      if Shots(CurPlayer,CurSong(CurPlayer)) Mod 12 = 0 then 
        SongComplete()
      end if
    case 6:
      ShowShot(300000)
      if i58.state=LightStateOff and i98.state=LightStateoff Then ' Both Orbits complete
        RestoreRGB(i68):RestoreRGB(i73):RestoreRGB(i79):RestoreRGB(i92)
        i68.state=LightStateBlinking:i73.state=LightStateBlinking:i79.state=LightStateOn:i92.state=LightStateBlinking
      end if
    case 7:
      ShowShot(100000):RestoreRGB(i73):i73.state=LightStateBlinking
    case 8:
       ShowShot(100000)
       if i58.state=LightStateOff then 
          RestoreRGB(i68):RestoreRGB(i68):i68.state=LightStateBlinking
          RestoreRGB(i79):RestoreRGB(i79):i79.state=LightStateBlinking
          RestoreRGB(i92):RestoreRGB(i92):i92.state=LightStateBlinking
       end if
    case 9:
    case 10:
  End Select
End Sub

Sub sw43HoleExit()
    If BallInHole > 0 Then
' check extraball
' check kiss Army
' check rock City
        ProcessScoop()
    End If
End Sub

Sub NewTrackTimer_Timer
   NewTrackTimer.enabled=False
   ChooseSongMode=False
   ScoopDelay.interval=500
   ScoopDelay.Enabled = True
   vpmtimer.addtimer 500, "FlashForMs FlasherExitHole, 1500, 30, 0 '"
   vpmtimer.addtimer 500, "FlashForMs SmallFlasher1, 1500, 30, 0 '"
End Sub

Sub ProcessScoop
  InScoop=True
  ScoopDelay.interval=4000

  if Not i40.state=LightStateOff then ' Kiss ArmyBonus  
    i40.state=LightStateOff 
    msgbox "NOT IMPLEMENTED YET"
  else
    if NOT i43.state=LightStateOff then ' Rock City
      debug.print "Rock City"
      i43.state=LightStateOff
      i25.state=LightStateOff:i26.state=LightStateOff:i27.state=LightStateOff:i28.state=LightStateOff ' Characters
      ' Light All Jackpots
      StopSound Track(cursong(CurPlayer))
      cursong(CurPlayer)=1
      PlaySound Track(cursong(CurPlayer)) ' "bg_music1" 'PlayMusic track(1) ' Detroit Rock City
    else
      if NOT i44.state=LightStateOff then ' New Track
        debug.print "New Track"
        KissHurryUp.enabled=False  ' Need to stop the hurry up and the countdown graphics
        ArmyHurryUp.enabled=False
        DMDTextPause "Select","New Track",3000
        EndMusic
        ChooseSongMode=TRUE
        NewTrackTimer.interval=10000:NewTrackTimer.enabled=True  'auto plunge in NewTrack Mode
        i44.state=LightStateOff
        AddScore(10000)
        Exit Sub
      end if 
      if NOT i39.state=LightStateOff then ' Extra BallHit 
          i39.state=LightStateOff
          debug.print "Extra Ball"
          AwardExtraBall
          LightRockAgain.state=LightStateON ' Rock Again lit
      else ' check backstage pass
        if NOT i41.state=LightStateOff then
          debug.print "BackStage Pass (video)"
          UDMDTimer.enabled=False
          i41.state=LightStateOff:i42.state=LightStateOff
          AddScore(5000)
          BSP.interval=2000:BSP.enabled=True
          Exit Sub ' 
        Else
          debug.print "Scoop - nothing lit"
          AddScore(5000)
          ScoopDelay.Interval = 1500
        End If
      End If
    End if
  End If
  ScoopDelay.Enabled = True
  vpmtimer.addtimer 500, "FlashForMs FlasherExitHole, 1500, 30, 0 '"
  vpmtimer.addtimer 500, "FlashForMs SmallFlasher1, 1500, 30, 0 '"

End Sub

Sub BSP_Timer 
  BSP.enabled=False
  vpmtimer.addtimer 1500, "FlashForMs FlasherExitHole, 1000, 30, 0 '"
  vpmtimer.addtimer 1500, "FlashForMs SmallFlasher1, 1000, 30, 0 '"
  ScoopCB() ' After Backstage Pass video is shown
  DMDGif "scene38.gif", "", "", slen(38)
End Sub

Sub Resync_targets()
    ' resync target lights in case we may have chose Hotter Than Hell
    debug.print "resync_targets"
    if sw38.isdropped=True then i35.state=LightStateOn else i35.state=LightStateOff end if
    if sw39.isdropped=True then i36.state=LightStateOn else i36.state=LightStateOff end if
    if sw40.isdropped=True then i37.state=LightStateOn else i37.state=LightStateOff end if
    if sw41.isdropped=True then i38.state=LightStateOn else i38.state=LightStateOff end if
End Sub

Sub ScoopCB() ' Play after BSP Scene completes
dim rr
   debug.print "ScoopCB()"
   rr=Int(Rnd*6)+1
   if rr > 6 then rr=6
   debug.print "Random Choice of " & rr
   ' Check that we dont already have the random choice
   if rr=1 and i118.state=LightStateBlinking then 
     rr=2
   end if
   if rr=2 and i95.state=LightStateBlinking then 
      rr=3
   end if
   if rr=3 and i122.state=LightStateBlinking then 
      rr=4
   end if
   if rr=5 and i29.state=LightStateOn and i30.state=LightStateOn and i31.state=LightStateOn then 
     rr=6
   end if
   if rr=6 and i61.state=LightStateOn then 
     rr = 4
   end if
Select case rr
     		  Case 1 : DMDGif "scene40.gif", "", "", slen(40)
                            ScoopDelay.interval=4000:ScoopDelay.enabled=True
                            debug.print "Superramps"
                            SetLightColor i118, "white", 2   ' Super Ramps
                            RampCnt(CurPlayer)=0
	   	  Case 2 : DMDGif "scene42.gif", "", "", slen(42)
                           ScoopDelay.interval=4000:ScoopDelay.enabled=True
                           SetLightColor i95, "white", 2   '
                           debug.print "SuperTargets"
                           TargetCnt(CurPlayer)=0
		  Case 3 : DMDGif "scene43.gif", "", "", slen(43)
                           ScoopDelay.interval=4000:ScoopDelay.enabled=True
                           debug.print "Super Spinner"
                           SetLightColor i122, "white", 2  
                           SpinCnt(CurPlayer)=0
                  Case 4 : DMDGif "scene62.gif", "2 MILLION", "", slen(62)
                           ScoopDelay.interval=slen(62)+500:ScoopDelay.enabled=True
                           debug.print "2 Million"
                           AddScore(2000000)
                  Case 5 : if i29.state=LightStateOff then
                             DMDGif "scene62.gif", "2X", "BONUS", slen(62)
                             ScoopDelay.interval=slen(62)+500:ScoopDelay.enabled=True
                             debug.print "2X"
                             BonusMultiplier(CurPlayer)=2
                             i29.state=LightStateOn
                           else
                             if i30.state=LightStateOff then
                               DMDGif "scene62.gif", "3X", "BONUS", slen(62)
                               ScoopDelay.interval=slen(62)+500:ScoopDelay.enabled=True
                               debug.print "3X"
                               BonusMultiplier(CurPlayer)=3
                               i30.state=LightStateOn
                             else
                               DMDGif "scene62.gif", "COLOSSAL", "BONUS", slen(62)
                               ScoopDelay.interval=slen(62)+500:ScoopDelay.enabled=True
                               debug.print "Colossal"
                               i31.state=LightStateOn
                               BonusMultiplier(CurPlayer)=5
                             end if
                           end if
                    Case 6 : DMDGif "scene41.gif", "", "", slen(41)
                           SetLightColor i61, "white", 2  
                           BumperCnt(CurPlayer)=0
                           debug.print "Super Bumpers"                       
                           ScoopDelay.interval=3000:ScoopDelay.enabled=True
     End Select
End Sub

Sub ScoopDelay_Timer
  debug.print "ScoopDelay_Timer()"
  ScoopDelay.Enabled = False
  BallInHole = BallInHole - 1
  sw43.CreateSizedball BallSize / 2
  PlaySound "fx_popper", 0, 1, -0.1, 0.25
  sw43.Kick 175, 14, 1
  vpmtimer.addtimer 1000, "sw43HoleExit '" 'repeat until all the balls are kicked out
  InScoop=False
End Sub

Sub SongComplete()
  DMDTextPause "Song Complete","Total " & ModeScore,1700
  ModeInProgress=False
  i44.state=LightStateBlinking ' Next Track
  for each xx in ShotsColl
    SetLightColor xx, "white", 0
  Next
  Resync_targets
End Sub

Sub ShowShot(points)
  DMDTextPause Title(cursong(CurPlayer)),points,1200
  AddScore(points)
End Sub

' STAR Targets
Sub CheckInstrument1 ' Pauls Guitar
  debug.print "CheckInstrument1"
  if NOT i72.state=LightStateOff then ' Collect Instrument
    debug.print "If STAR then collect Instrument"
    if i64.state=1 and i65.state=1 and i66.state=1 and i67.state=1 then 'STAR
      debug.print "STAR: Play Instrument Video"
      DMDGif "scene55.gif","","",slen(55) 
      i72.state=LightStateOff
      i21.state=LightStateOn  ' pf instrument light 
      i97.state=LightStateBlinking  ' Light Instrument Light
      instruments(CurPlayer)=instruments(CurPlayer)+1
      CheckInstruments()
    end if
  end if	      
End Sub

Sub CheckInstruments
  debug.print "CheckInstruments"
   if i21.state=LightStateOn and i22.state=LightStateOn and i23.state=LightStateOn and i24.state=LightStateOn then
     debug.print "Instrument Cnt is " & instruments(CurPlayer)
     if instruments(CurPlayer)=4 then
'       TextLine1.text="Extra Ball Light"
'       textline2.text="Instrument BONUS"
       I39.state=LightStateBlinking ' Light Extra Ball
       DisplayI(25) ' Extra Ball Lit
     end if
     i21.state=LightStateBlinking:i22.state=LightStateBlinking:i23.state=LightStateBlinking:i24.state=LightStateBlinking
   end if
   if instruments(CurPlayer)=6 then
     DMDGif "scene49.gif","","",slen(49)
     I26.blinkInterval=100
     I26.state=LightStateBlinking ' Spaceman
   end if
   if instruments(CurPlayer)=12 then
     DMDGif "scene49.gif","","",slen(49) 
     I26.blinkInterval=100
     I26.state=LightStateOn ' Spaceman
   end if
   if instruments(CurPlayer)=48 then ' Extra Ball
'     TextLine1.text="Extra Ball Light"
'     textline2.text="Instrument BONUS"
     I39.state=LightStateOn ' Light Extra Ball
   end if
End Sub

Sub CheckLoveGun
   debug.print "CheckLoveGun"
   if i64.state=1 and i65.state=1 and i66.state=1 and i67.state=1 then ' not already LG Ready
     debug.print "Turn off STAR"
     i64.state=0:i65.state=0:i66.state=0:i67.state=0
     if NOT i68.state=LightStateOff then
       processSC()
     end if
     if NOT LoveGunMode then   ' Dont turn on LG Mode while in LG
       if F116.state=LightStateOff then
         debug.print "Play LG Ready"
         F116.state=1                   ' star flasher
         DMDGif "scene26.gif","","",slen(26) 
       End If
    Else
      debug.print "Already in LG Mode"
    End if
   end if
End Sub

Sub CheckBumpers  
  if i17.state=LightStateBlinking and i18.state=LightStateBlinking and i19.state=LightStateBlinking and i20.state=LightStateBlinking and i28.state = LightStateOff then
    debug.print "CatMAN"
    DMDGif "scene47.gif","","",slen(47) 
    i28.blinkinterval=100
    i28.state=LightStateBlinking
    PlaySound "audio377" ' crowd noise
    CheckKissArmy()
  end if
  if i17.state=LightStateOn and i18.state=LightStateOn and i19.state=LightStateOn and i20.state=LightStateOn and NOT i28.state=LightStateOn then
    debug.print "CatMAN"
    DMDGif "scene47.gif","","",slen(47)
    i28.state=LightStateOn
    PlaySound "audio377" ' crowd noise
    CheckRockCity()
  end if
End Sub



' *************************************
'    KISS Game Rules
'
Const KissRulesV=1.05
' *************************************
' 20161030 - Super targets,bumpers
' 20161031 - save demon lock state between players
' 20161113 - strengthen demon kicker and enable ball save during multiball
' 20161119 - ball audio on right ramp
' 20161122 - front row bug 

Dim cRGB,xx
cRGB=1 ' Starting Colour

Sub InitMode(tr)
  debug.print "InitMode " & tr
  for each xx in ShotsColl
    xx.state=LightStateOff
  Next
  i35.state=LightStateOff ' KISS Targets
  i36.state=LightStateOff
  i37.state=LightStateOff
  i38.state=LightStateOff
  ModeInProgress=True
  Select Case tr
    case 1: ' two shots  lo-58, sc-68
      SetLightColor i58, RGBColors(cRGB), 2
      SetLightColor i68, RGBColors(cRGB), 2
    case 2: ' lo, sc
      SetLightColor i58, RGBColors(cRGB), 2
      SetLightColor i68, RGBColors(cRGB), 2
    case 3: ' ro
      SetLightColor i98, RGBColors(cRGB), 2
    case 4: ' kiss targets
      debug.print "KISS Targets Flash here..."
      SetLightColor i35, "white", 2
      SetLightColor i36, "white", 2
      SetLightColor i37, "white", 2
      SetLightColor i38, "white", 2
    case 5: ' sc, mr
      for each xx in ShotsColl
        SetLightColor xx, RGBColors(cRGB), 2
      Next
    case 6: ' lo, ro
      SetLightColor i58, RGBColors(cRGB), 2
      SetLightColor i98, RGBColors(cRGB), 2
    case 7: ' mid ramp
      SetLightColor i73, RGBColors(cRGB), 2
    case 8: ' mid ramp
      SetLightColor i73, RGBColors(cRGB), 2
    End Select
   cRGB=cRGB+1
   if cRGB>5 then cRGB=1
End Sub

'   ******************************************************
'   Rollover Switches

Sub sw1_hit  ' Leftmost Inlane
  DOF 133, DOFPulse
  PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub
  If i7.state <> LightStateOff then 'Light Bumper Shot
    SetLightColor i7, "white", 0
   ' light one of the "Light bumper" lights   i101,102,103,104, pf 17,18,19,20
   ' if any pf & upper lights are off then choose it
    if i17.state=LightStateOff or i18.state=lightstateoff or i19.state=lightstateoff or i20.state=lightstateoff then
      if i17.state=lightstateoff and i101.state=lightstateoff then
           SetLightColor i101, "white", 1
      else
        if i18.state=lightstateoff and i102.state=lightstateoff then
             SetLightColor i102, "white", 1
        else
          if i19.state=lightstateoff and i103.state=lightstateoff then
                SetLightColor i103, "white", 1
          else
            if i20.state=lightstateoff and i104.state=lightstateoff then
                SetLightColor i104, "white", 1
            end if
          end if
        end if
      end if
    else ' all of the lights were already done atleast once
      if i17.state=LightStateBlinking or i18.state=lightstateBlinking or i19.state=lightstateBlinking or i20.state=lightstateBlinking then
        if i17.state=lightstateBlinking and i101.state=lightstateoff then
                SetLightColor i101, "white", 1
        else
          if i18.state=lightstateBlinking and i102.state=lightstateoff then
                SetLightColor i102, "white", 1
          else
            if i19.state=lightstateBlinking and i103.state=lightstateoff then
                SetLightColor i103, "white", 1
            else
              if i20.state=lightstateBlinking and i104.state=lightstateoff then
                SetLightColor i104, "white", 1
              end if
            end if
          end if
        end if
      else ' already got all the pf lights so lets just build them up for increased bumper points and light colours
        if i101.state=lightstateoff and bumpercolor(4) < 4 then   ' probably should light them based on bumper colours!!
                SetLightColor i101, "white", 1
        else
          if i102.state=lightstateoff and bumpercolor(3) < 4  then
              SetLightColor i102, "white", 1
          else
            if i103.state=lightstateoff and bumpercolor(1) < 4 then
                SetLightColor i103, "white", 1
            else
              if i104.state=lightstateoff and bumpercolor(2) < 4 then
                SetLightColor i104, "white", 1
              end if
            end if
          end if
        end if
      end if
    end if
  End if
  AddScore(5000)
End Sub

Sub sw2_hit ' Right Inlane Kiss Combo
  DOF 134, DOFPulse
  PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub

  If i33.state=LightStateOff then
    KissCombo.Interval=10000
    KissCombo.enabled=True
    SetLightColor i33, "white", 2
    if i9.state = LightStateOff or i10.state = LightStateOff or i11.state = LightStateOff or i12.state = LightStateOff then
      DisplayI(5)
    else  ' you hit all the targets so now its a hurry up bonus
      SetLightColor i115, "white", 2
      DMDTextI "KISS COMBO","HURRY UP!", bgi                          ' right Lane
      KISSBonus=1000000
      KISSHurryUp.Interval=400
      KISSHurryUp.Enabled=True
    End if
  End if
  AddScore(5000)
End Sub

Sub sw3_hit
  DOF 132, DOFPulse
  PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub
  if LightRockAgain.state=LightStateOff then ' Save BallHit 
    if i6.state <> LightStateOff then
      if not DemonMBMode then
        DisplayI(13)
        PlaySound "audio113"
        Debug.print "Exit via Front Row"
        FrontRowSave=True ' Next Drain is just a drain .. 
      End IF
    end if
  end if
  AddScore(5000)
End Sub

Sub sw4_hit
  DOF 135, DOFPulse
  PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub
  AddScore(5000)
End Sub

Sub sw14_hit   ' left orbit
  DOF 130, DOFPulse
  PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
  If Tilted Then Exit Sub
  If BallLooping.enabled=False then ' Ignore Initial ball plunge
    BallLooping.enabled=True ' Ignore next switch if looping around 
    if KISSHurryUp.enabled=True Then ' Score KISSHurryUp
      KISSHurryUp.Enabled=False
      AddScore(KISSBONUS) ' HurryUp Bonus Award
      KissCombo.enabled=False
      debug.print "kiss Hurry Up"
      DMDFlush()
      DMDTextI ">KISS Hurry Up!<", KISSBONUS, bgi
      ComboCnt(CurPlayer)=ComboCnt(CurPlayer)+1
      i33.state=LightStateOff
    else
      DMDGif "scene51.gif","YOU ROCK!","",slen(51)
    end if
    AddScore(10000)
    if NOT i58.state=LightStateOff then
      processLO()
    else
      LastShot(CurPlayer)=-1
    end if
    if NOT i62.state=LightStateOff then ' Collect Instrument Drums
      DMDGif "scene58.gif","","",slen(58)
      i62.state=LightStateOff  
      SetLightColor i24, "white", 1  ' pf instrument light 
      i97.state=LightStateBlinking  ' Light Instrument Light
      instruments(CurPlayer)=instruments(CurPlayer)+1
      CheckInstruments()
    end if
    if i7.state=LightStateOff then ' Light Bumper Shot
      SetLightColor i7, "white", 1
    end if
  end if
End Sub

Sub sw24_hit
  debug.print "sw24_hit()"
  DOF 131, DOFPulse
  PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
  If Tilted Then Exit Sub
  if BallLooping.enabled=False then ' Ignore if this is the initial ball plunge
    BallLooping.enabled=True ' Ignore next switch if looping around 
    AddScore(10000)
    if NOT i98.state=LightStateOff then
      processRO()
    else
      LastShot(CurPlayer)=-1
    end if

    if i101.state=LightStateOn then  ' for catman bumper4 
      i101.state=LightStateOff
      if i17.state=LightStateOff then
          SetLightColor i17, "white", 2
    else
      if i17.state=LightStateBlinking then
            SetLightColor i17, "white", 1
      end if
    end if

    ' add one to the bumper light!

    if BumperColor(4)=0 then  ' None
       flasher4.visible=True
       flasher4.color=RGB(255, 255, 255) ' white
       SetLightColor B4L, "white", 1
    else
      if BumperColor(4)=1 Then
         debug.print "green"
         flasher4.color=RGB(0, 255, 0) ' green
         SetLightColor B4L, "green", 1
      else
         if BumperColor(4)=2 then
           flasher4.color=RGB(0, 0, 128) ' blue
           SetLightColor B4L, "blue", 1
           debug.print "purple"
         else
           if BumperColor(4)=3 then
             flasher4.color=RGB(255,0,0) ' red
             SetLightColor B4L, "red", 1
             debug.print "red"
           end if
         end if
      end if
    End If

    BumperColor(4)=BumperColor(4)+1
  end if

  if i102.state=LightStateOn then  ' for Spaceman  bumper3 
    i102.state=LightStateOff
    if i18.state=LightStateOff then
          SetLightColor i18, "white", 2
    else
      if i18.state=LightStateBlinking then
            SetLightColor i18, "white", 1
      end if
    end if

    ' add one to the bumper light!

    if BumperColor(3)=0 then  ' None
       flasher3.visible=True
       flasher3.color=RGB(255, 255, 255) ' white
       SetLightColor B3L, "white", 1
    else
      if BumperColor(3)=1 Then
         debug.print "green"
         flasher3.color=RGB(0, 255, 0) ' green
         SetLightColor B3L, "green", 1
      else
         if BumperColor(3)=2 then
           flasher3.color=RGB(0, 0, 128) ' blue
           SetLightColor B3L, "blue", 1
           debug.print "purple"
         else
           if BumperColor(3)=3 then
             flasher3.color=RGB(255,0,0) ' red
             SetLightColor B3L, "red", 1
             debug.print "red"
           end if
         end if
      end if
    End If

    BumperColor(3)=BumperColor(3)+1
  end if

  if i103.state=LightStateOn then   ' for Starchild   bumper1 
    i103.state=LightStateOff
    if i19.state=LightStateOff then
          SetLightColor i19, "white", 2
    else
      if i19.state=LightStateBlinking then
            SetLightColor i19, "white", 1
      end if
    end if

    ' add one to the bumper light!

    if BumperColor(1)=0 then  ' None
       flasher1.visible=True
       flasher1.color=RGB(255, 255, 255) ' white
       SetLightColor B1L, "white", 1
    else
      if BumperColor(1)=1 Then
         debug.print "green"
         flasher1.color=RGB(0, 255, 0) ' green
         SetLightColor B1L, "green", 1
      else
         if BumperColor(1)=2 then
           flasher1.color=RGB(0, 0, 128) ' blue
           SetLightColor B1L, "blue", 1
           debug.print "purple"
         else
           if BumperColor(1)=3 then
             flasher1.color=RGB(255,0,0) ' red
             SetLightColor B1L, "red", 1
             debug.print "red"
           end if
         end if
      end if
    End If

    BumperColor(1)=BumperColor(1)+1
  end if
'
   if i104.state=LightStateOn then    ' for demon   bumper2 
    i104.state=LightStateOff
    if i20.state=LightStateOff then
          SetLightColor i20, "white", 2
    else
      if i20.state=LightStateBlinking then
            SetLightColor i20, "white", 1
      end if
    end if

    ' add one to the bumper light!

    if BumperColor(2)=0 then  ' None
       flasher2.visible=True
       flasher2.color=RGB(255, 255, 255) ' white
       SetLightColor B2L, "white", 1
    else
      if BumperColor(2)=1 Then
         debug.print "green"
         flasher2.color=RGB(0, 255, 0) ' green
         SetLightColor B2L, "green", 1
      else
         if BumperColor(2)=2 then
           flasher2.color=RGB(0, 0, 128) ' blue
           SetLightColor B2L, "blue", 1
           debug.print "purple"
         else
           if BumperColor(2)=3 then
             flasher2.color=RGB(255,0,0) ' red
             SetLightColor B2L, "red", 1
             debug.print "red"
           end if
         end if
      end if
    End If

    BumperColor(2)=BumperColor(2)+1
  end if
   CheckBumpers()
 end if
End Sub

Sub BallLooping_Timer()
  BallLooping.enabled=False
End Sub

' ***********************
' STAR Targets
' ***********************
Sub sw25_hit  ' S-tar
  PlaySound SoundFXDOF("audio22",117,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
  If Tilted Then Exit Sub
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End if
  Else
    AddScore(5000)
  End If

  i64.state=LightStateOn   ' S
  CheckInstrument1()
  CheckLoveGun()
End Sub

Sub sw26_hit  ' s-T-tar
  PlaySound SoundFXDOF("audio22",117,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
  If Tilted Then Exit Sub
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End If
  Else
    AddScore(5000)
  End If

  I65.state=LightStateOn   ' S
  CheckInstrument1()
  CheckLoveGun()
End Sub
Sub sw27_hit  ' st-A-r
  PlaySound SoundFXDOF("audio22",117,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
  If Tilted Then Exit Sub
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End if
  Else
    AddScore(5000)
  End If

  I66.state=LightStateOn   ' S
  CheckInstrument1()
  CheckLoveGun()
End Sub
Sub sw28_hit  ' sta-R
  PlaySound SoundFXDOF("audio22",117,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
  If Tilted Then Exit Sub
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End if
  Else
    AddScore(5000)
  End If

  I67.state=LightStateOn   ' S
  CheckInstrument1()
  CheckLoveGun()
End Sub

'***************
' Ramp Switches
'***************

Sub sw56_Hit() ' Right Ramp1  RightRampDone
  PlaySound "fx_metalrolling", 0, 1, pan(ActiveBall)
  DOF 131, DOFPulse
  PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
  If Tilted Then Exit Sub
  PlaySound "audio167",0,0.2

  if NOT i118.state=LightStateOff then ' Super Ramps
      debug.print "Ramp Cnt is " & RampCnt(CurPlayer)
      RampCnt(CurPlayer)=RampCnt(CurPlayer)+1
      if lastShot(CurrentPlayer)=3 then ' if prior shot was mid ramp
        AddScore(150000)
      else
        AddScore(75000)
      End if
      if RampCnt(CurPlayer) > 9 Then  
        i118.state=LightStateOff
        DisplayI(27) ' Completed

        SpinCnt(CurPlayer) = 0
      else
        DisplayI(31)
      end if
  else
      AddScore(10000)
  end if

  if NOT i92.state=LightStateOff then
    processRR()
  else
    LastShot(CurPlayer)=-1
  end if

  if i97.state=LightStateBlinking then ' Light Instrument
    debug.print "Light an instrument"
    i97.state=LightStateOff
    if NOT i21.state=LightStateOn then
      SetLightColor i72, "white", 2   ' Light  star child
      DisplayI(14)
    else
      if NOT i22.state=LightStateOn then
        SetLightColor i78, "white", 2   ' Light center Ramp
        DisplayI(14)
      else
        if NOT i23.state=LightStateOn then
          SetLightColor i91, "white", 2   ' Light  Demon 
          DisplayI(3)
        else
          if NOT i24.state=LightStateOn then
            SetLightColor i62, "white", 2   ' Light 
            DisplayI(4)
          end if
        end if
      end if
    end if
  end if
  if i7.state=LightStateOff then ' Light Bumper Shot
    SetLightColor i7, "white", 2   ' Light Bumper
  end if
  RandomScene()
End sub

Sub sw64_Hit  ' Center Ramp
  debug.print "sw64_hit"
  PlaySound "fx_metalrolling", 0, 1, pan(ActiveBall)
  If Tilted Then Exit Sub
  PlaySound "audio166",0,0.2

  if NOT i118.state=LightStateOff then ' Super Ramps
      debug.print "Ramp Cnt is " & RampCnt(CurPlayer)
      RampCnt(CurPlayer)=RampCnt(CurPlayer)+1
      if lastShot(CurrentPlayer)=5 then  ' If prior shot was RR
        AddScore(150000)
      else
        AddScore(75000)
      End if
      if RampCnt(CurPlayer) > 9 Then  
        i118.state=LightStateOff
        DisplayI(27) ' Completed
        SpinCnt(CurPlayer) = 0
      else
        DisplayI(31) ' SuperRamps
      end if
  else
      AddScore(10000)
  end if

  if NOT i73.state=LightStateOff then
    debug.print "process MR()"
    processMR()
  else
    LastShot(CurPlayer)=-1
  end if

  if ArmyHurryUp.enabled then ' ArmyHurryUp 
    AddScore(ArmyBonus)
    DMDFlush()
    DMDTextPause ">Army Hurry Up!<", ArmyBonus, 500
    i77.state=LightStateOff
    ArmyHurryUp.enabled=False
  end if

  if NOT i78.state=LightStateOff then ' Collect Instrument Aces Guitar
    DMDGif "scene53.gif","","",slen(53) 
    i78.state=LightStateOff
    i22.state=LightStateOn        ' pf instrument light 
    i97.state=LightStateBlinking  ' Light Instrument Light
    instruments(CurPlayer)=instruments(CurPlayer)+1
    CheckInstruments()
  end if	
  if i7.state=LightStateOff then ' Light Bumper Shot
    SetLightColor i7,"white",2
  end if      
  RandomScene()
End sub

'   *********************************
'   Demon Lock Targets
Sub sw44_Hit
  PlaySound SoundFXDOF("audio22",128,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub
  if I82.state=LightStateBlinking then
    if Not i95.state=LightStateOff then
      targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
      if targetCnt(CurPlayer)=100 then
        DisplayI(30)
        AddScore(40000)
        i95.state=LightStateOff
      else
        if TargetCnt(CurPlayer) < 100 then
            AddScore(40000)
            if rnd*10 > 4 then 
              UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
            Else  
              DisplayI(34)
            end if
        Else  
            AddScore(5000)
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
        End If
      End if
    Else
      AddScore(5000)
    End If
    SetLightColor i82, "green", 1
    CheckLock
  else
    AddScore(100)
  end if
End Sub

Sub sw45_Hit
  PlaySound SoundFXDOF("audio22",128,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub
  if I84.state=LightStateBlinking then
    if Not i95.state=LightStateOff then
      targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
      if targetCnt(CurPlayer)=100 then
        DisplayI(30)
        AddScore(40000)
        i95.state=LightStateOff
      else
        if TargetCnt(CurPlayer) < 100 then
            AddScore(40000)
            if rnd*10 > 4 then 
              UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
            Else  
              DisplayI(34)
            end if
        Else  
            AddScore(5000)
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
        End If
      End if
    Else
      AddScore(5000)
    End If
    SetLightColor i84, "green", 1
    CheckLock
  else
    AddScore(100)
  end if
End Sub

Sub CheckLock()
    If bLockEnabled Then Exit Sub
    If i82.State + i84.State = 2 Then
        bLockEnabled = True
        SetLightColor i90, "green", 2 
    End If
End Sub

' ***********************************
Sub sw46_hit ' Rightmost Left Inlane Army Combo
  DOF 133, DOFPulse
  PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub
  ArmyCombo.Interval=10000
  ArmyCombo.enabled=True
  SetLightColor i8, "white", 2

  if i13.state=LightStateOff or i14.state=LightStateOff or i15.state=LightStateOff or i16.state=LightStateOff then
    AddScore(5000) ' just a normal combo shot lit
    DisplayI(11)
  else
    DMDTextI "ARMY", "HURRYUP!", bgi
    SetLightColor i77, "white", 2 ' Kiss Army Lights 
        ' start hurry Update  
     ArmyBonus=900000
     ArmyHurryUp.Interval=400
     ArmyHurryUp.Enabled=True
  end if
End Sub

Sub ArmyCombo_Timer()
  i8.state=LightStateOff
  ArmyCombo.enabled=False
End Sub

Sub KissCombo_Timer()
  i33.state=LightStateOff
  KissCombo.enabled=False
End Sub

'  *******************************************************
'     Right StandUp Targets
Sub sw58_hit
  PlaySound SoundFXDOF("audio134",116,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub
  FlashForMs SmallFlasher2, 1000, 50, 0
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End if
  Else
    AddScore(5000)
  End If
  SetLightColor i106, "white", 1
  SetLightColor i13, "yellow", 2
' Check for a Combo bonus
  if NOT i8.state = lightStateOff then ' check if this is a hurryup or just a combo 
    if i13.state=LightStateOff or i14.state=LightStateOff or i15.state=LightStateOff or i16.state=LightStateOff then
      ArmyCombo.enabled=False
      debug.print "army combo"
      DMDTextI "ARMY COMBO", "100000", bgi
      ComboCnt(CurPlayer)=ComboCnt(CurPlayer)+1
      i8.state=LightStateOff
      AddScore(100000)
    End If
  end if
  CheckBackStage()
End Sub

Sub sw59_hit
  PlaySound SoundFXDOF("audio134",116,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End if
  Else
    AddScore(5000)
  End If
  FlashForMs SmallFlasher2, 1000, 50, 0
  SetLightColor i107, "white", 1
  SetLightColor i14, "yellow", 2
' Check for a Combo bonus
  if NOT i8.state = lightStateOff then ' check if this is a hurryup or just a combo 
    if i13.state=LightStateOff or i14.state=LightStateOff or i15.state=LightStateOff or i16.state=LightStateOff then
      ArmyCombo.enabled=False
      debug.print "army combo"
      DMDTextI "ARMY COMBO", "100000", bgi
      ComboCnt(CurPlayer)=ComboCnt(CurPlayer)+1
      i8.state=LightStateOff
      AddScore(100000)
    End If
  end if
  CheckBackStage()
End Sub

Sub sw60_hit
  PlaySound SoundFXDOF("audio134",116,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End if
  Else
    AddScore(5000)
  End If
  FlashForMs SmallFlasher2, 1000, 50, 0
  SetLightColor i108, "white", 1
  SetLightColor i15, "yellow", 2
' Check for a Combo bonus
  if NOT i8.state = lightStateOff then ' check if this is a hurryup or just a combo 
    if i13.state=LightStateOff or i14.state=LightStateOff or i15.state=LightStateOff or i16.state=LightStateOff then
      ArmyCombo.enabled=False
      debug.print "army combo"
      DMDTextPause "ARMY COMBO", ArmyBonus,500
      ComboCnt(CurPlayer)=ComboCnt(CurPlayer)+1
      i8.state=LightStateOff
      AddScore(100000)
    End If
  end if
  CheckBackStage()
End Sub

Sub sw61_hit
  PlaySound SoundFXDOF("audio134",116,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End if
  Else
    AddScore(5000)
  End If
  FlashForMs SmallFlasher2, 1000, 50, 0
  SetLightColor i109, "white", 1
  SetLightColor i16, "yellow", 2
' Check for a Combo bonus
  if NOT i8.state = lightStateOff then ' check if this is a hurryup or just a combo 
    if i13.state=LightStateOff or i14.state=LightStateOff or i15.state=LightStateOff or i16.state=LightStateOff then
      ArmyCombo.enabled=False
      debug.print "army combo"
      DMDTextI "ARMY COMBO", "100000", bgi
      ComboCnt(CurPlayer)=ComboCnt(CurPlayer)+1
      i8.state=LightStateOff
      AddScore(100000)
    End If
  End if
  CheckBackStage()
End Sub

'  *******************************************************
'     Left Drop Targets

Sub sw38_hit
  PlaySound SoundFXDOF("audio22",116,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
End Sub

Sub sw38_Dropped
  debug.print "Target Dropped"
  if Tilted Then Exit Sub
  if NOT i33.state=LightStateOff then
    if i9.state = LightStateOff or i10.state = LightStateOff or i11.state = LightStateOff or i12.state = LightStateOff then
      AddScore(100000)  ' Combo but No HurryUp
      KissCombo.enabled=False
      debug.print "kiss combo"
      DMDTextI "KISS COMBO", "100000", bgi
      ComboCnt(CurPlayer)=ComboCnt(CurPlayer)+1
      i33.state=LightStateOff
    end if 
  end if

  if i35.state=LightStateBlinking then 'Hotter Than Hell Mode
    ProcessTarget()
  end if

  SetLightColor i35, "white", 1
  SetLightColor i9, "white", 1
  if B2SOn=True then
    Controller.B2SSetData 1, 1
  end if
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End if
  Else
    AddScore(5000)
  End If
  If sw38.IsDropped=1 and sw39.IsDropped=1 and sw40.IsDropped=1 and sw41.IsDropped=1 then 
    debug.print "All targets dropped"
    KissTargetStack=KissTargetStack+1
    FrontRowStack=FrontRowStack+1
    ResetTargets()
  End if
  CheckBackStage():CheckFrontRow()
End Sub

Sub sw39_hit
  PlaySound SoundFXDOF("audio22",116,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
End Sub

Sub sw39_Dropped
  debug.print "Target Dropped"
  if Tilted Then Exit Sub
  if NOT i33.state=LightStateOff then
    if i9.state = LightStateOff or i10.state = LightStateOff or i11.state = LightStateOff or i12.state = LightStateOff then
      AddScore(100000)  ' Combo but No HurryUp
      KissCombo.enabled=False
      debug.print "kiss combo"
      DMDTextI "KISS COMBO", "100000", bgi
      ComboCnt(CurPlayer)=ComboCnt(CurPlayer)+1
      i33.state=LightStateOff
    end if 
  end if

  if i36.state=LightStateBlinking then 'Hotter Than Hell Mode
    ProcessTarget()
  end if

  SetLightColor i36, "white", 1
  SetLightColor i10, "white", 1
  if B2SOn=True then
    Controller.B2SSetData 2, 1
  end if
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End if
  Else
    AddScore(5000)
  End If
  If sw38.IsDropped=1 and sw39.IsDropped=1 and sw40.IsDropped=1 and sw41.IsDropped=1 then 
    debug.print "All targets dropped"
    KissTargetStack=KissTargetStack+1
    FrontRowStack=FrontRowStack+1
    ResetTargets()
  End if
  CheckBackStage():CheckFrontRow()
End Sub

Sub sw40_hit
  PlaySound SoundFXDOF("audio22",116,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
End Sub

Sub sw40_Dropped
  debug.print "Target Dropped"
  if Tilted Then Exit Sub
  if NOT i33.state=LightStateOff then
    if i9.state = LightStateOff or i10.state = LightStateOff or i11.state = LightStateOff or i12.state = LightStateOff then
      AddScore(100000)  ' Combo but No HurryUp
      KissCombo.enabled=False
      debug.print "kiss combo"
      DMDTextI "KISS COMBO", "100000", bgi
      ComboCnt(CurPlayer)=ComboCnt(CurPlayer)+1
      i33.state=LightStateOff
    end if 
  end if

  if i37.state=LightStateBlinking then 'Hotter Than Hell Mode
    ProcessTarget()
  end if

  SetLightColor i37, "white", 1
  SetLightColor i11, "white", 1
  if B2SOn=True then
    Controller.B2SSetData 3, 1
  end if
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End if
  Else
    AddScore(5000)
  End If
  If sw38.IsDropped=1 and sw39.IsDropped=1 and sw40.IsDropped=1 and sw41.IsDropped=1 then 
    debug.print "All targets dropped"
    KissTargetStack=KissTargetStack+1
    FrontRowStack=FrontRowStack+1
    ResetTargets()
  End if
  CheckBackStage():CheckFrontRow()
End Sub

Sub sw41_hit
  PlaySound SoundFXDOF("audio22",116,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
End Sub

Sub sw41_Dropped
  debug.print "Target Dropped"
  if Tilted Then Exit Sub
  if NOT i33.state=LightStateOff then
    if i9.state = LightStateOff or i10.state = LightStateOff or i11.state = LightStateOff or i12.state = LightStateOff then
      AddScore(100000)  ' Combo but No HurryUp
      KissCombo.enabled=False
      debug.print "kiss combo"
      DMDTextI "KISS COMBO", "100000", bgi
      ComboCnt(CurPlayer)=ComboCnt(CurPlayer)+1
      i33.state=LightStateOff
    end if 
  end if

  if i38.state=LightStateBlinking then 'Hotter Than Hell Mode
    ProcessTarget()
  end if

  SetLightColor i38, "white", 1
  SetLightColor i12, "white", 1
  if B2SOn=True then
    Controller.B2SSetData 4, 1
  end if
  if Not i95.state=LightStateOff then
    targetCnt(CurPlayer)=targetCnt(CurPlayer)+1
    if targetCnt(CurPlayer)=100 then
      DisplayI(30)
      AddScore(40000)
      i95.state=LightStateOff
    else
      if TargetCnt(CurPlayer) < 100 then
          AddScore(40000)
          if rnd*10 > 4 then 
            UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("40k",INT(RND*8)+3," "), 15, 14, 100, 14
          Else  
            DisplayI(34)
          end if
      Else  
          AddScore(5000)
          UltraDMD.DisplayScene00 "scene19.gif", LPad("",INT(RND*3)," "), 10, Lpad("5k",INT(RND*8)+3," "), 15, 14, 100, 14
      End If
    End if
  Else
    AddScore(5000)
  End If
  If sw38.IsDropped=1 and sw39.IsDropped=1 and sw40.IsDropped=1 and sw41.IsDropped=1 then 
    debug.print "All targets dropped"
    KissTargetStack=KissTargetStack+1
    FrontRowStack=FrontRowStack+1
    ResetTargets()
  End if
  CheckBackStage():CheckFrontRow()
End Sub


' *********************************************************************
'                        General Routines
' *********************************************************************

Sub CheckBackStage
   debug.print "CheckBackStage"
   if I106.state=LightStateOn and I107.state=LightStateOn and I108.state=LightStateOn and I109.state=LightStateOn then
     ArmyTargetStack=ArmyTargetStack+1
     I106.state=LightStateOff:I107.state=LightStateOff:I108.state=LightStateOff:I109.state=LightStateOff
   end if

   if i41.state=LightStateOff then ' Not already lit
     if KissTargetStack > 0 then
       debug.print "Kiss Targets Complete"
       KissTargetStack=KissTargetStack-1
       DMDGif "scene37.gif"," ","LIT",slen(37)
       i41.state=LightStateBlinking:i42.state=LightStateBlinking
       SetLightColor i41, "white", 2
       SetLightColor i42, "white", 2
     else
       if ArmyTargetStack > 0 then 
         debug.print "Army targets complete"
         ArmyTargetStack=ArmyTargetStack-1
         DMDGif "scene37.gif"," ","LIT",slen(37)
         SetLightColor i41, "white", 2
         SetLightColor i42, "white", 2
       '  textline1.text="backstage pass"
       end if
     end if
   end if
End Sub

Sub CheckFrontRow
   if i6.state=LightStateOff and FrontRowStack > 0 and Not LoveGunMode and Not DemonMBMode then
     FrontRowStack=FrontRowStack-1
     SetLightColor i6, "orange", 2
     DMDTextI "FRONT ROW", "", bgi
   end if
End Sub

Sub ArmyHurryUp_timer
  ArmyBonus=ArmyBonus-15000
  DMDTextPause "ARMY Hurry Up!", ArmyBonus, 0
  if ArmyBonus <=0 then
    ArmyHurryUp.Enabled=False
    i77.state=LightStateOff
  end if
End Sub


Sub KISSHurryUp_timer
  KISSBonus=KISSBonus- 25000
  DMDTextPause "KISS Hurry Up!", KISSBonus,0
  if KISSBonus <=0 then
    KISSHurryUp.Enabled=False
    i115.state=LightStateOff  ' Doublecheck that light
  end if
End Sub

Sub sctarget_hit
  PlaySound SoundFXDOF("audio22",116,DOFPulse,DOFTargets), 0, 1, pan(ActiveBall)
  if Tilted Then Exit Sub
  AddScore(15000)
End Sub

Sub ResetTargets()
  debug.print "ResetTargets()"
  sw38.isDropped = False ' KISS Drop Targets
  sw39.isDropped = False
  sw40.isDropped = False
  sw41.isDropped = False
  if cursong(CurPlayer)=4 then ' Hotter Than Hell means they need to flash on Reset 
    debug.print "Target lights blink for Hotter Than Hell"
    I35.state=LightStateBlinking:I36.state=LightStateBlinking:I37.state=LightStateBlinking:I38.state=LightStateBlinking
  else
    I35.state=LightStateOff:I36.state=LightStateOff:I37.state=LightStateOff:I38.state=LightStateOff
  end if
  if B2SOn=True then
    Controller.B2SSetData 1, 0
    Controller.B2SSetData 2, 0
    Controller.B2SSetData 3, 0
    Controller.B2SSetData 4, 0
  end if
End Sub

'***********************************
'Demon Lock
'***********************************

Dim bLockEnabled

Sub sw42_Hit()
    PlaySound "fx_kicker_enter", 0, 1, 0.1
    If Tilted Then Exit Sub
    if NOT i79.state=LightStateOff then
      processD()
    else
      LastShot(CurPlayer)=-1
    end if
    Addscore 10000
    if NOT i91.state=LightStateOff then ' Collect Instrument Genes Bass
           DMDGif "scene56.gif","","",slen(56) 
           i91.state=LightStateOff
           SetLightColor i23, "white", 1  ' pf instrument light 
           SetLightColor i97, "white", 2  ' Light Instrument Light
           instruments(CurPlayer)=instruments(CurPlayer)+1
           CheckInstruments()
    end if
    if i85.state=lightstateoff then  'D-E-M-O-N lights
          SetLightColor i85, "orange", 1:DisplayI(6)
    else
      if i86.state=lightstateoff then
            SetLightColor i86, "orange", 1:DisplayI(6)
      else
            if i87.state=lightstateoff then
              SetLightColor i87, "orange", 1:DisplayI(6)
            else
              if i88.state=lightstateoff then
                SetLightColor i88, "orange", 1:DisplayI(6)
              else
                if i89.state=lightstateoff then
                  SetLightColor i89, "orange", 1:DisplayI(6)
                  PlaySound "audio429"
                end if
              end if
            end if
      end if
    end if
    If bLockEnabled and NOT demonMBMode and NOT LoveGunMode Then
            LockedBalls(CurPlayer) = LockedBalls(CurPlayer) + 1
            bLockEnabled=False
            Playsound "audio" & 657 + LockedBalls(CurPlayer)
            SetLightColor i82, "orange", 2
            SetLightColor i84, "orange", 2
            i90.State = 0
            DMDTextI "BALL " & LockedBalls(CurPlayer), "IS LOCKED", bgi
            If LockedBalls(CurPlayer) = 3 Then
                MBPauseTimer.interval=2000:MBPauseTimer.enabled=True
                vpmtimer.addtimer 2000, "DemonKickBall '"
            else
                vpmtimer.addtimer 1500, "DemonKickBall '"
            End IF
    Else
       vpmtimer.addtimer 1500, "DemonKickBall '"
    End If

End Sub

Sub MBPauseTimer_Timer()
  MBPauseTimer.enabled=False
  DemonMultiball
End Sub

Sub DemonMultiball()
    FlashForMs Flasher9, 1000, 50, 0
    FlashForMs Flasher10, 1000, 50, 0 
    DemonMBMode=True
    SaveSong=cursong(CurPlayer)
    debug.print "Saving Song #" & SaveSong
    SaveShots()
    AddMultiball 3
    if Not bBallSaverActive then
      EnableBallSaver(9) ' Multiball BallSaver
    end if

    LockedBalls(CurPlayer) = 0
    bLockEnabled = False
    i85.State = 0
    i86.State = 0
    i87.State = 0
    i88.State = 0
    i89.State = 0

    ' turn off the lock lights
    SetLightColor i82, "orange", 0
    SetLightColor i84, "orange", 0
    i90.State = 0

   ' turn off front row ball Save
    i6.state=lightStateOff
    i44.state=lightstateOff
   
   'Turn On the Jackpot lights
    for each xx in ShotsColl
      SetLightColor xx, "red", 2
    Next
    for xx = 1 to 4
      LastShot(xx)=-1  ' Clear JackPot Shots
    Next
   
    debug.print "Start DEMON MB" 
    DMDGif "scene64.gif","","",slen(64)
    StopSound Track(cursong(CurPlayer))
  ' play speach
    PlaySound "audio662"
    cursong(CurPlayer)=9  ' Calling DR Love
    Song = track(CurSong(CurPlayer))
    PlayMusic Song
    debug.print "Show Saved Song #" & SaveSong
End Sub


Sub Start_LoveGun()
    SetLightColor i25, "white", 2    ' Flash Start Child
    if NOT LoveGunMode and NOT DemonMBMode then ' Not already in MB
      debug.print "Start LoveGun()"
      LoveGunMode=True
      DMDGif "scene50.gif","LOVE GUN","",slen(50) 
      EndMusic
      ' save the previous states
      SaveSong=cursong(CurPlayer)
      SaveShots()
      AddMultiball 3
      if Not bBallSaverActive then
        EnableBallSaver(9) ' Multiball BallSaver
      end if

      'Turn Off Demon Locks
      SetLightColor i82, "orange", 0
      SetLightColor i84, "orange", 0
      i90.State = 0
      bLockEnabled = False    ' Lose the locks 
      
      PlaySound "audio472" ' Love Gun

      for each xx in ShotsColl
        SetLightColor xx, "blue", 1
      Next
      for xx = 1 to 4
        LastShot(xx)=-1  ' Clear JackPot Shots
      Next
      cursong(CurPlayer)=10
      Song = track(CurSong(CurPlayer))
      PlayMusic Song
    end if
End Sub

Sub DemonKickBall()
    PlaySound "fx_kicker", 0, 1, 0.1
    debug.print "DemonKickball Destroyball"
    sw42.destroyball
    sw42top.createball
    sw42top.Kick 200, 8	
    PlaySound "audio65"
End Sub

Sub ResetJackpotLights()
    debug.print "ResetJackpotLights()"
    i118.State = 0
    RightRampLight.State = 0

    LoveGunMode=False ' Turn off the All Jackpot Mode
    DemonMBMode=False

    PlaySound "audio501" ' that was insane
    debug.print "Done MB .. start original track #" & SaveSong

    EndMusic
    cursong(CurPlayer)=SaveSong
    Song = track(CurSong(CurPlayer))
    PlayMusic Song

    LoadShots() ' should restore last state here
    SetLightColor i82, "orange", 2  ' Demon Lock Light
    SetLightColor i84, "orange", 2  ' Demon Lock Light

    if cursong(CurPlayer)=4 then ' Hotter Than Hell means they need to flash on Reset 
      if sw38.isdropped=1 then i35.state=LightStateOn else i35.state=2 end if
      if sw39.isdropped=1 then i36.state=LightStateOn else i36.state=2 end if
      if sw40.isdropped=1 then i37.state=LightStateOn else i37.state=2 end if
      if sw41.isdropped=1 then i38.state=LightStateOn else i38.state=2 end if
    else
      if sw38.isdropped=1 then i35.state=LightStateOn else i35.state=LightStateOff end if
      if sw39.isdropped=1 then i36.state=LightStateOn else i36.state=LightStateOff end if
      if sw40.isdropped=1 then i37.state=LightStateOn else i37.state=LightStateOff end if
      if sw41.isdropped=1 then i38.state=LightStateOn else i38.state=LightStateOff end if
    end if
End Sub

Sub CheckRockCity
  if i25.state=LightStateOn and i26.state=LightStateOn and i27.state=LightStateOn and i28.state=LightStateOn then
    SetLightColor i43, "yellow", 2   ' Rock City
  end if
End Sub

Sub CheckKissArmy
  if i25.state=LightStateBlinking or i26.state=LightStateBlinking or i27.state=LightStateBlinking or i28.state=LightStateBlinking then '1 flashing
    if i25.state <> LightStateOff and i26.state <> LightStateOff and i27.state <> LightStateOff and i28.state <> LightStateOff then
      PlaySound "audio560"
      SetLightColor i40, "yellow", 2 ' KissArmy
    end if
  end if
End Sub

' *********************************************************************
'                        Game Choices
' *********************************************************************

Sub NextCity
  CurCity(CurPlayer)=CurCity(CurPlayer)+1
  if CurCity(CurPlayer) > 15 then
    CurCity(CurPlayer)=1
  end if
  debug.print "Change of city to " & CurCity(CurPlayer)
  PlaySound "audio" & 429+CurCity(CurPlayer),0,1,0.25,0.25
End Sub

Sub NextSong
  debug.print "Stopping "  & Track(cursong(CurPlayer)) 
  EndMusic   
  if NewTrackTimer.enabled then NewTrackTimer.enabled=False:NewTrackTimer.enabled=True  ' auto plunge in NewTrack Mode
  cursong(CurPlayer) = cursong(CurPlayer) + 1
  if cursong(CurPlayer) > 8 then 
    cursong(CurPlayer) = 1
  end if
  If UseUDMD Then UltraDMD.SetScoreboardBackgroundImage CurSong(CurPlayer) & ".png",15,13:UltraDMD.Clear:OnScoreboardChanged()
  if cursong(CurPlayer) = SaveShot(CurPlayer,9) then
    LoadShots()
  else
    InitMode(cursong(CurPlayer))
  End if
  Song=Track(CurSong(CurPlayer))
  SongPause.interval=500
  SongPause.enabled=True ' Play song after short pause in case they change their mind

 End Sub

' ***********************************************************************************
' Save and Load the Shots from player to player
' ***********************************************************************************

Dim SaveShot(4,25), SaveCol(4,6), SaveColf(4,6)

' Save the status of the 6 shots for MB start
Sub SaveShots() 'after MB
  debug.print "SaveShots for player " & CurPlayer
  SaveShot(CurPlayer,1)=i58.state
  SaveShot(CurPlayer,2)=i68.state
  SaveShot(CurPlayer,3)=i73.state
  SaveShot(CurPlayer,4)=i79.state
  SaveShot(CurPlayer,5)=i92.state
  SaveShot(CurPlayer,6)=i98.state
  SaveShot(CurPlayer,7)=i6.state ' front row
  SaveShot(CurPlayer,8)=i44.state 'next track
  SaveShot(CurPlayer,9)=CurSong(CurPlayer)
  if i35.state=LightStateBlinking or i36.state=LightStateBlinking or i37.state=LightStateBlinking or i38.state=LightStateBlinking then
    debug.print "SaveShots .. blinking"
    SaveShot(CurPlayer,10)=1
  else
    SaveShot(CurPlayer,10)=0
  end if
  SaveCol(CurPlayer,1)=i58.color
  SaveCol(CurPlayer,2)=i68.color
  SaveCol(CurPlayer,3)=i73.color
  SaveCol(CurPlayer,4)=i79.color
  SaveCol(CurPlayer,5)=i92.color
  SaveCol(CurPlayer,6)=i98.color
  SaveColf(CurPlayer,1)=i58.colorfull
  SaveColf(CurPlayer,2)=i68.colorfull
  SaveColf(CurPlayer,3)=i73.colorfull
  SaveColf(CurPlayer,4)=i79.colorfull
  SaveColf(CurPlayer,5)=i92.colorfull
  SaveColf(CurPlayer,6)=i98.colorfull
End Sub

' Restore the status of the 6 shots after MB Ends
Sub LoadShots() ' after MB
  debug.print "Load Shots for Player " & CurPlayer
  i58.state=SaveShot(CurPlayer,1)
  i68.state=SaveShot(CurPlayer,2)
  i73.state=SaveShot(CurPlayer,3)
  i79.state=SaveShot(CurPlayer,4)
  i92.state=SaveShot(CurPlayer,5)
  i98.state=SaveShot(CurPlayer,6)
  i6.state =SaveShot(CurPlayer,7)
  i44.state=SaveShot(CurPlayer,8)

  debug.print "Loadshots - targets if 1" & SaveShot(CurPlayer,10)
  if SaveShot(CurPlayer,10)=1 then
    i35.state=LightStateBlinking 
    i36.state=LightStateBlinking 
    i37.state=LightStateBlinking 
    i38.state=LightStateBlinking 
  end if
  i58.color = SaveCol(CurPlayer,1)
  i68.color = SaveCol(CurPlayer,2)
  i73.color = SaveCol(CurPlayer,3)
  i79.color = SaveCol(CurPlayer,4)
  i92.color = SaveCol(CurPlayer,5)
  i98.color = SaveCol(CurPlayer,6)

  i58.colorfull = SaveColf(CurPlayer,1)
  i68.colorfull = SaveColf(CurPlayer,2)
  i73.colorfull = SaveColf(CurPlayer,3)
  i79.colorfull = SaveColf(CurPlayer,4)
  i92.colorfull = SaveColf(CurPlayer,5)
  i98.colorfull = SaveColf(CurPlayer,6)
End Sub

Sub SaveState 'at ball end   ' BSP, SuperTargets, Super Ramps, SuperBumpers, SuperSpinner
  if i41.state=LightStateBlinking then ' BSP
    SaveShot(CurPlayer,11)=1
  else
    SaveShot(CurPlayer,11)=0
  end if
  if i122.state=LightStateBlinking then ' 
    SaveShot(CurPlayer,12)=1
  else
    SaveShot(CurPlayer,12)=0
  end if
  if i95.state=LightStateBlinking then ' 
    SaveShot(CurPlayer,13)=1
  else
    SaveShot(CurPlayer,13)=0
  end if
  if i118.state=LightStateBlinking then ' BSP
    SaveShot(CurPlayer,14)=1
  else
    SaveShot(CurPlayer,14)=0
  end if
  if i61.state=LightStateBlinking then ' Bumpers
    SaveShot(CurPlayer,15)=1
  else
    SaveShot(CurPlayer,15)=0
  end if
  SaveShot(CurPlayer,16)=i90.state
End Sub

Sub LoadState ' At InitBall
  if SaveShot(CurPlayer,11)=1 then
    i41.state=LightStateBlinking 
    i42.state=LightStateBlinking 
  end if
  if SaveShot(CurPlayer,12)=1 then
    i122.state=LightStateBlinking
  end if
  if SaveShot(CurPlayer,13)=1 then
    i95.state=LightStateBlinking
  end if
  if SaveShot(CurPlayer,14)=1 then
    i118.state=LightStateBlinking
  end if
  if SaveShot(CurPlayer,15)=1 then
    i61.state=LightStateBlinking
  end if
  i90.state=SaveShot(CurPlayer,16)
End Sub



