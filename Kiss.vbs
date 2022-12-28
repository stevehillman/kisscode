'On the DOF website put Exxx
'101 Left Flipper
'102 Right Flipper
'103 left slingshot
'104 right slingshot
'105
'106
'107 Center Bumper
'108 RIGHT Bumper
'109 Left bumper
'110 Rear Elevator VUK
'111 Iron Throne VUK
'112 Dragon Kicker
'118
'119 Reset drop Targets
'120 AutoFire
'122 knocker
'123 ballrelease
'
' Lamps
' 140 Start button
' 141 Action button white
' 142 Action button yellow
' 143  " " red
' 144  " " purple
' 145  " " orange
' 146  " " green
' 147  " " blue
'
' 151 action button blink white
' 152 action button blink yellow
'  <etc>

Option Explicit
Randomize

Const bDebug = False

'***********TABLE VOLUME LEVELS ********* 
' [Value is from 0 to 1 where 1 is full volume. 
' NOTE: you can go past 1 to amplify sounds]
Const cVolBGMusic = 0.15  ' Volume for table background music  
Const cVolCallout = 0.9   ' Volume of voice callouts
Const cVolDef = 0.5		 ' Volume for GoT sound effects
Const cVolSfx = 0.2		 ' Volume for table "physical" Sound effects (non Fleep sounds)
Const cVolTable = 0.2   ' Used by Daphishbowl's Service menu
'*** Fleep ****
'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.5

'///////////////////// ---- Physics Options ---- /////////////////////
Const Rubberizer = 3				'Enhance micro bounces on flippers. 0 - disable, 1 - rothbauerw version, 2 - iaakki version, 3 - apophis version (default)
Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)

'////// Cabinet Options
Const bHaveLockbarButton = False    ' Set to true if you have a lockdown bar button. Will disable the flasher on the apron
Const bCabinetMode = False          ' Set to true to hide side rails
Const DMDMode = 1                   ' Use FlexDMD (currently the only option supported)
Const bUsePlungerForSternKey = False ' If true, use the plunger key for the Action/Select button. 

'///// Game of Thrones Options - most from the real ROM
Const bUseDragonFire = True             ' Whether to use dragon fire effect on the upper playfield. Looks cool but isn't true to real table

'/////// No User Configurable options beyond this point ////////

Const BallSize = 50     ' 50 is the normal size used in the core.vbs, VP kicker routines uses this value divided by 2
Const BallMass = 1      ' standard ball mass

' Load the core.vbs for supporting Subs and functions
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
    On Error Goto 0
End Sub

' Define any Constants
Const cGameName = "sternkiss"  ' Used by B2S for DOF
Const myGameName = "sternkiss" ' Used for FlexDMD, High score storage, etc
Const myVersion = "0.9"
Const MaxPlayers = 4          ' from 1 to 4 - don't change
Const MaxMultiplier = 5       ' limit playfield multiplier
Const MaxBonusMultiplier = 20 'limit Bonus multiplier
Dim BallsPerGame
Const MaxMultiballs = 6       ' max number of balls during multiballs

'*****************************************************************************************************
' FlexDMD constants
Const 	FlexDMD_RenderMode_DMD_GRAY_2 = 0, _
		FlexDMD_RenderMode_DMD_GRAY_4 = 1, _
		FlexDMD_RenderMode_DMD_RGB = 2, _
		FlexDMD_RenderMode_SEG_2x16Alpha = 3, _
		FlexDMD_RenderMode_SEG_2x20Alpha = 4, _
		FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num = 5, _
		FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num_4x1Num = 6, _
		FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num = 7, _
		FlexDMD_RenderMode_SEG_2x7Num_2x7Num_10x1Num = 8, _
		FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num_gen7 = 9, _
		FlexDMD_RenderMode_SEG_2x7Num10_2x7Num10_4x1Num = 10, _
		FlexDMD_RenderMode_SEG_2x6Num_2x6Num_4x1Num = 11, _
		FlexDMD_RenderMode_SEG_2x6Num10_2x6Num10_4x1Num = 12, _
		FlexDMD_RenderMode_SEG_4x7Num10 = 13, _
		FlexDMD_RenderMode_SEG_6x4Num_4x1Num = 14, _
		FlexDMD_RenderMode_SEG_2x7Num_4x1Num_1x16Alpha = 15, _
		FlexDMD_RenderMode_SEG_1x16Alpha_1x16Num_1x7Num = 16

Const 	FlexDMD_Align_TopLeft = 0, _
		FlexDMD_Align_Top = 1, _
		FlexDMD_Align_TopRight = 2, _
		FlexDMD_Align_Left = 3, _
		FlexDMD_Align_Center = 4, _
		FlexDMD_Align_Right = 5, _
		FlexDMD_Align_BottomLeft = 6, _
		FlexDMD_Align_Bottom = 7, _
		FlexDMD_Align_BottomRight = 8

Const   UltraDMD_Animation_None = 14
Const   FDsep = ","     ' The latest, not yet published, version has switched to "|"
Dim FlexDMD
'********* End FlexDMD **************


' Define Global Variables that aren't game-specific 
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim PuPlayer                ' Not currently used
Dim plungerIM 'used mostly as an autofire plunger
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim Score(4)
Dim HighScore(16)
Dim HighScoreName(16)
Dim ReplayScore
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)     ' I don't think we need an array here - EBs can't be carried over
Dim ReplayAwarded(4)
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim BallsOnPlayfield        ' Active balls on playfield, including real locked ones
Dim RealBallsInLock			' These are actually locked on the table
Dim BallsInLock             ' Number of balls the current player has locked
Dim BallSearchCnt
Dim LastSwitchHit
Dim LightSaveState(100,4)   ' Array for saving state of lights. We save state, colour, fade to restore after Sequences (sequences only save state)
Dim LightNames(300)         ' Saves the mapping of light names (li<xxx>) to index into LightSaveState
Dim TotalPlayfieldLights

' flags
Dim bMultiBallMode
Dim bAutoPlunger
Dim bAutoPlunged
Dim bJustPlunged
Dim bBallSaved
Dim bInstantInfo
Dim bAttractMode
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bTableReady
Dim	bUseFlexDMD
Dim	bUsePUPDMD
Dim	bPupStarted
Dim bBonusHeld
Dim bJustStarted            ' Not sure what this is used for
Dim bShowMatch
Dim bSkillShotReady
Dim bMysteryAwardActive     ' Table specific - used by Mystery Award for flipper control
Dim WHCFlipperFrozen        ' table-specific - used to "freeze" the left flipper during Winter-Has-Come


' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    LoadEM
    Dim i
    Randomize

	vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, bumper2, bumper3, LeftSlingshot, RightSlingshot)
	
	bTableReady=False
	bUseFlexDMD=False
	bUsePUPDMD=False
	bPupStarted=False
	if DMDMode = 1 then 
		bUseFlexDMD= True
		set PuPlayer = New PinupNULL
	elseif DMDMode = 2 Then
		bUsePUPDMD = True
	Else 
		set PuPlayer = New PinupNULL
	End if 

	if b2son then 
		Controller.B2SSetData 1,1
		Controller.B2SSetData 2,1
		Controller.B2SSetData 3,1
		Controller.B2SSetData 4,1
		Controller.B2SSetData 5,1
		Controller.B2SSetData 6,1
		Controller.B2SSetData 7,1
		Controller.B2SSetData 8,1
	End if 

    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 45 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 1.5
        .InitExitSnd SoundFX("fx_autoplunger", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    If bCabinetMode Then lrail.Visible = 0 : rrail.Visible = 0

    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    ' Creat the array of light name indexes
    SavePlayfieldLightNames

    'load saved values, highscore, names, jackpot
    Loadhs
    DMDSettingsInit() ' Init the Service Menu

    ' initalise the DMD display
    DMD_Init

    ' initialse any other flags
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    bGameInPlay = False
	'bGameInPlayHidden = False 
	bShowMatch = False
	'bCreatedBall = False
    bAutoPlunger = False
	'bAutoPlunged = False
    BallsOnPlayfield = 0
	RealBallsInLock=0
    BallsInLock = 0
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJustStarted = True
    bInstantInfo = False
    bMysteryAwardActive = False

	
	bTableReady = True
	' set any lights for the attract mode
    GiOff
    TurnOffPlayfieldLights
    StartAttractMode

    ' Start the RealTime timer
    RealTime.Enabled = 1

    ' Load table color
    LoadLut

End Sub

Private Function BigMod(Value1, Value2)
    BigMod = Value1 - (Int(Value1 / Value2) * Value2)
End Function

Sub Table1_Exit()
    If Not IsEmpty(FlexDMD) Then
        If Not FlexDMD is Nothing Then
            FlexDMD.Show = False
            FlexDMD.Run = False
            FlexDMD = NULL
        End If
    End If
    If B2SOn Then Controller.Stop
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

    If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft()
    If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight()
    If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter()

    If keycode = LeftMagnaSave Then 
        bLutActive = True
        if bDebug And bGameInPlay And bBallInPlungerLane=False Then StartMidnightMadnessMBIntro
    End if
    If keycode = RightMagnaSave Then
        If bLutActive Then
            NxtLUT
            Exit Sub
        End If
    End If

    If Keycode = AddCreditKey Then
        Credits = Credits + 1
        if DMDStd(kDMDStd_FreePlay) = False Then DOF 140, DOFOn
        If(Tilted = False)Then GameAddCredit
    End If

    If keycode = PlungerKey Then Plunger.Pullback : SoundPlungerPull()

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    If bMysteryAwardActive Then
        UpdateMysteryAward(keycode)
        Exit Sub
    End If

    StartServiceMenu keycode

    ' Normal flipper action

    If bGameInPlay AND NOT Tilted Then

        If bMadnessMB = 1 Then Exit Sub

        If keycode = LeftTiltKey Then CheckTilt 'only check the tilt during game
        If keycode = RightTiltKey Then CheckTilt
        If keycode = CenterTiltKey Then CheckTilt

        If keycode = LeftFlipperKey Then 
            If WHCFlipperFrozen = False Then FlipperActivate LeftFlipper, LFPress : FlipperActivate LeftUFlipper, ULFPress : SolLFlipper 1
            InstantInfoTimer.Enabled = True
            RotateLaneLights 1
            If InstantInfoTimer.UserValue = 0 Then 
                InstantInfoTimer.UserValue = keycode ' Record which key triggered the InstantInfo
            ElseIf InstantInfoTimer.UserValue <> keycode And bInstantInfo Then
                InfoPage = InfoPage + 1:InstantInfo
            End If
        ElseIf keycode = RightFlipperKey Then 
            FlipperActivate RightFlipper, RFPress :  FlipperActivate RightUFlipper, URFPress : SolRFlipper 1
            InstantInfoTimer.Enabled = True
            RotateLaneLights 0
            If InstantInfoTimer.UserValue = 0 Then 
                InstantInfoTimer.UserValue = keycode ' Record which key triggered the InstantInfo
            ElseIf InstantInfoTimer.UserValue <> keycode And bInstantInfo Then
                InfoPage = InfoPage + 1:InstantInfo
            End if
        End If
        '  Action Button, Start Mode, fire ball
		If keycode = RightMagnaSave or keycode = LockBarKey or _  
			(keycode = PlungerKey and bUsePlungerForSternKey) Then

			if bAutoPlunger=False and bBallInPlungerLane = True then	' Auto fire ball with stern key
				plungerIM.Strength = 60
				'plungerIM.InitImpulseP swplunger, 60, 0		' Change impulse power while we are here
				PlungerIM.AutoFire
				DOF 112, DOFPulse
				plungerIM.Strength = 45
				'plungerIM.InitImpulseP swplunger, 45, 1.1
			Else	
				CheckActionButton				
			end if
		End if

        If CheckLocalKeydown(keycode) Then Exit Sub

        If keycode = StartGameKey Then
            If LFPress=1 And bBallInPlungerLane and (DMDStd(kDMDStd_LeftStartReset)=1 Or (DMDStd(kDMDStd_LeftStartReset)=2 And DMDStd(kDMDStd_FreePlay) = True)) Then  ' Quick Restart
                Dim b
                PlungerIM.AutoFire
                For each b in GetBalls : b.DestroyBall : next
                bGameInPLay = False	
                bShowMatch = False
                tmrBallSearch.Enabled = False
                
                ' ensure that the flippers are down
                SolLFlipper 0
                SolRFlipper 0
                GiOff
                StartAttractMode
                Exit Sub
            End If
            If((PlayersPlayingGame <MaxPlayers)AND(bOnTheFirstBall = True))Then

                If(DMDStd(kDMDStd_FreePlay) = True)Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                    DMD "_", CL(1, PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, False, ""
                    ScoreScene = Empty
                Else
                    If(Credits> 0)then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
                        DMD "_", CL(1, PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 1000, False, ""
                        ScoreScene = Empty
                        If Credits <1 And DMDStd(kDMDStd_FreePlay) = False Then DOF 140, DOFOff
                    Else
                        ' Not Enough Credits to start a game.
                        DisplayDMDText "CREDITS 0", "INSERT COINS", 1000
                    End If
                End If
            End If
        End If
    Else ' If (GameInPlay)

        If keycode = StartGameKey Then
            If(DMDStd(kDMDStd_FreePlay) = True)Then
                If(BallsOnPlayfield = 0)Then
                    ResetForNewGame()
                End If
            Else
                If(Credits> 0)Then
                    If(BallsOnPlayfield = 0)Then
                        Credits = Credits - 1
                        If Credits <1 And DMDStd(kDMDStd_FreePlay) = False Then DOF 140, DOFOff
                        ResetForNewGame()
                    End If
                Else
                    ' Not Enough Credits to start a game.
                    ' This is Table-specific:
                    PlaySoundVol "gotfx-nocredits",VolDef
                    If bAttractMode Then
                        tmrAttractModeScene.UserValue = 5
                        tmrAttractModeScene_Timer
                    Else
                        DisplayDMDText "CREDITS 0","INSERT COINS",1000
                    End If
                End If
            End If
        ElseIf bAttractMode And (keycode = LeftFlipperKey Or keycode = RightFlipperKey)  Then
            doFlipperSpeech keycode
        End If
    End If ' If (GameInPlay)
End Sub

Sub Table1_KeyUp(ByVal keycode)

    If keycode = LeftMagnaSave Then bLutActive = False

    If keycode = PlungerKey Then 
        Plunger.Fire
        If bBallInPlungerLane Then SoundPlungerReleaseBall() Else SoundPlungerReleaseNoBall()
    End If

    If hsbModeActive Then
        Exit Sub
    End If

    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then
            FlipperDeactivate LeftFlipper, LFPress : FlipperDeactivate LeftUFlipper, ULFPress : SolLFlipper 0
            If InstantInfoTimer.UserValue = keycode Then
                InstantInfoTimer.UserValue = 0
                InstantInfoTimer.Enabled = False
                InstantInfoTimer.Interval = 10000
                If bInstantInfo Then
                    tmrDMDUpdate.Enabled = true
                    DMDFlush : AddScore 0
                    bInstantInfo = False
                End If
            End If
        ElseIf keycode = RightFlipperKey Then
            FlipperDeactivate RightFlipper, RFPress : FlipperDeactivate RightUFlipper, URFPress : SolRFlipper 0
            If InstantInfoTimer.UserValue = keycode Then
                InstantInfoTimer.UserValue = 0
                InstantInfoTimer.Enabled = False
                InstantInfoTimer.Interval = 10000
                If bInstantInfo Then
                    tmrDMDUpdate.Enabled = true
                    DMDFlush : AddScore 0
                    bInstantInfo = False
                End If
            End If
        End If
    End If
End Sub

Dim InfoPage
Sub InstantInfoTimer_Timer
    InstantInfoTimer.Enabled = False
    If  hsbModeActive = False And bMultiBallMode = False Then
        bInstantInfo = True
        tmrDMDUpdate.Enabled = False
        DMDFlush
        InfoPage = 0
        InstantInfo
    End If
End Sub

'********************
'     Flippers
'********************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
        'PlaySoundAt SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), LeftFlipper

        DOF 101, DOFOn
        If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then 
            RandomSoundReflipUpLeft LeftFlipper
            RandomSoundReflipUpLeft LeftUFlipper
        Else 
            SoundFlipperUpAttackLeft LeftFlipper
            RandomSoundFlipperUpLeft LeftFlipper
            SoundFlipperUpAttackLeft LeftUFlipper
            RandomSoundFlipperUpLeft LeftUFlipper
        End If  
        LF.Fire
        ULF.Fire
        'LeftFlipper.EOSTorque = 0.75:LeftFlipper.RotateToEnd
        'LeftUFlipper.EOSTorque = 0.75:LeftUFlipper.RotateToEnd

    Else
        'PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
        DOF 101, DOFOff
		If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
			RandomSoundFlipperDownLeft LeftFlipper
            RandomSoundFlipperDownLeft LeftUFlipper
		End If
		FlipperLeftHitParm = FlipperUpSoundLevel
        LeftFlipper.EOSTorque = 0.2:LeftFlipper.RotateToStart
        LeftUFlipper.EOSTorque = 0.2:LeftUFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        'PlaySoundAt SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
        DOF 102, DOFOn
        If rightflipper.currentangle > leftflipper.endangle - ReflipAngle Then 
            RandomSoundReflipUpRight RightFlipper
            RandomSoundReflipUpRight RightUFlipper
        Else 
            SoundFlipperUpAttackRight RightFlipper
            RandomSoundFlipperUpRight RightFlipper
            SoundFlipperUpAttackRight RightUFlipper
            RandomSoundFlipperUpRight RightUFlipper
        End If  
        RF.Fire
        URF.Fire
        'RightFlipper.EOSTorque = 0.75:RightFlipper.RotateToEnd
        'RightUFlipper.EOSTorque = 0.75:RightUFlipper.RotateToEnd
    Else
        'PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
        DOF 102, DOFOff
        If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
			RandomSoundFlipperDownRight RightFlipper
            RandomSoundFlipperDownRight RightUFlipper
		End If	
		FlipperRightHitParm = FlipperUpSoundLevel
        RightFlipper.EOSTorque = 0.2:RightFlipper.RotateToStart
        RightUFlipper.EOSTorque = 0.2:RightUFlipper.RotateToStart
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    'PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
    LeftFlipperCollide parm
    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
End Sub

Sub RightFlipper_Collide(parm)
    'PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
    RightFlipperCollide parm
    CheckLiveCatch Activeball, RightFlipper, RFCount, parm
End Sub

Sub LeftUFlipper_Collide(parm)
    'PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
    LeftFlipperCollide parm
End Sub

Sub RightUFlipper_Collide(parm)
    'PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
    RightFlipperCollide parm
End Sub

'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                  'Called when table is nudged
    Tilt = Tilt + TiltSensitivity              'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt> TiltSensitivity)AND(Tilt <15)Then 'show a warning
        DMD "_", CL(1, "CAREFUL"), "_", eNone, eBlinkFast, eNone, 1500, True, "fx-tiltwarning"
        ' Table specific:
        PlaySoundVol "say-tiltwarning"&RndNbr(12),VolCallout
    End if
    If Tilt> 15 Then 'If more than 15 then TILT the table
        Tilted = True
        'display Tilt
        DMDFlush
        DMD "", "TILT", "", eNone, eNone, eNone, 5000, True, "fx-tilt"
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt> 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        if bDebug Then Wall075.collidable=0
        If Song <> "" Then StopSound Song : Song=""
        'turn off GI and turn off all the lights
        GiOff
        LightSeqTilt.Play SeqAllOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        LeftUFlipper.RotateToStart
        RightUFlipper.RotateToStart
        Bumper1.Threshold = 100
        Bumper2.Threshold = 100
        Bumper3.Threshold = 100
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        Bumper1.Threshold = 1
        Bumper2.Threshold = 1
        Bumper3.Threshold = 1
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        'clean up the buffer display
        DMDFlush
        Wall075.collidable=bDebug
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained then..
    If(BallsOnPlayfield-RealBallsInLock = 0)Then
        ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
        vpmtimer.Addtimer 2000, "EndOfBall() '"
        TiltRecoveryTimer.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub



Sub PlaySoundVol(soundname, Volume)
  PlaySound soundname, 1, Volume
End Sub

' Play an already playing sound (starts anew if not playing). Restart=whether to restart the sound. Presumably 0 = just let it keep playing
Sub PlayExistingSoundVol(soundname, Volume, Restart)
  PlaySound soundname, 1, Volume, 0, 0, 0, 1, Restart
End Sub

Sub PlaySoundLoopVol(soundname, Volume)
  PlaySound soundname, -1, Volume
End Sub

'********************
' Music as wav sounds
'********************

Dim Song
Song = ""

Sub ThemeSong
	PlaySong("Song-1")
	SongNum=1
End Sub

Sub RotateSong()
'debug.print "Rotate " & SongNum
	PlaySong "Song-" & SongNum
	SongNum=SongNum+1
	if (SongNum>=4) then SongNum=1
End Sub


dim bPlayPaused
bPlayPaused = False
Sub PlaySong(name)
'debug.print "PlaySong " & name & " " & song
	dim PlayLength
	if bUsePUPDMD then 			' Use Pup if we have it so we can pause the music
'		PlaySongPup(name)
		exit sub
	End If 
	StopSound "m_wait"
	StopSound Song	' Stop the old song
	if name <> "" then Song = name
	PlayLength = -1
	If Song = "m_end" Then PlayLength = 0
	bPlayPaused=False
	PlaySound Song, PlayLength, VolBGMusic 'this last number is the volume, from 0 to 1
End Sub


Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'********* Fleep's mech sounds ***********

'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                                                                                                                'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                                                                                                        'volume level; range [0, 1]
NudgeRightSoundLevel = 1                                                                                                'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                                                                                                'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                                                                                                'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                                                                                        'volume level; range [0, 1]
PlungerPullSoundLevel = 1                                                                                                'volume level; range [0, 1]
RollingSoundFactor = 1.1/5                

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                                                           'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                                                                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                                                        'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                                                      'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
SlingshotSoundLevel = 0.95                                                                                                'volume level; range [0, 1]
BumperSoundFactor = 4.25                                                                                                'volume multiplier; must not be zero
KnockerSoundLevel = 4 '1/VolumeDial                                                                                                         'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                                                                        'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                                                                                        'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                                                                                        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                                                                                'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                                                                        'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                                                                        'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                                                                        'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                                                                                        'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                                                                                                        'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                                                                                        'volume multiplier; must not be zero
DTSoundLevel = 0.25                                                                                                                'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                                                      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor 

DrainSoundLevel = 0.8                                                                                                                'volume level; range [0, 1]
BallReleaseSoundLevel = 1                                                                                                'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                                                                        'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                                                                                'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                                                                                                        'volume multiplier; must not be zero


'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
    PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
    Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
    Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
    PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


' *********************************************************************
'                     Fleep  Supporting Ball & Sound Functions
' *********************************************************************

Dim tablewidth, tableheight : tablewidth = table1.width : tableheight = table1.height

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
    tmp = 1
    On Error Resume Next
	tmp = tableobj.y * 2 / tableheight-1
    On Error Goto 0

	if tmp > 7000 Then
		tmp = 7000
	elseif tmp < -7000 Then
		tmp = -7000
	end if

    	If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
	Else
		AudioFade = Csng(-((- tmp) ^10) )
	End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
	Dim tmp
    tmp = 1
    On Error Resume Next
	tmp = tableobj.x * 2 / tablewidth-1
    On Error Goto 0

	if tmp > 7000 Then
		tmp = 7000
	elseif tmp < -7000 Then
		tmp = -7000
	end if

	If tmp > 0 Then
		AudioPan = Csng(tmp ^10)
	Else
		AudioPan = Csng(-((- tmp) ^10) )
	End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
    Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
    VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
    PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
    PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
    PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
    PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
    PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub


Sub SoundPlungerPull()
    PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
    PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger        
End Sub

Sub SoundPlungerReleaseNoBall()
    PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub


'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////
Sub KnockerSolenoid()
    PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
    PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
    PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
    PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
    PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
    PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
    PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
    PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
    FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
    PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
    FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
    PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
    PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
    PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
    PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
    PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
    PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
    PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
    FlipperLeftHitParm = parm/10
    If FlipperLeftHitParm > 1 Then FlipperLeftHitParm = 1
    FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
    RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
    FlipperRightHitParm = parm/10
    If FlipperRightHitParm > 1 Then FlipperRightHitParm = 1
    FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
    RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
    PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
        PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
        RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 5 then                
                 RandomSoundRubberStrong 1
        End if
        If finalspeed <= 5 then
                 RandomSoundRubberWeak()
         End If        
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
    Select Case Int(Rnd*10)+1
        Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
        Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
        Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
        Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
        Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
        Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
        Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
        Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
        Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
        Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
    End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
    PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 5 then RandomSoundRubberStrong 1 
    If finalspeed <= 5 then RandomSoundRubberWeak()
End Sub

Sub RandomSoundWall()
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then 
        Select Case Int(Rnd*5)+1
                Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
                Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
                Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
                Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
                Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
        End Select
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        Select Case Int(Rnd*4)+1
            Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
            Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
            Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
            Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
        End Select
    End If
    If finalspeed < 6 Then
        Select Case Int(Rnd*3)+1
            Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
            Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
            Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
        End Select
    End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
    PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
    RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
    RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////
Sub RandomSoundBottomArchBallGuide()
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then 
        PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        Select Case Int(Rnd*2)+1
            Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
            Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
        End Select
    End If
    If finalspeed < 6 Then
        Select Case Int(Rnd*2)+1
            Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
            Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
        End Select
    End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
    PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
    If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
        RandomSoundBottomArchBallGuideHardHit()
    Else
        RandomSoundBottomArchBallGuide
    End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then 
        Select Case Int(Rnd*2)+1
            Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
            Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
        End Select
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
    End If
    If finalspeed < 6 Then
        PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
    End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
    PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()                
    PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 10 then
        RandomSoundTargetHitStrong()
        RandomSoundBallBouncePlayfieldSoft Activeball
    Else 
        RandomSoundTargetHitWeak()
    End If        
End Sub

Sub Targets_Hit (idx)
    PlayTargetSound        
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
    Select Case Int(Rnd*9)+1
        Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
        Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
        Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
        Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
        Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
        Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
        Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
        Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
        Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
    End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
    PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
    Select Case Int(Rnd*5)+1
        Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
        Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
        Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
        Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
        Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()                        
    PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
    PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
    SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)        
    SoundPlayfieldGate        
End Sub        

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
    PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
    PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
    If Activeball.velx > 1 Then SoundPlayfieldGate
    StopSound "Arch_L1"
    StopSound "Arch_L2"
    StopSound "Arch_L3"
    StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
    If activeball.velx < -8 Then
        RandomSoundRightArch
    End If
End Sub

Sub Arch2_hit()
    If Activeball.velx < 1 Then SoundPlayfieldGate
    StopSound "Arch_R1"
    StopSound "Arch_R2"
    StopSound "Arch_R3"
    StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
    If activeball.velx > 10 Then
        RandomSoundLeftArch
    End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock(saucer)
    PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, saucer
End Sub

Sub SoundSaucerKick(scenario, saucer)
    Select Case scenario
        Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
        Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
    End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
    Dim snd
    Select Case Int(Rnd*7)+1
        Case 1 : snd = "Ball_Collide_1"
        Case 2 : snd = "Ball_Collide_2"
        Case 3 : snd = "Ball_Collide_3"
        Case 4 : snd = "Ball_Collide_4"
        Case 5 : snd = "Ball_Collide_5"
        Case 6 : snd = "Ball_Collide_6"
        Case 7 : snd = "Ball_Collide_7"
    End Select

    PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'/////////////////////////////////////////////////////////////////
'                                        End Mechanical Sounds
'/////////////////////////////////////////////////////////////////


'******************************************************
'		BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 10 ' total number of balls
Const maxvel = 50
ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
	Dim i
	For i = 0 to tnob
		rolling(i) = False
	Next
End Sub

Sub RollingUpdate
	Dim BOT, b
	BOT = GetBalls

	' stop the sound of deleted balls
	For b = UBound(BOT) + 1 to tnob
		rolling(b) = False
		StopSound("BallRoll_" & b)
        aBallShadow(b).Y = 3000
	Next

	' exit the sub if no balls on the table
	If UBound(BOT) = -1 Then Exit Sub

	' play the rolling sound for each ball

	For b = 0 to UBound(BOT)
		If BallVel(BOT(b)) > 1 Then ' AND BOT(b).z < 30 Then
			rolling(b) = True
			PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

		Else
			If rolling(b) = True Then
				StopSound("BallRoll_" & b)
				rolling(b) = False
			End If
		End If

        ' JPs ball shadows
        aBallShadow(b).X = BOT(b).X
        aBallShadow(b).Y = BOT(b).Y
        aBallShadow(b).Height = BOT(b).Z -24

		'***Ball Drop Sounds***
		If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
			If DropCount(b) >= 5 Then
				DropCount(b) = 0
				If BOT(b).velz > -7 Then
					RandomSoundBallBouncePlayfieldSoft BOT(b)
				Else
					RandomSoundBallBouncePlayfieldHard BOT(b)
				End If				
			End If
		End If
		If DropCount(b) < 5 Then
			DropCount(b) = DropCount(b) + 1
		End If
	Next
End Sub

'******************************************************
'		TRACK ALL BALL VELOCITIES
' 		FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
	public ballvel, ballvelx, ballvely

	Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub 

	Public Sub Update()	'tracks in-ball-velocity
		dim str, b, AllBalls, highestID : allBalls = getballs

		for each b in allballs
			if b.id >= HighestID then highestID = b.id
		Next

		if uBound(ballvel) < highestID then redim ballvel(highestID)	'set bounds
		if uBound(ballvelx) < highestID then redim ballvelx(highestID)	'set bounds
		if uBound(ballvely) < highestID then redim ballvely(highestID)	'set bounds

		for each b in allballs
			ballvel(b.id) = BallSpeed(b)
			ballvelx(b.id) = b.velx
			ballvely(b.id) = b.vely
		Next
	End Sub
End Class

Sub RDampen_Timer()
	Cor.Update
End Sub


'********************
'     FlippersPol
'********************


dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity
dim ULF : Set ULF = New FlipperPolarity
dim URF : Set URF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
	dim x, a : a = Array(LF, RF, ULF, URF)
	for each x in a
		x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
		x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
		x.enabled = True
		x.TimeDelay = 60
	Next

	AddPt "Polarity", 0, 0, 0
	AddPt "Polarity", 1, 0.05, -5.5
	AddPt "Polarity", 2, 0.4, -5.5
	AddPt "Polarity", 3, 0.6, -5.0
	AddPt "Polarity", 4, 0.65, -4.5
	AddPt "Polarity", 5, 0.7, -4.0
	AddPt "Polarity", 6, 0.75, -3.5
	AddPt "Polarity", 7, 0.8, -3.0
	AddPt "Polarity", 8, 0.85, -2.5
	AddPt "Polarity", 9, 0.9,-2.0
	AddPt "Polarity", 10, 0.95, -1.5
	AddPt "Polarity", 11, 1, -1.0
	AddPt "Polarity", 12, 1.05, -0.5
	AddPt "Polarity", 13, 1.1, 0
	AddPt "Polarity", 14, 1.3, 0

	Addpt "Velocity", 0, 0,         1
	Addpt "Velocity", 1, 0.16, 1.06
	Addpt "Velocity", 2, 0.41,         1.05
	Addpt "Velocity", 3, 0.53,         1'0.982
	Addpt "Velocity", 4, 0.702, 0.968
	Addpt "Velocity", 5, 0.95,  0.968
	Addpt "Velocity", 6, 1.03,         0.945

	LF.Object = LeftFlipper        
	LF.EndPoint = EndPointLp
	RF.Object = RightFlipper
	RF.EndPoint = EndPointRp
    ULF.Object = LeftUFlipper
    ULF.Endpoint = EndPointULp
    URF.Object = RightUFlipper
    URF.Endpoint = EndPointURp
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

Sub TriggerULF_Hit() : ULF.Addball activeball : End Sub
Sub TriggerULF_UnHit() : ULF.PolarityCorrect activeball : End Sub
Sub TriggerURF_Hit() : URF.Addball activeball : End Sub
Sub TriggerURF_UnHit() : URF.PolarityCorrect activeball : End Sub

'******************************************************
'           FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
        dim a : a = Array(LF, RF)
        dim x : for each x in a
                x.addpoint aStr, idx, aX, aY
        Next
End Sub

Class FlipperPolarity
	Public DebugOn, Enabled
	Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
	Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
	private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
	Private Balls(20), balldata(20)
        
	dim PolarityIn, PolarityOut
	dim VelocityIn, VelocityOut
	dim YcoefIn, YcoefOut
	Public Sub Class_Initialize 
		redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
		Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next 
	End Sub
        
	Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
	Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
	Public Property Get StartPoint : StartPoint = FlipperStart : End Property
	Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
	Public Property Get EndPoint : EndPoint = FlipperEnd : End Property        
	Public Property Get EndPointY: EndPointY = FlipperEndY : End Property
        
	Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out) 
		Select Case aChooseArray
			case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
			Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
			Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
		End Select
		if gametime > 100 then Report aChooseArray
	End Sub 

	Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
		if not DebugOn then exit sub
		dim a1, a2 : Select Case aChooseArray
			case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
			Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
			Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut 
			case else :tbpl.text = "wrong string" : exit sub
		End Select
		dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
		tbpl.text = str
	End Sub
        
	Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

	Private Sub RemoveBall(aBall)
		dim x : for x = 0 to uBound(balls)
		if TypeName(balls(x) ) = "IBall" then 
			if aBall.ID = Balls(x).ID Then
				balls(x) = Empty
				Balldata(x).Reset
			End If
		End If
		Next
	End Sub
        
	Public Sub Fire() 
		Flipper.RotateToEnd
		processballs
	End Sub

	Public Property Get Pos 'returns % position a ball. For debug stuff.
		dim x : for x = 0 to uBound(balls)
			if not IsEmpty(balls(x) ) then
				pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
			End If
		Next                
	End Property

	Public Sub ProcessBalls() 'save data of balls in flipper range
		FlipAt = GameTime
		dim x : for x = 0 to uBound(balls)
			if not IsEmpty(balls(x) ) then
				balldata(x).Data = balls(x)
			End If
		Next
		PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
		PartialFlipCoef = abs(PartialFlipCoef-1)
	End Sub
	Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect
        
	Public Sub PolarityCorrect(aBall)
		if FlipperOn() then 
			dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

			'y safety Exit
			if aBall.VelY > -8 then 'ball going down
				RemoveBall aBall
				exit Sub
			end if

			'Find balldata. BallPos = % on Flipper
			for x = 0 to uBound(Balls)
				if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then 
					idx = x
					BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
					if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
				end if
			Next

			If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
				BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
				if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
 			End If

			'Velocity correction
			if not IsEmpty(VelocityIn(0) ) then
				Dim VelCoef
				VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

				if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

				if Enabled then aBall.Velx = aBall.Velx*VelCoef
				if Enabled then aBall.Vely = aBall.Vely*VelCoef
			End If

                        'Polarity Correction (optional now)
			if not IsEmpty(PolarityIn(0) ) then
				If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
				dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
        
				if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
			End If
		End If
		RemoveBall aBall
	End Sub
End Class


'******************************************************
'                FLIPPER POLARITY AND RUBBER DAMPENER
'                        SUPPORTING FUNCTIONS 
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
	dim x, aCount : aCount = 0
	redim a(uBound(aArray) )
	for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
		if not IsEmpty(aArray(x) ) Then
			if IsObject(aArray(x)) then 
				Set a(aCount) = aArray(x)
			Else
				a(aCount) = aArray(x)
			End If
			aCount = aCount + 1
		End If
	Next
	if offset < 0 then offset = 0
	redim aArray(aCount-1+offset)        'Resize original array
	for x = 0 to aCount-1                'set objects back into original array
		if IsObject(a(x)) then 
			Set aArray(x) = a(x)
		Else
			aArray(x) = a(x)
		End If
	Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
	ShuffleArray aArray1, offset
	ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
	dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
	Y = M*x+b
	PSlope = Y
End Function

' Used for flipper correction
Class spoofball 
	Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius 
	Public Property Let Data(aBall)
		With aBall
			x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
			id = .ID : mass = .mass : radius = .radius
		end with
	End Property
	Public Sub Reset()
		x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty 
		id = Empty : mass = Empty : radius = Empty
	End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
	dim y 'Y output
	dim L 'Line
	dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
		if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
	Next
	if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
	Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

	if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
	if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

	LinearEnvelope = Y
End Function



'******************************************************
'                        FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
	FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
	FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
	FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
	FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle

    FlipperTricks LeftUFlipper, ULFPress, ULFCount, ULFEndAngle, ULFState
	FlipperTricks RightUFlipper, URFPress, URFCount, URFEndAngle, URFState
	FlipperNudge RightUFlipper, URFEndAngle, URFEOSNudge, LeftUFlipper, ULFEndAngle
	FlipperNudge LeftUFlipper, ULFEndAngle, ULFEOSNudge,  RightUFlipper, URFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge
Dim ULFEOSNudge, URFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
	Dim b

	If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
		EOSNudge1 = 1
        Dim gBOT : gBOT = GetBalls
		'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
		If Flipper2.currentangle = EndAngle2 Then 
			For b = 0 to Ubound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
					'Debug.Print "ball in flip1. exit"
					exit Sub
				end If
			Next
			For b = 0 to Ubound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
					gBOT(b).velx = gBOT(b).velx / 1.3
					gBOT(b).vely = gBOT(b).vely - 0.5
				end If
			Next
		End If
	Else 
		If Flipper1.currentangle <> EndAngle1 then 
			EOSNudge1 = 0
		end if
	End If
End Sub

'*****************
' Maths
'*****************
Dim PI: PI = 4*Atn(1)

Function dSin(degrees)
	dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
	dcos = cos(degrees * Pi/180)
End Function

Function Atn2(dy, dx)
	If dx > 0 Then
		Atn2 = Atn(dy / dx)
	ElseIf dx < 0 Then
		If dy = 0 Then 
			Atn2 = pi
		Else
			Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
		end if
	ElseIf dx = 0 Then
		if dy = 0 Then
			Atn2 = 0
		else
			Atn2 = Sgn(dy) * pi / 2
		end if
	End If
End Function

Function max(a,b)
	if a > b then 
		max = a
	Else
		max = b
	end if
end Function

Function min(a,b)
	if a > b then 
		min = b
	Else
		min = a
	end if
end Function


'*************************************************
' Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
	Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
	DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
	Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
	AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
	DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
	Dim DiffAngle
	DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
	If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

	If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
		FlipperTrigger = True
	Else
		FlipperTrigger = False
	End If        
End Function


'*************************************************
' End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

dim ULFPress, URFPress, ULFCount, URFCount
dim ULFState, URFState
dim URFEndAngle, ULFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8 
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode 
	Case 0:
		SOSRampup = 2.5
	Case 1:
		SOSRampup = 6
	Case 2:
		SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

ULFEndAngle = LeftUflipper.endangle
URFEndAngle = RightUFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
	FlipperPress = 1
	Flipper.Elasticity = FElasticity

	Flipper.eostorque = EOST         
	Flipper.eostorqueangle = EOSA         
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
	FlipperPress = 0
	Flipper.eostorqueangle = EOSA
	Flipper.eostorque = EOST*EOSReturn/FReturn

        
	If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
		Dim b
        Dim gBOT : gBOT = GetBalls         
		For b = 0 to UBound(gBOT)
			If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
				If gBOT(b).vely >= -0.4 Then gBOT(b).vely = -0.4
			End If
		Next
	End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState) 
	Dim Dir
	Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

	If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
		If FState <> 1 Then
			Flipper.rampup = SOSRampup 
			Flipper.endangle = FEndAngle - 3*Dir
			Flipper.Elasticity = FElasticity * SOSEM
			FCount = 0 
			FState = 1
		End If
	ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
		if FCount = 0 Then FCount = GameTime

		If FState <> 2 Then
			Flipper.eostorqueangle = EOSAnew
			Flipper.eostorque = EOSTnew
			Flipper.rampup = EOSRampup                        
			Flipper.endangle = FEndAngle
			FState = 2
		End If
	Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then 
		If FState <> 3 Then
			Flipper.eostorque = EOST        
			Flipper.eostorqueangle = EOSA
			Flipper.rampup = Frampup
			Flipper.Elasticity = FElasticity
			FState = 3
		End If

	End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

'######################### Add new dampener to CheckLiveCatch 
'#########################    Note the updated flipper angle check to register if the flipper gets knocked slightly off the end angle

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
    Dim Dir
    Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
    Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
    Dim CatchTime : CatchTime = GameTime - FCount

    if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
		if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
			LiveCatchBounce = 0
		else
			LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
		end If

		If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
		ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
		ball.angmomx= 0
		ball.angmomy= 0
		ball.angmomz= 0
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm, Rubberizer
    End If
End Sub

'****************************************************************************
'PHYSICS DAMPENERS
'****************************************************************************

'These are data mined bounce curves, 
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx) 
	RubbersD.dampen Activeball
	'TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx) 
	SleevesD.Dampen Activeball
	'TargetBouncer Activeball, 0.7
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'#########################    Adjust these values to increase or lessen the elasticity

dim FlippersD : Set FlippersD = new Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False	
FlippersD.addpoint 0, 0, 1.1	
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

'######################### Add Dampenf to Dampener Class 
'#########################    Only applies dampener when abs(velx) < 2 and vely < 0 and vely > -3.75  


Class Dampener
	Public Print, debugOn 'tbpOut.text
	public name, Threshold 	'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
	Public ModIn, ModOut
	Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub 

	Public Sub AddPoint(aIdx, aX, aY) 
		ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
		if gametime > 100 then Report
	End Sub

	public sub Dampen(aBall)
		if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
		dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		if cor.ballvel(aBall.id) = 0 then
			RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.001) 'hack
		Else
			RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
		end If
		coef = desiredcor / realcor 
		if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
		"actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline 
		if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)
		
		aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
		if debugOn then TBPout.text = str
	End Sub

	public sub Dampenf(aBall, parm, ver)
		If ver = 1 Then
			dim RealCOR, DesiredCOR, str, coef
			DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
			if cor.ballvel(aBall.id) = 0 then
                RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.001) 'hack
            Else
                RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
            end If
			coef = desiredcor / realcor 
			If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then 
				aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
			End If
		Elseif ver = 2 Then
			If parm < 10 And parm > 2 And Abs(aball.angmomz) < 15 And aball.vely < 0 then
				aball.angmomz = aball.angmomz * 1.2
				aball.vely = aball.vely * (1.1 + (parm/50))
				'debug.print "high"
			Elseif parm <= 2 and parm > 0.2 And aball.vely < 0 Then
				if (aball.velx > 0 And aball.angmomz > 0) Or (aball.velx < 0 And aball.angmomz < 0) then
			        	aball.angmomz = aball.angmomz * -0.7
				Else
					aball.angmomz = aball.angmomz * 1.2
				end if
				aball.vely = aball.vely * (1.2 + (parm/10))
				'debug.print "low"
			End if
		Elseif ver = 3 Then
			if parm < 10 And parm > 2 And Abs(aball.angmomz) < 10 then
				aball.angmomz = aball.angmomz * 1.2
				aball.vely = aball.vely * 1.2
			Elseif parm <= 2 and parm > 0.2 and aball.vely < 0 Then
				aball.angmomz = aball.angmomz * (-0.5)
				aball.vely = aball.vely * (1.2 + rnd(1)/3 )
			end if
		End If
	End Sub

	Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
		dim x : for x = 0 to uBound(aObj.ModIn)
			addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
		Next
	End Sub


	Public Sub Report() 	'debug, reports all coords in tbPL.text
		if not debugOn then exit sub
		dim a1, a2 : a1 = ModIn : a2 = ModOut
		dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
		TBPout.text = str
	End Sub
	

End Class



Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate4":End Sub
Sub aRollovers_Hit(idx):PlaySoundAtBall "fx_sensor":End Sub

' Slingshots has been hit

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    'PlaySoundAt SoundFXDOF("CrispySlingLeft", 103, DOFPulse, DOFContactors), Lemk
    RandomSoundSlingshotLeft Lemk
    DOF 103, DOFPulse
    LeftSling4.Visible = 1:LeftSling1.Visible = 0
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 170
    ' add some effect to the table?
    'FlashForMs l20, 1000, 50, 0:FlashForMs l20f, 1000, 50, 0
    'FlashForMs l21, 1000, 50, 0:FlashForMs l21f, 1000, 50, 0
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
    Dim i : i = RndNbr(3)
    PlaySoundVol "gotfx-sling"&i,VolDef
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:LeftSling1.Visible = 1:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = False
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If Tilted Then Exit Sub
    'PlaySoundAt SoundFXDOF("CrispySlingRight", 104, DOFPulse, DOFContactors),Remk
    RandomSoundSlingshotRight Remk
    DOF 104,DOFPulse
    RightSling4.Visible = 1:RightSling1.Visible = 0
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 170
    ' add some effect to the table?
    'FlashForMs l22, 1000, 50, 0:FlashForMs l22f, 1000, 50, 0
    'FlashForMs l23, 1000, 50, 0:FlashForMs l23f, 1000, 50, 0
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
    Dim i : i = RndNbr(3)
    PlaySoundVol "gotfx-sling"&i,VolDef
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14:
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2:
        Case 3:RightSLing2.Visible = 0:RightSLing1.Visible = 1:Remk.RotX = -10:RightSlingShot.TimerEnabled = False
    End Select
    RStep = RStep + 1
End Sub


'************************************
' High Score support
'************************************
'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim i,x
    For i = 1 to 16
        x = LoadValue(myGameName, "HighScore"&i)
        If(x <> "")Then HighScore(i-1) = CDbl(x)Else HighScore(i-1) = (17-i)*1000000 End If
        x = LoadValue(myGameName, "HighScore"&i&"Name")
        If(x <> "")Then HighScoreName(i-1) = x Else HighScoreName(i-1) = Chr(64+i)&Chr(64+i)&Chr(64+i) End If
    Next
    x = LoadValue(myGameName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0
    If Credits = 0 And DMDStd(kDMDStd_FreePlay) = False Then DOF 140, DOFOff Else DOF 140, DOFOn
    x = LoadValue(myGameName, "TotalGamesPlayed")
    If(x <> "")then TotalGamesPlayed = CInt(x)Else TotalGamesPlayed = 0 End If
End Sub

Sub reseths
    Dim i
    For i = 0 to 15
        HighScore(i) = (16-i)*10000000
        HighScoreName(i) = Chr(64+i)&Chr(64+i)&Chr(64+i)
    Next
    HighScore(0) = DMDStd(kDMDStd_HighScoreGC)
    HighScore(1) = DMDStd(kDMDStd_HighScore1)
    HighScore(2) = DMDStd(kDMDStd_HighScore2)
    HighScore(3) = DMDStd(kDMDStd_HighScore3)
    HighScore(4) = DMDStd(kDMDStd_HighScore4)
    savehs
End Sub

Sub Clearhs : reseths : End Sub
Sub ClearAll
    reseths
    Dim i
    For i = 0 to kDMDStd_LastStdSetting
        SaveValue myGameName, "DMDStd_"&i,""
    Next
    For i =  kDMDStd_LastStdSetting+1 to 200
        SaveValue myGameName, "DMDFet_"&i,""
    Next
End Sub

Sub Savehs
    Dim i
    For i = 1 to 16
        SaveValue myGameName, "HighScore"&i, HighScore(i-1)
        SaveValue myGameName, "HighScore"&i&"Name", HighScoreName(i-1)
    Next
    SaveValue myGameName, "Credits", Credits
    SaveValue myGameName, "TotalGamesPlayed", TotalGamesPlayed
End Sub

' ***********************************************************
'  High Score Initals Entry Functions - based on Black's code
' ***********************************************************

Dim hsbModeActive
Dim hsEnteredName
Dim hsEnteredDigits(10)
Dim hsCurrentDigit
Dim hsMaxDigits
Dim hsX
Dim hsStart, hsEnd
Dim hsValidLetters
Dim hsCurrentLetter
Dim hsLetterFlash
Dim HSscene

Sub CheckHighscore()
    Dim tmp
    tmp = GameCheckHighScore(Score(CurrentPlayer))

    Select Case tmp
        Case 0: EndOfBallComplete()
        Case 1: HighScoreEntryInit()
        Case 2,3,4,5,6
            Credits = Credits + tmp-1
            DOF 140, DOFOn
            KnockerSolenoid
            DOF 121, DOFPulse
            If tmp > 2 Then
                vpmTimer.AddTimer 500+100*(tmp-3), "HighScoreEntryInit() '"
                Do While tmp > 2
                    vpmTimer.AddTimer 200*(tmp-2),"KnockerSolenoid : DOF 121, DOFPulse '"
                    tmp = tmp - 1
                Loop
            Else
                vpmTimer.AddTimer 300, "HighScoreEntryInit() '"
            End If
    End Select
End Sub

Sub HighScoreEntryInit()
    Dim i
    hsbModeActive = True
    ' Table-specific
    i = RndNbr(4)
    PlaySoundVol "say-enterinitials"&i,VolCallout
    PlaySoundVol "got-track-hstd",VolDef/8
    hsLetterFlash = 0

    hsMaxDigits = DMDStd(kDMDStd_Initials)
    hsText = "> ___ <" : hsX = 40 : hsStart = "> " : hsEnd = " <"
    For i = 0 to hsMaxDigits-1 : hsEnteredDigits(i) = " " : Next
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<" ' < is back arrow
    if hsMaxDigits > 3 Then 
        hsValidLetters = hsValidLetters + "~" ' Add the ~ key to the end to exit long initials entry
        hsText = "" : hsStart="" : hsEnd=""
        hsX = 10
    End If
    hsCurrentLetter = 1
    If bUseFlexDMD Then
        Set HSscene = FlexDMD.NewGroup("highscore")
        tmrDMDUpdate.Enabled = False
        ' Note, these fonts aren't included with FlexDMD. Change to stock fonts for other tables
        HSscene.AddActor FlexDMD.NewLabel("name",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbWhite, 0),"YOUR NAME:")
        HSScene.GetLabel("name").SetAlignedPosition 2,2,FlexDMD_Align_TopLeft
        HSscene.AddActor FlexDMD.NewLabel("initials",FlexDMD.NewFont("skinny10x12.fnt", vbWhite, vbWhite, 0),hsText)
        HSScene.GetLabel("initials").SetAlignedPosition hsX,16,FlexDMD_Align_TopLeft
        DMDFlush()
        DMDDisplayScene HSscene
    End If
    HighScoreDisplayNameNow()

    HighScoreFlashTimer.Interval = 250
    HighScoreFlashTimer.Enabled = True
End Sub

Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        'playsound "fx_Previous"
        'Table-specific
        PlaySoundVol "gotfx-hstd-left",VolDef
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0)then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        'playsound "fx_Next"
        'Table-specific
        PlaySoundVol "gotfx-hstd-right",VolDef
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters))then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = PlungerKey OR keycode = StartGameKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<")then
            'playsound "fx_Enter"
            'Table-specific
            PlaySoundVol "gotfx-hstd-enter",VolDef
            if hsMaxDigits = 10 And (mid(hsValidLetters, hsCurrentLetter, 1) = "~") Then HighScoreCommitName() : Exit Sub
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3 And hsMaxDigits = 3)then
                HighScoreCommitName()
            else
                HighScoreDisplayNameNow()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit> 0)then
                hsCurrentDigit = hsCurrentDigit - 1
            end if
            HighScoreDisplayNameNow()
        end if
    end if
End Sub

Sub HighScoreDisplayNameNow()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreDisplayName()
    Dim i
    Dim TempTopStr
    Dim TempBotStr

    TempTopStr = "YOUR NAME:"
    TempBotStr = hsStart
    For i = 0 to hsMaxDigits-1
        if(hsCurrentDigit> i)then TempBotStr = TempBotStr & hsEnteredDigits(i)
    Next

    if(hsCurrentDigit <> hsMaxDigits)then
        if(hsLetterFlash <> 0)then
            TempBotStr = TempBotStr & "_"
        else
            TempBotStr = TempBotStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    For i = 1 to hsMaxDigits-1
        if(hsCurrentDigit <i)then TempBotStr = TempBotStr & hsEnteredDigits(i)
    Next

    TempBotStr = TempBotStr & hsEnd
    
    If bUseFlexDMD Then
        FlexDMD.LockRenderThread
        With HSscene.GetLabel("initials")
            .Text = TempBotStr
            .SetAlignedPosition hsX,16,FlexDMD_Align_TopLeft
        End With
        FlexDMD.UnlockRenderThread
    Else
        dLine(0) = ExpandLine(TempTopStr, 0)
        dLine(1) = ExpandLine(TempBotStr, 1)
        'DMDUpdate 0
        'DMDUpdate 1
    End If
End Sub

Sub HighScoreFlashTimer_Timer()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = hsLetterFlash + 1
    if(hsLetterFlash = 2)then hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreCommitName()
    Dim i,bBlank
    HighScoreFlashTimer.Enabled = False
    hsbModeActive = False

    hsEnteredName="":bBlank=true
    For i = 0 to hsMaxDigits-1
        hsEnteredName = hsEnteredName & hsEnteredDigits(i) 
        if hsEnteredDigits(i) <> " " Then bBlank=False
    Next
    if bBlank then hsEnteredName = "YOU"

    GameSaveHighScore Score(CurrentPlayer),hsEnteredName
    SortHighscore
    Savehs
    StopSound "got-track-hstd"
    tmrDMDUpdate.Enabled = True
    EndOfBallComplete()
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 0 to 4
        For j = 0 to 3
            If HighScore(j) <HighScore(j + 1)Then
                tmp = HighScore(j + 1)
                tmp2 = HighScoreName(j + 1)
                HighScore(j + 1) = HighScore(j)
                HighScoreName(j + 1) = HighScoreName(j)
                HighScore(j) = tmp
                HighScoreName(j) = tmp2
            End If
        Next
    Next
End Sub


'********************************************************
' DMD Support. Updated to FlexDMD-specific API calls
' and eventually PuP
'******************************************************** 

Const eNone = 0        ' Instantly displayed
Const eBlink = 1
Const eBlinkFast = 2

'Dim FlexPath
Dim UltraDMD
Sub LoadFlexDMD
    Dim curDir
	Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
    If FlexDMD is Nothing Then
        MsgBox "No FlexDMD found. This table will not be as good without it."
        bUseFlexDMD = False
        Exit Sub
    End If
	SetLocale(1033)
	With FlexDMD
		.GameName = myGameName
		.TableFile = Table1.Filename & ".vpx"
        .ProjectFolder = ".\"&myGameName&".FlexDMD\"
		.Color = RGB(255, 88, 32)
		.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
		.Width = 128
		.Height = 32
		.Clear = True
		.Run = True
	End With	

    'Dim fso
    'Set fso = CreateObject("Scripting.FileSystemObject")
    'curDir = fso.GetAbsolutePathName(".")
    'FlexPath = curDir & "\"&myGameName &".FlexDMD\"
End Sub

Sub DMD_Clearforhighscore()
	DMDClearQueue
End Sub

Sub DMDClearQueue				
	if bUseFlexDMD Then
		DMDqHead=0:DMDqTail=0
        FlexDMD.LockRenderThread
        FlexDMD.Stage.RemoveAll
        FlexDMD.UnlockRenderThread
        bDefaultScene = False
        DisplayingScene = Empty
	End If
End Sub

Sub PlayDMDScene(video, timeMs)
	debug.print "PlayDMDScene called"
End Sub

Sub DisplayDMDText(Line1, Line2, duration)
	debug.print "OldDMDText " & Line1 & " " & Line2
	if bUseFlexDMD Then
		'UltraDMD.DisplayScene00 "", Line1, 15, Line2, 15, 14, duration, 14
	Elseif bUsePUPDMD Then
		If bPupStarted then 
			if Line1 = "" or Line1 = "_" then 
				pupDMDDisplay "-", Line2, "" ,Duration/1000, 0, 10
			else
				pupDMDDisplay "-", Line1 & "^" & Line2, "" ,Duration/1000, 0, 10
			End If 
		End If
	End If
End Sub

Sub DisplayDMDText2(Line1, Line2, duration, pri, blink)
	if bUseFlexDMD Then
		'UltraDMD.DisplayScene00 "", Line1, 15, Line2, 15, 14, duration, 14
	Elseif bUsePUPDMD Then
		If bPupStarted then 
			if Line1 = "" or Line1 = "_" then 
				pupDMDDisplay "-", Line2, "" ,Duration/1000, blink, pri
			else
				pupDMDDisplay "-", Line1 & "^" & Line2, "" ,Duration/1000, blink, pri
			End If 
		End If 
	End If
End Sub



Sub DMDId(id, toptext, bottomtext, duration) 'used in the highscore entry routine
	if bUseFlexDMD then 
		'UltraDMD.DisplayScene00ExwithID id, false, "", toptext, 15, 0, bottomtext, 15, 0, 14, duration, 14
	Elseif bUsePUPDMD Then
		If bPupStarted then pupDMDDisplay "default", toptext & "^" & bottomtext, "" ,Duration/1000, 0, 10
	End If 
End Sub

Sub DMDMod(id, toptext, bottomtext, duration) 'used in the highscore entry routine
	if bUseFlexDMD then 
		'UltraDMD.ModifyScene00Ex id, toptext, bottomtext, duration
	Elseif bUsePUPDMD Then
		If bPupStarted then pupDMDDisplay "default", toptext & "^" & bottomtext, "" ,Duration/1000, 0, 10
	End If 
End Sub


Dim dCharsPerLine(2)
Dim dLine(2)

Sub DMD_Init() 'default/startup values
    Dim i, j

	if bUseFlexDMD then LoadFlexDMD()

    DMDFlush()
    dCharsPerLine(0) = 12
    dCharsPerLine(1) = 12
    dCharsPerLine(2) = 12
    For i = 0 to 2
        dLine(i) = Space(dCharsPerLine(i))
    Next
    DMD dLine(0), dLine(1), dLine(2), eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDFlush()
	DMDClearQueue
End Sub

Sub DMDBlank()
    Dim scene
    DMDClearQueue
    Set DefaultScene = FlexDMD.NewGroup("blank")
End Sub

Sub DMDScore()
    Dim tmp, tmp1, tmp2
	Dim TimeStr

	if bUseFlexDMD Then
        debug.log "Call to old DMDScore routine" 
		'	DisplayDMDText RL(0,FormatScore(Score(CurrentPlayer))), "", 1000 
	End If
End Sub

Sub DMDScoreNow
    DMDFlush
    DMDScore
End Sub

Sub DMD(Text0, Text1, Text2, Effect0, Effect1, Effect2, TimeOn, bFlush, Sound)
    Dim scene,line0
    if bUseFlexDMD Then
        Set scene = FlexDMD.NewGroup("dmd")
        If Text0 <> "" Then
            scene.AddActor FlexDMD.NewLabel("line0",FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0), Text0)
            Set line0 = scene.GetLabel("line0")
            line0.SetAlignedPosition 64,5,FlexDMD_Align_Center
            if Effect0 = eBlink Then 
                BlinkActor line0,0.1,int(TimeOn/200)
            Elseif Effect0 = eBlinkFast Then
                BlinkActor line0,0.05,int(TimeOn/100)
            End If
        End If
        If Text1 <> "" Then
            scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", vbWhite, vbWhite, 0), Text1)
            Set line0 = scene.GetLabel("line1")
            line0.SetAlignedPosition 64,16,FlexDMD_Align_Center
            if Effect1 = eBlink Then 
                BlinkActor line0,0.1,int(TimeOn/200)
            Elseif Effect1 = eBlinkFast Then
                BlinkActor line0,0.05,int(TimeOn/100)
            End if
        End If
        If Text2 <> "" Then
            scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0), Text2)
            Set line0 = scene.GetLabel("line2")
            line0.SetAlignedPosition 64,26,FlexDMD_Align_Center
            if Effect2 = eBlink Then 
                BlinkActor line0,0.1,int(TimeOn/200)
            Elseif Effect2 = eBlinkFast Then
                BlinkActor line0,0.05,int(TimeOn/100)
            End if
        End If
        if bFlush Then DMDClearQueue
        DMDEnqueueScene scene,0,TimeOn,TimeOn,1000,Sound
    Else
        DisplayDMDText Text0, Text1, TimeOn
        'if bUsePUPDMD and bPupStarted Then pupDMDDisplay "default", Text0 & "^" & Text1, "" ,2, 0, 10
        'if (bUsePUPDMD) Then pupDMDDisplay "attract", Text1 & "^" Text2, "@vidIntro.mp4" ,9, 1,		10
        PlaySoundVol Sound, VolDef
    End If
End Sub

Function ExpandLine(TempStr, id) 'id is the number of the dmd line
    If TempStr = "" Then
        TempStr = Space(dCharsPerLine(id))
    Else
        if(Len(TempStr) > Space(dCharsPerLine(id)))Then
            TempStr = Left(TempStr, Space(dCharsPerLine(id)))
        Else
            if(Len(TempStr) < dCharsPerLine(id))Then
                TempStr = TempStr & Space(dCharsPerLine(id)- Len(TempStr))
            End If
        End If
    End If
    ExpandLine = TempStr
End Function

Function FormatScore(ByVal Num) 'it returns a string with commas (as in Black's original font)
    dim i
    dim NumString

    NumString = CStr(abs(Num))
    i = InStr(NumString,"+")
    If i > 0 Then
        Dim exp
        ' We got a scientific notation number, convert to a string
        exp = right(Numstring,Len(NumString)-i)
        Numstring = left(NumString,i-1)
        'Get rid of the period between the first and second digit
        NumString = Replace(NumString,".","")
        NumString = Replace(NumString,"E","")
        ' And add 0s to the right length
        For i = Len(NumString) to exp : NumString = NumString & "0" : Next
    End If

    For i = Len(NumString)-3 to 1 step -3
        if IsNumeric(mid(NumString, i, 1))then
            NumString = left(NumString, i) & "," & right(NumString, Len(NumString)-i)
        end if
    Next
    FormatScore = NumString
End function

Function CL(id, NumString) 'center line
    Dim Temp, TempStr
	NumString = LEFT(NumString, dCharsPerLine(id))
    Temp = (dCharsPerLine(id)- Len(NumString)) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function

Function RL(id, NumString) 'right line
    Dim Temp, TempStr
	NumString = LEFT(NumString, dCharsPerLine(id))
    Temp = dCharsPerLine(id)- Len(NumString)
    TempStr = Space(Temp) & NumString
    RL = TempStr
End Function

Function FL(id, aString, bString) 'fill line
    Dim tmp, tmpStr
	aString = LEFT(aString, dCharsPerLine(id))
	bString = LEFT(bString, dCharsPerLine(id))
    tmp = dCharsPerLine(id)- Len(aString)- Len(bString)
	If tmp <0 Then tmp = 0
    tmpStr = aString & Space(tmp) & bString
    FL = tmpStr
End Function


'*************************
' FlexDMD Queue Management
'*************************
'
' FlexDMD supports queued scenes using its built-in Sequence class. However, there's no way to set priorities
' to allow new scenes to override playing scenes. In addition, there's no support for 'minimum play time' vs
' 'total play time', or for playing a sound with a scene. We want the ability to let a scene of a given priority play for at least 'minimum play time'
' as long as no scene of higher priority gets queued. If another scene of equal priority is queued, the playing scene
' will be replaced once it has played for 'minimum play time' ms.
' Queued higher priority scenes immediately replace playing lower priority scenes
' When no scenes are queued, show default scene (Score, "Choose..." or GameOver)
'
' If a scene gets queued that would take too long before it can be played due to items ahead of it, it gets dropped

Dim DMDSceneQueue(64,6)     ' Queue of scenes. Each entry has 7 fields: 0=Scene, 1=priority, 2=mintime, 3=maxtime, 4=waittime, 5=optionalSound, 6=timestamp
Dim DMDqHead,DMDqTail
Dim DMDtimestamp

'Queue up a FlexDMD scene in a virtual queue. 

Sub DMDEnqueueScene(scene,pri,mint,maxt,waitt,sound)
    Dim i
    If bDefaultScene = False And Not IsEmpty(DisplayingScene) Then
        If DisplayingScene is scene Then
            ' Already playing. Update it
            DMDSceneQueue(DMDqHead,1) = pri
            DMDSceneQueue(DMDqHead,2) = mint
            DMDSceneQueue(DMDqHead,3) = maxt
            DMDSceneQueue(DMDqHead,4) = DMDtimestamp + waitt
            DMDSceneQueue(DMDqHead,5) = sound
            DMDSceneQueue(DMDqHead,6) = DMDtimestamp
            Exit Sub
        End If
    End If

    ' Check to see whether the scene is worth queuing
    If Not DMDCheckQueue(pri,waitt) Then 
        ' debug.print "Discarding scene request with priority " & pri & " and waitt "&waitt
        Exit Sub
    End If

    ' Check to see if this is an update to an existing queued scene (e.g pictopops)
    Dim found:found=False
    If DMDqTail <> 0 Then
        For i = DMDqHead to DMDqTail-1
            If DMDSceneQueue(i,0) Is scene Then 
                Found=True
                Exit For
            End If
        Next
    End If
    'Otherwise add to end of queue
    If Found = False Then i = DMDqTail:DMDqTail = DMDqTail + 1
    Set DMDSceneQueue(i,0) = scene
    DMDSceneQueue(i,1) = pri
    DMDSceneQueue(i,2) = mint
    DMDSceneQueue(i,3) = maxt
    DMDSceneQueue(i,4) = DMDtimestamp + waitt
    DMDSceneQueue(i,5) = sound
    DMDSceneQueue(i,6) = 0
    If DMDqTail > 64 Then       ' Ran past the end of the queue!
        DMDqTail = 64
    End if
    debug.print "Enqueued scene at "&i& " name: "&scene.Name
End Sub

' Check the queue to see whether a scene willing to wait 'waitt' time would play
Function DMDCheckQueue(pri,waitt)
    Dim i,wait:wait=0
    If DMDqTail=0 Then DMDCheckQueue = True: Exit Function
    DMDCheckQueue = False
    For i = DMDqHead to DMDqTail
        If DMDSceneQueue(i,4) > DMDtimestamp Then 
            If DMDSceneQueue(i,1) = pri Then        'equal priority queued scene
                wait = wait + DMDSceneQueue(i,2)    ' so use mintime
            ElseIf DMDSceneQueue(i,1) < pri Then    'higher priority queued scene
                wait = wait + DMDSceneQueue(i,3)
            End If
            If wait > waitt Then Exit Function
        End If
        
    Next
    DMDCheckQueue = True
End Function

' Return the total length of the display queue, in ms
' Mainly used by end-of-ball processing to delay Bonus until all scenes have shown
' This isn't 100% accurate, as the last scene at a given priority level will play for maxtime
' before allowing a scene at a lower priority level play. We just add up all the mintimes
Function DMDGetQueueLength
    Dim i,j,len
    DMDGetQueueLength = 0
    j=0:len=0
    If DMDqTail = 0 Then Exit Function
    For j = 0 to 3  ' We don't care about really low priority scenes in this context
        For i = DMDqHead to DMDqTail
            If DMDSceneQueue(i,4) > DMDtimestamp Then 
                If DMDSceneQueue(i,1) = j And DMDSceneQueue(i,4) > DMDtimestamp+len Then        'equal priority queued scene
                    len = len + DMDSceneQueue(i,2)    ' so use mintime
                End If
            End If
        Next
    Next
    DMDGetQueueLength = len
End Function      

' Update DMD Scene. Called every 100ms
' Most of the work is done here. If scene queue is empty, display default scene (score, Game Over, etc)
' If scene queue isn't empty, check to see whether current scene has been on long enough or overwridden by a higher priority scene
' If it has, move to next spot in queue and search all of the queue for scene with highest priority, skipping any scenes that have timed out while waiting
Dim bDefaultScene,DefaultScene
Sub tmrDMDUpdate_Timer
    Dim i,j,bHigher,bEqual
    tmrDMDUpdate.Enabled = False
    DMDtimestamp = DMDtimestamp + 100   ' Set this to whatever frequency the timer uses
    If DMDqTail <> 0 Then
        ' Process queue
        ' Check to see if queue is idle (default scene on). If so, immediately play first item
        If bDefaultScene or (IsEmpty(DisplayingScene) And DMDqHead = 0) Then
            bDefaultScene = False
            debug.print "Idle: Displaying scene at " & DMDqHead & " Tail: "&DMDqTail
            DMDDisplayScene DMDSceneQueue(DMDqHead,0)
            DMDSceneQueue(DMDqHead,6) = DMDtimestamp
            If DMDSceneQueue(DMDqHead,5) <> ""  Then 
                PlaySoundVol DMDSceneQueue(DMDqHead,5),VolDef
            ' Note, code below is game-specific - triggers the SuperJackpot scene update timer as soon as the scene plays
            ElseIf DMDSceneQueue(DMDqHead,0).Name = "bwsjp" Then
                tmrSJPScene.UserValue = 0
                tmrSJPScene.Interval = 300
                tmrSJPScene.Enabled = True
            End If
        Else
            ' Check to see whether there are any queued scenes with equal or higher priority than currently playing one
            bEqual = False: bHigher = False
            If DMDqTail > DMDqHead+1 Then
                For i = DMDqHead+1 to DMDqTail-1
                    If DMDSceneQueue(i,1) < DMDSceneQueue(DMDqHead,1) Then bHigher=True:Exit For
                    If DMDSceneQueue(i,1) = DMDSceneQueue(DMDqHead,1) Then bEqual = True:Exit For
                Next
            End If
            If bHigher Or (bEqual And (DMDSceneQueue(DMDqHead,6)+DMDSceneQueue(DMDqHead,2) <= DMDtimestamp)) Or _ 
                    (DMDSceneQueue(DMDqHead,6)+DMDSceneQueue(DMDqHead,3) <= DMDtimestamp) Then 'Current scene has played for long enough

                ' Skip over any queued scenes whose wait times have expired
                Do 
                    DMDqHead = DMDqHead+1
                Loop While DMDSceneQueue(DMDqHead,4) < DMDtimestamp And DMDqHead < DMDqTail
                    
                If DMDqHead > 64 Then       ' Ran past the end of the queue!
                    debug.print "DMDSceneQueue too big! Resetting"
                    DMDqHead = 0:DMDqTail = 0
                    tmrDMDUpdate.Enabled = True
                ElseIf DMDqHead = DMDqTail Then ' queue is empty
                    DMDqHead = 0:DMDqTail = 0
                    tmrDMDUpdate.Enabled = True
                Else
                    ' Find the next scene with the highest priority
                    j = DMDqHead
                    For i = DMDqHead to DMDqTail-1
                        If DMDSceneQueue(i,1) < DMDSceneQueue(j,1) Then j=i:DMDqHead=i
                    Next

                    ' Play the scene, and a sound if there's one to accompany it
                    bDefaultScene = False
                    debug.print "Displaying scene at " &j & " name: "&DMDSceneQueue(j,0).Name & " Head: "&DMDqHead & " Tail: "&DMDqTail
                    DMDSceneQueue(j,6) = DMDtimestamp
                    DMDDisplayScene DMDSceneQueue(j,0)
                    If DMDSceneQueue(j,5) <> ""  Then 
                        PlaySoundVol DMDSceneQueue(j,5),VolDef
                    ' Note, code below is game-specific - triggers the SuperJackpot scene update timer as soon as the scene plays
                    ElseIf DMDSceneQueue(j,0).Name = "bwsjp" Then
                        tmrSJPScene.UserValue = 0
                        tmrSJPScene.Interval = 600
                        tmrSJPScene.Enabled = True
                    End If
                End If
            End If
        End If
    End If
    If DMDqTail = 0 Then ' Queue is empty
        ' Exit fast if defaultscene is already showing
        if bDefaultScene or IsEmpty(DefaultScene) then tmrDMDUpdate.Enabled = True : Exit Sub
        bDefaultScene = True
        If TypeName(DefaultScene) = "Object" Then
            DMDDisplayScene DefaultScene
        Else
            debug.print "DefaultScene is not an object!"
        End If
    End If
    tmrDMDUpdate.Enabled = True
End Sub
    
Dim DisplayingScene     ' Currently displaying scene
Sub DMDDisplayScene(scene)
    If TypeName(scene) <> "Object" then
		debug.print "DMDDisplayScene: scene is not an object! Type=" & TypeName(scene)
		exit sub
	ElseIf scene Is Nothing or IsEmpty(scene) Then
		debug.print "DMDDisplayScene: scene is empty!"
		exit Sub
	End If
    If Not IsEmpty(DisplayingScene) Then If DisplayingScene Is scene Then Exit Sub
    FlexDMD.LockRenderThread
    FlexDMD.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
    FlexDMD.Stage.RemoveAll
    FlexDMD.Stage.AddActor scene
    FlexDMD.Show = True
    FlexDMD.UnlockRenderThread
    Set DisplayingScene = scene
End Sub

' Create a new scene with a video file. If the video file
' is not found, look for an image file. If that's not found, 
' create a new blank scene
Function NewSceneWithVideo(name,videofile)
    Dim actor
    Set NewSceneWithVideo = FlexDMD.NewGroup(name)
    Set actor = FlexDMD.NewVideo(name&"vid",videofile & ".gif")
    If actor is Nothing Then
        Set actor = FlexDMD.NewImage(name&"img",videofile&".png")
        if actor is Nothing Then 
            debug.print "Warning: "&videofile&" image not found"
            Exit Function
        End if
    End If
    NewSceneWithVideo.AddActor actor
End Function

' Create a new scene with an image file. If that's not found, 
' create a new blank scene
Function NewSceneWithImage(name,imagefile)
    Dim actor
    Set NewSceneWithImage = FlexDMD.NewGroup(name)
    Set actor = FlexDMD.NewImage(name&"img",imagefile&".png")
    if actor is Nothing Then Exit Function
    NewSceneWithImage.AddActor actor
End Function

' Create a scene from a series of images. The only reason to use this
' function is if you need to use transparent images. If you don't, use
' an animated GIF - much easier. However, this does have one other advantage over
' an animated GIF: FlexDMD will loop animated GIFs, regardless of what the loop attribute is set to in the GIF
'  name   - name of the scene object returned
'  imgdir - directory inside the FlexDMD project folder where the images are stored
'  start  - number of first image
'  num    - number of images, numbered from image1..image<num>
'  fps    - frames per second - a delay of 1/fps is used between frames
'  hold   - if non-zero, how long to hold the last frame visible. If 0, the last scene will end with the last frame visible
'  repeat - Number of times to repeat. 0 or 1 means don't repeat
Function NewSceneFromImageSequence(name,imgdir,num,fps,hold,repeat)
    Set NewSceneFromImageSequence = NewSceneFromImageSequenceRange(name,imgdir,1,num,fps,hold,repeat)
End Function

Function NewSceneFromImageSequenceRange(name,imgdir,start,num,fps,hold,repeat)
    Dim scene,i,actor,af,blink,total,delay,e
    total = num/fps + hold
    delay = 1/fps
    e = start+num-1
    Set scene = FlexDMD.NewGroup(name)
    For i = start to e
        Set actor = FlexDMD.NewImage(name&i,imgdir&"\image"&i&".png")
        actor.Visible = 0
        Set af = actor.ActionFactory
        Set blink = af.Sequence()
        blink.Add af.Wait((i-start)*delay)
        blink.Add af.Show(True)
        blink.Add af.Wait(delay*1.2)    ' Slightly longer than one frame length to ensure no flicker
        if i=e And hold > 0 Then blink.Add af.Wait(hold)
        if repeat > 1 or i<e Then 
            blink.Add af.Show(False)
            blink.Add af.Wait((e-i)*delay)
        End If
        If repeat > 1 Then actor.AddAction af.Repeat(blink,repeat) Else actor.AddAction blink
        scene.AddActor actor
    Next
    Set NewSceneFromImageSequenceRange = scene
End Function


' Add a blink action to an Actor in a FlexDMD scene. 
' Usage: BlinkActor scene.GetActor("name"),blink-interval-in-seconds,repetitions 
' Blink action is only natively supported in FlexDMD 1.9+
' poplabel.AddAction af.Blink(0.1, 0.1, 5)
Sub BlinkActor(actor,interval,times)
    Dim af,blink
    Set af = actor.ActionFactory
    Set blink = af.Sequence()
    blink.Add af.Show(True)
    blink.Add af.Wait(interval)
    blink.Add af.Show(False)
    blink.Add af.Wait(interval)
    actor.AddAction af.Repeat(blink,times)
End Sub

' Add action to delay toggling the state of an actor
' If on=true then delay then show, otherwise, delay then hide
Sub DelayActor(actor,delay,bOn)
    Dim af,blink
    Set af = actor.ActionFactory
    Set blink = af.Sequence()
    blink.Add af.Wait(delay)
    blink.Add af.Show(bOn)
    actor.AddAction blink
End Sub

' Add action to an actor to delay turning it on, wait, then turn it off
Sub FlashActor(actor,delayon,delayoff)
    Dim af,blink
    Set af = actor.ActionFactory
    Set blink = af.Sequence()
    blink.Add af.Wait(delayon)
    blink.Add af.Show(true)
    blink.Add af.Wait(delayoff)
    blink.Add af.Show(false)
    actor.AddAction blink
End Sub

' Add action to delay turning on an actor, then blink it 
Sub DelayBlinkActor(actor,delay,blinkinterval,times)
    Dim af,seq,blink
    Set af = actor.ActionFactory
    Set seq = af.Sequence()
    Set blink = af.Sequence()
    seq.Add af.Wait(delay)
    blink.Add af.Show(True)
    blink.Add af.Wait(blinkinterval)
    blink.Add af.Show(False)
    blink.Add af.Wait(blinkinterval)
    seq.Add af.Repeat(blink,times)
    actor.AddAction seq
End Sub


Dim TableName : TableName = myGameName
' ***********************************************************************
' DAPHISHBOWL - STERN Service Menu
' - Modified by Skillman for SPIKE1/FlexDMD
'   - Refactored to remove distinction between Std and Game-specific features in feature array
'   - Added variables to indicate where game-specific features start in the array
'   - Moved all display logic outside the ServiceMenu sub
'   - Added Defaults to Settings array
'   - Optimized initialization of settings
'   - Added support for Boolean options in the Service Menu Adjustment
' ***********************************************************************

' SERVICE Key Definitions
Const ServiceCancelKey 	= 8			' 7 key
Const ServiceUpKey 		= 9			' 8 key
Const ServiceDownKey 	= 10		' 9 key
Const ServiceEnterKey 	= 11		' 0 key

Dim serviceSaveAttract
dim serviceIdx
Dim serviceLevel
dim bServiceMenu
dim bServiceVol
dim bServiceEdit
dim serviceOrigValue
Dim bInService:bInService=False
Dim MasterVol:MasterVol=100
'Dim VolBGVideo:VolBGVideo = cVolBGVideo
Dim VolBGMusic :VolBGMusic = cVolBGMusic
Dim VolCallout : VolCallout = cVolCallout
Dim VolDef:VolDef = cVolDef
Dim VolSfx:VolSfx = cVolSfx
Dim VolTable:VolTable = cVolTable
Dim TopArray:TopArray = Array("GO TO DIAGNOSTIC MENU","GO TO AUDITS MENU","GO TO ADJUSTMENTS MENU","GO TO UTILITIES MENU","GO TO TOURNAMENT MENU","GO TO REDEMPTION MENU","EXIT SERVICE MENU")
Dim AdjArray:AdjArray = Array("STANDARD ADJUSTMENTS","FEATURE ADJUSTMENTS","PREVIOUS MENU","EXIT SRVICE MENU","HELP")
Const kMenuNone 	= -1
Const kMenuTop	 	= 0
Const kMenuAdj	 	= 1
Const kMenuAdjStd	= 2
Const kMenuAdjGame  = 3

Sub StartServiceMenu(keycode)
	Dim i
	Dim maxVal
	Dim minVal
	Dim dataType
	dim TmpStr
	Dim direction
	Dim values
	Dim valuesLst
	Dim NewVal
	Dim displayText
    Dim SvcFontSm
    Dim SvcFontBg
    Dim line3
	
	if keycode=ServiceEnterKey and bInService=False then 
'debug.print "Start Service:" & bInService
		PlaySoundVol "stern-svc-enter", VolSfx
		serviceSaveAttract=bAttractMode
		bAttractMode=False
		bServiceMenu=False
		bServiceEdit=False 
		bInService=True
		bServiceVol=False
		serviceIdx=-1
		serviceLevel=kMenuNone
		
        tmrDMDUpdate.Enabled = False
        tmrAttractModeScene.Enabled = False

        CreateServiceDMDScene serviceIdx

	elseif keycode=ServiceCancelKey then 					' 7 key starts the service menu
		if bInService then
			PlaySoundVol "stern-svc-cancel", VolSfx
			Select case serviceLevel:
				case kMenuTop, kMenuNone: 
					bInService=False
					bServiceVol=False
					bAttractMode=serviceSaveAttract
					if bAttractMode then
						StartAttractMode
                    Else
                        tmrDMDUpdate.Enabled = True
					End if 
					
				case kMenuAdj:
					serviceLevel=kMenuNone
					serviceIdx=0
					StartServiceMenu 11
				case kMenuAdjStd,kMenuAdjGame:
					if bServiceEdit then 
						bServiceEdit=False 
                        DMDStd(DMDMenu(serviceIdx).StdIdx)=serviceOrigValue
                        UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
					else 
                        serviceLevel=kMenuTop
						serviceIdx=2
						StartServiceMenu 11
					End if 
			End Select  
		End if 
	elseif bInService Then
		if keycode=ServiceEnterKey then 	' Select 
			PlaySoundVol "stern-svc-enter", VolSfx
			select case serviceLevel
				case kMenuNone:
					serviceLevel=kMenuTop
					serviceIdx=0
                    CreateServiceDMDScene serviceLevel
				case kMenuTop: 
					if serviceIdx=2 then 
						serviceLevel=kMenuAdj : serviceIdx=0
						CreateServiceDMDScene serviceLevel
					elseif serviceIdx=6 then 
						StartServiceMenu 8		' Exit 
						Exit sub 
					else 
						PlaySoundVol "sfx-deny", VolSfx
					End if 
				case kMenuAdj:
                    Select Case serviceIdx
					    Case 0,1:
                            serviceLevel=kMenuAdjStd+serviceIdx
                            If serviceLevel = kMenuAdjStd Then serviceIdx=0 Else serviceIdx = MinFetDMDSetting
                            CreateServiceDMDScene serviceLevel
                            StartServiceMenu 0
					    Case 2: 		' Go Up
                            serviceLevel=kMenuNone : serviceIdx=0
                            StartServiceMenu 11
                            Exit sub 
					    Case 3:
                            StartServiceMenu 8		' Exit 
                            Exit sub 
					    Case else: 
						    PlaySoundVol "sfx-deny", VolSfx
					End Select
				case kMenuAdjStd,kMenuAdjGame:
					if DMDMenu(serviceIdx).ValType="FUN" then		' Function
                        Select Case DMDMenu(serviceIdx).StdIdx
                            Case 0: reseths
                            Case 1: ClearAll
                        End Select
						UpdateServiceDMDScene "<DONE>"
					else 
						if bServiceEdit=False then 	' Start Editing  
							bServiceEdit=True
                            serviceOrigValue=DMDStd(DMDMenu(serviceIdx).StdIdx)
							UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
						else 							' Save the change 
							bServiceEdit=False
							UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
						End if 
					End if 
			End Select 
		elseif serviceLevel<>kMenuNone then 
			if bServiceEdit = False then 
				if keycode=ServiceUpKey then 	' Left
					PlaySoundVol "stern-svc-minus", VolSfx
					serviceIdx=serviceIdx-1
					if serviceIdx<0 then serviceIdx=0
				elseif keycode=ServiceDownKey then 	' Right 
					PlaySoundVol "stern-svc-plus", VolSfx
					serviceIdx=serviceIdx+1
				End if

				select case serviceLevel
					case kMenuTop:
						if serviceIdx>6 then serviceIdx=0
						UpdateServiceDMDScene ""
					case kMenuAdj:
						if serviceIdx>4 then serviceIdx=0
						UpdateServiceDMDScene ""
					case kMenuAdjStd
                        if serviceIdx > MaxStdDMDSetting Then serviceIdx = MaxStdDMDSetting
                        if DMDMenu(serviceIdx).ValType<>"FUN" then
							UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
						else 
							UpdateServiceDMDScene "<EXECUTE>"
						End if 
                    case kMenuAdjGame:
                        if serviceIdx < MinFetDMDSetting Then serviceIdx = MinFetDMDSetting
                        If serviceIdx > MaxFetDMDSetting Then serviceIdx = MaxFetDMDSetting
                        UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
				End Select
			else
        
            ' ******************************************************    HANDLE EDIT MODE
				dataType=mid(DMDMenu(serviceIdx).ValType, 1, 3)
				minVal=-1
				maxVal=-1
				if Instr(DMDMenu(serviceIdx).ValType, "[")<> 0 then 
					TmpStr =mid(DMDMenu(serviceIdx).ValType, 5, Len(DMDMenu(serviceIdx).ValType)-5)
					values = Split(TmpStr, ":")
					if dataType="INT" or dataType="PCT" then 
						minVal=CLNG(values(0))
						maxVal=CLNG(values(1))
					End if 
				End if

				direction=0
				if keycode=ServiceUpKey then 	' Left
					PlaySoundVol "stern-svc-minus", VolSfx
					direction=-1
				elseif keycode=ServiceDownKey then 	' Right 
					PlaySoundVol "stern-svc-plus", VolSfx
					direction=1
				End if
				if direction<>0 then 
					if dataType="INT" or dataType="PCT" then                            
                        NewVal=DMDStd(DMDMenu(serviceIdx).StdIdx) + (direction*DMDMenu(serviceIdx).Increment)
                        if minVal=-1 or (NewVal <= maxVal and NewVal >= minVal) then 
                            DMDStd(DMDMenu(serviceIdx).StdIdx)=DMDStd(DMDMenu(serviceIdx).StdIdx) + (direction*DMDMenu(serviceIdx).Increment)
                            SaveValue TableName, "DMDStd_"&DMDMenu(serviceIdx).StdIdx, DMDStd(DMDMenu(serviceIdx).StdIdx)	' SAVE 
                            if DMDMenu(serviceIdx).StdIdx<>kDMDStd_Initials and DMDMenu(serviceIdx).StdIdx<>kDMDStd_LeftStartReset then 
                                SaveValue TableName, "dmdCriticalChanged", "True"		' SAVE
                                dmdCriticalChanged=True
                            End if 
                        end if 
                        displayText=DMDStd(DMDMenu(serviceIdx).StdIdx)
						
					elseif dataType="LST" then 
                        For i = 0 to ubound(values)
                            valuesLst=Split(values(i), ",")
'debug.print "EDIT LST:" & valuesLst(0) & " " & valuesLst(1) & " val: " & DMDStd(DMDMenu(serviceIdx).StdIdx) & " idx:" & i & " ubound:" & ubound(values)
                            if DMDStd(DMDMenu(serviceIdx).StdIdx)&"" = valuesLst(0) and direction=1 then
                                If i <  ubound(values) Then valuesLst=Split(values(i+1), ",") Else  valuesLst=Split(values(0), ",")
                                DMDStd(DMDMenu(serviceIdx).StdIdx)=valuesLst(0)
                                SaveValue TableName, "DMDStd_"&DMDMenu(serviceIdx).StdIdx, DMDStd(DMDMenu(serviceIdx).StdIdx)	' SAVE 
                                displayText=valuesLst(1)
                                Exit For
                            elseif DMDStd(DMDMenu(serviceIdx).StdIdx)&"" = valuesLst(0) and direction=-1 then 
                                If i > lbound(values) Then valuesLst=Split(values(i-1), ",") Else valuesLst=Split(values(ubound(values)), ",")
                                DMDStd(DMDMenu(serviceIdx).StdIdx)=valuesLst(0)
                                SaveValue TableName, "DMDStd_"&DMDMenu(serviceIdx).StdIdx, DMDStd(DMDMenu(serviceIdx).StdIdx)	' SAVE 
                                displayText=valuesLst(1)
                                Exit For
                            End if 
                        Next 
                    elseif dataType="BOO" then 'Boolean
                        If DMDStd(DMDMenu(serviceIdx).StdIdx) = 1 Then
                            DMDStd(DMDMenu(serviceIdx).StdIdx) = 0 : displayText = "OFF"
                        Else
                            DMDStd(DMDMenu(serviceIdx).StdIdx) = 1 : displayText = "ON"
                        End if
                        SaveValue TableName, "DMDStd_"&DMDMenu(serviceIdx).StdIdx, DMDStd(DMDMenu(serviceIdx).StdIdx)	' SAVE
					End if 

					if displayText<>"" then UpdateServiceDMDScene displayText
				End If
				if DMDMenu(serviceIdx).StdIdx = kDMDStd_BallsPerGame Then BallsPerGame = DMDStd(kDMDStd_BallsPerGame) 
			End if 
		End if
	elseif keycode=ServiceUpKey or keycode=ServiceDownKey then 		' If you press 8 & 9 without being in service you do volume 
		if keycode=ServiceUpKey and MasterVol>0 then 	' Left
			MasterVol=MasterVol-1
			PlaySoundVol "stern-svc-minus", VolSfx
		elseif keycode=ServiceDownKey and MasterVol<100 then 	' Right 
			PlaySoundVol "stern-svc-plus", VolSfx
			MasterVol=MasterVol+1
		End if

		'VolBGVideo = cVolBGVideo * (MasterVol/100.0)
		VolBGMusic = cVolBGMusic * (MasterVol/100.0)
		VolDef 	 = cVolDef * (MasterVol/100.0)
		VolSfx 	 = cVolSfx * (MasterVol/100.0)

		if bServiceVol=False then
			bServiceVol=True 
			serviceSaveAttract=bAttractMode
			bAttractMode=False
            tmrDMDUpdate.Enabled = False
            tmrAttractModeScene.Enabled = False
            CreateServiceDMDScene 0
		End if 

		UpdateServiceDMDScene

		tmrService.Enabled = False 
		tmrService.Interval = 5000
		tmrService.Enabled = True 
	End if 
End Sub

Sub tmrService_Timer()
	tmrService.Enabled = False 	

	bServiceVol=False 
	bAttractMode=serviceSaveAttract
	if bAttractMode then 
		StartAttractMode
    Else
        tmrDMDUpdate.Enabled = true
	End if 
End Sub

' Decode the ValType from the feature definition array to extract what kind of setting it is
' If it's INT or PCT, just pass back the current value as the text
' If it's LST, look up the current value in the LST to find the string equiv
' If it's BOOL, convert to "OFF" or "ON"
Function GetTextForAdjustment(idx)
    Dim txt,dataType,values,valuesLst,TmpStr,i
    txt = DMDStd(DMDMenu(idx).StdIdx) ' For INT and PCT types

    dataType=mid(DMDMenu(serviceIdx).ValType, 1, 3) ' First 3 chars of valType
    if Instr(DMDMenu(serviceIdx).ValType, "[")<> 0 then 
        TmpStr =mid(DMDMenu(serviceIdx).ValType, 5, Len(DMDMenu(serviceIdx).ValType)-5)
        values = Split(TmpStr, ":")	
    End if
    
    Select Case dataType
        Case "BOO": If DMDStd(DMDMenu(idx).StdIdx) <> 0 Then txt = "ON" Else txt = "OFF"
        Case "LST"
            For i = 0 to ubound(values)
				valuesLst=Split(values(i), ",")
                If DMDStd(DMDMenu(idx).StdIdx)&"" = valuesLst(0) Then txt = valuesLst(1) : Exit For
            Next
    End Select
    debug.print "DataType: "&dataType&"  Returning "&txt
    GetTextForAdjustment = txt
End Function

Class DMDSettings
	Public Name 
	Public StdIdx
	Public Deflt
	Public ValType					' bool, pct
	Public Increment			' value to increment by
    Sub Class_Initialize
        me.Name="EMPTY"
    End Sub
	Public sub Setup(Name, StdIdx, Deflt, ValType, Increment)
		me.name=name
		me.StdIdx=StdIdx
        me.Deflt = Deflt
		me.ValType=ValType
		me.Increment=Increment
	End sub 
End Class 

' Set this according to the size of your settings list
Dim MaxStdDMDSetting,MinFetDMDSetting,MaxFetDMDSetting : MaxStdDMDSetting = 27
Dim DMDMenu(60)				' MaxDMDSetting
Dim dmdChanged:dmdChanged=False							' Did we change a value 
Dim dmdCriticalChanged:dmdCriticalChanged=False			' Did we change a critical value 

' Indexes into the DMDStd array, where the actual value is stored. These indexes could just as easily
' count up from 0, but these values were taken from Pinball Browser and match the actual game
Const kDMDStd_ExtraBallLimit = &H3B
Const kDMDStd_ExtraBallPCT = &H3C
Const kDMDStd_MatchPCT = &H3D           '√
Const kDMDStd_TiltWarn = &H3F
Const kDMDStd_TiltDebounce = &H40
Const kDMDStd_LeftStartReset = &H46     '√
Const kDMDStd_BallSave = &H4C           '√
Const kDMDStd_BallSaveExtend = &H4D     '√
Const kDMDStd_ReplayType = &H2C         '√
Const kDMDStd_AutoReplayStart = &H30    '√
Const kDMDStd_BallsPerGame = &H3E       '√
Const kDMDStd_FreePlay = &H42           '√
Const kDMDStd_HighScoreGC = &H54        '√
Const kDMDStd_HighScore1 = &H55         '√
Const kDMDStd_HighScore2 = &H56         '√
Const kDMDStd_HighScore3 = &H57         '√
Const kDMDStd_HighScore4 = &H58         '√
Const kDMDStd_HighScoreAward = &H59     '√
Const kDMDStd_GCAwards = &H5A           '√
Const kDMDStd_HS1Awards = &H5B          '√
Const kDMDStd_HS2Awards = &H5C          '√
Const kDMDStd_HS3Awards = &H5D          '√
Const kDMDStd_HS4Awards = &H5E          '√
Const kDMDStd_Initials = &H5F           '√
Const kDMDStd_MusicAttenuation = &H60
Const kDMDStd_SpeechAttenuation = &H61

Const kDMDStd_LastStdSetting = &H67

'Table-specific indexes below here
Const kDMDFet_HouseCompleteExtraBall = &H68     '√
Const kDMDFet_CasualPlayerMode = &H6C
Const kDMDFet_PlayerHouseChoice = &H6D          '√
Const kDMDFet_DisableRampDropTarget = &H71
Const kDMDFet_ActionButtonAction = &H73   ' Toggle between house actions or "Iron Bank" (pre v1.37 behaviour)
Const kDMDFet_LannisterButtonsPerGame = &H74    '√
Const kDMDFet_TargaryenFreezeTime = &H75        '√
Const kDMDFet_TargaryenHousePower = &H76        '√ toggle between 1, 2 or all dragons completed. 1 or 2 scores 30M per at start of game
Const kDMDFet_DireWolfFrequency = &H77 ' Start Action button: 1-5 per game, 1 per ball, or 0
Const kDMDFet_SwordsUnlockMultiplier = &H78     '√
Const kDMDFet_RamHitsForMult = &H79
Const kDMDFet_TwoBankDifficulty = &H7A
Const kDMDFet_BWMB_SaveTimer = &H7D             '√
Const kDMDFet_WallMB_SaveTimer = &H7E           '√
Const kDMDFet_HOTKMB_SaveTimer = &H7F           '√
Const kDMDFet_ITMB_SaveTimer = &H80             '√
Const kDMDFet_MMMB_SaveTimer = &H81             '√
Const kDMDFet_WHCMB_SaveTimer = &H82            '√
Const kDMDFet_CMB_SaveTimer = &H83              '√
Const kDMDFet_CastleLoopCompletesHouse = &H85
Const kDMDFet_InactivityPauseTimer = &H96
Const kDMDFet_ReduceWHCLamps = &H97
Const kDMDFet_MidnightMadness = &H98


Sub DMDSettingsInit()
	Dim i
	Dim x

	For i = 0 to 60
		set DMDMenu(i)=new DMDSettings
    Next
		
        '  Name, Index, Default, Variable type
    DMDMenu(0).Setup "HSTD  INITIALS", 						kDMDStd_Initials	,	3,		"LST[3,3 INITIALS:10,10 LETTER NAME]", 1 
    DMDMenu(1).Setup "EXTRA  BALL  LIMIT", 					kDMDStd_ExtraBallLimit,	5,		"INT[0:9]", 1 
    DMDMenu(2).Setup "EXTRA  BALL  PERCENT", 				kDMDStd_ExtraBallPCT,	25,		"PCT[0:50]", 1 
    DMDMenu(3).Setup "MATCH  PERCENT", 						kDMDStd_MatchPCT,		9,		"PCT[0:10]", 1 
    DMDMenu(4).Setup "LEFT+START  RESETS", 					kDMDStd_LeftStartReset,	2,		"LST[0,DISABLED:1,ALWAYS:2,FREE PLAY]", 1 
    DMDMenu(5).Setup "BALL  SAVE  SECONDS", 					kDMDStd_BallSave,		5,		"INT[0:15]", 1 
    DMDMenu(6).Setup "BALL  SAVE  EXTEND  SEC", 				kDMDStd_BallSaveExtend,	3,		"INT", 1 
    DMDMenu(7).Setup "TILT  WARNINGS", 						kDMDStd_TiltWarn	,	2,		"INT[0:3]", 1 
    DMDMenu(8).Setup "TILT  DEBOUNCE", 						kDMDStd_TiltDebounce,	1000,	"INT[750:1500]", 1 
    DMDMenu(9).Setup "REPLAY  TYPE", 						kDMDStd_ReplayType,	    1,		"LST[1,EXTRA GAME:2,EXTRA BALL]", 1 
    DMDMenu(10).Setup "AUTO  REPLAY  START", 					kDMDStd_AutoReplayStart,300000000,"INT[10000000:1000000000]", 5000000
    DMDMenu(11).Setup "BALL  PER  GAME",                      kDMDStd_BallsPerGame,   3,     "INT[3:5]", 1
    DMDMenu(12).Setup "FREE  PLAY",                          kDMDStd_FreePlay,       0,     "BOOL", 1
    DMDMenu(13).Setup "GRAND  CHAMPION  SCORE",               kDMDStd_HighScoreGC,    750000000,     "INT[10000000:1000000000]", 5000000
    DMDMenu(14).Setup "HIGH  SCORE  1",                       kDMDStd_HighScore1,     500000000,     "INT[10000000:1000000000]", 5000000
    DMDMenu(15).Setup "HIGH  SCORE  2",                       kDMDStd_HighScore2,     400000000,     "INT[10000000:1000000000]", 5000000
    DMDMenu(16).Setup "HIGH  SCORE  3",                       kDMDStd_HighScore3,     300000000,     "INT[10000000:1000000000]", 5000000
    DMDMenu(17).Setup "HIGH  SCORE  4",                       kDMDStd_HighScore4,     200000000,     "INT[10000000:1000000000]", 5000000
    DMDMenu(18).Setup "HIGH  SCORE  AWARDS",                  kDMDStd_HighScoreAward, 1,      "BOOL", 1
    DMDMenu(19).Setup "GRAND CHAMPION  AWARDS",              kDMDStd_GCAwards,       1,     "LST[0,0 CREDITS:1,1 CREDIT:2,2 CREDITS:3,3 CREDITS:4,4 CREDITS]", 1
    DMDMenu(20).Setup "HIGH  SCORE  1  AWARDS",                kDMDStd_HS1Awards,      1,     "LST[0,0 CREDITS:1,1 CREDIT:2,2 CREDITS:3,3 CREDITS]", 1
    DMDMenu(21).Setup "HIGH  SCORE  2  AWARDS",                kDMDStd_HS2Awards,      0,     "LST[0,0 CREDITS:1,1 CREDIT:2,2 CREDITS]", 1
    DMDMenu(22).Setup "HIGH  SCORE  3  AWARDS",                kDMDStd_HS3Awards,      0,     "LST[0,0 CREDITS:1,1 CREDIT]", 1
    DMDMenu(23).Setup "HIGH  SCORE  4  AWARDS",                kDMDStd_HS4Awards,      0,     "LST[0,0 CREDITS:1,1 CREDIT]", 1
    DMDMenu(24).Setup "MUSIC  ATTENUATION",                  kDMDStd_MusicAttenuation,0,    "INT[-60:60]", 5
    DMDMenu(25).Setup "SPEECH  ATTENUATION",                 kDMDStd_SpeechAttenuation,0,   "INT[-60:60]", 5
    DMDMenu(26).Setup "CLEAR  HIGHSCORE", 					0,		0,						"FUN", 1 
	DMDMenu(27).Setup "RESET  ALL  SETTINGS", 				1,		1,			            "FUN", 1 

    '***********************
    'Table-Specific Settings
    '***********************

    MinFetDMDSetting = 30
    DMDMenu(30).Setup "HOUSE  COMPLETE  EXTRA  BALL",         kDMDFet_HouseCompleteExtraBall,   8,      "LST[2,2:3,3:4,4:5,5:6,6:7,7:8,AUTO]", 1
    DMDMenu(31).Setup "CASUAL  PLAYER  MODE",                 kDMDFet_CasualPlayerMode,         0,      "BOOL", 1
    DMDMenu(32).Setup "PLAYER  HOUSE  CHOICE",                kDMDFet_PlayerHouseChoice,        0,      "LST[0,PLAYERS CHOICE:1,STARK:2,BARATHEON:3,LANNISTER:4,GREYJOY:5,TYRELL:6,MARTELL:7,TARGARYEN:8,RANDOM]", 1
    DMDMenu(33).Setup "DISABLE  RAMP  DROP  TARGET",          kDMDFet_DisableRampDropTarget,    0,      "BOOL", 1
    DMDMenu(34).Setup "ACTION  BUTTON  ACTION",               kDMDFet_ActionButtonAction,       1,      "LST[0,IRON BANK:1,HOUSE POWER]", 1
    DMDMenu(35).Setup "LANNISTER  BUTTONS  PER  GAME",        kDMDFet_LannisterButtonsPerGame,  8,      "INT[4:12]", 1
    DMDMenu(36).Setup "TARGARYEN  FREEZE  TIME",              kDMDFet_TargaryenFreezeTime,      12,     "INT[4:16]", 1
    DMDMenu(37).Setup "TARGARYEN  HOUSE  POWER",              kDMDFet_TargaryenHousePower,      3,      "LST[1,ONE DRAGON:2,TWO DRAGONS:3,COMPLETED]", 1
    DMDMenu(38).Setup "DIRE  WOLF  FREQUENCY",                kDMDFet_DireWolfFrequency,        6,      "LST[0,DISABLED:1,1 PER GAME:2,2 PER GAME:3,3 PER GAME:4,4 PER GAME:5,5 PER GAME:6,1 PER BALL", 1
    DMDMenu(39).Setup "SWORDS  UNLOCK  MULTIPLIER  X",        kDMDFet_SwordsUnlockMultiplier,   1,      "INT[0:2]", 1
    DMDMenu(40).Setup "BATTERING  RAM  HITS  FOR  MULTIPLIER",kDMDFet_RamHitsForMult,           1,      "INT[0:2]", 1
    DMDMenu(41).Setup "TWO  BANK  DIFFICULTY",                kDMDFet_TwoBankDifficulty,        1,      "LST[0,EASY:1,NORMAL:2,HARD]", 1
    DMDMenu(42).Setup "BLACKWATER  BALL  SAVE  TIMER",        kDMDFet_BWMB_SaveTimer,           10,     "INT[5:60]", 1
    DMDMenu(43).Setup "WALL  MULTIBALL  BALL  SAVE  TIMER",   kDMDFet_WallMB_SaveTimer,         20,     "INT[5:60]", 1
    DMDMenu(44).Setup "HAND  OF  THE  KING  BALL  SAVE  TIMER",kDMDFet_HOTKMB_SaveTimer,        20,     "INT[5:60]", 1
    DMDMenu(45).Setup "IRON  THRONE  BALL  SAVE  TIMER",      kDMDFet_ITMB_SaveTimer,           25,     "INT[5:60]", 1
    DMDMenu(46).Setup "MIDNIGHT  MADNESS  BALL  SAVE  TIMER", kDMDFet_MMMB_SaveTimer,           12,     "INT[5:60]", 1
    DMDMenu(47).Setup "WINTER  HAS  COME  BALL  SAVE  TIMER", kDMDFet_WHCMB_SaveTimer,          15,     "INT[5:60]", 1
    DMDMenu(48).Setup "CASTLE  MULTIBALL  BALL  SAVE  TIMER", kDMDFet_CMB_SaveTimer,            15,     "INT[5:60]", 1
    DMDMenu(49).Setup "CASTLE  LOOP  COMPLETES  HOUSE",       kDMDFet_CastleLoopCompletesHouse, 1,      "BOOL", 1
    DMDMenu(50).Setup "ALLOW  INACTIVITY  TO  PAUSE  TIMERS", kDMDFet_InactivityPauseTimer,     1,      "BOOL", 1
    DMDMenu(51).Setup "REDUCE  WINTER  IS  COMING  FLASHERS", kDMDFet_ReduceWHCLamps,           0,      "BOOL", 1
    DMDMenu(52).Setup "MIDNIGHT  MADNESS",                    kDMDFet_MidnightMadness,          1,      "BOOL", 1
    MaxFetDMDSetting = 52

    '** End Table-specific settings **

    ' Load Values from NVRAM. If not set, load default from array
    For i = 0 to MaxFetDMDSetting
        If DMDMenu(i).name <> "EMPTY" Then
            x = LoadValue(TableName, "DMDStd_"&DMDMenu(i).StdIdx)
            If (x <> "") Then 
                DMDStd(DMDMenu(i).StdIdx) = x
            Else
                DMDStd(DMDMenu(i).StdIdx) = DMDMenu(i).Deflt
            End If
        End if
    Next
  	x = LoadValue(TableName, "dmdCriticalChanged"):	If(x<>"") Then dmdCriticalChanged=True

    BallsPerGame = DMDStd(kDMDStd_BallsPerGame)


End Sub 
Dim DMDStd(200)

' Create the FlexDMD scene based on what service menu level we're at

Dim SvcScene
Sub CreateServiceDMDScene(lvl)
    Dim scene,i,max,offset,SvcFontSm,SvcFontBg
    offset=0
    Set SvcFontSm = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0)
    Set SvcFontBg = FlexDMD.NewFont("udmd-f7by10.fnt",vbWhite,vbWhite,0)
    Set scene = FlexDMD.NewGroup("service")
    If bServiceVol Then ' Not in Service menu, just adjusting volume
        scene.AddActor FlexDMD.NewLabel("svcl1",SvcFontSm,"USE +/- TO ADJUST VOLUME")
        scene.GetLabel("svcl1").SetAlignedPosition 64,0,FlexDMD_Align_Top
        scene.AddActor FlexDMD.NewLabel("svcl2",SvcFontBg,"VOLUME "&MasterVol)
        scene.GetLabel("svcl2").SetAlignedPosition 64,14,FlexDMD_Align_Top
    Else
        Select Case lvl
            Case kMenuNone
                scene.AddActor FlexDMD.NewLabel("svcl1",SvcFontSm,"GAME OF THRONES PREMIUM"&vbLf&"V1.37.0    "&myVersion&"  SYS. 2.31.0")
                scene.GetLabel("svcl1").SetAlignedPosition 64,0,FlexDMD_Align_Top
                scene.AddActor FlexDMD.NewLabel("svcl2",SvcFontBg,"SERVICE MENU")
                scene.GetLabel("svcl2").SetAlignedPosition 64,14,FlexDMD_Align_Top
                scene.AddActor FlexDMD.NewLabel("svcl3",SvcFontSm,"PRESS SELECT TO CONTINUE")
                scene.GetLabel("svcl3").SetAlignedPosition 64,25,FlexDMD_Align_Top
            Case kMenuTop,kMenuAdj  ' one of the icon levels. Create row of icons
                If lvl=kMenuTop Then 
                    max = 6 
                    scene.AddActor FlexDMD.NewLabel("svcl4",SvcFontSm,TopArray(0))
                else 
                    max = 4 : offset=17
                    scene.AddActor FlexDMD.NewLabel("svcl4",SvcFontSm,AdjArray(0))
                End if
                For i = 0 to max
                    scene.AddActor FlexDMD.NewImage("iconoff"&i,"got-svciconoff"&lvl&i&".png")
                    With scene.GetImage("iconoff"&i)
                        .SetAlignedPosition offset+i*17,0,FlexDMD_Align_TopLeft
                        if i = 0 Then .Visible = 0
                    End With
                    scene.AddActor FlexDMD.NewImage("iconon"&i,"got-svciconon"&lvl&i&".png")
                    With scene.GetImage("iconon"&i)
                        .SetAlignedPosition offset+i*17,0,FlexDMD_Align_TopLeft
                        If i > 0 Then .Visible = 0
                    End With
                Next
                scene.GetImage("iconoff0").Visible = 0
                scene.GetLabel("svcl4").SetAlignedPosition 64,25,FlexDMD_Align_Top
            Case kMenuAdjStd,kMenuAdjGame ' Adjustment level, create the base adjustment screen
                scene.AddActor FlexDMD.NewLabel("svcl1",SvcFontSm,"ADJUSTMENT")
                scene.GetLabel("svcl1").SetAlignedPosition 64,0,FlexDMD_Align_Top
                scene.AddActor FlexDMD.NewLabel("svcl2",FlexDMD.NewFont("udmd-f3by7.fnt",vbWhite,vbWhite,0),"SETTING NAME")
                scene.GetLabel("svcl2").SetAlignedPosition 64,7,FlexDMD_Align_Top
                scene.AddActor FlexDMD.NewLabel("svcl3",FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt",vbWhite,vbWhite,0),"VALUE")
                scene.GetLabel("svcl3").SetAlignedPosition 64,16,FlexDMD_Align_Top
                scene.AddActor FlexDMD.NewLabel("svcl4",SvcFontSm,"INSTALLED/FACTORY DEFAULT")
                scene.GetLabel("svcl4").SetAlignedPosition 64,25,FlexDMD_Align_Top
        End Select
    End if
    Set SvcScene = scene
    DMDDisplayScene SvcScene
End Sub

Sub UpdateServiceDMDScene(line3)
    Dim max,i
    If serviceLevel = kMenuTop Then max = 6 else max = 4
    FlexDMD.LockRenderThread
    If bServiceVol Then ' Just update the Volume level
        With scene.GetLabel("svcl2")
            .Text = "VOLUME "&MasterVol
            .SetAlignedPosition 64,14,FlexDMD_Align_Top
        End With
    Else
        If serviceLevel = kMenuTop Or serviceLevel = kMenuAdj Then
            For i = 0 to max
                With SvcScene.GetImage("iconon"&i)
                    If i=serviceIdx Then .Visible=1 Else .Visible=0
                End With
                With SvcScene.GetImage("iconoff"&i)
                    If i=serviceIdx Then .Visible=0 Else .Visible=1
                End With
            Next
        End If
        Select Case serviceLevel
            Case kMenuTop: 
                SvcScene.GetLabel("svcl4").Text = TopArray(serviceIdx)
                SvcScene.GetLabel("svcl4").SetAlignedPosition 64,25,FlexDMD_Align_Top
            Case kMenuAdj: 
                SvcScene.GetLabel("svcl4").Text = AdjArray(serviceIdx)
                SvcScene.GetLabel("svcl4").SetAlignedPosition 64,25,FlexDMD_Align_Top
            Case kMenuAdjStd,kMenuAdjGame
                With SvcScene.GetLabel("svcl1")
                    if serviceIdx < MinFetDMDSetting Then
                        .Text = "STANDARD ADJUSTMENT #" & serviceIdx & "(" & DMDMenu(serviceIdx).StdIdx & ")"
                    Elseif DMDMenu(serviceIdx).ValType<>"FUN" then 
                        .Text = "GAME ADJUSTMENT #" & serviceIdx & "(" & DMDMenu(serviceIdx).StdIdx & ")"
                    Else
                        .Text = "GAME FUNCTION #" & serviceIdx
                    End If
                    .SetAlignedPosition 64,0,FlexDMD_Align_Top
                End With
                With SvcScene.GetLabel("svcl2")
                    .Text = DMDMenu(serviceIdx).Name
                    .SetAlignedPosition 64,7,FlexDMD_Align_Top
                End With
                With SvcScene.GetLabel("svcl3")
                    .Text = line3
                    .SetAlignedPosition 64,16,FlexDMD_Align_Top
                End with
                If bServiceEdit Then 
                   With SvcScene.GetLabel("svcl2")
                        .ClearActions()
                        .Visible=1
                    End With
                    With SvcScene.GetLabel("svcl3")
                        .ClearActions()
                        .Visible=1
                    End With
                    DelayBlinkActor SvcScene.GetLabel("svcl3"),0.75,0.25,999
                Else
                    With SvcScene.GetLabel("svcl3")
                        .ClearActions()
                        .Visible=1
                    End With
                    With SvcScene.GetLabel("svcl2")
                        .ClearActions()
                        .Visible=1
                    End With
                    DelayBlinkActor SvcScene.GetLabel("svcl2"),0.75,0.25,999
                End If
                With SvcScene.GetLabel("svcl4")
                    If CInt(DMDStd(DMDMenu(serviceIdx).StdIdx)) = DMDMenu(serviceIdx).Deflt Then .Visible = 1 Else .Visible = 0
                End With
        End Select
    End if
    FlexDMD.UnlockRenderThread
End Sub


'  END DAPHISHBOWL - STERN DMD

    

'*********
'   LUT
'*********

Dim bLutActive, LUTImage
Sub LoadLUT
    Dim x
    bLutActive = False
    x = LoadValue(myGameName, "LUTImage")
    If(x <> "")Then LUTImage = x Else LUTImage = 0
    table1.ColorGradeImage = "LUT"&LUTImage
End Sub

Sub SaveLUT
    SaveValue myGameName, "LUTImage", LUTImage
End Sub

Sub NxtLUT:LUTImage = (LUTImage + 1)MOD 16: table1.ColorGradeImage = "LUT"&LUTImage : SaveLUT : End Sub

Sub UpdateLUT
    table1.ColorGradeImage = "LUT"&LUTImage
End Sub




'********************************************************************************************
' Only for VPX 10.2 and higher.
' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first myVersion

    If TypeName(MyLight) = "Light" Then

        If FinalState = 2 Then
            FinalState = MyLight.State 'Keep the current light state
        End If
        MyLight.BlinkInterval = BlinkPeriod
        MyLight.Duration 2, TotalPeriod, FinalState
    ElseIf TypeName(MyLight) = "Flasher" Then

        Dim steps

        ' Store all blink information
        steps = Int(TotalPeriod / BlinkPeriod + .5) 'Number of ON/OFF steps to perform
        If FinalState = 2 Then                      'Keep the current flasher state
            FinalState = ABS(MyLight.Visible)
        End If
        MyLight.UserValue = steps * 10 + FinalState 'Store # of blinks, and final state

        ' Start blink timer and create timer subroutine
        MyLight.TimerInterval = BlinkPeriod
        MyLight.TimerEnabled = 0
        MyLight.TimerEnabled = 1
        ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
    End If
End Sub

'******************************************
' Change light color - simulate color leds
' changes the light color and state
' 11 colors: red, orange, amber, yellow...
'******************************************
' in this table this colors are use to keep track of the progress during the modes

'colors
Const magenta = 14
Const cyan = 13
Const midblue = 12
Const ice = 11
Const red = 10
Const orange = 9
Const amber = 8
Const yellow = 7
Const darkgreen = 6
Const green = 5
Const blue = 4
Const darkblue = 3
Const purple = 2
Const white = 1
Const teal = 0
'******************************************
' Change light color - simulate color leds
' changes the light color and state
' colors: red, orange, yellow, green, blue, white, purple, amber
' Note: Colors tweaked slightly to match GoT color scheme
'******************************************
' Modified to handle changing light colors while LightSeqPlayfield is running
' If it's running then
'  - If state is "-1", or sequence isn't running a coloured sequence, then change the colour right on the light
'  - Otherwise, save the colour change in the LightSaveState table and let it be restored once the sequence ends

Sub SetLightColor(n, col, stat) 'stat 0 = off, 1 = on, 2 = blink, -1= no change
    Dim color,colorfull
    Select Case col
        Case magenta
            color = RGB(16,4,16)
            colorfull = RGB(255,64,224)
        Case cyan
            color = RGB(0,18,18)
            colorfull = RGB(0, 224, 240)
        Case midblue
            color = RGB(0,0,18)
            colorfull = RGB(0,0,255)
        Case ice
            color = RGB(0, 18, 18)
            colorfull = RGB(192, 255, 255)
        Case red
            color = RGB(18, 0, 0)
            colorfull = RGB(255, 0, 0)
        Case orange
            color = RGB(18, 3, 0)
            colorfull = RGB(255, 64, 0)
        Case amber
            color = RGB(193, 49, 0)
            colorfull = RGB(255, 153, 0)
        Case yellow
            color = RGB(18, 18, 0)
            colorfull = RGB(255, 255, 0)
        Case darkgreen
            color = RGB(0, 8, 0)
            colorfull = RGB(0, 64, 0)
        Case green
            color = RGB(0, 16, 0)
            colorfull = RGB(0, 192, 0)
        Case blue
            color = RGB(0, 18, 18)
            colorfull = RGB(4, 192, 255)
        Case darkblue
            color = RGB(0, 8, 8)
            colorfull = RGB(0, 64, 64)
        Case purple
            color = RGB(64, 0, 96)
            colorfull = RGB(128, 0, 192)
        Case white
            color = RGB(255, 197, 143)
            colorfull = RGB(255, 252, 224)
        Case teal
            color = RGB(1, 64, 62)
            colorfull = RGB(2, 128, 126)
    End Select
    If stat = -1 Or LightSeqPlayfield.UserValue = 0 Then n.color = color : n.colorfull = colorfull Else SavePlayfieldLightColor n,color,colorfull

    If stat <> -1  Then
        n.State = 0
        n.State = stat
    End If
End Sub

Sub SetFlashColor(n, col, stat) 'stat 0 = off, 1 = on, -1= no change - no blink for the flashers
    Select Case col
        Case red
            n.color = RGB(255, 0, 0)
        Case orange
            n.color = RGB(255, 64, 0)
        Case amber
            n.color = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 64, 64)
        Case purple
            n.color = RGB(128, 0, 192)
        Case white
            n.color = RGB(255, 252, 224)
        Case teal
            n.color = RGB(2, 128, 126)
    End Select
    If stat <> -1 Then
        n.Visible = stat
    End If
End Sub

Sub ChangeGi(col) 'changes the gi color
    Dim bulb
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
End Sub

Sub GiOn
    DOF 118, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
    GameGiOn
End Sub

Sub GiOff
    DOF 118, DOFOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
    GameGiOff
End Sub

' GI, light & flashers sequence effects

Sub GiEffect(n)
    Dim ii
    LightSeqGi.StopPlay
    Select Case n
        Case 0 'all off
            LightSeqGi.Play SeqAlloff
        Case 1 'all blink
            LightSeqGi.UpdateInterval = 20
            LightSeqGi.Play SeqBlinking, , 15, 10
        Case 2 'random
            LightSeqGi.UpdateInterval = 20
            LightSeqGi.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqGi.UpdateInterval = 20
            LightSeqGi.Play SeqBlinking, , 10, 10
        Case 4 'seq up
            LightSeqGi.UpdateInterval = 3
            LightSeqGi.Play SeqUpOn, 25, 3
        Case 5 'seq down
            LightSeqGi.UpdateInterval = 3
            LightSeqGi.Play SeqDownOn, 25, 3
        Case 6 ' blink once
            LightSeqGi.UpdateInterval = 20
            LightSeqGi.Play SeqBlinking,0,1,10
            'LightSeqGi.Play SeqAlloff
    End Select
End Sub

Sub GiIntensity(i)
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = i
    Next
    For each bulb in aFiLights
        bulb.IntensityScale = i
    Next
End Sub

' In order to be able to manipulate RGB lamp colour and fade rate during sequences,
' we have to save all of that information into arrays so it can be restored after the sequence
' (the built-in VPX light sequencer only restores lamp on/off state, not colour/fade, etc)
' Right now we only save colour and fade speed, but we could save anything 

' This is called before a sequence that requires custom colours is run
Sub SavePlayfieldLightState
    Dim i,a
    i = 0
    For each a in aPlayfieldLights
        LightSaveState(i,0) = a.State
        LightSaveState(i,1) = a.Color
        LightSaveState(i,2) = a.Colorfull
        LightSaveState(i,3) = a.FadeSpeedUp
        LightSaveState(i,4) = a.FadeSpeedDown
        i = i + 1
    Next
End Sub

' This is called once at table start. It saves all lamp names into an array, whose index is equal to the number after "li"
' For this to work, all lamp objects that you want to be saved must be named "li<x>" where x is a not-zero-padded number
Sub SavePlayfieldLightNames
    Dim i,j,a
    i = 0
    On Error Resume Next
    For each a in aPlayfieldLights
        If Left(a.Name,2) = "li" Then      
            j = CInt(Mid(a.Name,3))
            If j > 0 And Not Err Then LightNames(j) = i
        End If
        i = i + 1
    Next
    On Error Goto 0
End Sub

' If SetLightColor is called while a sequence is running, it calls this sub to save the desired colour into an array so it can be
' restored once the sequence ends
Sub SavePlayfieldLightColor(a,col,colf)
    On Error Resume Next
    Dim j
    If Left(a.Name,2) = "li" Then      
        j = CInt(Mid(a.Name,3))
        If j > 0 And Not Err Then LightSaveState(LightNames(j),1) = col : LightSaveState(LightNames(j),2) = colf
    End If
    On Error Goto 0
End Sub

' This is called by the LightSeq_PlayDone handler to restore light states after the sequence ends
Sub RestorePlayfieldLightState(state)
    Dim i,a
    i = 0
    For each a in aPlayfieldLights
        If state Then a.State = LightSaveState(i,0)
        a.Color = LightSaveState(i,1)
        a.ColorFull = LightSaveState(i,2)
        a.FadeSpeedUp = LightSaveState(i,3)
        a.FadeSpeedDown = LightSaveState(i,4)
        i = i + 1
    Next
    For i = 0 to 69
        Lampz.FadeSpeedUp(i) = LightSaveState(0,3)
        Lampz.FadeSpeedDown(i) = LightSaveState(0,4)
    Next
End Sub

' Set Playfield lights to slow fade a color
Sub PlayfieldSlowFade(color,fadespeed)
    Dim a
    For each a in aPlayfieldLights
        If color >= 0 Then SetLightColor a,color,-1
    Next
    For a = 0 to 69 '	for x = 0 to 140 : Lampz.FadeSpeedUp(x) = 1/2 : Lampz.FadeSpeedDown(x) = 1/10 : next

        Lampz.FadeSpeedUp(a) = fadespeed
        Lampz.FadeSpeedDown(a) = fadespeed/4
    Next
End Sub

'********************
' Real Time updates
'********************
'used for all the real time updates

Sub Realtime_Timer
    RollingUpdate
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
    LeftUFlipperTop.RotZ = LeftUFlipper.CurrentAngle
    RightUFlipperTop.RotZ = RightUFlipper.CurrentAngle
' add any other real time update subs, like gates or diverters, flippers
End Sub


'*******************************
' Attract mode support
' (should be table-independent)
'*******************************

Sub StartAttractMode
    'StartLightSeq
    GameStartAttractMode
End Sub

Sub StopAttractMode
    GameStopAttractMode
    'LightSeqAttract.StopPlay
End Sub

Sub LightSeqAttract_PlayDone()
    RestorePlayfieldLightState False    ' Restore color and fade speed but not state. Sequencer takes care of that
End Sub

Sub LightSeqTilt_PlayDone()
    LightSeqTilt.Play SeqAllOff
End Sub

Sub LightSeqSkillshot_PlayDone()
    LightSeqSkillshot.Play SeqAllOff
End Sub

Sub LightSeqPlayfield_PlayDone()
    If LightSeqPlayfield.UserValue <> 0 Then RestorePlayfieldLightState False
    LightSeqPlayfield.UserValue = 0
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
    UPFFlasher001.visible = 0 : UPFFlasher002.visible = 0
    fl242.state = 0 : fl237.state = 0
End Sub

'**************************************
' Non-game-specific Light sequences
'**************************************

Sub LightEffect(n)
    Select Case n
        Case 0 ' all off
            LightSeqInserts.Play SeqAlloff
        Case 1 'all blink
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 15, 10
        Case 2 'random
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 10, 10
        Case 5 'top down
            LightSeqPlayfield.UpdateInterval = 4
            LightSeqPlayfield.Play SeqDownOn, 15, 2
        Case 6 'down to top
            LightSeqPlayfield.UpdateInterval = 4
            LightSeqPlayfield.Play SeqUpOn, 15, 1
    End Select
End Sub


' Lamp Fader code. Right now this is only used by a couple of Flashers but could be used for any lamp
' Copied from Ghostbuster ESD table.

Dim bulb
Dim LampState(10), FadingLevel(10), chgLamp(10)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
Lamp2Timer.Interval = -1 'lamp fading speed
Lamp2Timer.Enabled = 1

Sub InitLamps()
    Dim x
    For x = 0 to 10
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
    Next
End Sub

Sub Lamp2Timer_Timer()
    Dim num, chg, ii, vid, prevlvl, targetlvl
 
    For ii = 0 To 10
        prevlvl = LampState(ii)		'previous frame value
        targetlvl = chgLamp(ii)					'latest frame value

        'Fadinglevel filtering
        ' 0 = static off -> skipped in subs
        ' 1 = static on -> skipped in subs
        ' 3 = not used
        ' 4 = fading down -> subs have equation that will fade the level down
        ' 5 = fading up -> This is not actually in use with vpinspa as it provides slow fade ups already

        if prevlvl <> targetlvl then
            if prevlvl < targetlvl Then
                FadingLevel(ii) = 5							'fading up...
                LampState(ii) = targetlvl						'as vpinspa output is modulated, we write the level here
            end if

            if prevlvl > targetlvl Then
                FadingLevel(ii) = 4 							'fading down...
                'LampState(ii) = (prevlvl + targetlvl) / 2		'Skipping level set, let the subs to fade it slow
            end if
        Else							'no change in intensity. Vpinspa outputs these occasionally, but we should let them skip in fading subs
            if targetlvl = 255 Then
                FadingLevel(ii ) = 1 'on already
            elseif targetlvl = 0 Then
                FadingLevel(ii ) = 0 'off already
            Else
                'debug.print "same level outputted for consecutive frames -> don't care"
            end if
        end if
    Next
    
    UpdateLamps
End Sub


Sub SingleLamp(lampobj, id)
	dim intensity

	if FadingLevel(id) > 1 then

		Intensity = LampState(id)
'		debug.print "SingleLamp: " & intensity
		lampobj.IntensityScale = LampState(id) / 255

		if FadingLevel(id) = 4 Then
			Intensity = LampState(id) * 0.7 - 0.01
			if Intensity <= 0 then : Intensity = 0 : FadingLevel(id) = 0 : lampobj.state = 0
		end if

		if FadingLevel(id) = 5 Then
			if Intensity = 255 then FadingLevel(id) = 1
            lampobj.state = 1
		end if

		LampState(id) = Intensity
	end if
End Sub 

Sub SingleLampM(lampobj, id)
    if FadingLevel(id) > 1 then
		lampobj.IntensityScale = LampState(id) / 255
        if LampState(id) > 0 Then lampobj.state=1 Else lampobj.state=0
	end If
End Sub 

Sub SetLamp(id,state)
    Select Case state
        Case 0: id.state = 0
        Case 1: id.state = 1
        Case 3: id.state = 1 : vpmTimer.addTimer 16,id.name&".state=0 '"
    End Select
End Sub

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' lost a ball :-( check to see how many balls are on the playfield.
' if only one then decrement the remaining count AND test for End of game
' if more than 1 ball (multi-ball) then kill of the ball but don't create
' a new one. This is not table-specific code - calls out to subs specific to this game
'
Sub Drain_Hit()
    ' Destroy the ball
    Drain.DestroyBall
    BallsOnPlayfield = BallsOnPlayfield - 1 : If BallsOnPlayfield < 0 Then BallsOnPlayfield = 0
    ' Exit Sub ' only for debugging - this way you can add balls from the debug window

    ' pretend to knock the ball into the ball storage mech
    'PlaySoundAt "fx_drain", Drain
    RandomSoundDrain Drain

    ' Table-specific
    If bGameInPLay = False Or bMadnessMB = 1 Then Exit Sub 'don't do anything, just delete the ball

    'if Tilted the end Ball Mode
    If Tilted Then
        StopEndOfBallModes
    End If

    ' if there is a game in progress AND it is not Tilted
    If(bGameInPLay = True)AND(Tilted = False)Then

        ' is the ball saver active, (or ready to be activated but the ball sewered right away)
        If(bBallSaverActive = True and bEarlyEject = False) Or (bBallSaverReady = True AND DMDStd(kDMDStd_BallSave) <> 0 And bBallSaverActive = False) Then
            DoBallSaved 0
        Else
        	bEarlyEject = False
            ' cancel any multiball if on last ball (ie. lost all other balls)
            ' NOTE: This is GoT-specific. Remove tmrBWMultiBallRelease check for other tables
            If BallsOnPlayfield-RealBallsInLock < 2 And bMultiBallMode Then
                If tmrBWmultiballRelease.Enabled Then 
                    debug.print "Ball drained in MB mode but tmrBWMultiballRelease is enabled"
                Else
                    ' not in multiball mode any more
                    bMultiBallMode = False
                    ' you may wish to change any music over at this point and

                    ' turn off any multiball specific lights
                    ChangeGi white
                    'stop any multiball modes
                    StopMBmodes
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield-RealBallsInLock = 0 And tmrBWmultiballRelease.Enabled = False)Then
                ' End Mode and timers

                ChangeGi white
                ' Show the end of ball animation
                ' and continue with the end of ball
                ' DMD something?
                StopEndOfBallModes
                vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, if there is no animation then move to the end of ball
            End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
    'debug.print "ball in plunger lane"
    ' some sound according to the ball position
    'PlaySoundAt "fx_sensor", swPlungerRest
    bBallInPlungerLane = True
    ' turn on Launch light is there is one
    'LaunchLight.State = 2

    'be sure to update the Scoreboard after the animations, if any

    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"
        vpmtimer.addtimer 1000, "PlungerIM.AutoFire:DOF 120, DOFPulse:PlaySoundAt ""fx_kicker"", swPlungerRest:bAutoPlunger = False:bAutoPlunged = True '"
    End If
    'Start the Selection of the skillshot if ready
    ' GoT Premium/LE has no skill shot
    If bSkillShotReady Then
        PlaySong "mu_shooterlane"
        UpdateSkillshot()
        ' show the message to shoot the ball in case the player has fallen sleep
        SwPlungerCount = 0
        swPlungerRest.TimerEnabled = 1
    End If
    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    If bSkillShotReady Then
        ChangeSong
        ResetSkillShotTimer.Enabled = 1
    End If
    bJustPlunged = True
    tmrJustPlunged.Interval = 2500
    tmrJustPlunged.Enabled = 1

    GameDoBallLaunched
    bAutoPlunged = False
    bBallSaved = False
End Sub

Sub tmrJustPlunged_Timer : bJustPlunged = False : End Sub

' Not used in this game. GoT has its own BallSaver logic
' Sub EnableBallSaver(seconds)
'     'debug.print "Ballsaver started"
'     ' set our game flag
'     bBallSaverActive = True
'     bBallSaverReady = False
'     ' start the timer
'     BallSaverTimerExpired.Interval = 1000 * seconds
'     BallSaverTimerExpired.Enabled = True
'     BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
'     BallSaverSpeedUpTimer.Enabled = True
'     ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
'     LightShootAgain.BlinkInterval = 160
'     LightShootAgain.State = 2
' End Sub

' ' The ball saver timer has expired.  Turn it off AND reset the game flag
' '
' Sub BallSaverTimerExpired_Timer()
'     'debug.print "Ballsaver ended"
'     BallSaverTimerExpired.Enabled = False
'     ' clear the flag
'     bBallSaverActive = False
'     ' if you have a ball saver light then turn it off at this point
'     LightShootAgain.State = 0
' End Sub

' Sub BallSaverSpeedUpTimer_Timer()
'     'debug.print "Ballsaver Speed Up Light"
'     BallSaverSpeedUpTimer.Enabled = False
'     ' Speed up the blinking
'     LightShootAgain.BlinkInterval = 80
'     LightShootAgain.State = 2
' End Sub



















'*********************************************
' Kiss-Specific code begins here
'*********************************************

Const Duece=1
Const Hotter=2
Const LickItUp=3
Const ShoutIt=4
Const DetroitRockCity=5
Const RockAllNight=6
Const LoveItLoud=7
Const BlackDiamond=8

Const LeftOrbit=1
Const StarTargets=2
Const CenterRamp=3
Const DemonShot=4
Const RightRamp=5
Const RightOrbit=6


Dim PlayerMode
Dim LockLit
Dim PlayerState(4)
Dim CompletedModes              ' bitmask of completed modes (songs)

Dim bKissTargets(4)
Dim bArmyTargets(4)
Dim bLockTargets(2)
Dim bStarTargets(4)
Dim bKissLights(4)
Dim bArmyLights(4)
Dim KissTargetsCompleted        ' Number of times KISS target bank has been completed
Dim ArmyTargetsCompleted

Dim ModeSongs
ModeSongs = Array("DUECE","HOTTER THAN HELL","LICK IT UP","SHOUT IT OUT LOUD","DETROIT ROCK CITY","ROCK & ROLL ALL NIGHT","LOVE IT LOUD","BLACK DIAMOND")


Class cPState
    Dim i
    Dim myKissTargets(4)
    Dim myArmyTargets(4)
    Dim myStarTargets(4)
    Dim myLockTargets(2)
    Dim myCompletedModes
    Dim myKissTargetsCompleted
    Dim myArmyTargetsCompleted

    Sub Save
        Dim i
        For i = 1 to 4
            myKissTargets(i)=bKissTargets(i)
            myArmyTargets(i)=bArmyTargets(i)
            myStarTargets(i)=bStarTargets(i)
        Next
        For i=0 to 1: myLockTargets(i)=bLockTargets(i):Next
        myCompletedModes=CompletedModes
        myKissTargetsCompleted = KissTargetsCompleted
        myArmyTargetsCompleted = ArmyTargetsCompleted
    End Sub

    Sub Restore
        Dim i
        For i = 1 to 4
            bKissTargets(i)=myKissTargets(i)
            bArmyTargets(i)=myArmyTargets(i)
            bStarTargets(i)=myArmyTargets(i)
        Next
        For i=0 to 1: bLockTargets(i)=myLockTargets(i):Next
        CompletedModes=myCompletedModes
        KissTargetsCompleted = myKissTargetsCompleted
        ArmyTargetsCompleted = myArmyTargetsCompleted
    End Sub
End Class

Dim ModeLights
ModeLights = Array(li90,li90,li100,li105,li111,li124,li130)
Dim ModeColours
ModeColours = Array(ice,rockandroll=ice,detroitrockcity=yellow,lickitup=white,loveitloud=amber)

Class cMode
    Dim ModeState
    Dim CurrentMode
    Dim ShotMask
    Dim ModeShotsCompleted  ' Used for scoring - resets at beginning of ball
    Dim ShotsMade           ' Used to track progress in Detroit Rock City and Love It Loud
    Dim BaseScore
    Dim ModeScore
    Dim bModeComplete
    Dim ModeColour

    Public Property Get GetMode : GetMode = CurrentMode : End Property
    Public Property Let SetMode(m) : CurrentMode = m : End Property

    Sub SetModeLights
        Dim i,t
        t=False
        For i = 1 to 6
            If (ShotMask And (2^i)) > 0 And bModeComplete=False Then
                If t=False and CurrentMode=DetroitRockCity Then
                    SetLightColor ModeLights(i),ice,2
                    t=true
                Else
                    SetLightColor ModeLights(i),ice,1
                End If
            Else
                SetLightColor ModeLights(i),ice,0
            End If
        Next
        If (ShotMask And 128) > 0 Then SetKissLights True
    End Sub

    Sub StartMode
        Dim i
        ModeState = 1
        bModeComplete = False
        ModeShotsCompleted = 0
        BaseScore = 250000
        Select Case CurrentMode
            Case Duece:           ModeColour=blue   : BaseScore = 300000 : ShotMask=6      'two left-most shots lit
            Case Hotter:          ModeColour=white  : ShotMask=128   'KISS targets
            Case LickItUp:        ModeColour=magenta: BaseScore = 250000 : i=RndNbr(5) : ShotMask = 2^i + 2^(i+1)    ' two neighbouring shots lit
            Case ShoutIt:         ModeColour=cyan   : ShotMask = 2^CenterRamp   ' TODO: Figure out scoring for ShoutIt 
            Case DetroitRockCity: ModeColour=yellow : BaseScore = 200000 : ShotMask = 126    ' All major shots lit
            Case RockAllNight:    ModeColour=ice    : BaseScore = 500000 : ShotMask = 2^CenterRamp
            Case LoveItLoud:      ModeColour=amber  : BaseScore = 300000 : ShotMask = 2^LeftOrbit + 2^RightOrbit
            Case BlackDiamond:    ModeColour=green  : BaseScore = 300000 : Shotmask = 2^RightOrbit
        End Select

        SetModeLights
    End Sub

    Sub ResetForNewBall
        ModeScore = 0
        ModeShotsCompleted = 0  ' didn't reset for Rock&Roll All Night
    End Sub

    ' Check switch hits against mode (song) rules
    Sub CheckModeHit(sw)
        Dim AwardScore,i

        if (ShotMask And (2^sw)) = 0 Or bModeComplete Then Exit Sub' Shot wasn't lit

        ModeShotsCompleted = ModeShotsCompleted + 1
        'TODO: Award score is doubled sometimes
        AwardScore = BaseScore*ModeShotsCompleted

        Select Case CurrentMode
            Case Duece:
                ModeState = ModeState+1
                If ModeState >= 9 Then
                    bModeComplete = True
                ElseIf ModeState < 6 Then
                    ShotMask = 2^ModeState+2^(ModeState+1)
                Else ' Set up Duece light timer to cycle lights across playfield
                    tmrDueceMode.Enabled = False
                    tmrDueceMode.Interval = (9-ModeState)*500
                    tmrDueceMode.Enabled = True
                    ShotMask = 6
                End If
            Case Hotter:
                ModeState = ModeState+1
                If ModeState >= 9 Then
                    bModeComplete = True
                ElseIf (ModeState MOD 2) = 0 Then
                    ShotMask = 2^RndNbr(6)
                Else
                    ShotMask = 128
                End If
                If ModeState > 7 Then AwardScore=AwardScore*2
            Case LickItUp:
                ModeState = ModeState+1
                If ModeState >= 9 Then
                    bModeComplete = True
                Else
                    i = RndNbr(5) : ShotMask = 2^i + 2^(i+1)
                End If
            Case ShoutIt:
                ' Mode needs fleshing out. This is all we have:
                ' "Shoot center ramp, then orbits, then STAR targets, right ramp and demon. Scoring is (?)"
                ShotMask = ShotMask Xor (2^sw)
                If ShotMask = 0 Then ModeState=ModeState+1
                If ModeState >= 4 Then
                    bModeComplete = True
                Else
                    Select Case ModeState
                        Case 2: 66  ' left orbit + right orbit
                        Case 3: 52  ' STAR target + demon + right ramp
                    End Select
                End If
            Case DetroitRockCity:
                ShotMask = ShotMask Xor (2^sw)
                ShotsMade=ShotsMade+1
                If ModeState=2 And ShotsMade >= 3 Then bModeComplete=True
                If ShotMask = 0 Then ModeState=ModeState+1 : ShotMask = 126 : ShotsMade=0
                If (ShotMask And ((2^sw)-1)) = 0 Then AwardScore=AwardScore*2   'Leftmost lit shot was hit
            Case RockAllNight:
                Select Case ModeState
                    Case 1: ShotMask=2^RightOrbit
                    Case 2: ShotMask=6 'Left orbit + STAR
                    Case 3: ShotMask=96 ' Right orbit + right ramp
                    Case 4,6: ShotMask=14 ' left orbit, STAR, Center ramp
                    Case 5,7: ShotMask=112 ' right orbit, right ramp, demon
                    Case 8: bModeComplete = True
                End Select
                ModeState=ModeState+1
            Case LoveItLoud:
                If ModeState=1 Then
                    ShotMask = ShotMask Xor (2^sw)
                    If ShotMask=0 Then ModeState=2 : ShotsMade=0 : ShotMask=60 ' 4 inner shots
                Else
                    ShotMask = 60 Xor (2^sw)
                    ShotsMade=ShotsMade+1
                    If ShotsMade >= 8 Then bModeComplete=True
                End If
            Case BlackDiamond:
                ModeState = ModeState+1
                If ModeState >= 9 Then
                    bModeComplete = True
                Else
                    If sw=RightOrbit Then ShotMask = 2^(RndNbr(4)+1) Else ShotMask=64
                End if
        End Select
        SetModeLights

        ' TODO: Do a 'hit' scene
        AddScore AwardScore
        ModeScore = ModeScore+AwardScore
    End Sub

    Sub BlackDiamondPopHit
        Dim OldMask
        If CurrentMode <> BlackDiamond Or (ModeState MOD 2) <> 0 Then Exit Sub
        OldMask = ShotMask
        Do
            ShotMask = 2^(RndNbr(4)+1)
        Loop While ShotMask=OldMask
        SetModeLights
    End Sub

    Sub DueceTimer
        ShotMask=ShotMask*2
        If ShotMask=96 Then ShotMask=6
        SetModeLights
    End Sub

End Class

Sub tmrDueceMode : Mode(CurrentPlayer).DueceTimer : End Sub




















Sub SetPlayfieldLights
    TurnOffPlayfieldLights
    SetKissLights False
End Sub



Dim KissLights
KissLights = Array(i67,i67,i68,i69,i70)
Sub SetKissLights(InMode)
    Dim i
    If InMode Then
        ' TODO: Change blink interval and pattern to mode 
        For i = 1 to 4 : SetLightColor KissLights(i),White,2 : Next
    Else
        For i = 1 to 4 : SetLightColor KissLights(i),White,ABS(bKissTargets(i)) : Next
    End if
End Sub









Sub VPObjects_Init
    Dim i
    ReplayScore = DMDStd(kDMDStd_AutoReplayStart)
    SetDefaultPlayfieldLights   ' Sets all playfield lights to their default colours
    'If table has a power-on sound, set it here
    'vpmTimer.AddTimer 1000,"PlaySoundVol ""say-whathouse"",VolCallout '"
End Sub

Sub ResetForNewGame
    Dim i
    For i = 0 to 4
        bKissTargets(i)=False
        bArmyTargets(i)=False
        bStarTargets(i)=False
    Next
    bLockTargets(0)=False:bLockTargets(1)=False
    LockLit=0
    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusHeldPoints(i) = 0
        BonusMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ReplayAwarded(i) = False        
        Set PlayerState(i) = New cPState
        Set Mode(i) = New cMode
    Next

    ' initialise any other flags
    bGameInPLay = True
    bMadnessMB = 0
    GameTimeStamp = 0
    'resets the score display, and turn off attract mode
    StopAttractMode
    StopSound "got-track-gameover"
    GiOn
    LightSeqPlayfield.UserValue = 0

    bAlternateScoreScene=False
    ScoreScene = Empty

    ' Just in case
    BallsOnPlayfield = 0
	RealBallsInLock=0
    BallsInLock = 0

    TotalGamesPlayed = TotalGamesPlayed + 1
    CurrentPlayer = 1
    PlayersPlayingGame = 1
    Tilt = 0

    ' initialise Game variables
    Game_Init()

    DMDBlank

    tmrGame.Interval = 100
    tmrGame.Enabled = 1

    vpmtimer.addtimer 1500, "FirstBall '"
End Sub




' Create a new ball on the Playfield

Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass

    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' kick it out..
    PlaySoundAt SoundFXDOF("fx_Ballrel", 123, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 90, 4

' if there is 2 or more balls then set the multiball flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If (BallsOnPlayfield-RealBallsInLock = 2) And LastSwitchHit = "OutlaneSW" And (bBallSaverActive = True Or bLoLLit = True) Then
        ' Preemptive ball save
        bAutoPlunger = True
        If bBallSaverActive = False Then bLoLLit = False: SetOutlaneLights
    ElseIf (BallsOnPlayfield-RealBallsInLock > 1) Then
        DOF 143, DOFPulse
        bMultiBallMode = True
        bAutoPlunger = True
    End If
End Sub


Sub DoBallSaved(l)
    ' create a new ball in the shooters lane
    ' we use the Addmultiball in case the multiballs are being ejected
    AddMultiball 1
    ' we kick the ball with the autoplunger
    bAutoPlunger = True
    bBallSaved = True
    If l Then
        PlaySoundVol "gotfx-lolsave",VolDef
        bLoLLit = False
    Else
        if bMultiBallMode = False Then 
            DisableBallSaver
            DMDDoBallSaved
        End If
    End If
End Sub

Sub EnableBallSaver(seconds)
    'debug.print "Ballsaver started"
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    ' start the timer
    SetGameTimer tmrBallSave,seconds*10
    SetGameTimer tmrBallSaveSpeedUp,(seconds-5)*10
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightShootAgain.BlinkInterval = 160
    LightShootAgain.State = 2
End Sub

Sub DisableBallSaver
    TimerFlags(tmrBallSave) = 0
    TimerFlags(tmrBallSaveSpeedUp) = 0
    BallSaveTimer
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaveTimer
    'debug.print "Ballsaver ended"
    ' clear the flag
    bBallSaverActive = False
    If ExtraBallsAwards(CurrentPlayer) = 0 Then LightShootAgain.State = 0 Else LightShootAgain.State = 1
End Sub

Sub BallSaverSpeedUpTimer
    'debug.print "Ballsaver Speed Up Light"
    ' Speed up the blinking
    LightShootAgain.BlinkInterval = 80
    LightShootAgain.State = 2
End Sub

Sub AddScore(points)
    AddScoreNoX points*PlayfieldMultiplierVal
End Sub

Sub AddScoreNoX(points)
    ResetBallSearch
    'TODO: GoT allows you to hit certain targets without validating the playfield and starting timers
    ThawAllGameTimers
    If bPlayfieldValidated = False Then bPlayfieldValidated = True : PlayModeSong
    If (Tilted = False) Then
        ' if there is a need for a ball saver, then start off a timer
        ' only start if it is ready, and it is currently not running, else it will reset the time period
        If(bBallSaverReady = True)AND(DMDStd(kDMDStd_BallSave) <> 0)And(bBallSaverActive = False)Then
            EnableBallSaver DMDStd(kDMDStd_BallSave)
        End If

        If Score(CurrentPlayer) < ReplayScore And Score(CurrentPlayer) + points > ReplayScore Then 
            ReplayAwarded(CurrentPlayer) = True
            ' If Bonus is being display, slightly delay Replay scene so it plays after bonus scene
            If bBonusSceneActive Then vpmTimer.AddTimer 1500,"DoAwardReplay '" : Else DoAwardReplay
        End If
        Score(CurrentPlayer) = Score(CurrentPlayer) + points

        ' update the score display
        DMDLocalScore

    End If
End Sub

sub ResetBallSearch()
	if BallSearchResetting then Exit Sub	' We are resetting jsut exit for now 
	'debug.print "Ball Search Reset"
	tmrBallSearch.Enabled = False	' Reset Ball Search
    tmrBallSearch.Interval = 20000
	BallSearchCnt=0
	tmrBallSearch.Enabled = True
End Sub 

dim BallSearchResetting:BallSearchResetting=False

Sub tmrBallSearch_Timer()	' We timed out
	' See if we are in mode select, a flipper is up that might be holding the ball or a ball is in the lane 

	'debug.print "Ball Search"
	if bGameInPlay and _ 
        PlayerMode >= 0 and _
		bBallInPlungerLane = False and _
        bMysteryAwardActive = False And _
        hsbModeActive = False And _
        tmrEndOfBallBonus.Enabled = False And _
		LeftFlipper.CurrentAngle = LeftFlipper.StartAngle and _
		RightFlipper.CurrentAngle = RightFlipper.StartAngle Then

		debug.print "Ball Search - NO ACTIVITY " & BallSearchCnt

		BallSearchResetting=True
		BallSearchCnt = BallSearchCnt + 1
        tmrBallSearch.Interval = 10000
		DisplayDMDText "BALL SEARCH","", 1000
        If BallSearchCnt = 1 Then
            If Target7.IsDropped = 0 And Target8.IsDropped = 0 And Target9.IsDropped = 0 Then
                Target7.IsDropped = 1 : Target8.IsDropped = 1 : Target9.IsDropped = 1
                vpmTimer.AddTimer 1000,"ResetDropTargets : BallSearchCnt = 0 '"
            End If
            KickerIT.Kick 180,3
        ElseIf BallSearchCnt = 2 Then
            ' Not on the playfield - try the lock wall
            ReleaseLockedBall 0
        Elseif BallSearchCnt >= 3 Then
			dim Ball
			debug.print "--- listing balls ---"
			For each Ball in GetBalls
				Debug.print "Ball: (" & Ball.x & "," & Ball.y & ")"
			Next
			debug.print "--- listing balls ---"

	        ' Delete all balls on the playfield and add back a new one
			BallsOnPlayfield = 0 : RealBallsInLock = 0
			for each Ball in GetBalls
				Ball.DestroyBall
			Next
			BallSearchCnt = 0
			
			AddMultiball(1)
			DisplayDMDText "BALL SEARCH FAIL","", 1000
			Exit sub
		End if
        
		vpmtimer.addtimer 3000, "BallSearchResetting = False '"
	Else 
		ResetBallSearch
	End if 
End Sub 

'************************************
' End of ball/game processing
'************************************

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded. This is a mini state machine. State is
' stored in timer.UserValue. States:
' "BONUS" 0.25s
' Base Bonus 0.45s
' X Houses complete 0.45s
' Swords 150k each 0.45s
' Castles 0.45s
' Gold (405=121500) 0.45s  e.g “405 GOLD” on line 1, “121,500” on line 2
' Wildfire (22=121000) 0.55s
' Then bonus X times all added together 1.2s
' Then Total bonus. 1.4s
'
' Note, there are 7 "notes" in the ROM, but only 6 bonuses. Perhaps note 7 is reserved.
' In any case, we skip over it below, in case we want it in the future.

'TODO: Add support for fastforwarding bonus if flipper is held in
Dim tmpBonusTotal
dim bonusCnt
Dim BonusScene
Sub tmrEndOfBallBonus_Timer()
    Dim line1,line2,ol,skip,font,i,j,incr
    ol = False
    skip = False
    incr = 1 : i = 0
    tmrEndOfBallBonus.Enabled = False
    tmrEndOfBallBonus.Interval = 500
    j = Int(tmrEndOfBallBonus.UserValue)
    ' State machine based on GoTG line 2549 onwards
    Select Case j
        Case 0
            bonusCnt = 0
            StopSound Song : Song = ""
            line1 = "BONUS"
            ol = True
            'tmrEndOfBallBonus.Interval = 250
        Case 1
            line1 = "BASE BONUS" : line2 = FormatScore(BonusPoints(CurrentPlayer))
            BonusCnt = BonusPoints(CurrentPlayer)
        Case 2 
            If CompletedHouses > 0 Then
                line1 = CompletedHouses & " HOUSES COMPLETE":line2= FormatScore(175000*CompletedHouses)
                BonusCnt = BonusCnt + (175000*CompletedHouses)
            Else
                Skip = True
            End If
        Case 3 
            If SwordsCollected > 0 Then
                line1 = SwordsCollected & " SWORD"
                If SwordsCollected > 1 Then line1 = line1&"S"
                line2 = FormatScore(150000*SwordsCollected)
                BonusCnt = BonusCnt + (150000*SwordsCollected)
            Else
                Skip = True
            End If
        Case 4
            If CastlesCollected > 0 Then
                line1 = CastlesCollected & " CASTLE"
                If CastlesCollected > 1 Then line1 = line1&"S" 
                line2 = FormatScore(7500000*CastlesCollected)
                BonusCnt = BonusCnt + (7500000*CastlesCollected)
            Else
                Skip = True
            End If
        Case 5
            If TotalGold > 0 Then
                line1 = FormatScore(TotalGold) & " GOLD" : line2 = FormatScore(300*TotalGold)
                BonusCnt = BonusCnt + (300*TotalGold)
            Else
                Skip = True
            End If
        Case 6
            If TotalWildfire > 0 Then
                line1 = FormatScore(TotalWildfire) & " WILDFIRE" : line2 = FormatScore(5500*TotalWildfire)
                BonusCnt = BonusCnt + (5500*TotalWildfire)
                tmrEndOfBallBonus.Interval = 600
            Else
                Skip = True
            End If
        Case 8
            i = Int(tmrEndOfBallBonus.UserValue*100) - 799  
            line1 = i&"X" : line2 = FormatScore(i*BonusCnt)
            If i = BonusMultiplier(CurrentPlayer) Then 
                tmrEndOfBallBonus.Interval = 1200
                incr = 0
                tmrEndOfBallBonus.UserValue = 9
            Else
                tmrEndOfBallBonus.Interval = 125
                incr = 0.01
            End If
        Case 9
            line1 = "TOTAL BONUS" : line2 = FormatScore(BonusCnt * BonusMultiplier(CurrentPlayer))
            PlayfieldMultiplierVal = 1
            AddScore BonusCnt * BonusMultiplier(CurrentPlayer)
            tmrEndOfBallBonus.Interval = 1700
        Case 10
            vpmtimer.addtimer 100, "EndOfBall2 '"
            Exit Sub
        Case Else
            Skip = True
    End Select

    If Skip Then
        tmrEndOfBallBonus.Interval = 10
    Else
        ' Do Bonus Scene
        If bUseFlexDMD Then
            If j=8 And i > 1 Then
                FlexDMD.LockRenderThread
                With BonusScene.GetLabel("line1")
                    .Text = line1 & vbLf & line2
                    .SetAlignedPosition 64,16,FlexDMD_Align_CENTER
                End With
                FlexDMD.UnlockRenderThread
                PlaySoundVol "gotfx-spike-count8",VolDef
            Else
                Set BonusScene = FlexDMD.NewGroup("bonus")
                If ol Then
                    Set font = FlexDMD.NewFont("FlexDMD.Resources.udmd-f7by13.fnt", vbWhite, vbWhite, 0)
                Else
                    line1 = line1 & vbLf & line2
                    Set font = FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite, vbWhite, 0)
                End if
                BonusScene.AddActor FlexDMD.NewFrame("bonusbox")
                With BonusScene.GetFrame("bonusbox")
                    .Thickness = 1
                    .SetBounds 0, 0, 128, 32
                End With
                BonusScene.AddActor FlexDMD.NewLabel("line1",font,line1)
                BonusScene.GetLabel("line1").SetAlignedPosition 64,16,FlexDMD_Align_CENTER
                DMDClearQueue
                DMDEnqueueScene BonusScene,0,3000,3000,10,"gotfx-spike-count"&j
            End If
        Else
            If ol Then 
                DisplayDMDText "",line1,tmrEndOfBallBonus.Interval
            Else
                DisplayDMDText line1,line2,tmrEndOfBallBonus.Interval
            End If
            PlaySoundVol "gotfx-spike-count"&j, VolDef
        End If
    End If

    tmrEndOfBallBonus.UserValue = tmrEndOfBallBonus.UserValue + incr
    tmrEndOfBallBonus.Enabled = True
End Sub

Sub EndOfBall()
    Dim AwardPoints, TotalBonus,delay
    AwardPoints = 0
    TotalBonus = 0
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False

    If bHotkMBReady Then ResetHotkLights
	
    StopGameTimers
    If HouseBattle1 > 0 Then BattleModeTimer1
    If HouseBattle2 > 0 Then BattleModeTimer2
    EndHurryUp
    PlayerMode = 0
    
    ' only process any of this if the table is not tilted.  (the tilt recovery
    ' mechanism will handle any extra balls or end of game)

    If(Tilted = False)Then
        TurnOffPlayfieldLights
        DMDflush

        delay = 0
        If tmrMultiballCompleteScene.Enabled Then
            delay=delay+2000
            tmrMultiballCompleteScene_Timer
        ElseIf bITMBActive And ITScore > 0 Then 
            delay = delay + 2000
        End If
        If tmrBattleCompleteScene.Enabled Then
            delay=delay+3000
            tmrBattleCompleteScene_Timer
        End If
        
        ' Delay for a Battle Total screen to be shown
        tmrEndOfBallBonus.Interval = delay + 400
        ' Start the Bonus timer - this timer calls the Bonus Display code when it runs down
		tmrEndOfBallBonus.UserValue = 0
		tmrEndOfBallBonus.Enabled = true
        ' Bonus will start EndOfBall2 when it is done
    Else
        vpmtimer.addtimer 100, "EndOfBall2 '"
    End If
End Sub

' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the currentplayer)
'
Sub EndOfBall2()
	dim i
	dim thisMode
    ' if were tilted, reset the internal tilted flag (this will also
    ' set TiltWarnings back to zero) which is useful if we are changing player LOL
    Tilted = False
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots


    ' has the player won an extra-ball ? (might be multiple outstanding)
    If(ExtraBallsAwards(CurrentPlayer) <> 0) and bBallInPlungerLane=False Then	' Save Extra ball for later if there is a ball in the plunger lane
        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

        ' if no more EB's then turn off any shoot again light
        If(ExtraBallsAwards(CurrentPlayer) = 0)Then
            LightShootAgain.State = 0
        End If

        If bUseFlexDMD Then
            Dim Scene
            Set Scene = NewSceneWithVideo("shootagain","got-shootagain")
            DMDEnqueueScene Scene,0,4000,8000,2000,"" : DoDragonRoar 1
            vpmTimer.addTimer 1500,"PlaySoundVol ""say-shoot-again"",VolCallout '"
        Else
            DMD "_", CL(1, ("EXTRA BALL")), "_", eNone, eBlink, eNone, 1000, True, "say-shoot-again"
        End if

         ' Save the current player's state - needed so that when it's restored in a moment, it won't screw everything up
        PlayerState(CurrentPlayer).Save
        ' reset the playfield for the new player (or new ball) (also restores player state)
        ResetForNewPlayerBall()
        ' Create a new ball in the shooters lane
        CreateNewBall()
    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0) Then				' GAME OVER
            debug.print "No More Balls, High Score Entry"

			' Turn off DOF so we dont accidently leave it on
			PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
			LeftFlipper.RotateToStart
            LeftUFlipper.RotateToStart

			PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
			RightFlipper.RotateToStart
			RightUFlipper.RotateToStart

            ' Submit the currentplayers score to the High Score system
            CheckHighScore()
			' you may wish to play some music at this point
        Else
            ' not the last ball (for that player)
            ' if multiple players are playing then move onto the next one
            EndOfBallComplete()
        End If
    End If
End Sub

' This function is called when the end of bonus display
' (or high score entry finished) AND it either end the game or
' move onto the next player (or the next ball of the same player)
'all of the same player)
'
Sub EndOfBallComplete()
    Dim NextPlayer
	dim Match
    Dim i

    debug.print "EndOfBall - Complete"

    ' Save the current player's state
    PlayerState(CurrentPlayer).Save

    ' are there multiple players playing this game ?
    If(PlayersPlayingGame > 1)Then
        ' then move to the next player
        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If(NextPlayer > PlayersPlayingGame)Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    debug.print "Next Player = " & NextPlayer

    ' is it the end of the game ? (all balls been lost for all players)
    If((BallsRemaining(CurrentPlayer) <= 0)AND(BallsRemaining(NextPlayer) <= 0))Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode

		DisplayDMDText2 "_", "GAME OVER", 20000, 5, 0

        ' Drop the lock walls to release any locked balls
        SwordWall.collidable = False : MoveSword True
        LockWall.collidable = False : Lockwall.Uservalue = 0 : LockPost.TransZ = 0
        BallsInLock = 0 : RealBallsInLock = 0

		bGameInPLay = False									' EndOfGame sets this but need to set early to disable flippers 
		bShowMatch = True

		' Do Match end score code
		Match=10 * INT(RND * 9)
		'Match = Score(CurrentPlayer) mod 100		' Force Match for testing 
        For i = 1 to PlayersPlayingGame
            if BigMod(Score(CurrentPlayer), 100) = Match then									' Handles large scores 
                vpmtimer.addtimer 5500+i*200, "PlayYouMatched '"
            End If
        Next

        DMDDoMatchScene Match
		'if Match = 0 then 
	'		playmedia "Match-00.mp4", "PupVideos", pOverVid, "", -1, "", 1, 1
		'Else
		'	playmedia "Match-"&Match&".mp4", "PupVideos", pOverVid, "", -1, "", 1, 1
		'End If
		'vpmtimer.addtimer 100, "PlaySoundVol ""Match-Score"", VolDef '"
        ' set the machine into game over mode
		'if osbactive <> 0 then 	' Orbital takes more time 
		'	vpmtimer.addtimer 9000, "if bShowMatch then EndOfGame() '"
		'else 
			vpmtimer.addtimer 9500, "if bShowMatch then EndOfGame() '"
		'End If 

    ' you may wish to put a Game Over message on the desktop/backglass

    Else
        ' set the next player
        CurrentPlayer = NextPlayer
		'UpdateNumberPlayers				' Update the Score Sizes
        ' make sure the correct display is up to date
        AddScore 0

        ' reset the playfield for the new player (or new ball)
        ResetForNewPlayerBall()

        ' AND create a new ball
        CreateNewBall()

        ' play a sound if more than 1 player
        If PlayersPlayingGame > 1 Then
            i = RndNbr(2)
            PlaySoundVol "say-player" &CurrentPlayer&"-youre-up"&i, VolCallout
            DMD "", CL(1, "PLAYER " &CurrentPlayer), "", eNone, eNone, eNone, 800, True, ""
        End If
    End If
End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    Dim i
    If bGameInPLay = True then Exit Sub ' In case someone pressed 'Start' during Match sequence
    debug.print "End Of Game"	
	
    bGameInPLay = False	
	bShowMatch = False
	tmrBallSearch.Enabled = False
    
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

	' Drop the lock walls just in case the ball is behind it (just in Case)
	SwordWall.collidable = False : MoveSword True
    LockWall.collidable = False : Lockwall.Uservalue = 0 : LockPost.TransZ = 0

    PlaySoundVol "got-track-gameover",VolDef/8
    i = RndNbr(12)
    PlaySoundVol "say-gameover"&i,VolDef
    DMD "_", "GAME OVER", "",eNone,eNone,eNone,6000,true,""

    GiOff

    vpmTimer.AddTimer 3000,"StartAttractMode '"

End Sub

Sub PlayYouMatched
	'PlaySoundVol "YouMatchedPlayAgain", VolDef
	DOF 140, DOFOn
	DMDFlush
	DMD "_", CL(1, "CREDITS: " & Credits), "", eNone, eNone, eNone, 500, True, "fx_knocker"
End Sub 







' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
End Sub
Sub AddMultiballFast(nballs)
	if CreateMultiballTimer.Enabled = False then 
		CreateMultiballTimer.Interval = 100		' shortcut the first time through 
	End If 
   AddMultiball(nballs)
End Sub

' Eject the ball after the delay, AddMultiballDelay
Sub CreateMultiballTimer_Timer()
    ' wait if there is a ball in the plunger lane
	CreateMultiballTimer.Interval = 2000
    If bBallInPlungerLane Then
        Exit Sub
    Else
        If (BallsOnPlayfield-RealBallsInLock) < MaxMultiballs Then
            CreateNewBall()
            mBalls2Eject = mBalls2Eject -1
            If mBalls2Eject = 0 Then 'if there are no more balls to eject then stop the timer
                Me.Enabled = False
            End If
        Else 'the max number of multiballs is reached, so stop the timer
            mBalls2Eject = 0
            Me.Enabled = False
        End If
    End If
End Sub


'************************
' Switch hits
'************************

Sub sw1_hit 'Left inlane (L)
    If Tilted Then Exit Sub
End Sub

Sub sw46_hit 'Left inlane (R)
    If Tilted Then Exit Sub
End Sub

Sub sw2_hit 'Right inlane
    If Tilted Then Exit Sub
End Sub

Sub sw3_hit 'Left outlane
    If Tilted Then Exit Sub
End Sub

Sub sw4_hit 'Right outlane
    If Tilted Then Exit Sub
End Sub

Sub sw14_Hit 'Left Orbit
    If Tilted Then Exit Sub
End Sub

Sub sw24_Hit 'Right Orbit
    If Tilted Then Exit Sub
End Sub

Sub sw34_Hit 'Star Entrance trigger
    If Tilted Then Exit Sub
End Sub

Sub sw49_Hit 'Demon entrance trigger
    If Tilted Then Exit Sub
End Sub

Sub sw56_Hit 'Right Ramp Exit
    If Tilted Then Exit Sub
End Sub

Sub sw64_Hit 'Left Ramp Exit
    If Tilted Then Exit Sub
End Sub


'*********************
' Target Hits
'*********************
Sub sw25_Hit '(S)tar
    If Tilted Then Exit Sub
End Sub
Sub sw26_Hit 'S(t)ar
    If Tilted Then Exit Sub
End Sub
Sub sw27_Hit 'St(a)r
    If Tilted Then Exit Sub
End Sub
Sub sw28_Hit 'Sta(r)
    If Tilted Then Exit Sub
End Sub

Sub sw37_Hit 'Star entrance drop target
    If Tilted Then Exit Sub
End Sub

Sub sw38_Hit '(K)iss
    If Tilted Then Exit Sub
End Sub
Sub sw39_Hit 'K(i)ss
    If Tilted Then Exit Sub
End Sub
Sub sw40_Hit 'Ki(s)s
    If Tilted Then Exit Sub
End Sub
Sub sw41_Hit 'Kis(s)
    If Tilted Then Exit Sub
End Sub

Sub sw44_Hit 'Left lock target
    If Tilted Then Exit Sub
End Sub

Sub sw45_Hit 'Right lock target
    If Tilted Then Exit Sub
End Sub


Sub sw58_Hit '(A)rmy
    If Tilted Then Exit Sub
End Sub
Sub sw59_Hit 'A(r)my
    If Tilted Then Exit Sub
End Sub
Sub sw60_Hit 'Ar(m)y
    If Tilted Then Exit Sub
End Sub
Sub sw61_Hit 'Arm(y)
    If Tilted Then Exit Sub
End sub



'***********************
' Bumper hits
'***********************
'**************
'Bumper Hits
'**************
Sub Bumper1_Hit
    If Tilted then Exit Sub
    RandomSoundBumperTop Bumper1
    DOF 107, DOFPulse
    doBumperHit 0
End Sub

Sub Bumper2_Hit
    If Tilted then Exit Sub
    RandomSoundBumperTop Bumper2
    DOF 108, DOFPulse
    doBumperHit 1
End Sub

Sub Bumper3_Hit
    If Tilted then Exit Sub
    RandomSoundBumperTop Bumper3
    DOF 109, DOFPulse
    doBumperHit 2
End Sub

Sub Bumper4_Hit
    If Tilted then Exit Sub
    RandomSoundBumperTop Bumper4
    DOF 109, DOFPulse
    doBumperHit 3
End Sub




Dim ChampionNames
ChampionNames = Array("CITY COMBO","HEAVENS ON FIRE","ROCK CITY")

Sub GameStartAttractMode
    If bGameInPLay Then GameStopAttractMode : Exit Sub
    tmrDMDUpdate.Enabled = False
    SetUPFFlashers 0,cyan
    bAttractMode = True
    tmrAttractModeScene.UserValue = 0
    tmrAttractModeScene.Interval = 10
    tmrAttractModeScene.Enabled = True

    SavePlayfieldLightState
    tmrAttractModeLighting.UserValue = 0
    tmrAttractModeLighting.Interval = 10
    tmrAttractModeLighting.Enabled = True

    tmrFlipperSpeech.UserValue=0
End Sub

Sub GameStopAttractMode
    bAttractMode = False
    tmrAttractModeScene.Enabled = False
    tmrAttractModeLighting.Enabled = False
    LightSeqAttract.StopPlay
    RestorePlayfieldLightState True
    DMDClearQueue
    tmrDMDUpdate.Enabled = True
End Sub

Sub tmrFlipperSpeech_Timer
    Dim i
    Select Case tmrFlipperSpeech.UserValue
        Case 0: i=6
        Case 6: i=7
        Case 7: i=8
        Case 8: i=15
        Case Else: i=20     ' Default delay between speeches, in seconds
    End Select
    tmrFlipperSpeech.UserValue = i
    tmrFlipperSpeech.Enabled = 0
End Sub

Sub doFlipperSpeech(keycode)
    tmrAttractModeScene.Enabled = False
    If keycode = LeftFlipperKey Then
        tmrAttractModeScene.UserValue = tmrAttractModeScene.UserValue - 2
        If tmrAttractModeScene.UserValue < 0 Then tmrAttractModeScene.UserValue =  tmrAttractModeScene.UserValue + 24
    End If
    tmrAttractModeScene_Timer

    if tmrFlipperSpeech.Enabled <> 0 Then Exit Sub
    tmrFlipperSpeech.Interval = tmrFlipperSpeech.UserValue*1000
    tmrFlipperSpeech.Enabled=1
    PlaySoundVol "say-flipper"&RndNbr(43),VolCallout
End Sub

' To launch attract mode, disable DMDUpdateTimer and enble tmrAttractModeScene
' Attract mode is a big state machine running through various scenes in a loop. The timer is called
' after the scene has displayed for the set time, to move onto the next scene
' Scenes:
' 0: Stern logo - 3 seconds
' 1: PRESENTS - 2 seconds
' 2: GoT logo video - 21 seconds
' 3: part of Nights Watch oath on winter storm bg - 9 seconds
' 4: most recent score (just p1?) - 2 sec
' 5: credits - 2 sec
' 6: Replay at <x> - 2 sec
' 7: more oath - 9 sec
' 8: game logo - 3 sec
' 9-24: various high scores - 2 seconds each

' Format tells us how to format each scene
'  1: 1 line of text
'  2: 2 lines of text (same size)
'  3: 3 lines of text (small, big, medium)
'  4: image only, no text
'  5: video only, no text
'  6: video, scrolling text
'  7: 2 lines of text (big, small)
'  8: 3 lines of text (same size)
'  9: scrolling image
' 10: 1 line of text with outline
Sub tmrAttractModeScene_Timer
    Dim scene,scene2,img,line1,line2,line3
    Dim skip,font,format,scrolltime,y,delay,skipifnoflex,i
    Dim font1,font2,font3
    skip = False
    tmrAttractModeScene.Enabled = False
    If bGameInPLay Or bAttractMode = False Then GameStopAttractMode : Exit Sub
    delay = 2000
    skipifnoflex = True  ' Most scenes won't render without FlexDMD
    scrolltime = 0
    i = tmrAttractModeScene.UserValue
    Select Case tmrAttractModeScene.UserValue
        Case 0:img = "got-sternlogo":format=9:scrolltime=3:y=73:delay=3000
        Case 1:line1 = "PRESENTS":format=10:font="udmd-f7by10.fnt":delay=2000
        Case 2:img = "got-intro":format=5:delay=17200
        Case 3,7:img = "got-winterstorm":format=6:line1 = GetNextOath():delay=9000:font = "skinny10x12.fnt":scrolltime=9:y=100   ' Oath Text
        Case 4
            format=7:font="udmd-f11by18.fnt":line1=FormatScore(Score(1)):skipifnoflex=False  ' Last score
            If Score(1) > 999999999 Then font="udmd-f7by10.fnt"
            If DMDStd(kDMDStd_FreePlay) Then line2 = "FREE PLAY" Else Line2 = "CREDITS "&Credits
        Case 5
            format=1:font="udmd-f7by10.fnt":skipifnoflex=False
            If DMDStd(kDMDStd_FreePlay) Then 
                line1 = "FREE PLAY" 
            ElseIf Credits > 0 Then 
                Line1 = "CREDITS "&Credits
            Else 
                Line1 = "INSERT COINS"
            End if
        Case 6:format=1:font="udmd-f7by10.fnt":line1 = "REPLAY AT" & vbLf & FormatScore(ReplayScore)
        Case 8:img = "got-introframe":format=4:delay=3000
        Case 9:format=8:line1="GRAND CHAMPION":line2=HighScoreName(0):line3=FormatScore(HighScore(0)):font="udmd-f6by8.fnt":skipifnoflex=False
        Case 10:format=8:line1="HIGH SCORE #1":line2=HighScoreName(1):line3=FormatScore(HighScore(1)):font="udmd-f6by8.fnt":skipifnoflex=False
        Case 11:format=8:line1="HIGH SCORE #2":line2=HighScoreName(2):line3=FormatScore(HighScore(2)):font="udmd-f6by8.fnt":skipifnoflex=False
        Case 12:format=8:line1="HIGH SCORE #3":line2=HighScoreName(3):line3=FormatScore(HighScore(3)):font="udmd-f6by8.fnt":skipifnoflex=False
        Case 13:format=8:line1="HIGH SCORE #4":line2=HighScoreName(4):line3=FormatScore(HighScore(4)):font="udmd-f6by8.fnt":skipifnoflex=False
        Case 14,15,16,17,18,19,20,21,22,23
            If HighScore(i-9) = 0 Then delay = 5
            format=3:skipifnoflex=False
            line1 = ChampionNames(i-14)&" CHAMPION"
            line2 = HighScoreName(i-9):line3 = FormatScore(HighScore(i-9))
    End Select
    If i = 23 Then tmrAttractModeScene.UserValue = 0 Else tmrAttractModeScene.UserValue = i + 1
    If bUseFlexDMD=False And skipifnoflex=True Then tmrAttractModeScene.Interval = 10 Else tmrAttractModeScene.Interval = delay
    tmrAttractModeScene.Enabled = True

    ' Create the scene
    if bUseFlexDMD Then
        If format=4 or Format=5 or Format=6 or Format=9 Then
            Set scene = NewSceneWithVideo("attract"&i,img)
        Else
            Set scene = FlexDMD.NewGroup("attract"&i)
        End If

        Select Case format
            Case 1
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                scene.GetLabel("line1").SetAlignedPosition 64,16,FlexDMD_Align_Center
            Case 3
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line1)
                scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont("udmd-f7by10.fnt",vbWhite,vbWhite,0),line2)
                scene.AddActor FlexDMD.NewLabel("line3",FlexDMD.NewFont("udmd-f6by8.fnt",vbWhite,vbWhite,0),line3)
                scene.GetLabel("line1").SetAlignedPosition 64,3,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,15,FlexDMD_Align_Center
                scene.GetLabel("line3").SetAlignedPosition 64,27,FlexDMD_Align_Center
            Case 6
                scene.AddActor FlexDMD.NewGroup("scroller")
                scene.SetBounds 0,0-y,128,32+(2*y)  ' Create a large canvas for the text to scroll through
                With scene.GetGroup("scroller")
                    .SetBounds 0,y+32,128,y
                    .AddAction scene.GetGroup("scroller").ActionFactory().MoveTo(0,0,scrolltime)
                    .AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                End With
                scene.GetVideo("attract"&i&"vid").SetAlignedPosition 0,y,FlexDMD_Align_TopLeft ' move image to screen
                scene.GetLabel("line1").SetAlignedPosition 64,0,FlexDMD_Align_Top        
            Case 7
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line2)
                scene.GetLabel("line1").SetAlignedPosition 64,14,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,27,FlexDMD_Align_Center
            Case 8
                Set font1 = FlexDMD.NewFont(font,vbWhite,vbWhite,0)
                scene.AddActor FlexDMD.NewLabel("line1",font1,line1)
                scene.AddActor FlexDMD.NewLabel("line2",font1,line2)
                scene.AddActor FlexDMD.NewLabel("line3",font1,line3)
                scene.GetLabel("line1").SetAlignedPosition 64,5,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,16,FlexDMD_Align_Center
                scene.GetLabel("line3").SetAlignedPosition 64,27,FlexDMD_Align_Center
            Case 9
                scene.SetBounds 0,0-y,128,32+(2*y)  ' Create a large canvas for the image to scroll through
                With scene.GetImage("attract"&i&"img")
                    .SetAlignedPosition 0,y+32,FlexDMD_Align_TopLeft
                    .AddAction scene.GetImage("attract"&i&"img").ActionFactory().MoveTo(0,0,scrolltime)
                End With
            Case 10
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,RGB(64, 64, 64),1),line1)
                scene.GetLabel("line1").SetAlignedPosition 64,16,FlexDMD_Align_Center
        End Select

        DMDDisplayScene scene
    End If
End Sub