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
        If(x <> "")Then 
            HighScore(i-1) = CDbl(x)
        Else
            if i < 6 Then HighScore(i-1) = DMDStd(kDMDStd_HighScoreGC + i - 1) Else HighScore(i-1) = (17-i)*1000000 
        End If
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
            DOF 122, DOFPulse
            If tmp > 2 Then
                vpmTimer.AddTimer 500+100*(tmp-3), "HighScoreEntryInit() '"
                Do While tmp > 2
                    vpmTimer.AddTimer 200*(tmp-2),"KnockerSolenoid : DOF 122, DOFPulse '"
                    tmp = tmp - 1
                Loop
            Else
                vpmTimer.AddTimer 300, "HighScoreEntryInit() '"
            End If
    End Select
End Sub

Sub HighScoreEntryInit()
    Dim i
    hsbModeActive = 1
    ' Table-specific
    i = RndNbr(4)
    PlaySoundVol "say-enterinitials"&i,VolCallout
    PlaySoundVol "kiss-track-hstd",VolDef/8
    hsLetterFlash = 0

    hsMaxDigits = DMDStd(kDMDStd_Initials)
    For i = 0 to hsMaxDigits-1 : hsEnteredDigits(i) = "" : Next
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<" ' < is back arrow
    if hsMaxDigits > 3 Then 
        hsValidLetters = hsValidLetters + "~" ' Add the ~ key to the end to exit long initials entry
        hsX = 10
    Else
        hsX = 49
    End If
    hsCurrentLetter = 1
    If bUseFlexDMD Then
        Set HSscene = FlexDMD.NewGroup("highscore")
        tmrDMDUpdate.Enabled = False
        ' Note, these fonts aren't included with FlexDMD. Change to stock fonts for other tables
        HSscene.AddActor FlexDMD.NewLabel("name",FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbWhite, 0),"YOUR NAME:")
        HSScene.GetLabel("name").SetAlignedPosition 2,2,FlexDMD_Align_TopLeft
        HSscene.AddActor FlexDMD.NewLabel("initials",FlexDMD.NewFont("skinny10x12.fnt", vbWhite, vbWhite, 0),"")
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
        PlaySoundVol "kissfx-hstd-left",VolDef
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0)then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        'playsound "fx_Next"
        'Table-specific
        PlaySoundVol "kissfx-hstd-right",VolDef
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter> len(hsValidLetters))then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = PlungerKey OR keycode = StartGameKey Or keycode = RightMagnaSave Or keycode = LockBarKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<")then
            'playsound "fx_Enter"
            'Table-specific
            PlaySoundVol "kissfx-hstd-enter",VolDef
            if hsMaxDigits = 10 And (mid(hsValidLetters, hsCurrentLetter, 1) = "~") Then HighScoreCommitName() : Exit Sub
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3 And hsMaxDigits = 3) Or (hsCurrentDigit = 10 And hsMaxDigits = 10)then
                HighScoreCommitName()
            else
                HighScoreDisplayNameNow()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = ""
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
    Dim i,bBlank,bm,scores
    HighScoreFlashTimer.Enabled = False
    hsbModeActive = 2

    hsEnteredName="":bBlank=true
    For i = 0 to hsMaxDigits-1
        hsEnteredName = hsEnteredName & hsEnteredDigits(i) 
        if hsEnteredDigits(i) <> " " Then bBlank=False
    Next
    if bBlank then hsEnteredName = "YOU"

    bm = GameSaveHighScore(Score(CurrentPlayer),hsEnteredName)
    SortHighscore
    Savehs
    For i = 0 to 4
        If Score(CurrentPlayer) = HighScore(i) then scores = 2^i
    Next
    scores = scores Or bm
    GameDoDMDHighScoreScene scores
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

	elseif keycode=ServiceCancelKey then 					' 7 key cancels the service menu
		if bInService then
			Select case serviceLevel:
				case kMenuTop, kMenuNone: 
					bInService=False
					bServiceVol=False
					bAttractMode=serviceSaveAttract
					if bAttractMode then
						StartAttractMode
                    Else
                        bDefaultScene = False
                        tmrDMDUpdate.Enabled = True
					End if 
					
				case kMenuAdj:
					serviceLevel=kMenuNone
					serviceIdx=0
					StartServiceMenu 11
                    Exit Sub
				case kMenuAdjStd,kMenuAdjGame:
					if bServiceEdit then 
						bServiceEdit=False 
                        DMDStd(DMDMenu(serviceIdx).StdIdx)=serviceOrigValue
                        UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
					else 
                        serviceLevel=kMenuTop
						serviceIdx=2
						StartServiceMenu 11
                        Exit Sub
					End if 
			End Select  
            PlaySoundVol "stern-svc-cancel", VolSfx
		End if 
	elseif bInService Then
		if keycode=ServiceEnterKey then 	' Select 
            Dim svcsound : svcsound = "stern-svc-enter"
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
						svcsound = "sfx-deny"
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
						    svcsound = "sfx-deny"
					End Select
				case kMenuAdjStd,kMenuAdjGame:
					if DMDMenu(serviceIdx).ValType="FUN" then		' Function
                        Select Case DMDMenu(serviceIdx).StdIdx
                            Case 0: reseths
                            Case 1: ClearAll
                        End Select
						UpdateServiceDMDScene "<DONE>"
                        svcsound = "stern-svc-set"
					else 
						if bServiceEdit=False then 	' Start Editing  
							bServiceEdit=True
                            serviceOrigValue=DMDStd(DMDMenu(serviceIdx).StdIdx)
							UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
						else 							' Save the change 
							bServiceEdit=False
							UpdateServiceDMDScene GetTextForAdjustment(serviceIdx)
                            svcsound = "stern-svc-set"
						End if 
					End if 
			End Select 
            If svcsound <> "" Then PlaySoundVol svcsound, VolSfx
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

		SvcSetVolume

		if bServiceVol=False then
			bServiceVol=True 
			serviceSaveAttract=bAttractMode
			bAttractMode=False
            tmrDMDUpdate.Enabled = False
            tmrAttractModeScene.Enabled = False
            CreateServiceDMDScene 0
		End if 

		UpdateServiceDMDScene ""

		tmrService.Enabled = False 
		tmrService.Interval = 5000
		tmrService.Enabled = True 
	End if 
End Sub

Sub SvcSetVolume
    'VolBGVideo = cVolBGVideo * (MasterVol/100.0)
    VolBGMusic = cVolBGMusic * (MasterVol/100.0)
    VolDef 	 = cVolDef * (MasterVol/100.0)
    VolSfx 	 = cVolSfx * (MasterVol/100.0)
    VolCallout = cVolCallout * (MasterVol/100.0)
End Sub

Sub tmrService_Timer()
	tmrService.Enabled = False 	

    SaveValue TableName,"MasterVol",MasterVol
	bServiceVol=False 
	bAttractMode=serviceSaveAttract
	if bAttractMode then 
		StartAttractMode
    Else
        bDefaultScene = False ' Force a refresh of the DMD
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
Const kDMDStd_ExtraBallLimit = &H29
Const kDMDStd_ExtraBallPCT = &H2A
Const kDMDStd_MatchPCT = &H2D           
Const kDMDStd_TiltWarn = &H2F
Const kDMDStd_TiltDebounce = &H30
Const kDMDStd_LeftStartReset = &H33     
Const kDMDStd_BallSave = &H37           
Const kDMDStd_BallSaveExtend = &H36     
Const kDMDStd_ReplayType = &H2C         
Const kDMDStd_AutoReplayStart = &H1E    
Const kDMDStd_BallsPerGame = &H3E       '√
Const kDMDStd_FreePlay = &H32           '√
Const kDMDStd_HighScoreGC = &H43        '√
Const kDMDStd_HighScore1 = &H44         '√
Const kDMDStd_HighScore2 = &H45         '√
Const kDMDStd_HighScore3 = &H46         '√
Const kDMDStd_HighScore4 = &H47         '√
Const kDMDStd_HighScoreAward = &H48     '√
Const kDMDStd_GCAwards = &H49           '√
Const kDMDStd_HS1Awards = &H4A          '√
Const kDMDStd_HS2Awards = &H4B          '√
Const kDMDStd_HS3Awards = &H4C          '√
Const kDMDStd_HS4Awards = &H4D          '√
Const kDMDStd_Initials = &H4E           '√
Const kDMDStd_MusicAttenuation = &H50
Const kDMDStd_SpeechAttenuation = &H51

Const kDMDStd_LastStdSetting = &H51

'Table-specific indexes below here
Const kDMDFet_BallSaveTimerDemon = &H5F     
Const kDMDFet_DemonBallSaveEnabled = &H60
Const kDMDFet_DemonSpelloutDifficulty = &H61
Const kDMDFet_DemonMBDifficulty = &H62
Const kDMDFet_DemonMBVirtualLock = &H63
Const kDMDFet_DemonMBBallSaveTimer = &H64
Const kDMDFet_DemonSaveMBProgress = &H65
Const kDMDFet_DemonLockTimer = &H66
Const kDMDFet_GridsNeededForGridMB1 = &H67
Const kDMDFet_GridsNeededForGridMB2 = &H68
Const kDMDFet_GridsNeededForGridMBMax = &H69
Const kDMDFet_GridMBBallSaveTimer = &H6A
Const kDMDFet_GridMBAABBallSaveTimer = &H6B
Const kDMDFet_GridResetAtBallStart = &H6C
Const kDMDFet_KissArmyMBBallSaveTimer = &H6D
Const kDMDFet_FaceRuleDifficulty = &H6E
Const kDMDFet_StarDifficulty = &H6F
Const kDMDFet_LoveGunMBBallSaveTimer = &H70
Const kDMDFet_LoveGunMBSaveProgress = &H71
Const kDMDFet_LoveGunMBLockTimer = &H72
Const kDMDFet_RockCityMBBallSaveTimer = &H73
Const kDMDFet_DoubleScoringTimer = &H74
Const kDMDFet_SuperScoringTimer = &H75
Const kDMDFet_KissComboTimer = &H76
Const kDMDFet_ArmyComboTimer = &H77
Const kDMDFet_FrontRowDifficulty = &H78
Const kDMDFet_CityComboEBat = &H7A
Const kDMDFet_CityComboChampion = &H7B
Const kDMDFet_CityComboChampionAward = &H7C
Const kDMDFet_HeavensOnFireChampion = &H7D
Const kDMDFet_HeavensOnFireChampionAward = &H7E
Const kDMDFet_KissArmyChampion = &H81
Const kDMDFet_KissArmyChampionAward = &H82
Const kDMDFet_RockCityChampion = &H84
Const kDMDFet_RockCityChampionAward = &H85
Const kDMDFet_ResetShotXatBallStart = &H87
Const kDMDFet_AllShotXreqdForRecollet = &H88
Const kDMDFet_ResetPFXatBallStart = &H89
Const kDMDFet_SongCompletionsForPFX = &H8A
Const kDMDFet_MultShotIllumination = &H8B
Const kDMDFet_SongSelectTimerOff = &H8C
Const kDMDFet_SongSelectWithMultLit = &H8D
Const kDMDFet_InstrumentRuleDifficulty = &H8E
Const kDMDFet_FirstInstrumentEB = &H90
Const kDMDFet_InstrumentEBEvery = &H91
Const kDMDFet_InstrToBlinkSpaceman = &H92
Const kDMDFet_InstrToLightSpaceman = &H93
Const kDMDFet_BackStagePassBallSaveTimer = &H94


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
    DMDMenu(30).Setup "BALL  SAVE  TIMER  DEMON",             kDMDFet_BallSaveTimerDemon,       2,      "INT[0:5]", 1
    DMDMenu(31).Setup "DEMON  BALL  SAVE  DURING  MBALL",     kDMDFet_DemonBallSaveEnabled,     1,      "BOOL", 1
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
    x = LoadValue(TableName,"MasterVol") : If x<>"" Then MasterVol=CInt(x) : SvcSetVolume

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
        With SvcScene.GetLabel("svcl2")
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
                    If CLng(DMDStd(DMDMenu(serviceIdx).StdIdx)) = DMDMenu(serviceIdx).Deflt Then .Visible = 1 Else .Visible = 0
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




'******************************************************
'****  DROP TARGETS by Rothbauerw
'******************************************************
' This solution improves the physics for drop targets to create more realistic behavior. It allows the ball 
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full drop target animation, including deflection on hit and a slight lift when the drop 
' targets raise, switch handling, bricking, and popping the ball up if it's over the drop target when it raises.
'
'Add a Timer named DTAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
tmrDTAnim.interval = 10
tmrDTAnim.enabled = True

Sub tmrDTAnim_Timer
	DoDTAnim
	DoSTAnim
	TargetMovableHelper
End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and   
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target. 
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield 
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded 
' target mesh in this table's example. It uses the same texture map as the VP drop targets.


'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

'Define a variable for each drop target
Dim DT7, DT8, DT9, DT90

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
' 	primary: 			primary target wall to determine drop
'	secondary:			wall used to simulate the ball striking a bent or offset target after the initial Hit
'	prim:				primitive target used for visuals and animation
'							IMPORTANT!!! 
'							rotz must be used for orientation
'							rotx to bend the target back
'							transz to move it up and down
'							the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
'	switch:				ROM switch number
'	animate:			Array slot for handling the animation instrucitons, set to 0
'						Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target 
'   isDropped:			Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.

DT7 = Array(target7, target7a, T7_BM_Dark_Room, 1, 0, false)
DT8 = Array(target8, target8a, T8_BM_Dark_Room, 2, 0, false)
DT9 = Array(target9, target9a, T9_BM_Dark_Room, 3, 0, false)
DT90= Array(target90, target90a, T90_BM_Dark_Room,  4, 0, true)

Dim DTArray
DTArray = Array(DT7, DT8, DT9, DT90)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 60 'in milliseconds
Const DTDropUpSpeed = 30 'in milliseconds
Const DTDropUnits = 50 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 5 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 6 'max degrees primitive rotates when hit
Const DTDropDelay = 10 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 10 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.6 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
	Dim i
	i = DTArrayID(switch)

	PlayTargetSound
    debug.print "dthit "&i&" pre-state: "&DTArray(i)(4)
	DTArray(i)(4) =  DTCheckBrick(Activeball,DTArray(i)(2))
	If DTArray(i)(4) = 1 or DTArray(i)(4) = 3 or DTArray(i)(4) = 4 Then
		DTBallPhysics Activeball, DTArray(i)(2).rotz, DTMass
	End If
    debug.print "post-state: "&DTArray(i)(4)
	'DoDTAnim - No need for this, it's called on a timer every 10ms
End Sub

Sub DTRaise(switch)
	Dim i
	i = DTArrayID(switch)

	DTArray(i)(4) = -1
	DoDTAnim
End Sub

Sub DTDrop(switch)
	Dim i
	i = DTArrayID(switch)

	DTArray(i)(4) = 1
	DoDTAnim
End Sub

Function DTArrayID(switch)
	Dim i
	For i = 0 to uBound(DTArray) 
		If DTArray(i)(3) = switch Then DTArrayID = i:Exit Function 
	Next
End Function


sub DTBallPhysics(aBall, angle, mass)
	dim rangle,bangle,calc1, calc2, calc3
	rangle = (angle - 90) * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

	calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
	calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
	calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

	aBall.velx = calc1 * cos(rangle) + calc2
	aBall.vely = calc1 * sin(rangle) + calc3
End Sub


'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim) 
	dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
	rangle = (dtprim.rotz - 90) * 3.1416 / 180
	rangle2 = dtprim.rotz * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	bangleafter = Atn2(aBall.vely,aball.velx)

	Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
	Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

	cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

	perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
	paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

	perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle) 
	paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

	If perpvel > 0 and  perpvelafter <= 0 Then
		If DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
			DTCheckBrick = 3
		Else
			DTCheckBrick = 1
		End If
	ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
		DTCheckBrick = 4
	Else 
		DTCheckBrick = 0
	End If
End Function


Sub DoDTAnim()
	Dim i
	For i=0 to Ubound(DTArray)
		DTArray(i)(4) = DTAnimate(DTArray(i)(0),DTArray(i)(1),DTArray(i)(2),DTArray(i)(3),DTArray(i)(4))
	Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
	dim transz, switchid
	Dim animtime, rangle

	switchid = switch

	Dim ind
	ind = DTArrayID(switchid)

	rangle = prim.rotz * PI / 180

	DTAnimate = animate

	if animate = 0  Then
		primary.uservalue = 0
		DTAnimate = 0
		Exit Function
	Elseif primary.uservalue = 0 then 
		primary.uservalue = gametime
	end if

	animtime = gametime - primary.uservalue

	If (animate = 1 or animate = 4) and animtime < DTDropDelay Then
		primary.collidable = 0
	    If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
		prim.rotx = DTMaxBend * cos(rangle)
		prim.roty = DTMaxBend * sin(rangle)
		DTAnimate = animate
		Exit Function
    elseif (animate = 1 or animate = 4) and animtime > DTDropDelay Then
		primary.collidable = 0
		If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
		prim.rotx = DTMaxBend * cos(rangle)
		prim.roty = DTMaxBend * sin(rangle)
		animate = 2
		SoundDropTargetDrop prim
	End If

	if animate = 2 Then
		
		transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
		if prim.transz > -DTDropUnits  Then
			prim.transz = transz
		end if

		prim.rotx = DTMaxBend * cos(rangle)/2
		prim.roty = DTMaxBend * sin(rangle)/2

		if prim.transz <= -DTDropUnits Then 
			prim.transz = -DTDropUnits
			secondary.collidable = 0
			DTArray(ind)(5) = true 'Mark target as dropped
			DTAction switchid
			primary.uservalue = 0
			DTAnimate = 0
			Exit Function
		Else
			DTAnimate = 2
			Exit Function
		end If 
	End If

	If animate = 3 and animtime < DTDropDelay Then
		primary.collidable = 0
		secondary.collidable = 1
		prim.rotx = DTMaxBend * cos(rangle)
		prim.roty = DTMaxBend * sin(rangle)
	elseif animate = 3 and animtime > DTDropDelay Then
		primary.collidable = 1
		secondary.collidable = 0
		prim.rotx = 0
		prim.roty = 0
		primary.uservalue = 0
		DTAnimate = 0
		Exit Function
	End If

	if animate = -1 Then
		transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

		If prim.transz = -DTDropUnits Then
			Dim b, BOT
			BOT = GetBalls

			For b = 0 to UBound(BOT)
				If InRotRect(BOT(b).x,BOT(b).y,prim.x, prim.y, prim.rotz, -25,-10,25,-10,25,25,-25,25) and BOT(b).z < prim.z+DTDropUnits+25 Then
					BOT(b).velz = 20
				End If
			Next
		End If

		if prim.transz < 0 Then
			prim.transz = transz
		elseif transz > 0 then
			prim.transz = transz
		end if

		if prim.transz > DTDropUpUnits then 
			DTAnimate = -2
			prim.transz = DTDropUpUnits
			prim.rotx = 0
			prim.roty = 0
			primary.uservalue = gametime
		end if
		primary.collidable = 0
		secondary.collidable = 1
		DTArray(ind)(5) = false 'Mark target as not dropped

	End If

	if animate = -2 and animtime > DTRaiseDelay Then
		prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits 
		if prim.transz < 0 then
			prim.transz = 0
			primary.uservalue = 0
			DTAnimate = 0

			primary.collidable = 1
			secondary.collidable = 0
		end If 
	End If
End Function

'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS 
'******************************************************


' Used for drop targets
'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
	Dim AB, BC, CD, DA
	AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
	BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
	CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
	DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

	If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
		InRect = True
	Else
		InRect = False       
	End If
End Function

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
    Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
    Dim rotxy
    rotxy = RotPoint(ax,ay,angle)
    rax = rotxy(0)+px : ray = rotxy(1)+py
    rotxy = RotPoint(bx,by,angle)
    rbx = rotxy(0)+px : rby = rotxy(1)+py
    rotxy = RotPoint(cx,cy,angle)
    rcx = rotxy(0)+px : rcy = rotxy(1)+py
    rotxy = RotPoint(dx,dy,angle)
    rdx = rotxy(0)+px : rdy = rotxy(1)+py

    InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function


'******************************************************
'****  END DROP TARGETS
'******************************************************

'******************************************************
'		STAND-UP TARGET INITIALIZATION
'******************************************************

'Define a variable for each stand-up target
Dim ST04, ST05, ST06, ST07, ST08, ST09, ST10, ST11, ST12, ST13

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
' 	primary: 			vp target to determine target hit
'	prim:				primitive target used for visuals and animation
'							IMPORTANT!!! 
'							transy must be used to offset the target animation
'	switch:				ROM switch number
'	animate:			Arrary slot for handling the animation instrucitons, set to 0
' 
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

ST04 = Array(Target32, targets_004_BM_Dark_Room, 4,  0)
ST05 = Array(Target33, targets_005_BM_Dark_Room, 5,  0)
ST06 = Array(Target34, targets_006_BM_Dark_Room, 6,  0)
ST07 = Array(Target35, targets_007_BM_Dark_Room, 7,  0)
ST08 = Array(Target36, targets_008_BM_Dark_Room, 8,  0)
ST09 = Array(Target43, targets_009_BM_Dark_Room, 9,  0)
ST10 = Array(Target44, targets_010_BM_Dark_Room, 10, 0)
ST11 = Array(Target82, targets_011_BM_Dark_Room, 11, 0)
ST12 = Array(Target81, targets_012_BM_Dark_Room, 12, 0)
ST13 = Array(Target80, targets_BM_Dark_Room    , 13, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST04, ST05, ST06, ST07, ST08, ST09, ST10, ST11, ST12, ST13)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5 				'vpunits per animation step (control return to Start)
Const STMaxOffset = 9 			'max vp units target moves when hit

Const STMass = 0.2				'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'				STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
	Dim i
	i = STArrayID(switch)

	PlayTargetSound
	STArray(i)(3) =  STCheckHit(Activeball,STArray(i)(0))

	If STArray(i)(3) <> 0 Then
        ' The line below slows down the ball, the same as a drop target does, but standup targets don't behave that way
		'DTBallPhysics Activeball, STArray(i)(0).orientation, STMass
	End If
	DoSTAnim
End Sub

Function STArrayID(switch)
	Dim i
	For i = 0 to uBound(STArray) 
		If STArray(i)(2) = switch Then STArrayID = i:Exit Function 
	Next
End Function

'Check if target is hit on it's face
Function STCheckHit(aBall, target) 
	dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
	rangle = (target.orientation - 90) * 3.1416 / 180	
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	bangleafter = Atn2(aBall.vely,aball.velx)

	perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
	paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

	perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle) 
	paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

	If perpvel > 0 and  perpvelafter <= 0 Then
		STCheckHit = 1
	ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
		STCheckHit = 1
	Else 
		STCheckHit = 0
	End If
End Function

Sub DoSTAnim()
	Dim i
	For i=0 to Ubound(STArray)
		STArray(i)(3) = STAnimate(STArray(i)(0),STArray(i)(1),STArray(i)(2),STArray(i)(3))
	Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
	Dim animtime

	STAnimate = animate

	if animate = 0  Then
		primary.uservalue = 0
		STAnimate = 0
		Exit Function
	Elseif primary.uservalue = 0 then 
		primary.uservalue = gametime
	end if

	animtime = gametime - primary.uservalue

	If animate = 1 Then
		primary.collidable = 0
		prim.transy = -STMaxOffset
        STAction switch
		STAnimate = 2
		Exit Function
	elseif animate = 2 Then
		prim.transy = prim.transy + STAnimStep
		If prim.transy >= 0 Then
			prim.transy = 0
			primary.collidable = 1
			STAnimate = 0
			Exit Function
		Else 
			STAnimate = 2
		End If
	End If	
End Function

Sub STAction(Switch)
	'TODO add Select Case for each standup target
End Sub




'******************************************************
'		END STAND-UP TARGETS
'******************************************************



'***********************************************************************
'* TABLE OPTIONS *******************************************************
'***********************************************************************

Dim Cabinetmode				'0 - Siderails On, 1 - Siderails Off
Dim OutPostMod
Dim LightLevel : LightLevel = 50
VolumeDial = 0.8			'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim RampRollVolume : RampRollVolume = 0.5 	'Level of ramp rolling volume. Value between 0 and 1
DynamicBallShadowsOn = 1		'0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow

' Base options
Const Opt_Light = 0
'Const Opt_Outpost = 1
Const Opt_Flames = 1
Const Opt_Volume = 2
Const Opt_Volume_Ramp = 3
Const Opt_Volume_Ball = 4
' Table mods & toys
Const Opt_Cabinet = 5
Const Opt_LockBar = 6
' Advanced options
Const Opt_DynBallShadow = 7
' Informations
Const Opt_Info_1 = 9
Const Opt_Info_2 = 10
Const Opt_LUT = 8

Const NOptions = 11

Dim OptionDMD: Set OptionDMD = Nothing
Dim bOptionsMagna, bInOptions, optionsSaveAttract : bOptionsMagna = False
Dim OptPos, OptSelected, OptN, OptTop, OptBot, OptSel
Dim OptFontHi, OptFontLo
Dim Options_Flames_Enabled

Sub Options_Open
    optionsSaveAttract=bAttractMode
    bAttractMode=False
    tmrDMDUpdate.Enabled = False
    tmrAttractModeScene.Enabled = False
	
    bOptionsMagna = False
	
	Debug.Print "Option UI opened"
	bInOptions = True
	OptPos = 0
	OptSelected = False
	
	Dim a, scene, font
	Set scene = FlexDMD.NewGroup("Options")
	Set OptFontHi = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
	Set OptFontLo = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(100, 100, 100), RGB(100, 100, 100), 0)
	Set OptSel = FlexDMD.NewGroup("Sel")
	Set a = FlexDMD.NewLabel(">", OptFontLo, ">>>")
	a.SetAlignedPosition 1, 16, FlexDMD_Align_Left
	OptSel.AddActor a
	Set a = FlexDMD.NewLabel(">", OptFontLo, "<<<")
	a.SetAlignedPosition 127, 16, FlexDMD_Align_Right
	OptSel.AddActor a
	scene.AddActor OptSel
	OptSel.SetBounds 0, 0, 128, 32
	OptSel.Visible = False
	
	Set a = FlexDMD.NewLabel("Info1", OptFontLo, "MAGNA EXIT/ENTER")
	a.SetAlignedPosition 1, 32, FlexDMD_Align_BottomLeft
	scene.AddActor a
	Set a = FlexDMD.NewLabel("Info2", OptFontLo, "FLIPPER SELECT")
	a.SetAlignedPosition 127, 32, FlexDMD_Align_BottomRight
	scene.AddActor a
	Set OptN = FlexDMD.NewLabel("Pos", OptFontLo, "LINE 1")
	Set OptTop = FlexDMD.NewLabel("Top", OptFontLo, "LINE 1")
	Set OptBot = FlexDMD.NewLabel("Bottom", OptFontLo, "LINE 2")
	scene.AddActor OptN
	scene.AddActor OptTop
	scene.AddActor OptBot
	Options_OnOptChg
	DMDDisplayScene scene
End Sub

Sub Options_Close
	bInOptions = False
	bAttractMode=optionsSaveAttract
    if bAttractMode then
        StartAttractMode
    Else
        bDefaultScene = False
        tmrDMDUpdate.Enabled = True
    End if 
End Sub

Function Options_OnOffText(opt)
	If opt Then
		Options_OnOffText = "ON"
	Else
		Options_OnOffText = "OFF"
	End If
End Function

Sub Options_OnOptChg
	FlexDMD.LockRenderThread
	OptN.Text = (OptPos+1) & "/" & NOptions
	If OptSelected Then
		OptTop.Font = OptFontLo
		OptBot.Font = OptFontHi
		OptSel.Visible = True
	Else
		OptTop.Font = OptFontHi
		OptBot.Font = OptFontLo
		OptSel.Visible = False
	End If
	If OptPos = Opt_Light Then
		OptTop.Text = "LIGHT LEVEL"
		OptBot.Text = "LEVEL " & LightLevel
		SaveValue cGameName, "LIGHT", LightLevel
	' ElseIf OptPos = Opt_Outpost Then
	' 	OptTop.Text = "OUT POST DIFFICULTY"
	' 	If OutPostMod = 0 Then
	' 		OptBot.Text = "EASY"
	' 	ElseIf OutPostMod = 1 Then
	' 		OptBot.Text = "MEDIUM"
	' 	ElseIf OutPostMod = 2 Then
	' 		OptBot.Text = "HARD"
	' 	ElseIf OutPostMod = 3 Then
	' 		OptBot.Text = "HARDEST"
	' 	End If
	' 	SaveValue cGameName, "OUTPOST", OutPostMod
    ElseIf OptPos = Opt_Flames Then
        OptTop.Text = "DRAGON FLAMES"
        OptBot.Text = Options_OnOffText(Options_Flames_Enabled)
        SaveValue cGameName, "DRAGONFLAMES", Options_Flames_Enabled
	ElseIf OptPos = Opt_Volume Then
		OptTop.Text = "MECH VOLUME"
		OptBot.Text = "LEVEL " & CInt(VolumeDial * 100)
		SaveValue cGameName, "VOLUME", VolumeDial
	ElseIf OptPos = Opt_Volume_Ramp Then
		OptTop.Text = "RAMP VOLUME"
		OptBot.Text = "LEVEL " & CInt(RampRollVolume * 100)
		SaveValue cGameName, "RAMPVOLUME", RampRollVolume
	ElseIf OptPos = Opt_Volume_Ball Then
		OptTop.Text = "BALL VOLUME"
		OptBot.Text = "LEVEL " & CInt(BallRollVolume * 100)
		SaveValue cGameName, "BALLVOLUME", BallRollVolume
	ElseIf OptPos = Opt_Cabinet Then
		OptTop.Text = "CABINET MODE"
		OptBot.Text = Options_OnOffText(CabinetMode)
		SaveValue cGameName, "CABINET", CabinetMode
    ElseIf OptPos = Opt_LockBar Then
        OptTop.text = "PHYSICAL LOCKBAR BUTTON"
        If bHaveLockbarButton Then OptBot.Text = "YES" Else OptBot.Text = "NO"
        SaveValue cGameName, "LBBUT", bHaveLockbarButton
	ElseIf OptPos = Opt_DynBallShadow Then
		OptTop.Text = "DYN. BALL SHADOWS"
		OptBot.Text = Options_OnOffText(DynamicBallShadowsOn)
		SaveValue cGameName, "DYNBALLSH", DynamicBallShadowsOn
	ElseIf OptPos = Opt_Info_1 Then
		OptTop.Text = "VPX " & VersionMajor & "." & VersionMinor & "." & VersionRevision
		OptBot.Text = "GAME OF THRONES " & TableVersion
	ElseIf OptPos = Opt_Info_2 Then
		OptTop.Text = "RENDER MODE"
		If RenderingMode = 0 Then OptBot.Text = "DEFAULT"
		If RenderingMode = 1 Then OptBot.Text = "STEREO 3D"
		If RenderingMode = 2 Then OptBot.Text = "VR"
    ElseIf OptPos = Opt_LUT Then
        OptTop.Text = "LUT"
        OptBot.Text = LUTImage
	End If
	OptTop.SetAlignedPosition 127, 1, FlexDMD_Align_TopRight
	OptBot.SetAlignedPosition 64, 16, FlexDMD_Align_Center
	FlexDMD.UnlockRenderThread
	UpdateMods
End Sub

Sub Options_Toggle(amount)
	If OptPos = Opt_Light Then
		LightLevel = LightLevel + amount * 10
		If LightLevel < 0 Then LightLevel = 100
		If LightLevel > 100 Then LightLevel = 0
	' ElseIf OptPos = Opt_Outpost Then
	' 	OutPostMod = OutPostMod + amount
	' 	If OutPostMod < 0 Then OutPostMod = 3
	' 	If OutPostMod > 3 Then OutPostMod = 0
    ElseIf OptPos = Opt_Flames Then
        Options_Flames_Enabled = 1 - Options_Flames_Enabled
	ElseIf OptPos = Opt_Volume Then
		VolumeDial = VolumeDial + amount * 0.1
		If VolumeDial < 0 Then VolumeDial = 1
		If VolumeDial > 1 Then VolumeDial = 0
	ElseIf OptPos = Opt_Volume_Ramp Then
		RampRollVolume = RampRollVolume + amount * 0.1
		If RampRollVolume < 0 Then RampRollVolume = 1
		If RampRollVolume > 1 Then RampRollVolume = 0
	ElseIf OptPos = Opt_Volume_Ball Then
		BallRollVolume = BallRollVolume + amount * 0.1
		If BallRollVolume < 0 Then BallRollVolume = 1
		If BallRollVolume > 1 Then BallRollVolume = 0
	ElseIf OptPos = Opt_Cabinet Then
		CabinetMode = 1 - CabinetMode
    ElseIf OptPos = Opt_LockBar Then
        bHaveLockbarButton = 1 - bHaveLockbarButton
	ElseIf OptPos = Opt_DynBallShadow Then
		DynamicBallShadowsOn = 1 - DynamicBallShadowsOn
    Elseif OptPos = Opt_LUT Then
        LUTImage = LUTImage + amount
        If LUTImage < 0 Then LUTImage = 15
        LUTImage = ABS(LUTImage) MOD 16
        table1.ColorGradeImage = "LUT"&LUTImage
        SaveLUT
	End If
End Sub

Sub Options_KeyDown(ByVal keycode)
	If OptSelected Then
		If keycode = LeftMagnaSave Then ' Exit / Cancel
			OptSelected = False
		ElseIf keycode = RightMagnaSave Then ' Enter / Select
			OptSelected = False
		ElseIf keycode = LeftFlipperKey Then ' Next / +
			Options_Toggle	-1
		ElseIf keycode = RightFlipperKey Then ' Prev / -
			Options_Toggle	1
		End If
	Else
		If keycode = LeftMagnaSave Then ' Exit / Cancel
			Options_Close
		ElseIf keycode = RightMagnaSave Then ' Enter / Select
			If OptPos < Opt_Info_1 Then OptSelected = True
		ElseIf keycode = LeftFlipperKey Then ' Next / +
			OptPos = OptPos - 1
			If OptPos < 0 Then OptPos = NOptions - 1
		ElseIf keycode = RightFlipperKey Then ' Prev / -
			OptPos = OptPos + 1
			If OptPos >= NOPtions Then OptPos = 0
		End If
	End If
	Options_OnOptChg
End Sub

Sub Options_Load
	Dim x
    x = LoadValue(cGameName, "LIGHT")
    If x <> "" Then LightLevel = CInt(x) Else LightLevel = 50
    'x = LoadValue(cGameName, "OUTPOST")
    'If x <> "" Then OutPostMod = CInt(x) Else OutPostMod = 1
    x = LoadValue(cGameName,"DRAGONFLAMES")
    If x <> "" Then Options_Flames_Enabled = CInt(x) Else Options_Flames_Enabled = 1
    x = LoadValue(cGameName, "VOLUME")
    If x <> "" Then VolumeDial = CNCDbl(x) Else VolumeDial = 0.8
    x = LoadValue(cGameName, "RAMPVOLUME")
    If x <> "" Then RampRollVolume = CNCDbl(x) Else RampRollVolume = 0.5
    x = LoadValue(cGameName, "BALLVOLUME")
    If x <> "" Then BallRollVolume = CNCDbl(x) Else BallRollVolume = 0.5
    x = LoadValue(cGameName, "CABINET")
    If x <> "" Then CabinetMode = CInt(x) Else CabinetMode = 0
    x = LoadValue(cGameName,"LBBUT")
    If x <> "" Then bHaveLockbarButton = CInt(x) Else bHaveLockbarButton = False
    x = LoadValue(cGameName, "DYNBALLSH")
    If x <> "" Then DynamicBallShadowsOn = CInt(x) Else DynamicBallShadowsOn = 1
	UpdateMods
End Sub

Sub UpdateMods
	Dim y, enabled,nenabled

    if (CabinetMode <> 0  And bHaveLockbarButton = 0) then enabled=1 else enabled=0
    nenabled = 1-enabled
    For each y in ApronFButton_BL : y.visible=abs(enabled) : Next
    For each y in fb_base_BL : y.visible=abs(enabled) : Next
    For each y in LockdownButton_BL : y.visible = nenabled : Next
    For each y in fl_base_BL : y.visible = nenabled : Next
	
    if CabinetMode=0 then enabled=1 else enabled=0
    SideRailMapLeft.visible = enabled : SideRailMapRight.visible = enabled

    Select Case LUTImage
        Case 14: FlexDMD.Color = RGB(16,255,16)
        Case 15: FlexDMD.Color = RGB(255,255,255)
        Case Else: FlexDMD.Color = RGB(255,44,16)
    End Select

    ' Not Implemented
	' If OutPostMod = 0 Then ' Easy
	' 	routpost_BM_Dark_Room.x = 817.4415
	' 	routpost_BM_Dark_Room.y = 1469.287
	' 	loutpost_BM_Dark_Room.x = 51.87117
	' 	loutpost_BM_Dark_Room.y = 1469.451
	' 	rout_easy.collidable = true
	' 	rout_medium.collidable = false
	' 	rout_hard.collidable = false
	' 	rout_hardest.collidable = false
	' 	lout_easy.collidable = true
	' 	lout_medium.collidable = false
	' 	lout_hard.collidable = false
	' 	lout_hardest.collidable = false
	' ElseIf OutPostMod = 1 Then ' Medium
	' 	routpost_BM_Dark_Room.x = 823.0609
	' 	routpost_BM_Dark_Room.y = 1459.596
	' 	loutpost_BM_Dark_Room.x = 48.93166
	' 	loutpost_BM_Dark_Room.y = 1457.558
	' 	rout_easy.collidable = false
	' 	rout_medium.collidable = true
	' 	rout_hard.collidable = false
	' 	rout_hardest.collidable = false
	' 	lout_easy.collidable = false
	' 	lout_medium.collidable = true
	' 	lout_hard.collidable = false
	' 	lout_hardest.collidable = false
	' ElseIf OutPostMod = 2 Then ' Hard
	' 	routpost_BM_Dark_Room.x = 829.0158
	' 	routpost_BM_Dark_Room.y = 1449.152
	' 	loutpost_BM_Dark_Room.x = 45.53271
	' 	loutpost_BM_Dark_Room.y = 1446.529
	' 	rout_easy.collidable = false
	' 	rout_medium.collidable = false
	' 	rout_hard.collidable = true
	' 	rout_hardest.collidable = false
	' 	lout_easy.collidable = false
	' 	lout_medium.collidable = false
	' 	lout_hard.collidable = true
	' 	lout_hardest.collidable = false
	' ElseIf OutPostMod = 3 Then ' Hardest
	' 	routpost_BM_Dark_Room.x = 834.4323
	' 	routpost_BM_Dark_Room.y = 1439.318
	' 	loutpost_BM_Dark_Room.x = 42.86211
	' 	loutpost_BM_Dark_Room.y = 1435.624
	' 	rout_easy.collidable = false
	' 	rout_medium.collidable = false
	' 	rout_hard.collidable = false
	' 	rout_hardest.collidable = true
	' 	lout_easy.collidable = false
	' 	lout_medium.collidable = false
	' 	lout_hard.collidable = false
	' 	lout_hardest.collidable = true
	' End If

	' ' FIXME this should be setup in Blender
	' routpost_LM_GI.x = routpost_BM_Dark_Room.x
	' routpost_LM_GI.y = routpost_BM_Dark_Room.y
	' routpost_LM_Lit_Room.x = routpost_BM_Dark_Room.x
	' routpost_LM_Lit_Room.y = routpost_BM_Dark_Room.y
	' routpost_LM_flashers_l130.x = routpost_BM_Dark_Room.x
	' routpost_LM_flashers_l130.y = routpost_BM_Dark_Room.y
	' routpost_LM_flashers_l131.x = routpost_BM_Dark_Room.x
	' routpost_LM_flashers_l131.y = routpost_BM_Dark_Room.y
	' routpost_LM_flashers_l132.x = routpost_BM_Dark_Room.x
	' routpost_LM_flashers_l132.y = routpost_BM_Dark_Room.y

	' ' FIXME this should be setup in Blender
	' loutpost_LM_GI.x = loutpost_BM_Dark_Room.x
	' loutpost_LM_GI.y = loutpost_BM_Dark_Room.y
	' loutpost_LM_Lit_Room.x = loutpost_BM_Dark_Room.x
	' loutpost_LM_Lit_Room.y = loutpost_BM_Dark_Room.y
	' loutpost_LM_flashers_l130.x = loutpost_BM_Dark_Room.x
	' loutpost_LM_flashers_l130.y = loutpost_BM_Dark_Room.y
	' loutpost_LM_flashers_l131.x = loutpost_BM_Dark_Room.x
	' loutpost_LM_flashers_l131.y = loutpost_BM_Dark_Room.y

	Lampz.state(150) = LightLevel
End Sub

' Culture neutral string to double conversion (handles situation where you don't know how the string was written)
Function CNCDbl(str)
    Dim strt, Sep, i
    If IsNumeric(str) Then
        CNCDbl = CDbl(str)
    Else
        Sep = Mid(CStr(0.5), 2, 1)
        Select Case Sep
        Case "."
            i = InStr(1, str, ",")
        Case ","
            i = InStr(1, str, ".")
        End Select
        If i = 0 Then     
            CNCDbl = Empty
        Else
            strt = Mid(str, 1, i - 1) & Sep & Mid(str, i + 1)
            If IsNumeric(strt) Then
                CNCDbl = CDbl(strt)
            Else
                CNCDbl = Empty
            End If
        End If
    End If
End Function













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
Const Scoop=7
Const StarTargets=8


Dim PlayerMode
Dim LockLit
Dim PlayerState(4)
Dim CompletedModes              ' bitmask of completed modes (songs)

Dim PlayfieldMultiplierVal
Dim DemonPlayfieldMultiplierVal

Dim bKissTargets(4)
Dim bArmyTargets(4)
Dim bLockTargets(2)
Dim bStarTargets(4)
Dim bKissLights(4)
Dim bArmyLights(4)
Dim KissTargetsCompleted        ' Number of times KISS target bank has been completed
Dim ArmyTargetsCompleted
Dim DemonTargetsCompleted
Dim DemonLockLit
Dim DemonMBCompleted

Dim EBisLit

Dim ModeSongs
ModeSongs = Array("","DUECE","HOTTER  THAN  HELL","LICK  IT  UP","SHOUT  IT  OUT  LOUD","DETROIT  ROCK  CITY","ROCK  &  ROLL  ALL  NIGHT","LOVE  IT  LOUD","BLACK  DIAMOND")
Dim ModeWidth   ' Width the song title in pixels, for the progress bar
ModeWidth = Array(0,30,85,50,85,89,76,62,73)

Dim ModeLights
ModeLights = Array(li90,li90,li100,li105,li111,li124,li130)
Dim ModeColours
ModeColours = Array(ice,rockandroll=ice,detroitrockcity=yellow,lickitup=white,loveitloud=amber)
Dim ComboShotPattern(15)
Dim ComboShotProgress(15)
ComboShotPattern(1) = Array(2,"CHICAGO",CenterRamp,RightRamp)
ComboShotPattern(2) = Array(2,"PITTSBURGH",RightRamp,CenterRamp)
ComboShotPattern(3) = Array(2,"SEATTLE",LeftOrbit,Scoop)
ComboShotPattern(4) = Array(2,"PORTLAND",LeftOrbit,StarTargets)
ComboShotPattern(5) = Array(3,"LOS ANGELES",RightRamp,CenterRamp,DemonShot)
ComboShotPattern(6) = Array(3,"HOUSTON",RightRamp,CenterRamp,RightOrbit)
ComboShotPattern(7) = Array(3,"NEW ORLEANS",CenterRamp,RightRamp,LeftOrbit)
ComboShotPattern(8) = Array(3,"ATLANTA",CenterRamp,RightRamp,Scoop)
ComboShotPattern(9) = Array(3,"ORLANDO",CenterRamp,RightRamp,StarTargets)
ComboShotPattern(10) = Array(4,"MEXICO CITY",CenterRamp,RightRamp,LeftOrbit,Scoop)
ComboShotPattern(11) = Array(4,"TOKYO",CenterRamp,RightRamp,LeftOrbit,StarTargets)
ComboShotPattern(12) = Array(4,"LONDON",CenterRamp,RightRamp,CenterRamp,RightRamp)
ComboShotPattern(13) = Array(4,"NEW YORK",RightRamp,CenterRamp,RightRamp,CenterRamp)
ComboShotPattern(14) = Array(5,"SAN FRANCISCO",CenterRamp,CenterRamp,CenterRamp,CenterRamp,RightOrbit)
ComboShotPattern(15) = Array(10,"DETROIT",CenterRamp,RightRamp,CenterRamp,RightRamp,CenterRamp,RightRamp,CenterRamp,RightRamp,CenterRamp,RightRamp)


' The cPState class is to save and restore player state
Class cPState
    Dim i
    Dim myKissTargets(4)
    Dim myArmyTargets(4)
    Dim myStarTargets(4)
    Dim myLockTargets(2)
    Dim myCompletedModes
    Dim myKissTargetsCompleted
    Dim myArmyTargetsCompleted
    Dim myDemonTargetsCompleted
    Dim myDemonLockLit
    Dim myDemonMBCompleted
    Dim myEBisLit

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
        myDemonTargetsCompleted = DemonTargetsCompleted
        myDemonLockLit = DemonLockLit
        myDemonMBCompleted = DemonMBCompleted
        myEBisLit = EBisLit
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
        DemonTargetsCompleted = myDemonTargetsCompleted
        DemonLockLit = myDemonLockLit
        DemonMBCompleted = myDemonMBCompleted
        EBisLit = myEBisLit
    End Sub
End Class

' The cMode class is primarily for tracking Mode (Song) progress towards completion. However, since
' all songs are completed using the 6 primary shots, and these shots are also used for Combos, Jackpots,
' and Instrument collection, we track those in this class too

Class cMode
    Dim ModeState
    Dim CurrentMode
    Dim CurrentModeSong     ' Song continues to play when mode is completed, so track it separately
    Dim ShotMask
    Dim ModeShotsCompleted  ' Used for scoring - resets at beginning of ball
    Dim ShotsMade           ' Used to track progress in Detroit Rock City and Love It Loud
    Dim SongsCompletedThisBall
    Dim BaseScore
    Dim ScoreIncr
    Dim ModeScore
    Dim bModeComplete
    Dim ModeColour
    Dim ModeProgress(8)     ' Progress of each song, in case an unfinished song is resumed later
    Dim ComboShotProgress(15)
    Dim ComboShotsAwarded    ' bitmask of combo shots already awarded
    Dim bComboEBAwarded     ' Whether the Combo Extra Ball has been awarded yet
    Dim bOnFirstBall

    Private Sub Class_Initialize(  )
        Dim i
        For i = 1 to 8 : ModeProgress(i) = 0
        ComboShotsAwarded = 0
        bComboEBAwarded = False
        bOnFirstBall = True
        CurrentMode = 0
    End Sub

    Public Property Get GetMode : GetMode = CurrentMode : End Property
    Public Property Let SetMode(m) : CurrentMode = m : CurrentModeSong = m : ModeShotsCompleted = ModeProgress(m) : bOnFirstBall = False : End Property

    Function CanChooseNewSong
        If (ModeShotsCompleted > 6 And bBallInPlungerLane) Or CurrentMode = 0 Then CanChooseNewSong = True Else CanChooseNewSong = False
    End Function

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
        ModeShotsCompleted = ModeProgress(CurrentMode)
        BaseScore = 250000
        Select Case CurrentMode
            Case Duece:           ModeColour=blue   : BaseScore = 300000 : ScoreIncr = 300000 : ShotMask=6      'two left-most shots lit
            Case Hotter:          ModeColour=white  : ShotMask=128 : ScoreIncr = 300000  'KISS targets
            Case LickItUp:        ModeColour=magenta: BaseScore = 250000 : ScoreIncr = 300000 : i=RndNbr(5) : ShotMask = 2^i + 2^(i+1)    ' two neighbouring shots lit
            Case ShoutIt:         ModeColour=cyan   : BaseScore = 500000 : ScoreIncr = 200000 : ShotMask = 2^CenterRamp
            Case DetroitRockCity: ModeColour=yellow : BaseScore = 200000 : ScoreIncr = 200000 : ShotMask = 126    ' All major shots lit
            Case RockAllNight:    ModeColour=ice    : BaseScore = 500000 : ScoreIncr = 200000 : ShotMask = 28
            Case LoveItLoud:      ModeColour=amber  : BaseScore = 500000 : ScoreIncr = 200000 : ShotMask = 2^LeftOrbit + 2^RightOrbit
            Case BlackDiamond:    ModeColour=green  : BaseScore = 300000 : ScoreIncr = 300000 : Shotmask = 2^RightOrbit
        End Select
        
        SetModeLights
        DMDPlayModeIntroScene CurrentMode
    End Sub

    Sub ResetForNewBall
        Dim i
        For i = 1 to 15 : ComboShotProgress(i) = 0 : Next
        ModeScore = 0
        SongsCompletedThisBall = 0
    End Sub

    ' Check switch hits against mode (song) rules
    Sub CheckModeHit(sw)
        Dim AwardScore,i,scenetype,maxhits
        scenetype=1 ' Video followed by dimmed image with score over top

        if (ShotMask And (2^sw)) = 0 Or bModeComplete Then CheckComboShots sw : Exit Sub' Shot wasn't lit

        'TODO: Award score is doubled sometimes
        AwardScore = BaseScore+ScoreIncr*ModeShotsCompleted

        ModeShotsCompleted = ModeShotsCompleted + 1
        ModeProgress(CurrentMode) = ModeShotsCompleted

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
                scenetype=2
                ModeState = ModeState+1
                If ModeState >= 9 Then
                    bModeComplete = True
                ElseIf (ModeState MOD 2) = 0 Then
                    ShotMask = 2^RndNbr(6)
                Else
                    ShotMask = 128
                End If
                If ModeState > 7 Then AwardScore=AwardScore*2 'TODO is this right??
            Case LickItUp:
                scenetype=2
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
                    Case 2,4,6,8: ShotMask=14 ' left orbit, STAR, Center ramp
                    Case 1,3,5,7: ShotMask=112 ' right orbit, right ramp, demon
                    Case 10: bModeComplete = True
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
        DMDPlayHitScene 
        AddScore AwardScore
        ModeScore = ModeScore+AwardScore

        if bModeComplete Then 
            ModeProgress(CurrentMode) = 10
            CurrentMode = 0
            ' Check to see whether all songs have been completed. If so, reset them
            Dim t : t=true
            For i = 1 to 8
                If ModeProgress(i) < 10 then t=false : Exit For
            Next
            if t Then For i = 1 to 8 : ModeProgress(i) = 0 : Next
            SetScoopLights
            'TODO: Play a scene with song name and mode score
            SongsCompletedThisBall = SongsCompletedThisBall + 1
            if (SongsCompletedThisBall MOD DMDStd(kDMDFet_SongCompletionsForPFX)) = 0 Then
                Select Case PlayfieldMultiplierVal
                    Case 1: PlayfieldMultiplierVal = 2
                    Case 2: PlayfieldMultiplierVal = 3
                    Case Else: PlayfieldMultiplierVal = 5
                End Select
            End if                
            'TODO: What else happens when song is completed?
        End If

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

    ' Track shots made towards each City Combo. Each combo has its own progress counter.
    ' As long as shots made keep matching a City Combo's pattern, the progress counter increases for that combo.
    ' As soon as it doesn't match, the progress is reset. As long as combo shots continue to be made
    ' before the tmrComboShots elapses, progress will count up. When the timer fires, all progress resets to 0
    ' Combos can only be awarded once, until all cities have been collected, at which time they all reset 
    Sub CheckComboShots(shot)
        Dim i,awarded,found : found=false : awarded = 0
        For i = 1 to 15
            if (ComboShotsAwarded And 2^i) = 0 Then 'combo still in contention
                if ComboShotProgress(i) < ComboShotPattern(i)(0) Then
                    if ComboShotPattern(i)(ComboShotProgress(i)+2) <> sw Then ' not this combo pattern
                        ComboShotProgress(i) = 0   ' This combo is now out of contention, reset its shot progress to 0
                    Elseif ComboShotProgress(i) = ComboShotPattern(i)(0)-1 Then 
                        ' completed all shots in this Combo - award it!
                        ComboShotsAwarded = ComboShotsAwarded Or 2^i
                        DoComboShotAward i
                    Else
                        ' shot is part of combo, but combo not complete yet - keep counting up
                        found=true
                        ComboShotProgress(i) = ComboShotProgress(i)+1
                    End if
                End if
            End If
            if (ComboShotsAwarded And 2^i) > 0 Then awarded=awarded+1
        Next
        if found Then
            tmrComboShots.Enabled = False
            tmrComboShots.Interval = 5000 ' TODO: Verify how long we have before combo times out
            tmrComboShots.EnableBallSaver = True
        Else
            ResetComboShots
            tmrComboShots.Enabled = False
        End if

        if bComboEBAwarded=False And awarded = DMDStd(kDMDFet_CityComboEBat) Then bComboEBAwarded=True : EBisLit=EBisLit+1

        If ComboShotsAwarded = &HFFFE Then
            ' All Combos collected! Big bonus
            DoComboShotAward 16
            ComboShotsAwarded = 0
            bComboEBAwarded = False
            ResetComboShots
        End If
    End Sub

    Sub ResetComboShots : Dim i : For i = 1 to 15 : ComboShotProgress(i) = 0 : Next : End Sub

End Class

Sub tmrDueceMode_Timer : Mode(CurrentPlayer).DueceTimer : End Sub

Sub tmrComboShots_Timer : Mode(CurrentPlayer).ResetComboShots : End Sub



' Demon ball lock logic:
' - When DemonMBsCompleted=0
'   - hit both demon targets to light Lock. "Lock Lit" is stackable, so it can be lit "twice" at once
' - When DemonMBCompleted=1
'   - hit both demon targets to light lock if not already lit
' - When DemonMBCompleted>1
'   - Need to complete both demon targets twice to light lock
'
' - Shoot ball into Demon's mouth
'  - Kick ball up VUK into ramp and hold there with invisible wall
'  - If Lock is not lit, eject one ball from ramp after 500ms? delay
'  - If lock is lit:
'    - If less than 3 balls locked in ramp, release new ball from trough
'    - If 3 balls in ramp, release a ball from ramp
'    - If this is player's third locked ball, trigger start of Demon MB
'
' Ultimately, logic should be almost identical to BWMultiball in GoT - replace sword with Demon ramp

Dim bDemonTargets(2)
Dim DemonTargetLights(2)
DemonTargetLights = Array(li114,li116)

Sub DoDemonTargetHit(t)
    Dim t1 : t1=0
    if t=0 then t1=1
    If bDemonTargets(t) > 0 Then
        ' Target was already lit. TODO: Play a sound?
        Exit Sub
    End If
    If bDemonTargets(t1) = 1 Then
        ' Other target was lit. Complete the bank
        Dim DoLock : DoLock=False
        bDemonTargets(t1)=0
        if DemonMBCompleted < 2 Then
            If DemonMBCompleted < 1 Or DemonMBCompleted=0 Then DemonLockLit=DemonLockLit+1 : DoLock=True
        Else
            DemonTargetsCompleted = DemonTargetsCompleted + 1
            If DemonTargetsCompleted = 2 Then DemonLockLit=1 : DemonTargetsCompleted = 0 : DoLock=True
        End if
        If DoLock Then
            'TODO: Play Sound
            'TODO: Does Demon LOCK light flash or go solid when lit?
            SetLightColor li122,Green,2
        End If
    Else
        ' Other target wasn't lit. Light this target, play sound
        bDemonTargets(t) = 1
    End If
    SetDemonTargetLights
End Sub












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

Sub SetScoopLights
End Sub

Sub SetDemonTargetLights
    Dim i
    For i = 0 to 1 : SetLightColor DemonTargetLights(i),green,bDemonTargets(i) : Next
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

    bComboEBAwarded = False

    ' initialise Game variables
    Game_Init()

    DMDBlank

    tmrGame.Interval = 100
    tmrGame.Enabled = 1

    vpmtimer.addtimer 1500, "FirstBall '"
End Sub

Sub FirstBall
    ' reset the table for a new ball
    ResetForNewPlayerBall()
    ' create a new ball in the shooters lane
    CreateNewBall()
End Sub

' (Re-)Initialise the Table for a new ball (either a new ball after the player has
' lost one or we have moved onto the next player (if multiple are playing))

Sub ResetForNewPlayerBall()
    ' make sure the correct display is upto date
    AddScore 0

    ' set the current players bonus multiplier back down to 1X
    SetBonusMultiplier 1
    BonusPoints(CurrentPlayer) = 0

    'Reset any table specific
    ResetNewBallVariables

    'This is a new ball, so activate the ballsaver
    bBallSaverReady = True

    'This table doesn't use a skill shot
    bSkillShotReady = False

    FreezeAllGameTimers ' First score will thaw them

    if <onfirstball> Then
        
    Else 
        PlayerState(CurrentPlayer).Restore
        PlayerMode = 0
        Mode(CurrentPlayer).ResetForNewBall
    End If
    SetPlayfieldLights
    PlayModeSong
End Sub

Sub ResetNewBallVariables() 'reset variables for a new ball or player
    dim i
    DemonLockLit=0
    DemonTargetsCompleted=0
    bDemonTargets(0)=0 : bDemonTargets(1)=0
    DemonMBCompleted=0

    'turn on or off the needed lights before a new ball is released
    ResetPictoPops
    ResetDropTargets

    bBallReleasing = False
    
    bPlayfieldValidated = False
    
End Sub


' Create a new ball on the Playfield

Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass

    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' kick it out..
    'PlaySoundAt SoundFXDOF("fx_Ballrel", 123, DOFPulse, DOFContactors), BallRelease
    RandomSoundBallRelease BallRelease
    DOF 123, DOFPulse
    BallRelease.Kick 90, 4

' if there is 2 or more balls then set the multiball flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If (BallsOnPlayfield-RealBallsInLock = 2) And LastSwitchHit = "OutlaneSW" And (bBallSaverActive = True Or bLoLLit = True) Then
        ' Preemptive ball save
        bAutoPlunger = True
        If bBallSaverActive = False Then bLoLLit = False: SetOutlaneLights
    ElseIf (BallsOnPlayfield-RealBallsInLock > 1 Or bMadnessMB > 0) Then
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
    TimerFlags(tmrBallSaverGrace) = 0
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightShootAgain.BlinkInterval = 160
    LightShootAgain.BlinkPattern = "10"
    SetShootAgainLight
End Sub

Sub DisableBallSaver
    TimerFlags(tmrBallSave) = 0
    TimerFlags(tmrBallSaveSpeedUp) = 0
    TimerFlags(tmrBallSaverGrace) = 0
    BallSaveTimer
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaveTimer
    ' clear the flag 
    SetGameTimer tmrBallSaverGrace, DMDStd(kDMDStd_BallSaveExtend)*10  ' 3 second grace for Ball Saver
    SetShootAgainLight
   
End Sub

Sub BallSaverGraceTimer
    bBallSaverActive = False
    SetShootAgainLight
End Sub

Sub BallSaverSpeedUpTimer
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
	
    StopGameTimers
    
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

        DMDShootAgainScene

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
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
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
End Sub' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
End Sub
Sub AddMultiballFast(nballs)
	if CreateMultiballTimer.Enabled = False then 
		CreateMultiballTimer.Interval = 100		' shortcut the first time through 
	End If 
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
End Sub

' Check for key presses specific to this game.
' If PlayerMode is < 0, in a 'Select' state, so use flippers to toggle
Function CheckLocalKeydown(ByVal keycode)
    CheckLocalKeydown = False
    if (keycode = LeftFlipperKey or keycode = RightFlipperKey) Then
        If PlayerMode < 0 Then CheckLocalKeydown = True
        if PlayerMode = -1 Then 
            ChooseSong(keycode)
        End If
        ' Check for both flippers pushed
        If ((LFPress=1 and keycode = RightFlipperKey) Or (RFPress=1 And keycode = LeftFlipperKey)) Then
            'TODO: If there's anywhere where both flips can skip scenes, add the logic here
        End If    
    End If
End Function

'***********************
' Song selection
'***********************
Sub ChooseSong(ByVal keycode)
    If keycode = LeftFlipperKey or keycode = RightFlipperKey Then
        If LFPress=1 And RFPress=1 Then ' Both flippers pressed - start song
         '  TODO Start song
        Else
            If keycode = LeftFlipperKey Then
                PlaySoundVol "gotfx-choosebattle-left",VolDef
                CurrentSongChoice = CurrentSongChoice - 1
                if CurrentSongChoice < 0 Then CurrentSongChoice = TotalSongChoices - 1
                
            Else
                PlaySoundVol "gotfx-choosebattle-right",VolDef
                CurrentSongChoice = CurrentSongChoice + 1
                if CurrentSongChoice >= TotalSongChoices Then CurrentSongChoice = 0
            End If
            UpdateChooseBattle
            ResetBallSearch
        End If
    End If
End Sub

Dim SongChoices(8)   ' Max possible number of choices
Dim TotalSongChoices,CurrentSongChoice
Sub StartChooseSong
    Dim i,j,tmrval,done

    PlayerMode = -1

    TurnOffPlayfieldLights
    
    DMDCreateChooseSongScene

    TotalSongChoices = 0
    For i = 1 to 8
        If Mode(CurrentPlayer).CanPlaySong(i) Then SongChoices(TotalSongChoices) = i : TotalSongChoices = TotalSongChoices + 1
    Next
    CurrentSongChoice = RndNbr(TotalSongChoices)
    UpdateChooseSong ' Update the scene with the currently selected songname and progress, if any

    PlaySong "kiss-songinst"&SongChoices(CurrentSongChoice)
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
    Mode(CurrentPlayer).CheckModeHit LeftOrbit
End Sub

Sub sw24_Hit 'Right Orbit
    If Tilted Then Exit Sub
    Mode(CurrentPlayer).CheckModeHit RightOrbit
End Sub

Sub sw34_Hit 'Star Entrance trigger
    If Tilted Then Exit Sub
End Sub

Sub sw49_Hit 'Demon entrance trigger
    If Tilted Then Exit Sub
    Mode(CurrentPlayer).CheckModeHit DemonShot
End Sub

Sub sw56_Hit 'Right Ramp Exit
    If Tilted Then Exit Sub
    Mode(CurrentPlayer).CheckModeHit RightRamp
End Sub

Sub sw64_Hit 'Left Ramp Exit
    If Tilted Then Exit Sub
    Mode(CurrentPlayer).CheckModeHit CenterRamp
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
    DoDemonTargetHit(0)
End Sub

Sub sw45_Hit 'Right lock target
    If Tilted Then Exit Sub
    DoDemonTargetHit(1)
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



Sub tmrGame_Timer
    Dim i
    GameTimeStamp = GameTimeStamp + 1
    If Timer() < 180 And bMadnessMB = 0 And bGameInPlay And bBallInPlungerLane = False And (PlayerMode=0 Or PlayerMode=1) And bMultiBallMode=False Then StartMidnightMadnessMBIntro
    if bGameTimersEnabled = False Then Exit Sub
    bGameTimersEnabled = False
    For i = 1 to MaxTimers
        If ((TimerFlags(i) AND 2) = 2 Or (TimerFlags(i) And 4) = 4) And i > 5 Then TimerTimestamp(i) = TimerTimestamp(i) + 1 ' "Frozen" timer - increase its expiry by 1 step (Timers 1 - 5 can't be frozen.)
        If (TimerFlags(i) AND 1) = 1 Then 
            bGameTimersEnabled = True
            If TimerTimestamp(i) <= GameTimeStamp Then
                TimerFlags(i) = TimerFlags(i) AND 254   ' Set bit0 to 0: Disable timer
                Execute(TimerSubroutine(i))
            End If
        End if
    Next
End Sub

Sub SetGameTimer(tmr,val)
    TimerTimestamp(tmr) = GameTimeStamp + val
    TimerFlags(tmr) = TimerFlags(tmr) or 1
    bGameTimersEnabled = True
End Sub

Sub StopGameTimers
    Dim i
    For i = 1 to MaxTimers:TimerFlags(i) = TimerFlags(i) And 254: Next
End Sub

' Freeze all timers, including ball saver timers
' This should only be called when no balls are in play - i.e. the only "live" ball is held somewhere
Dim bAllGameTimersFrozen
Sub FreezeAllGameTimers
    Dim i
    For i = 1 to MaxTimers
        TimerFlags(i) = TimerFlags(i) Or 2
    Next
    bAllGameTimersFrozen = True
End Sub

Sub ThawAllGameTimers
    Dim i
    If bAllGameTimersFrozen = False Then Exit Sub
    For i = 1 to MaxTimers
        TimerFlags(i) = TimerFlags(i) And 253
    Next
    ' If mode timers are running, set up a mode pause timer that will pause them 2 seconds from now
    ' This timer is reset every time this sub is called (i.e. every time there's a scoring event)
    If (TimerFlags(tmrBattleMode1) And 1) > 0 Or (TimerFlags(tmrBattleMode2) And 1) > 0 Then SetGameTimer tmrModePause,20
End Sub



'*********************
' HurryUp Support
'*********************

' Called every 200ms by the GameTimer to update the HurryUp value
Sub HurryUpTimer
    Dim lbl
    If bHurryUpActive Then HurryUpCounter = HurryUpCounter + 1
    If bTGHurryUpActive Then TGHurryUpCounter = TGHurryUpCounter + 1
    If bUseFlexDMD Then
        If bHurryUpActive And Not IsEmpty(HurryUpScene) Then
            If Not (HurryUpScene is Nothing) Then
                Set lbl = HurryUpScene.GetLabel("HurryUp")
                If Not lbl is Nothing Then
                    FlexDMD.LockRenderThread
                    lbl.Text = FormatScore(HurryUpValue)
                    FlexDMD.UnlockRenderThread
                End If
            End if
        End If
        If bTGHurryUpActive And Not IsEmpty(TGHurryUpScene) Then
            If Not (TGHurryUpScene is Nothing) Then
                Set lbl = TGHurryUpScene.GetLabel("TGHurryUp")
                If Not lbl is Nothing Then
                    FlexDMD.LockRenderThread
                    lbl.Text = FormatScore(TGHurryUpValue)
                    FlexDMD.UnlockRenderThread
                End If
            End if
        End If
    End if
    if bHurryUpActive And HurryUpCounter > HurryUpGrace Then 
        HurryUpValue = HurryUpValue - HurryUpChange
        If HurryUpValue <= 0 Then HurryUpValue = 0 : EndHurryUp
    End If
    if bTGHurryUpActive And TGHurryUpCounter > TGHurryUpGrace Then 
        TGHurryUpValue = TGHurryUpValue - TGHurryUpChange
        If TGHurryUpValue <= 0 Then TGHurryUpValue = 0 : EndTGHurryUp
    End If
    If bTGHurryUpActive or bHurryUpActive Then SetGameTimer tmrHurryUp,2
End Sub

' Start a HurryUp
'  value: Starting value of HurryUp
'  scene: A FlexDMD scene containing a Label named "HurryUp". The text of the label will be
'         updated every HurryUp period (200ms)
'  grace: Grace period in 200ms ticks. Value will start declining after this many ticks have elapsed
'
' HurryUpChange value calculated by watching change value of numerous Hurry Ups on real GoT tables. Ratio of change to original value was always the same
' The real GoT table sometimes introduces variability to the change value (e.g  alternating between +10K and -10K from base value) but we're not
' going to bother
Sub StartHurryUp(value,scene,grace)
    if bHurryUpActive Then
        debug.print "HurryUp already active! Can't have two!"
        Exit Sub
    End If
    If bUseFlexDMD Then
        Set HurryUpScene = scene
    End If
    HurryUpGrace = grace
    HurryUpValue = value
    HurryUpCounter = 0
    HurryUpChange = Int(HurryUpValue / 1033.32) * 10
    bHurryUpActive = True
    SetGameTimer tmrHurryUp,2
End Sub

' Called when the HurryUp runs down. Ends battle mode if running
Sub EndHurryUp
    StopHurryUp
    If PlayerMode = 1 Then
        If HouseBattle2 > 0 Then House(CurrentPlayer).BattleState(HouseBattle2).BEndHurryUp
        If HouseBattle1 > 0 Then House(CurrentPlayer).BattleState(HouseBattle1).BEndHurryUp
    End if
    If bHotkMBActive Then House(CurrentPlayer).HotkHurryUpExpired
End Sub

' Called when the HurryUp has been scored. Ends the HurryUp but doesn't end the battle (if applicable)
Sub StopHurryUp
    Dim lbl
    bHurryUpActive = False
    if bTGHurryUpActive = False Then TimerFlags(tmrHurryUp) = TimerFlags(tmrHurryUp) And 254
    If PlayerMode = 2 Then
        PlayerMode = 0
        House(CurrentPlayer).SetUPFState False
        StopSound "gotfx-long-wind-blowing"
        tmrWiCLightning.Enabled = False
        GiIntensity 1
        SetPlayfieldLights
        SetTopGates
        PlayModeSong
        If bBWMultiballActive Then DMDCreateBWMBScoreScene : Else DMDResetScoreScene
        DMDFlush
        AddScore 0
    End If
    if bUseFlexDMD Then
        If Not IsEmpty(HurryUpScene) Then
            If Not (HurryUpScene is Nothing) Then 
                Set lbl = HurryUpScene.GetLabel("HurryUp")
                If Not lbl is Nothing Then
                    FlexDMD.LockRenderThread
                    lbl.Visible = False
                    FlexDMD.UnlockRenderThread
                End if
            End If
        End if
    End if
End Sub


' Do the combo shot scene and award points
' Scene is 3 lines, first line is city, second line is score, third is "COMBO"
' first and third lines are 6x8, middle is 9x12
Sub DoComboShotAward(city)
    Dim award,cityname
    ' TODO: If the final shot was on a currently-doubled shot, presumably score is doubled too
    award = 1000000*ComboShotPattern(city)(0)
    cityname = ComboShotPattern(city)(1)
    AddScore award
    'TODO create DMD scene for Combo award
End Sub

Dim ChampionNames
ChampionNames = Array("CITY COMBO","HEAVENS ON FIRE","KISS ARMY","ROCK CITY")

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
        Case 0:img = "kiss-sternlogo":format=9:scrolltime=3:y=73:delay=3000
        Case 1:line1 = "PRESENTS":format=10:font="udmd-f7by10.fnt":delay=2000
        Case 2:img = "kiss-intro":format=5:delay=17200
        Case 3 ' last score
            format=7:font="udmd-f11by18.fnt":line1=FormatScore(Score(1)):skipifnoflex=False  ' Last score
            If Score(1) > 999999999 Then font="udmd-f7by10.fnt"
            If DMDStd(kDMDStd_FreePlay) Then line2 = "FREE PLAY" Else Line2 = "CREDITS "&Credits
        Case 5 ' Credits
            format=1:font="udmd-f7by10.fnt":skipifnoflex=False
            If DMDStd(kDMDStd_FreePlay) Then 
                line1 = "FREE PLAY" 
            ElseIf Credits > 0 Then 
                Line1 = "CREDITS "&Credits
            Else 
                Line1 = "INSERT COINS"
            End if
        Case 6:format=1:font="udmd-f7by10.fnt":line1 = "REPLAY AT" & vbLf & FormatScore(ReplayScore)
        Case 9:format=8:line1="GRAND CHAMPION":line2=HighScoreName(0):line3=FormatScore(HighScore(0)):font="udmd-f6by8.fnt":skipifnoflex=False
        Case 10:format=8:line1="HIGH SCORE #1":line2=HighScoreName(1):line3=FormatScore(HighScore(1)):font="udmd-f6by8.fnt":skipifnoflex=False
        Case 11:format=8:line1="HIGH SCORE #2":line2=HighScoreName(2):line3=FormatScore(HighScore(2)):font="udmd-f6by8.fnt":skipifnoflex=False
        Case 12:format=8:line1="HIGH SCORE #3":line2=HighScoreName(3):line3=FormatScore(HighScore(3)):font="udmd-f6by8.fnt":skipifnoflex=False
        Case 13:format=8:line1="HIGH SCORE #4":line2=HighScoreName(4):line3=FormatScore(HighScore(4)):font="udmd-f6by8.fnt":skipifnoflex=False
        Case 14,15,16,17
            If HighScore(i-9) = 0 Then delay = 5
            format=3:skipifnoflex=False
            line1 = ChampionNames(i-14)&" CHAMPION"
            line2 = HighScoreName(i-9):line3 = FormatScore(HighScore(i-9))
        Case Else: delay=10
    End Select
    If i = 17 Then tmrAttractModeScene.UserValue = 0 Else tmrAttractModeScene.UserValue = i + 1
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

Dim ChooseSongScene
' ChooseSong scene has player scores in each of up to 4 corners, clockwise starting at top left
' Song name is below top line of scores, in 5x7 font, except Rock&Roll which is 3x7 (any others?)
' Under the song name is a 1 pixel line that is dark grey, but with a white portion equal to the mode progress (nothing at start)
Sub DMDCreateChooseSongScene
    Dim ScoreFont,tinyfont,x,y,i,j,align
    Set ChooseSongScene = FlexDMD.NewGroup("choosesong")
    Set ScoreFont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0) 
    Set tinyfont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt",RGB(128,128,128),RGB(128,128,128),0)
    ' Score text
    For i = 1 to PlayersPlayingGame
        If i = CurrentPlayer Then
            j=""
            ScoreScene.AddActor FlexDMD.NewLabel("Score", ScoreFont, FormatScore(Score(i)))
        Else
            j=i
            ScoreScene.AddActor FlexDMD.NewLabel("Score"&i, tinyfont, FormatScore(Score(i)))
        End If
        x = 8 : y = 0
        If CurrentPlayer > 2 Then y = 20
        If (i MOD 2) = 0 Then x = 127 : align=FlexDMD_Align_TopRight : Else align = FlexDMD_Align_TopLeft
        
        ScoreScene.GetLabel("Score"&j).SetAlignedPosition x,y, align
    Next
    ' Ball, credits
    ScoreScene.AddActor FlexDMD.NewLabel("Ball", tinyfont, "BALL 1")
    ScoreScene.AddActor FlexDMD.NewLabel("Credit", tinyfont, "CREDITS 0")
    If DMDStd(kDMDStd_FreePlay) Then ScoreScene.GetLabel("Credit").Text = "Free Play"
    ' Align them
    ScoreScene.GetLabel("Ball").SetAlignedPosition 32,28, FlexDMD_Align_Center
    ScoreScene.GetLabel("Credit").SetAlignedPosition 96,28, FlexDMD_Align_Center
    ' Divider
    ScoreScene.AddActor FlexDMD.NewFrame("HSeparator")
    ScoreScene.GetFrame("HSeparator").Thickness = 1
    ScoreScene.GetFrame("HSeparator").SetBounds 0, 24, 128,1
'End Sub

' Create a "Hit" scene which plays every time a qualifying or battle target is hit
'  vid   - the name of the video for the first part of the scene
'  sound - the sound to play with the video
'  delay - How long to wait before cutting to the second part of the scene, in seconds (float)
'  line1-3 - Up to 3 lines of text
'  combo - If 0, text is full width. Otherwise, Combo multiplier is on the right side
'  format - The following formats (layouts) are supported
'       1 - video followed by dimmed image with 2 lines of text over image. First line is 3x5 font, second is 7x10?
'       2 - video with transparent background plays over top of single line of text

Sub DMDPlayHitScene(vid,sound,delay,line1,line2,line3,combo,format)
    Dim scene,scenevid,font1,font2,font3,x,y1,y2,y3,combotxt,pri,l3vis
    Set scenevid = Nothing
    If bUseFlexDMD Then
        If format = 2 Then  ' Create empty scene. Video is added after text
            Set scene = FlexDMD.NewGroup("hitscene")
        Else
            Set scene = NewSceneWithVideo("hitscene",vid)
            Set scenevid = scene.GetVideo("hitscenevid")
            If scenevid is Nothing Then Set scenevid = scene.getImage("hitsceneimg")
        End If
        y1 = 4: y2 = 15: y3 = 26
        x = 64
        l3vis = False
        ' TODO: Adjust fonts once we have them figured out
        Select Case format
            Case 0,6
                Set font1 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Set font2 = FlexDMD.NewFont("skinny7x12.fnt", vbWhite, vbWhite, 0)
                Set font3 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
                l3vis = True
            Case 1
                if Len(line1) > 12 Then
                    Set font1 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Else 
                    Set font1 = FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0)
                End if
                Set font2 = FlexDMD.NewFont("skinny7x12.fnt", vbWhite, vbWhite, 0)
                set font3 = FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", vbWhite, vbWhite, 0)
            Case 2
                Set font1 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Set font2 = font1
                Set font3 = font1
                l3vis = True
            Case 3,5,7,8,9
                Set font1 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Set font2 = FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbWhite, 0)
                Set font3 = font1
                If format = 8 or format = 9 then y1=10 : y2=23
            Case 4
                Set font1 = FlexDMD.NewFont("skinny10x12.fnt", vbWhite, vbWhite, 0)
                Set font2 = font1
                Set font3 = font1
                y1 = 8: y2 = 23
            Case 10
                Set font1 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
                Set font2 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Set font3 = font1
                x = 50
                l3vis = True
            Case 11
                Set font2 = FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbWhite, 0)
                Set font1 = font2
                Set font3 = font2
                x = 50
                l3vis = False
            Case 12
                Set font1 = FlexDMD.NewFont("udmd-f3by7.fnt", vbWhite, vbWhite, 0)
                Set font3 = FlexDMD.NewFont("udmd-f6by8.fnt", vbWhite, vbWhite, 0)
                Set font2 = font1
                l3vis = True
        End Select

        scene.AddActor FlexDMD.NewGroup("hitscenetext")
        'TODO If format=1, then AddActor for faded image behind the text
        With scene.GetGroup("hitscenetext")
            .AddActor FlexDMD.NewLabel("line1",font1,line1)
            .AddActor FlexDMD.NewLabel("line2",font2,line2)
            .AddActor FlexDMD.NewLabel("line3",font3,line3)
            .Visible = False
        End With
        
        With scene.GetGroup("hitscenetext")
            .GetLabel("line1").SetAlignedPosition x,y1,FlexDMD_Align_Center
            .GetLabel("line2").SetAlignedPosition x,y2,FlexDMD_Align_Center
            .GetLabel("line3").SetAlignedPosition x,y3,FlexDMD_Align_Center
        End With

        ' If line2 is a score, flash it
        If format = 1 Then BlinkActor scene.GetGroup("hitscenetext").GetLabel("line2"),100,10        

        scene.GetGroup("hitscenetext").GetLabel("line3").Visible = l3vis

        ' After delay, disable video/image and enable text
        If format <> 2 And delay > 0 And Not (scenevid Is Nothing) Then
            If format <> 9 then DelayActor scenevid,delay,False
            If vid="got-cmbsuperjp" then DelayActor scene.GetGroup("hitscenefire"),delay,False
            DelayActor scene.GetGroup("hitscenetext"),delay,True
        Else
            scene.GetGroup("hitscenetext").Visible = True
            delay = 0
        End If

        If format = 2 Then
            scene.AddActor NewSceneFromImageSequence("hitscenevid",vid,50,25,0,0) ' TODO: update the '50' which is how many images are in the vid, for Lick It Up and Hotter Than Hell flames
        End if

        DMDEnqueueScene scene,pri,delay*1000+1000,delay*1000+2000,3000,sound
    Else
        DisplayDMDText line1,line2,2000
        PlaySoundVol sound,VolDef
    End If

End Sub

Sub GameDoDMDHighScoreScene(scoremask)
    Dim scene,i,scenenum,actor,ondly,offdly
    Set scene = FlexDMD.NewGroup("hs")
    ondly = 0 : offdly = 4
    For i = 0 to 13 ' number of highscore slots
        if (scoremask And (2^i)) > 0 Then ' This is one of the player's high scores
            scenenum = scenenum+1
            scene.AddActor FlexDMD.NewLabel("scname"&scenenum, FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0), ChampionNames(i) )
            Set actor = scene.GetLabel("scname"&scenenum)
            With actor
                .SetAlignedPosition 64,3,FlexDMD_Align_Center
                .Visible = 0
            End With
            FlashActor actor,ondly,offdly
            scene.AddActor FlexDMD.NewLabel("scscore"&scenenum, FlexDMD.NewFont("udmd-f7by10.fnt",vbWhite,vbWhite,0),FormatScore(HighScore(i)))
            Set actor = scene.GetLabel("scscore"&scenenum)
            With actor
                .SetAlignedPosition 64,16,FlexDMD_Align_Center
                .Visible=0
            End With
            FlashActor actor,ondly,offdly
            ondly = ondly+offdly
            offdly = 2
        End If
    Next
    DMDFlush
    DMDDisplayScene scene
    vpmTimer.AddTimer 2000+2000*scenenum-100,"StopSound ""kiss-track-hstd"" : hsbModeActive=0 : DMDBlank : tmrDMDUpdate.Enabled = True : EndOfBallComplete() '"
End Sub

'***************
' INSTANT INFO
'***************
' Format tells us how to format each scene (copied from Attract mode)
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
' 11: 3 lines of text (small, medium, medium)
Sub InstantInfo
    Dim scene,format,font,skipifnoflex,y,img
    Dim line1,line2,line3,font1
    InstantInfoTimer.Enabled = False
    If Tilted Then Exit Sub
    font = "udmd-f7by10.fnt" ' Most common font
    Select Case InfoPage
        Case 0: format=1:line1="INSTANT INFO"
        Case 1 ' current ball/credits/player
            If bGameInPlay = False Then InfoPage = 29 : InstantInfo : Exit Sub 
            format=8:font="udmd-f6by8.fnt" 
            line1="BALL "&BallsPerGame-BallsRemaining(CurrentPlayer)+1
            If DMDStd(kDMDStd_FreePlay) Then 
                line2 = "FREE PLAY" 
            Else
                Line2 = "CREDITS "&Credits
            End if
            line3 = "PLAYER "&CurrentPlayer&" IS UP"
        Case 2,3,4,5 ' current scores
            If PlayersPlayingGame < InfoPage-1 Then InfoPage = 6 : InstantInfo : Exit Sub 
            format=2:line1="PLAYER "&InfoPage-1:line2=FormatScore(Score(InfoPage-1)):skipifnoflex=False
        Case 6 ' current player's gold
            format=1:line1=CurrentGold&" GOLD"
        Case 7 ' LoL status
            format=8:font="udmd-f3by7.fnt"
            line1="LORD   OF   LIGHT"
            line3=""
            If bLoLLit Then
                line2="LIT"
            ElseIf bLoLUsed Then
                line2="USED"
            Else
                line2="SHOOT   LEFT   TARGETS":line3="FOR   OUTLANE   BALL   SAVE"
            End If
        Case 8 ' Blackwater status
            format=11:font="FlexDMD.Resources.udmd-f4by5.fnt"
            line1="BLACKWATER MULTIBALL"
            If bLockIsLit Then line2="LOCK IS LIT" Else line2="SHOOT RIGHT TARGETS"
            If BallsInLock = 1 Then line3="1 BALL IN LOCK" Else line3 = BallsInLock & " BALLS IN LOCK"
        Case 9
            Dim i
            format=11:font="udmd-f6by8.fnt"
            line1 = "WALL MULTIBALL"
            If bWallMBReady Then
                line2 = "SHOOT RIGHT ORBIT"
            Else
                If WallMBCompleted = 0 Then i = 6-WallMBLevel Else i = 11-WallMBLevel
                line2 = i & " ADVANCEMENTS"
            End If
            line3 = "TO START MULTIBALL"
        Case 10
            If House(CurrentPlayer).WiCs >= 4 Then InfoPage=12:InstantInfo:Exit Sub
            format=3
            line1 = "WINTER IS COMING"
            line2 = 3-House(CurrentPlayer).WiCShots : If SelectedHouse = Greyjoy then line2 = line2 + 1
            If line2 = 1 Then line3 = "SHOT TO START" Else line3 = "SHOTS TO START"
        Case 11
            line1 = "CASTLE MULTIBALL"
            line2 = "CURRENT LEVEL"
            line3 = House(CurrentPlayer).UPFLevel - 1
            format=8:font="udmd-f6by8.fnt"
        Case 12
            If CompletedHouses >= 3 And EBisLit = 0 Then InfoPage=13:InstantInfo:Exit Sub
            format=2:font="udmd-f3by7.fnt"
            If EBisLit > 0 Then 
                line1 = "EXTRA   BALL   IS   LIT"
                line2 = "SHOOT   RIGHT   ORBIT"
            Else
                line1 = 3-CompletedHouses & "  HOUSE  COMPLETIONS  NEEDED"
                line2 = "FOR   EXTRA   BALL"
            End If
        Case 13
            format=8:font="udmd-f3by7.fnt"
            line1 = "SELECTED  HOUSE:  "&HouseToUCString(SelectedHouse)
            line2 = "ACTION   BUTTON"
            line3 = HouseAbility(House(CurrentPlayer).ActionAbility)
        Case 14,15,16,17,18,19,20
            format=2:font="udmd-f6by8.fnt"
            line1 = "HOUSE "&HouseToUCString(InfoPage-13)
            If House(CurrentPlayer).Completed(InfoPage-13) Then
                line2="COMPLETED"
            ElseIf House(CurrentPlayer).Qualified(InfoPage-13) Then 
                line2="IS LIT"
            Else line2 = (3 - House(CurrentPlayer).QualifyCount(InfoPage-13)) & " MORE TO LIGHT"
            End If
        Case 21
            format=12:font="udmd-f6by8.fnt"
            line1="SWORD COLLECTION"
            line2="SWORDS: "&SwordsCollected
            If SwordsCollected*DMDStd(kDMDFet_SwordsUnlockMultiplier) < 3 Then line3="NEXT UNLOCKS "&SwordsCollected*DMDStd(kDMDFet_SwordsUnlockMultiplier)+3&"X TIMES MULTIPLIER" Else line3=""
        Case 22
            format=11:font="udmd-f6by8.fnt"
            line1="SPINNER"
            line2="LEVEL "&Int((SpinnerLevel+2)/3)
            line3=SpinnerValue & " A SPIN"
        Case 23
            format=12:font="udmd-f6by8.fnt"
            line1="TOTAL BONUS"
            line2=FormatScore(BonusPoints(CurrentPlayer))
            line3="CURRENT MULTIPLIER "&BonusMultiplier(CurrentPlayer)&"X"
        'Could add more Info screens here. Real game doesn't have any more though
        Case 24: InfoPage=29:InstantInfo:Exit Sub

        Case 29:format=1:line1 = "REPLAY AT" & vbLf & FormatScore(ReplayScore)
        Case 30,31,32,33,34
            format = 8 : font="udmd-f6by8.fnt" : :skipifnoflex=False
            If InfoPage = 30 Then line1 = "GRAND CHAMPION" Else line1 = "HIGH SCORE #" & InfoPage-30
            line2 = HighScoreName(InfoPage-30) : line3 = FormatScore(HighScore(InfoPage-30))
        Case 35,36,37,38,39,40,41,42,43,44
            If bGameInPlay Then InfoPage=0:InstantInfo:Exit Sub
            format=3:skipifnoflex=False
            line1 = ChampionNames(i-35)&" CHAMPION"
            line2 = HighScoreName(i-30):line3 = HighScore(i-30)
    End Select
    If InfoPage >= 45 Then InfoPage = 0:InstantInfo:Exit Sub
    If bUseFlexDMD=False And skipifnoflex=True Then InfoPage=InfoPage+1:InstantInfo:Exit Sub

    ' Create the scene
    if bUseFlexDMD Then
        If format=4 or Format=5 or Format=6 or Format=9 Then
            Set scene = NewSceneWithVideo("attract"&InfoPage,img)
        Else
            Set scene = FlexDMD.NewGroup("attract"&InfoPage)
        End If

        ' Most of these modes aren't used for InstantInfo but we could probably combine the code with AttractMode to make it DRY
        Select Case format
            Case 1
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                scene.GetLabel("line1").SetAlignedPosition 64,16,FlexDMD_Align_Center
            Case 2
                Set font1 = FlexDMD.NewFont(font,vbWhite,vbWhite,0)
                scene.AddActor FlexDMD.NewLabel("line1",font1,line1)
                scene.AddActor FlexDMD.NewLabel("line2",font1,line2)
                scene.GetLabel("line1").SetAlignedPosition 64,9,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,22,FlexDMD_Align_Center
            Case 3
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line1)
                scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont("skinny10x12.fnt",vbWhite,vbWhite,0),line2)
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
            Case 11
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line1)
                scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line2)
                scene.AddActor FlexDMD.NewLabel("line3",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line3)
                scene.GetLabel("line1").SetAlignedPosition 64,3,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,13,FlexDMD_Align_Center
                scene.GetLabel("line3").SetAlignedPosition 64,25,FlexDMD_Align_Center
            Case 12
                scene.AddActor FlexDMD.NewLabel("line1",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line1)
                scene.AddActor FlexDMD.NewLabel("line2",FlexDMD.NewFont(font,vbWhite,vbWhite,0),line2)
                scene.AddActor FlexDMD.NewLabel("line3",FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0),line3)
                scene.GetLabel("line1").SetAlignedPosition 64,5,FlexDMD_Align_Center
                scene.GetLabel("line2").SetAlignedPosition 64,16,FlexDMD_Align_Center
                scene.GetLabel("line3").SetAlignedPosition 64,27,FlexDMD_Align_Center           
        End Select

        DMDDisplayScene scene
    End If
End Sub

 We support multiple score "scenes", depending on what mode the table is in. Not all modes
' support all fields, so define a SceneMask that decides which fields need to be updated
'  bit   data (Label name)
'   0    Score
'   1    Ball
'   2    Credits
'   4    HurryUp
'
' "scene" is a pre-created scene with all of the proper text labels already created. There MUST be a label
' corresponding with every bit set in the scenemask
Sub DMDSetAlternateScoreScene(scene,mask)
    bAlternateScoreScene = True
    Set ScoreScene = scene
    AlternateScoreSceneMask = mask
    DMDLocalScore
End Sub

' Set Score scene back to default for regular play
Sub DMDResetScoreScene
    bAlternateScoreScene = False
    If Not IsEmpty(DisplayingScene) Then
        If DisplayingScene Is ScoreScene Then DMDClearQueue
    End if
    ScoreScene = Empty
    AlternateScoreSceneMask = 0
    DMDLocalScore
End Sub

Dim ScoreScene,bAlternateScoreScene,AlternateScoreSceneMask,remaining
Sub DMDLocalScore
    Dim ComboFont,ScoreFont,i,j,x,y,font,tinyfont
    If bUseFlexDMD Then
        If IsEmpty(ScoreScene) And bAlternateScoreScene = False Then
            Set ScoreScene = FlexDMD.NewGroup("ScoreScene")
            Set ComboFont = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", vbWhite, vbWhite, 0)
            If Score(CurrentPlayer) < 1000000000 And PlayersPlayingGame < 3 Then font = "FlexDMD.Resources.udmd-f7by13.fnt" Else font = "udmd-f6by8.fnt"
            Set ScoreFont = FlexDMD.NewFont(font, vbWhite, vbWhite, 0) 
            Set tinyfont = FlexDMD.NewFont("tiny3by5.fnt",vbWhite,vbWhite,0)
            ' Score text
            For i = 1 to PlayersPlayingGame
                If i = CurrentPlayer Then
                    j=""
                    ScoreScene.AddActor FlexDMD.NewLabel("Score", ScoreFont, FormatScore(Score(i)))
                Else
                    j=i
                    ScoreScene.AddActor FlexDMD.NewLabel("Score"&i, tinyfont, FormatScore(Score(i)))
                End If
                x = 80 : y = 0
                If CurrentPlayer > 2 Then 
                    y=6
                Elseif i > 2 Then 
                    y = 10
                End If
                If (i MOD 2) = 0 Then 
                    x = 127
                ElseIf (CurrentPlayer MOD 2) = 0 Then
                    x = 46
                End if 
                ScoreScene.GetLabel("Score"&j).SetAlignedPosition x,y, FlexDMD_Align_TopRight
            Next
            ' Ball, credits
            ScoreScene.AddActor FlexDMD.NewLabel("Ball", ComboFont, "BALL 1")
            ScoreScene.AddActor FlexDMD.NewLabel("Credit", ComboFont, "CREDITS 0")
            If DMDStd(kDMDStd_FreePlay) Then ScoreScene.GetLabel("Credit").Text = "Free Play"
            ' Align them
            ScoreScene.GetLabel("Ball").SetAlignedPosition 32,20, FlexDMD_Align_Center
            ScoreScene.GetLabel("Credit").SetAlignedPosition 96,20, FlexDMD_Align_Center
            ' Divider
            ScoreScene.AddActor FlexDMD.NewFrame("HSeparator")
            ScoreScene.GetFrame("HSeparator").Thickness = 1
            ScoreScene.GetFrame("HSeparator").SetBounds 0, 24, 128,1
            ' Combo Multipliers
            For i = 1 to 5
                ScoreScene.AddActor FlexDMD.NewLabel("combo"&i, ComboFont, "0")
            Next
        End If
        FlexDMD.LockRenderThread
        ' Update fields
        If bAlternateScoreScene = False or (AlternateScoreSceneMask And 1) = 1 Then ScoreScene.GetLabel("Score").Text = FormatScore(Score(CurrentPlayer))
        If bAlternateScoreScene = False or (AlternateScoreSceneMask And 2) = 2 Then ScoreScene.GetLabel("Ball").Text = "BALL " & CStr(BallsPerGame - BallsRemaining(CurrentPlayer) + 1)
        If DMDStd(kDMDStd_FreePlay) = False And (bAlternateScoreScene = False or (AlternateScoreSceneMask And 4) = 4) Then 
            With ScoreScene.GetLabel("Credit")
                .Text = "CREDITS " & CStr(Credits)
                .SetAlignedPosition 96,20, FlexDMD_Align_Center
            End With
        End If
        
        ' Update score position
        If bAlternateScoreScene = False Then
            x = 80 : y = 0
            If CurrentPlayer = 2 Or CurrentPlayer = 4 Then x = 127
            If CurrentPlayer > 2 Then y = 10
            ScoreScene.GetLabel("Score").SetAlignedPosition x,y, FlexDMD_Align_TopRight
        End If
        

        ' Update special battlemode fields
        If bAlternateScoreScene Then
            If (AlternateScoreSceneMask And 16) = 16 Then ScoreScene.GetLabel("HurryUp").Text = FormatScore(HurryUpValue)
           
        End If
        FlexDMD.UnlockRenderThread
        Set DefaultScene = ScoreScene
    Else
        DisplayDMDText "",FormatScore(Score(CurrentPlayer)),0
    End If
End Sub
