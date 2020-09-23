VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTable 
   Appearance      =   0  'Flat
   BackColor       =   &H00008000&
   ClientHeight    =   7650
   ClientLeft      =   1170
   ClientTop       =   1845
   ClientWidth     =   10875
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   1
   Icon            =   "frmTable.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDisableSound 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Disable Sound"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8310
      TabIndex        =   0
      Top             =   7980
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   360
      Top             =   870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer timerAnimate 
      Interval        =   10
      Left            =   3000
      Top             =   300
   End
   Begin VB.Timer timerDemo 
      Enabled         =   0   'False
      Left            =   2250
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fast Win= when all cards in table are with face up the game is won."
      Height          =   255
      Index           =   1
      Left            =   1710
      TabIndex        =   2
      Top             =   7260
      Width           =   7785
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTable.frx":000C
      Height          =   495
      Index           =   0
      Left            =   1710
      TabIndex        =   1
      Top             =   6780
      Width           =   7785
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowStats 
         Caption         =   "S&how Stats"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoplay 
         Caption         =   "&AutoPlay"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuCpuMove 
         Caption         =   "&Cpu play solitaire"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFastWin 
         Caption         =   "&Fast Win"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSound 
         Caption         =   "&Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "&Setup"
      Begin VB.Menu mnuClickMove 
         Caption         =   "Click Auto Move"
      End
      Begin VB.Menu mnuAutoFaceUp 
         Caption         =   "Auto Face Up"
      End
      Begin VB.Menu mnuAutoSpeed 
         Caption         =   "Auto Speed"
      End
      Begin VB.Menu mnuAnimateSpeed 
         Caption         =   "Animate Speed"
      End
      Begin VB.Menu mnuBackground 
         Caption         =   "Background"
         Begin VB.Menu mnuBackgroundColor 
            Caption         =   "Color"
         End
      End
      Begin VB.Menu mnuBackDesign 
         Caption         =   "Card Back Design"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Solitaire"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
'**Solitaire 1.0
'**Copyright by Paulo Matoso
'**E-Mail: paulomt1@clix.pt
'**
'**
'**Last Modification ---> 19/10/2002
'**********************************************************************************



Option Explicit
Dim mGameStarted As Integer
Dim mCpuIsWork As Integer
Dim MoveDealAuto As Integer
Dim MoveFaceUpAuto As Integer
Dim MoveToFoundationAuto As Integer
Dim MultLayoutAuto As Integer
Dim mUserMenuAction As Integer
Dim mDoubleClick As Integer
Dim Paused As Integer

Dim mClickMove As Integer

Const MouseBtnLeft = 1
Const MouseBtnRight = 2
Const mSaveUndo = True
Const mUpdateBackground = False

Const mTablePile = 1
Const mFundationPile = 2
Const mDiscardPile = 3
Const mDealPile = 4

Const mClearForm = True
Const mDontClearForm = False
Const mCardWith = 72
Const TPM_CENTERALIGN = &H4&

Sub ResetAutoPlayVars()
    mCpuIsWork = False
    gMoveMultipleAuto = False
    MultLayoutAuto = False
    MoveDealAuto = False
End Sub

Sub SubSetAutoMoves()
    If gMoveMultipleAuto Then
        MoveFaceUpAuto = True
        MultLayoutAuto = True
        MoveToFoundationAuto = True
        MoveDealAuto = True
    Else
        ResetAutoPlayVars
    End If
End Sub


Sub Form_DblClick()
'Start the Cpu autoplay
    mDoubleClick = True
End Sub

Sub Form_Load()
gSoundDeal = App.Path + "\DrawUp.wav"
gSoundStack = App.Path + "\Stack.wav"
gSoundShuffle = App.Path + "\Shuffle.wav"
gSoundApplause = App.Path + "\Applause.wav"
gSoundTurn = App.Path + "\Turn.wav"
gSoundDrawUp = App.Path + "\DrawUp.wav"

    mGameStarted = False
    gFirstTime = True
    CPUPlay = False
    mDoubleClick = False
    mCpuIsWork = False
    gMoveMultipleAuto = False
    MoveFaceUpAuto = True
    mnuAutoFaceUp.Checked = True
    MultLayoutAuto = False
    
    MoveToFoundationAuto = True
    mnuAutoplay.Checked = True
    
    mUserMenuAction = False
    SubInit
    mClickMove = True
    mnuClickMove.Checked = True
    timerDemo.Enabled = True
    FuncSetAnimationSpeed 250 'set the speed for animate cards
End Sub

Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lCords As tCordinates

    mDoubleClick = False 'stop Cpu Auto Playing
    If CheckForClickInCard() Or CPUPlay Or mCpuIsWork Then
        Exit Sub
    End If
    Select Case Button
        Case MouseBtnLeft
            lCords.tX = CInt(x)
            lCords.tY = CInt(y)
            SubCheckIfCardIsMoveable lCords
        Case MouseBtnRight
            frmTable.Enabled = False
            PopupMenu mnuFile, TPM_CENTERALIGN
            frmTable.Enabled = True
    End Select
End Sub

Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lMouseCords As tCordinates


    If CheckForClickInCard() Then
    ' the player have clicked on a card
        If Button = 0 Then
            subRefreshPile
        Else
        'the player as drag a card(s)
            lMouseCords.tX = CInt(x)
            lMouseCords.tY = CInt(y)
            SubRefreshCard lMouseCords 'refresh form
        End If
    End If
End Sub

Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lMouseCords As tCordinates
Dim lTolerance As Single
Dim lPile As PILESINFORMATION
Dim lReturn As Integer

    
    If CPUPlay Or mCpuIsWork Then
    'player have stoped the autoplay mode
        ResetAutoPlayVars
        mUserMenuAction = True
        Exit Sub
    End If
    
 If mDoubleClick Then mnuCpuMove_Click
    
    mUserMenuAction = False
    Select Case Button
        Case MouseBtnLeft
            
            SubGetPile gPile
            
            If gPile.tnPileNumber = 0 Then
                Exit Sub 'player have clicked in green table, exit sub
            End If
            
            If gPile.tCardIndex = 0 Then 'if no cards in fundation
                If gPile.tnPileNumber = mDealPile Then
                    SubRefillDealPile gPile 'if the dealpile draw all cards from discardpile
                Else
                    gMoveMultipleAuto = False
                    timerDemo_Timer
                End If
                Exit Sub
            End If
        
            If Not CheckForClickInCard() Then
                If Not gPile.tnPileInfo.tPileCardBlocked Then
                'the card is blocked or not moveable
                    lReturn = SubSetFaceUp(gPile) 'check if is the first card in pile
                End If
                Exit Sub
            End If
            
            lMouseCords.tX = CInt(x)
            lMouseCords.tY = CInt(y)
            SubGetTolerance lMouseCords, gPileCardsPicts, lTolerance 'get where the
                        'card as droped
                            
            
            If lTolerance > gMidleOfCardHeight Then 'if the card pass across midle of
                                            'fundation or other card when the player drag
                If gPileCardsPicts.tnPileNumber = mTablePile And gPileCardsPicts.tCardIndex <> 0 Then
                    SubGetCardInfoFromTablePile gPileCardsPicts.tnPileNumberInPile, gPileCardsPicts
                End If
                'check if the card(s) is a valid droped
                If SubTablePileFindPlaceForCard(gPile, gPileCardsPicts, CheckPlayerCardsDroped()) Then
                    'yes, is a valid fundation or a destination
                    'card and a valid rule for solitaire game
                    SubHumanDrawCard gPileCardsPicts ' Draw the card
                    'actualize array
                    lReturn = SubCardMove(gPile, gPileCardsPicts, CheckPlayerCardsDroped(), mSaveUndo, mUpdateBackground)
                    SubUpdateMobealble gPile
                    SubUpdateMobealble gPileCardsPicts
                Else
                    subRefreshPile ' is a invalid drop, refresh screen
                End If
            Else
                subRefreshPile
                If Not mClickMove Then Exit Sub ' if clicmove go down, and find
                    'if card can go to a valid fundation or pile
                Select Case gPile.tnPileNumber
                Case mTablePile
                    If FuncCpuGetInfoMoveableTablePile(gPile.tnPileNumberInPile, gPile.tCardIndex) Then
                        gCardsInPile = SubHowManyCardsInPile(gPile.tnPileNumberInPile) - gPile.tCardIndex + 1
                        SubGetCardInfoFromTablePile gPile.tnPileNumberInPile, lPile
                        SubFindInFundationPile lPile.tnPileInfo.tPileInfoCard, gPileCardsPicts
                        
                        If SubTablePileFindPlaceForCard(gPile, gPileCardsPicts, gCardsInPile) Then
                            lReturn = FuncCardMove()
                            Exit Sub
                        End If
                        
                        SubFindInTablePile gPile, gPileCardsPicts
                        
                        If SubTablePileFindPlaceForCard(gPile, gPileCardsPicts, gCardsInPile) Then
                            lReturn = FuncCardMove()
                            Exit Sub
                        End If
                        
                        SubFindInDiscardPile gPileCardsPicts
                        If SubTablePileFindPlaceForCard(gPile, gPileCardsPicts, gCardsInPile) Then
                            lReturn = FuncCardMove()
                            Exit Sub
                        End If
                    End If
                    
                    SubGetCardInfoFromTablePile gPile.tnPileNumberInPile, gPile
                    SubFindInFundationPile gPile.tnPileInfo.tPileInfoCard, gPileCardsPicts
                    
                    If SubTablePileFindPlaceForCard(gPile, gPileCardsPicts, 1) Then
                        gCardsInPile = 1
                        lReturn = FuncCardMove()
                        Exit Sub
                    End If
                    
                    SubFindInTablePile gPile, gPileCardsPicts
                    If SubTablePileFindPlaceForCard(gPile, gPileCardsPicts, 1) Then
                        gCardsInPile = 1
                        lReturn = FuncCardMove()
                        Exit Sub
                    End If
                
                Case mFundationPile
                    SubFindInTablePile gPile, gPileCardsPicts
                    If gPileCardsPicts.tnPileNumberInPile <> 0 Then ' if found
                        gCardsInPile = 1
                        lReturn = FuncCardMove()
                    End If
                
                Case mDiscardPile
                    SubGetCardInfoFromDiscardPile gPile.tnPileNumberInPile, gPile
                
                    'first find in fundation pile if exist a valid move
                    SubFindInFundationPile gPile.tnPileInfo.tPileInfoCard, gPileCardsPicts
                    If gPileCardsPicts.tnPileNumberInPile <> 0 Then
                        gCardsInPile = 1
                        lReturn = FuncCardMove()
                        Exit Sub
                    End If
                    
                    'and find in table pile
                    SubFindInTablePile gPile, gPileCardsPicts
                    If gPileCardsPicts.tnPileNumberInPile <> 0 Then
                        gCardsInPile = 1
                        lReturn = FuncCardMove()
                        Exit Sub
                    End If
                
                Case mDealPile
                    SubDealPileFindPlaceForCard
                End Select
            End If
        End Select
End Sub

Sub Form_Paint()
    If Not gLoadPict Then
        SubRefreshPilesBackground mDontClearForm
    End If
End Sub


Sub Form_Resize()
'if With of form is bellow than mCardWith * 7 piles then stop resize
If frmTable.ScaleWidth < 7 * mCardWith Then Exit Sub

    If WindowState = SIZEICONIC Then 'if minimized, stop all activity
        timerDemo.Enabled = False
        timerAnimate.Enabled = False
    Else
        frmTable.Enabled = True
        timerDemo.Enabled = True
        timerAnimate.Enabled = True
        SubResize
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not FuncAreUSure Then Cancel = True

End Sub

Sub Form_Unload(Cancel As Integer)
    End
End Sub

Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub


Sub mnuAnimateSpeed_Click()
    frmScroll.VScroll1.Min = 100
    frmScroll.VScroll1.Max = 1000
    frmScroll.VScroll1.Value = CSng(FuncGetAnimationSpeed())
    frmScroll.VScroll1.LargeChange = 100
    frmScroll.VScroll1.SmallChange = 50
    frmScroll.lblScrollValue.Caption = CStr(frmScroll.VScroll1.Value)
    frmScroll.Caption = "Animate Speed"
    frmScroll.Show 1
    FuncSetAnimationSpeed CSng(frmScroll.VScroll1.Value)
    Unload frmScroll
End Sub


Sub mnuAutoFaceUp_Click()
    mnuAutoFaceUp.Checked = Not mnuAutoFaceUp.Checked
    MoveFaceUpAuto = Not MoveFaceUpAuto
End Sub


Sub mnuAutoSpeed_Click()
    frmScroll.VScroll1.Min = 500
    frmScroll.VScroll1.Max = 10
    frmScroll.VScroll1.Value = timerDemo.Interval
    frmScroll.VScroll1.LargeChange = (frmScroll.VScroll1.Min - frmScroll.VScroll1.Max) / 10
    frmScroll.VScroll1.SmallChange = 10
    frmScroll.lblScrollValue.Caption = CStr(frmScroll.VScroll1.Value)
    frmScroll.Caption = "Auto Move Rate"
    frmScroll.Show 1
    timerDemo.Interval = frmScroll.VScroll1.Value
    Unload frmScroll
End Sub

Private Sub mnuAutoplay_Click()
'enable or disable AutoPlay
MoveToFoundationAuto = Not MoveToFoundationAuto
If MoveToFoundationAuto Then
    mnuAutoplay.Checked = True
Else
    mnuAutoplay.Checked = False
End If
End Sub

Sub mnuBackDesign_Click()
    frmChooseBack.Show 1
    SubRefreshPilesBackground mClearForm
End Sub

Sub mnuBackgroundColor_Click()
    SubSetBackgroundColor
    mnuBackgroundColor.Checked = True
End Sub

Sub mnuClickMove_Click()
    mnuClickMove.Checked = Not mnuClickMove.Checked
    mClickMove = Not mClickMove
End Sub


Sub mnuExit_Click()
    Unload Me
End Sub

Sub mnuFastWin_Click()
If gFirstTime Then Exit Sub
    mnuFastWin.Checked = Not mnuFastWin.Checked
End Sub





Sub mnuCpuMove_Click()
If gFirstTime Then Exit Sub
    gMoveMultipleAuto = True
    MoveFaceUpAuto = True
    MoveToFoundationAuto = True
    MoveDealAuto = True
    MultLayoutAuto = True
    timerDemo_Timer
End Sub

Sub mnuNewGame_Click()
gNoMoreMoves = 0
If Not gFirstTime Then
    If MsgBox("Abandon current game and start another game?", MB_YESNO + MB_ICONEXCLAMATION, "Solitaire") = IDYES Then
        SubInitNewGame
    End If
Else
    SubInitNewGame
End If
End Sub


Sub mnuRefresh_Click()
    SubResize
    SubRefreshPilesBackground mClearForm
    SubUpdateMobealbleToAll
End Sub



Private Sub mnuSound_Click()
    If mnuSound.Checked = True Then
        SubSetSound False
        mnuSound.Checked = False
    Else
        SubSetSound True
        mnuSound.Checked = True
    End If
End Sub

Sub mnuShowStats_Click()
Dim lRet As Integer
    lRet = SubShowStats(False)
End Sub

Sub mnuUndo_Click()
    SubUndoMove
    SubUpdateMobealbleToAll
    mUserMenuAction = True
End Sub

Sub timerAnimate_Timer()
Dim lRet As Integer
    
    If CPUPlay Then
        If CheckIfCardAnimationIsFinish() Then 'if animation is finish
            lRet = SubCardMove(gPile, gPileCardsPicts, gCardsInPile, mSaveUndo, mUpdateBackground)
            SubUpdateMobealble gPile
            SubUpdateMobealble gPileCardsPicts
            CPUPlay = False
            If lRet Then
                ResetAutoPlayVars
            End If
        End If
    End If
End Sub

Sub timerDemo_Timer()
If gFirstTime Then
    SubInitNewGame
    mnuRefresh_Click
End If

    If Paused Or mUserMenuAction Or CPUPlay Then Exit Sub
    
    'Check if game is completed
    If FuncFastWin(CInt(mnuFastWin.Checked)) Then
        ResetAutoPlayVars
        Paused = True
        SubGameCompleted
        Paused = False
        Exit Sub
    End If
    
    mCpuIsWork = True
    
    'first find in DealPile and TablePile if the first card is face down
    If MoveFaceUpAuto Then
        If FuncAutoFaceUpDealPile() Then
            Exit Sub 'turned card to face up, now exit
        End If
    End If
    
    'and find in table if any card can go to fundation pile(AutoPlay Mode)
    If MoveToFoundationAuto Then
        If FuncAutoMoveToFoundation() Then
            SubSetAutoMoves
            gNoMoreMoves = 0
            Exit Sub
        End If
    End If
    
    'and find in table if any card can go to other table pile
    If MultLayoutAuto Then
        If FuncAutoMoveLayout() Then
            SubSetAutoMoves
            gNoMoreMoves = 0
            Exit Sub
        Else
            MultLayoutAuto = False
        End If
    End If
    
    'don´t find any card to move, go deal from Dealpile
    If MoveDealAuto Then
        If FuncCpuFindMove() Then
                If gMoveMultipleAuto Then
                    SubSetAutoMoves
                Else
                    If CPUPlay Then
                        MoveDealAuto = gPile.tCardIndex <> 1
                    Else
                        MoveDealAuto = gPile.tCardIndex <> 0
                    End If
                    If MoveDealAuto Then
                        MoveDealAuto = (gPileCardsPicts.tnPileNumber = mDiscardPile) Or (gPileCardsPicts.tnPileNumber = mDealPile)
                    End If
                End If
            Exit Sub
        Else
            MoveDealAuto = False
        End If
    End If
    'no more moves at this point
    If gMoveMultipleAuto Then
        MsgBox "Sorry, i can´t find more valid moves", MB_ICONASTERISK
    End If
    ResetAutoPlayVars
End Sub

