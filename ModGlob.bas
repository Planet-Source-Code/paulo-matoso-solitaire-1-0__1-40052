Attribute VB_Name = "ModGlob"
'**********************************************************************************
'**Solitaire 1.0
'**Copyright by Paulo Matoso
'**E-Mail: paulomt1@clix.pt
'**
'**
'**Last Modification ---> 17/08/2002
'**********************************************************************************



Option Explicit

' Dialog Box Command IDs
Declare Function GetTickCount Lib "kernel32" () As Long

Declare Function bitblt Lib "gdi32" _
        Alias "BitBlt" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As _
         Long, ByVal nWidth As Long, ByVal nHeight As _
         Long, ByVal hSrcDC As Long, ByVal xSrc As _
         Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex&) As Long


Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const IDYES = 6

' WM_SIZE message wParam values
Public Const SIZE_RESTORED = 0
Public Const SIZE_MINIMIZED = 1
Public Const SIZE_MAXIMIZED = 2

' Obsolete constant names
Public Const SIZENORMAL = SIZE_RESTORED
Public Const SIZEICONIC = SIZE_MINIMIZED
Public Const SIZEFULLSCREEN = SIZE_MAXIMIZED

Public Const MB_YESNO = &H4

Public Const MB_ICONEXCLAMATION = &H30
Public Const MB_ICONASTERISK = &H40

Global Const gTotalCardsInDeck = 52

Global Const gClubs = 1
Global Const gDiamonds = 2
Global Const gHearts = 3
Global Const gSpades = 4

Global Const gRedCard = 1
Global Const gBlackCard = 2


Global Const gFaceDown = False
Global Const gFaceUp = True

Global Const gFundationRule = 1
Global Const gTableRule = 2

Global Const TablePile = 1
Global Const FundationPile = 2
Global Const DiscardPile = 3
Global Const DealPile = 4

Global Const gNoSaveUndo = False
Global Const gSaveUndo = True

Global Const GameSolitaire = "Solitaire"
Global gBmpFile As String


Global gSolitaireTableRule As Integer
Global gMidleOfCardHeight As Single
Global gLoadPict As Variant
Global CPUPlay As Integer

Global gPile As PILESINFORMATION
Global gPileCardsPicts As PILESINFORMATION

Global gCardsInPile As Integer

Global gSoundTurnedOn As Integer

Global gSoundDeal As String
Global gSoundStack As String
Global gSoundShuffle As String
Global gSoundApplause As String
Global gSoundTurn As String
Global gSoundDrawUp As String

Global Const SND_ASYNC = &H1         '  play asynchronously
Global gNoMoreMoves As Integer
Global gFirstTime As Integer
Global gMoveMultipleAuto As Integer

Type tCordinates
  tX As Integer
  tY As Integer
End Type





Type CardInfo
  tCardSuit As Integer '1=Clubs, 2= Diamonds, 3=Hearts, 4=Spades
  tCardValue As Integer '1,2,3,4,5,6,7,8,9,10,11,12,13 --> 13=King
End Type

Type tPileInfo
  tPileInfoCords As tCordinates
  tPileInfoCard As CardInfo
  tPileInfoMoveable As Integer '  if card is moveable or not
  tPileCardBlocked As Integer ' if card is moveable or not for CPU
End Type


Type UNDOTYPE
  tUndoPileSource As Integer
  tUndoPileInPileSource As Integer
  tUndoPileDest As Integer
  tUndoPileInPileDest As Integer
  tUndoCardsInPileAfterMove As Integer
  tUndoTotalCardsMoved As Integer
End Type



Type PILESINFORMATION
  tnPileNumber As Integer 'pile number
  tnPileNumberInPile As Integer ' pile number in pile
  tCardIndex As Integer 'card selected in pile
  tnPileInfo As tPileInfo
End Type
