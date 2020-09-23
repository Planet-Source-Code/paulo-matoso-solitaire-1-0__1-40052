Attribute VB_Name = "modUtils"
'**********************************************************************************
'**Solitaire 1.0
'**Copyright by Paulo Matoso
'**E-Mail: paulomt1@clix.pt
'**
'**
'**Last Modification ---> 17/08/2002
'**********************************************************************************



Option Explicit
Dim mFile As String
Const mContinuePlaying = 1


Const SM_CYCAPTION = 4

Const mWindowHeight = 600
Const mWindowWidth = 800

Const cdlColor = 3

Const SoundDisable = 0
Const mClearForm = True

Const mAnimationSpeed = 500
Const mAnimationInterval = 30 / mAnimationSpeed
Const cdlOFNReadOnly = &H1 'Causes the Read Only check box to be initially checked when the dialog box is created. This flag also indicates the state of the Read Only check box when the dialog box is closed.
Const mExitGame = 2


Sub SubInit()
Dim lCard As tCordinates
    frmTable.chkDisableSound = SoundDisable
    SubSetSound (frmTable.chkDisableSound = SoundDisable)
    frmTable.mnuSound.Checked = (frmTable.chkDisableSound = SoundDisable)
    

    SubSetCardBack 1
    
    frmTable.WindowState = SIZENORMAL
    mFile = vbNullString
    gLoadPict = True
    frmTable.Picture = LoadPicture(mFile)
    gLoadPict = False
    frmTable.timerDemo.Interval = 250
    gBmpFile = App.Path + "\Cards.bmp"
    SubInitializeDeck frmTable, mAnimationInterval, mAnimationSpeed, gBmpFile
    SubNewGame
    SubGetCardMeasure lCard 'get card pixels values
    gMidleOfCardHeight = lCard.tY / 2 'save this for calculate the drop cards
End Sub

Function FuncAreUSure(Optional pWinFlag As Integer) As Integer
        Select Case SubShowDialogExitSave("Solitaire", "Are you sure?", "Exit", "Cancel Exit")
        Case 1
                FuncAreUSure = True ' exit
                If pWinFlag = True Then End
                Exit Function
        Case 2
                FuncAreUSure = False ' cancel exit
                Exit Function
        End Select
FuncAreUSure = True
End Function

Sub SubSetBackgroundColor()
    frmTable.CMDialog1.CancelError = True
    frmTable.CMDialog1.Flags = cdlOFNReadOnly
    frmTable.CMDialog1.Color = frmTable.BackColor
    frmTable.CMDialog1.Action = cdlColor
    frmTable.BackColor = frmTable.CMDialog1.Color
    gLoadPict = True
    frmTable.Picture = LoadPicture(vbNullString)
    DoEvents
    SubRefreshPilesBackground mClearForm
    gLoadPict = False
End Sub



Sub SubInitNewGame()
Dim lSound As Integer

    SubPutImgCardsInBuffer frmTable
    lSound = frmTable.mnuSound.Checked 'read from tableform
    
    'SubSetSound False 'set sound off for deal the cards into form
    SubDealCardsInTableGame 'draw all cards into form
    SubUpdateMobealbleToAll 'update the moveable cards
    
    'SubSetSound lSound
    
    If Not gFirstTime Then 'skip for the first game
        SubUpdateScore False 'add loose to stats
    Else
        gFirstTime = False
    End If
End Sub




Sub SubGameCompleted()
    SubPlayApplause
    SubUpdateScore True
    Select Case SubShowStats(True)
    Case mExitGame
        FuncAreUSure True
    Case mContinuePlaying
        SubInitNewGame
    End Select
End Sub

Sub SubActualizeUndosNumber(pUndosLeft As Integer, ByVal pValue As Integer, ByVal pMaxUndosPermited As Integer)
    If pValue > 0 Then
        If pUndosLeft = pMaxUndosPermited Then
            pUndosLeft = 1
        Else
            pUndosLeft = pUndosLeft + 1
        End If
    Else
        If pUndosLeft = 1 Then
            pUndosLeft = pMaxUndosPermited
        Else
            pUndosLeft = pUndosLeft - 1
        End If
    End If
End Sub

Function FuncGetCordXForBox() As Single
    FuncGetCordXForBox = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CYCAPTION)
End Function

Sub Delay(ByVal pMileseconds As Integer)
Dim lTime As Long

    lTime = GetTickCount()
    Do While GetTickCount() - lTime < pMileseconds
    Loop

End Sub
