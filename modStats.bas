Attribute VB_Name = "modStats"
'**********************************************************************************
'**Solitaire 1.0
'**Copyright by Paulo Matoso
'**E-Mail: paulomt1@clix.pt
'**
'**
'**Last Modification ---> 17/08/2002
'**********************************************************************************



Option Explicit
Dim mTotalGamesPlayed As Integer
Dim mTotalGamesWin As Integer
Dim mTotalGamesLosse As Integer
Dim mChoice As Integer

Sub SubResetStats()
'reset all vars for game stats
    mTotalGamesPlayed = 0
    mTotalGamesWin = 0
    mTotalGamesLosse = 0
End Sub

Sub SubUpdateScore(ByVal pWinLosse As Integer)
    mTotalGamesPlayed = mTotalGamesPlayed + 1
    
    If pWinLosse Then
        mTotalGamesWin = mTotalGamesWin + 1
    Else
        mTotalGamesLosse = mTotalGamesLosse + 1
    End If
End Sub



Sub SubUpdateStats()
    frmStats.lblTotalGames = CStr(mTotalGamesPlayed) 'total games played
    frmStats.lblWins = CStr(mTotalGamesWin) 'total games hard win
    frmStats.lblLosses = CStr(mTotalGamesLosse) 'total games losse
End Sub

Sub SubSetUserChoice(ByVal pValue As Integer)
    mChoice = pValue
End Sub

Function SubShowStats(ByVal pExitBtn As Integer) As Integer
    Load frmStats
    frmStats.cmdExit.Visible = pExitBtn
    frmStats.Caption = "Solitaire Stats"
    frmStats.Left = (Screen.Width - frmStats.Width) / 2
    frmStats.Top = (Screen.Height - frmStats.Height) / 2
    SubUpdateStats
    frmStats.Show 1
    SubShowStats = mChoice
End Function
