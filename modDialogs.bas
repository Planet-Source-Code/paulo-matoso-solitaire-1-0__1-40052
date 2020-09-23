Attribute VB_Name = "modDialogs"
'**********************************************************************************
'**Solitaire 1.0
'**Copyright by Paulo Matoso
'**E-Mail: paulomt1@clix.pt
'**
'**
'**Last Modification ---> 17/08/2002
'**********************************************************************************



Option Explicit
Dim mBtnChoice As Integer

Sub SubSetBtnChoice(pBtnChoice As Integer)
     mBtnChoice = pBtnChoice
     Unload frmDlg
End Sub


Function SubShowDialogExitSave(pFormCaption As String, ByVal pMessage As String, ByVal pMenuItem1 As String, ByVal pMenuItem2 As String) As Integer
Dim lBtnHeight As Variant
Dim lHeight As Single
    frmDlg.Caption = pFormCaption
    frmDlg.lblMessage = pMessage
    
    If LenB(pMenuItem1) <> 0 Then
        frmDlg.Command1.Caption = pMenuItem1 'menu1
        frmDlg.Command1.Enabled = True
    Else
        frmDlg.Command1.Enabled = False
    End If
    
    If LenB(pMenuItem2) <> 0 Then
        frmDlg.Command2.Caption = pMenuItem2 'menu2
        frmDlg.Command2.Enabled = True
    Else
        frmDlg.Command2.Enabled = False
    End If
    
    
    
    If frmDlg.Command2.Enabled Then
        lBtnHeight = frmDlg.Command2.Top + frmDlg.Command2.Height
    ElseIf frmDlg.Command1.Enabled Then
        lBtnHeight = frmDlg.Command1.Top + frmDlg.Command1.Height
    Else
        lBtnHeight = 0
    End If
    
    lHeight = frmDlg.lblMessage.Top + frmDlg.lblMessage.Height
    If lHeight > lBtnHeight Then
        frmDlg.Height = lHeight + 2 * FuncGetCordXForBox()
    Else
        frmDlg.Height = lBtnHeight + 2 * FuncGetCordXForBox()
    End If
    
    frmDlg.Move Screen.Width / 2 - frmDlg.Width / 2, Screen.Height / 2 - frmDlg.Height / 2
    frmDlg.Command2.Visible = frmDlg.Command2.Enabled
    frmDlg.Command1.Visible = frmDlg.Command1.Enabled
    frmDlg.Show 1
    SubShowDialogExitSave = mBtnChoice
    
End Function
