VERSION 5.00
Begin VB.Form frmChooseBack 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Card Back Design"
   ClientHeight    =   2085
   ClientLeft      =   1110
   ClientTop       =   1485
   ClientWidth     =   6600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   HelpContextID   =   50
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   139
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pictBacks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   5
      Left            =   5460
      ScaleHeight     =   1425
      ScaleWidth      =   1065
      TabIndex        =   9
      Top             =   60
      Width           =   1095
   End
   Begin VB.PictureBox pictBacks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   4
      Left            =   4380
      ScaleHeight     =   1425
      ScaleWidth      =   1065
      TabIndex        =   8
      Top             =   60
      Width           =   1095
   End
   Begin VB.PictureBox pictBacks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   3
      Left            =   3300
      ScaleHeight     =   1425
      ScaleWidth      =   1065
      TabIndex        =   7
      Top             =   60
      Width           =   1095
   End
   Begin VB.PictureBox pictBacks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   2
      Left            =   2220
      ScaleHeight     =   1425
      ScaleWidth      =   1065
      TabIndex        =   6
      Top             =   60
      Width           =   1095
   End
   Begin VB.PictureBox pictBacks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   1
      Left            =   1140
      ScaleHeight     =   1425
      ScaleWidth      =   1065
      TabIndex        =   5
      Top             =   60
      Width           =   1095
   End
   Begin VB.PictureBox pictBacks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   0
      Left            =   60
      ScaleHeight     =   1425
      ScaleWidth      =   1065
      TabIndex        =   4
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      HelpContextID   =   51
      Left            =   4830
      TabIndex        =   1
      Top             =   1650
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   52
      Left            =   630
      TabIndex        =   0
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Selected picture:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   3
      Top             =   1710
      Width           =   1470
   End
   Begin VB.Label lblSelection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3690
      TabIndex        =   2
      Top             =   1710
      Width           =   585
   End
End
Attribute VB_Name = "frmChooseBack"
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
'**Last Modification ---> 17/08/2002
'**********************************************************************************



Option Explicit
Dim mCardSize As tCordinates
Dim mSelectedBack As Integer
Const SCRCOPY = &HCC0020

Sub cmdCancel_Click()
SubSetCardBack mSelectedBack
SubRefreshPilesBackground False
Unload frmChooseBack
End Sub

Sub cmdOK_Click()
    Unload frmChooseBack
End Sub



Sub Form_Load()
Dim lReturn As Integer
    SubGetCardMeasure mCardSize
    lblSelection.Caption = CStr(FuncGetCardPicture())
    mSelectedBack = FuncGetCardPicture()
    lReturn = bitblt(pictBacks(0).hDC, 0, 0, mCardSize.tX, mCardSize.tY, frmImgCrdBuffers.pictCrdImage.hDC, 0, 0, SCRCOPY)
    lReturn = bitblt(pictBacks(1).hDC, 0, 0, mCardSize.tX, mCardSize.tY, frmImgCrdBuffers.pictCrdImage.hDC, 1 * mCardSize.tX, 0, SCRCOPY)
    lReturn = bitblt(pictBacks(2).hDC, 0, 0, mCardSize.tX, mCardSize.tY, frmImgCrdBuffers.pictCrdImage.hDC, 2 * mCardSize.tX, 0, SCRCOPY)
    lReturn = bitblt(pictBacks(3).hDC, 0, 0, mCardSize.tX, mCardSize.tY, frmImgCrdBuffers.pictCrdImage.hDC, 3 * mCardSize.tX, 0, SCRCOPY)
    lReturn = bitblt(pictBacks(4).hDC, 0, 0, mCardSize.tX, mCardSize.tY, frmImgCrdBuffers.pictCrdImage.hDC, 4 * mCardSize.tX, 0, SCRCOPY)
    lReturn = bitblt(pictBacks(5).hDC, 0, 0, mCardSize.tX, mCardSize.tY, frmImgCrdBuffers.pictCrdImage.hDC, 5 * mCardSize.tX, 0, SCRCOPY)
End Sub





Private Sub pictBacks_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
lblSelection = Str$(Index + 1)
SubSetCardBack Index + 1
SubRefreshPilesBackground False
DoEvents
End Sub

