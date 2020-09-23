VERSION 5.00
Begin VB.Form frmStats 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solitaire Stats"
   ClientHeight    =   1275
   ClientLeft      =   1395
   ClientTop       =   1320
   ClientWidth     =   3690
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
   HelpContextID   =   60
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   85
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   246
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      Caption         =   "Reset Scores"
      Height          =   375
      HelpContextID   =   63
      Left            =   2130
      TabIndex        =   2
      Top             =   450
      Width           =   1485
   End
   Begin VB.CommandButton cmdContinue 
      Appearance      =   0  'Flat
      Caption         =   "Continue Game"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   61
      Left            =   2130
      TabIndex        =   1
      Top             =   60
      Width           =   1485
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "Exit Game"
      Height          =   375
      HelpContextID   =   62
      Left            =   2130
      TabIndex        =   0
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label lblLosses 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1680
      TabIndex        =   8
      Top             =   840
      Width           =   150
   End
   Begin VB.Label lblWins 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Width           =   150
   End
   Begin VB.Label lblTotalGames 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   150
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Losses"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   840
      Width           =   1395
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Wins"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   45
      TabIndex        =   4
      Top             =   480
      Width           =   1410
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Games"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   45
      TabIndex        =   3
      Top             =   120
      Width           =   1410
   End
End
Attribute VB_Name = "frmStats"
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
Const mContinuePlaying = 1
Const mExitGame = 2

Sub cmdContinue_Click()
    SubSetUserChoice mContinuePlaying
    Unload frmStats
End Sub

Sub cmdExit_Click()
    SubSetUserChoice mExitGame
    Unload frmStats
    
End Sub

Sub cmdReset_Click()
    SubResetStats
    SubUpdateStats
End Sub
