VERSION 5.00
Begin VB.Form frmDlg 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1770
   ClientLeft      =   1110
   ClientTop       =   1485
   ClientWidth     =   2505
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1770
   ScaleWidth      =   2505
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Command2"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   660
      Width           =   1665
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   2085
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDlg"
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


Sub Command1_Click()
    SubSetBtnChoice (1)
End Sub

Sub Command2_Click()
    SubSetBtnChoice (2)
End Sub
