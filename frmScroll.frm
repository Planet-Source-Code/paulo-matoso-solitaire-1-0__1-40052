VERSION 5.00
Begin VB.Form frmScroll 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2820
   ClientLeft      =   1080
   ClientTop       =   1485
   ClientWidth     =   1335
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2820
   ScaleWidth      =   1335
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VScroll1 
      Height          =   2355
      Left            =   90
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   540
      TabIndex        =   2
      Top             =   2430
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   375
      Left            =   540
      TabIndex        =   1
      Top             =   2010
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Normal"
      Height          =   225
      Left            =   360
      TabIndex        =   6
      Top             =   1470
      Width           =   645
   End
   Begin VB.Label LblFast 
      Caption         =   "Fast"
      Height          =   195
      Left            =   30
      TabIndex        =   5
      Top             =   2610
      Width           =   435
   End
   Begin VB.Label lblScrollValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DefaultValue"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   570
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblSlow 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Slow"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   420
   End
End
Attribute VB_Name = "frmScroll"
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

Sub cmdCancel_Click()
    VScroll1.Value = CSng(lblScrollValue)
    Hide
End Sub

Sub cmdOK_Click()
    Hide
End Sub
