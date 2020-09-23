VERSION 5.00
Begin VB.Form frmImgCrdBuffers 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Work pictures"
   ClientHeight    =   1770
   ClientLeft      =   285
   ClientTop       =   2265
   ClientWidth     =   4725
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   315
   Begin VB.PictureBox pictBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   60
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   308
      TabIndex        =   0
      Top             =   30
      Width           =   4620
      Begin VB.PictureBox ImgCrdDragBuild 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000011&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   3450
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   4
         Top             =   30
         Width           =   1095
      End
      Begin VB.PictureBox ImgCrdDrag 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   2310
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   3
         Top             =   30
         Width           =   1095
      End
      Begin VB.PictureBox ImgCrdDragBackGround 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   1170
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   2
         Top             =   30
         Width           =   1095
      End
      Begin VB.PictureBox pictCrdImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   30
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   1
         Top             =   30
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmImgCrdBuffers"
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




