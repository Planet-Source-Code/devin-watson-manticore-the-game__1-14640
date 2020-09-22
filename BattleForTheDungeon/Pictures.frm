VERSION 5.00
Begin VB.Form Pictures 
   Caption         =   "Pictures (NOT DISPLAYED)"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form2"
   ScaleHeight     =   6090
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBlank 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1935
      Left            =   4680
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Canvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   3
      Top             =   2040
      Width           =   5655
   End
   Begin VB.PictureBox Sprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   2
      Left            =   1680
      Picture         =   "Pictures.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   1200
      Width           =   540
   End
   Begin VB.PictureBox Sprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   1
      Left            =   1080
      Picture         =   "Pictures.frx":13EE
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   1200
      Width           =   540
   End
   Begin VB.PictureBox Sprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   0
      Left            =   480
      Picture         =   "Pictures.frx":2956
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1200
      Width           =   540
   End
End
Attribute VB_Name = "Pictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

