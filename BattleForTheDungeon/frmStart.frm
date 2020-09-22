VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "MANTICORE! - THE DEMO"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Choose Your Destiny, Sort Of"
      Height          =   2895
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Game"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New Game"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    DestroyEverything
    Unload Me
    End
End Sub

Private Sub cmdLoad_Click()
    MsgBox "Feature Not Implemented. Yet.", vbOKOnly + vbInformation, "NAAAAAAAAAAAH!"
End Sub


Private Sub cmdNew_Click()
    Me.Hide
    frmMain.Visible = True
    frmMain.ZOrder 0
    DrawMap
End Sub


