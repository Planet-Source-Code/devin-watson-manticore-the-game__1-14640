VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   0  'None
   Caption         =   "&H80000009&"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hit Enter or Click the window to close."
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   2685
   End
   Begin VB.Label lblDialog 
      BackStyle       =   0  'Transparent
      Caption         =   "Dialog goes here!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Hide
    End If
End Sub

Private Sub Form_Paint()
    DialogForm.MakeTranslucent
End Sub


Private Sub lblDialog_Click()
    Hide
End Sub


