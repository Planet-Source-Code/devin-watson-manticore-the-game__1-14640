VERSION 5.00
Begin VB.Form frmBattle 
   AutoRedraw      =   -1  'True
   Caption         =   "BATTLE!"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox BackBuffer 
      BackColor       =   &H00000000&
      Height          =   3495
      Left            =   0
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CommandButton cmdAttack 
      Caption         =   "Attack!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   720
      TabIndex        =   1
      Top             =   4020
      Width           =   5895
   End
   Begin VB.PictureBox picScreen 
      BackColor       =   &H00000000&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is the main battle form.
'Here are some objects we will use here:
'TheBlitter - clsBlitter object
'picScreen - PictureBox object
'Pictures.Enemies - PictureBox array
'Pictures.Sprites - PictureBox array
'This will be animated!
Dim AttStr As Integer

Public Function DidHit() As Boolean
    If Player.HasSword Then
        DidHit = (Rnd < 0.65)
    Else
        DidHit = (Rnd < 0.35)
    End If
End Function


Public Sub DoAttack()
    'Performs a normal attack
    
    
    
    
End Sub


Public Sub DoSwordAttack()
    'Performs a sword attack.
    
End Sub

Private Sub cmdAttack_Click()
    'Performs an attack.

    
    If DidHit Then
        If Player.HasSword Then
            AttStr = 1 + Int(2 * Rnd + 0.5)
        Else
            AttStr = 1
        End If
    Else
        AttStr = 0
    End If

'    If Not Player.HasSword Then
'        'Do a regular attack
'        DoAttack
'    Else
'       'Attack with sword
 '       DoSwordAttack
''    End If
    RunTest
End Sub


