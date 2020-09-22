VERSION 5.00
Begin VB.Form frmIntro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MANTICORE! THE DEMO"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   26.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000220FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   516
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1800
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This just does a simple intro
Dim TheTile As Tile
Dim GameSounds As clsSound
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long




Public Sub FadeOut()
    Dim I As Long
    Dim J As Long
    picCanvas.Cls
    picCanvas.Width = 90
    picCanvas.Height = 90
    For I = 0 To Me.ScaleWidth Step TheTile.Width
        For J = 0 To Me.ScaleHeight - TheTile.Height Step 32
            TheTile.Blitter.Blt Me.hdc, I, J, 32, 32, picCanvas.hdc, 0, 0, TheTile.Blitter.SRCCOPY
            Me.Refresh
            Sleep 2
            DoEvents
        Next J
    Next I

End Sub


Public Sub RunDemo()
    Dim I As Long
    Dim J As Long
    
    Me.Show
    For I = 0 To Me.ScaleWidth - (TheTile.Width - 12) Step TheTile.Width - 1
        For J = 0 To Me.ScaleHeight - (TheTile.Height - 20) Step 32
            'Put the new one on
            TheTile.BltTile Me.hdc, I, J
            Me.Refresh
            Sleep 2
            DoEvents
        Next J
    Next I
    'Draw the MANTICORE! THE GAME text
    Me.ForeColor = vbBlack
    WriteText Me.hdc, "MANTICORE!", 110, 60
    Me.ForeColor = vbWhite
    WriteText Me.hdc, "MANTICORE!", 100, 50
    GameSounds.PlaySound "Whoomp", SND_ASYNC
    Me.ForeColor = vbRed
    WriteText Me.hdc, "MANTICORE!", 99, 50
    Me.ForeColor = vbWhite
    Me.Font.Size = 10
    WriteText Me.hdc, "THE DEMO", 150, 85
    GameSounds.PlaySound "Whoomp", SND_ASYNC
    Me.Refresh
    Sleep 100
    'Then draw the Hit Enter to Start text
    Me.Font.Size = 11
    Me.ForeColor = RGB(100, 255, 0)
    WriteText Me.hdc, "Hit Enter To Start", 175, 150
    GameSounds.PlaySound "Whoomp", SND_ASYNC
    Me.Refresh
    
End Sub


Public Sub WriteText(DC As Long, theString As String, x As Long, y As Long)
    TextOut DC, x, y, theString, Len(theString)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Unload Me
            frmStart.Visible = True
            frmStart.ZOrder 0
            Exit Sub
    End Select
End Sub

Private Sub Form_Load()
    Set TheTile = New Tile
    Set GameSounds = New clsSound
    TheTile.LoadTile App.Path & "\gfx\title.bmp", 32, 32
    GameSounds.LoadSound App.Path & "\snd\whoomp.wav", "Whoomp"
    RunDemo
End Sub


Private Sub Form_Unload(Cancel As Integer)
    FadeOut
    Set TheTile = Nothing
End Sub


