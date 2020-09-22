VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MANTICORE! - THE DEMO"
   ClientHeight    =   5790
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   523
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'frmMain - Where a lot of the action takes place.

Public Sub TreasureCheck(TileIdx As Long)
'First, we branch on which map we're currently on,
'then check the individual tile indices, to find out
'which generic treasure we got.
    Select Case CurMapIndex
        Case 1
            Select Case TileIdx
    
            'Found a key
            Case 3
                DrawMap
                CurMap.OverwriteTile Player.x, Player.y, "2"
                Sounds.PlaySound "GotTreasure", SND_ASYNC
                Player.NumKeys = Player.NumKeys + 1
                DialogForm.Client.lblDialog = "You got a key. With this, you can travel through portals."
                DialogForm.Show
                Exit Sub
            Case 6
            'Found a sword
                DrawMap
                CurMap.OverwriteTile Player.x, Player.y, "2"
                Sounds.PlaySound "GotTreasure", SND_ASYNC
                Player.HasSword = True
                DialogForm.Client.lblDialog = "You got the sword!"
                DialogForm.Show
                Exit Sub
            Case 5
            'Found a blue potion
                DrawMap
                CurMap.OverwriteTile Player.x, Player.y, "2"
                Sounds.PlaySound "GotTreasure", SND_ASYNC
                Player.HasPotion = True
                DialogForm.Client.lblDialog = "You got a Blue Potion. With this, you can heal yourself. Just press the 'P' key."
                DialogForm.Show
                Exit Sub
        End Select
        
        Case 2
            Select Case TileIdx
            'Found a key
            Case 3
                DrawMap
                CurMap.OverwriteTile Player.x, Player.y, "4"
                Sounds.PlaySound "GotTreasure", SND_ASYNC
                Player.NumKeys = Player.NumKeys + 1
                DialogForm.Client.lblDialog = "You got a key."
                DialogForm.Show
                Exit Sub
            End Select
        End Select

End Sub


Private Sub Form_Initialize()
    'Initialize all of the player-specific
    'data, and where we start.
    Player.x = 2
    Player.y = 5
    Rooms(1).TopX = 1
    Rooms(1).TopY = 1
    Player.Direction = 1
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Every action is determined here. Usually, if this were an
    'action-based game, we wouldn't even use this. I put
    'everything in this event so that you could see everything
    'happen on-screen step-by-step.
    Dim OldX As Long
    Dim OldY As Long
    Dim WorldX As Long
    Dim WorldY As Long
    Dim strSwitchName As String
    Dim NewTerrain As Long
    
    OldX = Player.x
    OldY = Player.y
    
    WorldX = CurMap.TopX
    WorldY = CurMap.TopY
    'This method handles direction control.
    'To do anything above and beyond that (i.e., special keys)
    'implement them below.
    Player.HandleKeystroke KeyCode
    
    'This handles taking a potion.
    If KeyCode = vbKeyP Then
        If Player.HasPotion Then
            Player.Stats("HP").CurrentValue = Player.Stats("HP").MaxValue
            Player.HasPotion = False
        End If
    End If
        
    'We get the next tile on the map, according to
    'what HandleKeystroke computed for us.
    NewTerrain = CurMap.TileAt(Player.x, Player.y)
    'Get the switch here, if there is any
    strSwitchName = CurMap.MapSwitches.SwitchAt(Player.x, Player.y)
    'If the return value is not a blank string, then
    'we can check specific switches. Right now,
    'there is only one switch.
    If Trim(strSwitchName) <> "" Then
        If strSwitchName = "DungeonDoor" And CurMap.MapSwitches(strSwitchName).Value = False Then
            'If the player has a key, then we let them
            'go to the next map.
            If Player.NumKeys > 0 Then
                Player.NumKeys = Player.NumKeys - 1
                CurMap.MapSwitches(strSwitchName).Value = True
                Sounds.PlaySound "Explosion", SND_ASYNC
                Set CurMap = Rooms(2)
                CurMap.TopX = 21
                CurMap.TopY = 1
                CurMapIndex = 2
                Set Player.CurrentMap = CurMap
                Player.x = 27
                Player.y = 2
                DrawMap
                Exit Sub
            Else
            'Tell the player they need a key to advance to the
            'next map.
            DialogForm.Client.lblDialog = "You'll need a key to go through the portal."
            DialogForm.Show
            End If
        End If
    End If
    
    'Otherwise, we check to see if the tile
    'is exclusive. If it is,
    'then we don't let the player go near it.
    'This is only one way to handle exlusive tiles. I'm
    'using them as a means to show which tile indices are
    'solid or impassible.
    If CurMap.Exclusions.IsExclusive(NewTerrain) Then
        Player.x = OldX
        Player.y = OldY
    End If
    'Call the TreasureCheck routine, just to
    'see if we hit a treasure tile.
    TreasureCheck NewTerrain
    
    'Adjust map to follow player, if the player
    'goes outside the extents.
    If ((Player.x - CurMap.TopX) < 3) And (Player.x > 3) Then
        CurMap.TopX = Player.x - 3
    ElseIf ((Player.x - CurMap.TopX) > 6) And (Player.x < 198) Then
        CurMap.TopX = Player.x - 6
    End If
    
    If ((Player.y - CurMap.TopY) < 3) And (Player.y > 3) Then
        CurMap.TopY = Player.y - 3
    ElseIf ((Player.y - CurMap.TopY) > 6) And (Player.y < 198) Then
        CurMap.TopY = Player.y - 6
    End If
    'And we render the map again.
    DrawMap
End Sub


Private Sub Form_Load()
    'Set all of our properties for the
    'dialog form, which will be translucent.
    
    Set DialogForm = New TranslucentForm
    Set DialogForm.Client = Dialog

    DialogForm.TranslucentColor = RGB(124, 167, 177)
    DialogForm.Client.Width = frmMain.Width + 35

    'This "primes the pump" so that an accurate
    'snapshot will be taken of the background
    'for translucency effects.
    DialogForm.Client.Visible = True
    DialogForm.Client.ZOrder 0
    DialogForm.MakeTranslucent
    DialogForm.Client.Visible = False
End Sub

Private Sub Form_Paint()
    'We call the render routine here
    'so the map stays consistently on the
    'form.
    DrawMap
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Destroy all of our structures in memory
    DestroyEverything
    Cancel = True
    'This is just to make sure the program exits.
    End
End Sub


Private Sub mnuExit_Click()
    'Well, if they want to exit, they
    'gotta exit...
    Unload Me
End Sub




