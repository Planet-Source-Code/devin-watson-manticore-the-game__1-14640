Attribute VB_Name = "Support"
Option Explicit
'Contains support routines and declarations for the game
'***EVERYTHING STARTS IN SUB MAIN() IN THIS MODULE!***

'Global declarations for references and
'other structures.
Global Sounds As clsSound
Global CurMap As TileMap
Global DialogForm As TranslucentForm
Global CurMapIndex As Integer
'We have 2 TileSets currently; one for each of
'the maps, which represent an individual "room"
Global MasterTiles(2) As TileSet
Global CurEnemy As Tile
Global MasterIcons As TileSet
Global BattleBackground As Tile
Global TheBlitter As clsBlitter
Global Rooms(2) As TileMap
Global NMEs As TileSet
Global ExplosionTiles As TileSet
Global Flags As GameFlags
Global CursorPos As Long 'Used for computing text output

Global Player As Character
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Sub DestroyEverything()
    'Called when the game is exited.
    Dim I As Long
    
    For I = 0 To UBound(Rooms)
        Set Rooms(I) = Nothing
    Next I
    
    Set TheBlitter = Nothing
    Set MasterIcons = Nothing
    Set MasterTiles(1) = Nothing
    Set MasterTiles(2) = Nothing
    Set CurMap = Nothing
End Sub

Public Sub DrawInventory()
    'Draw Inventory
        Dim strTemp As String
        'This would be more elegant, if there
        'were more keys in there. But there aren't.
        Select Case Player.NumKeys
            Case 0
                MasterIcons(1).BltTile Pictures.Canvas.hdc, CursorPos, -3
            Case 1
                MasterIcons(2).BltTile Pictures.Canvas.hdc, CursorPos, -3
            Case 2
                MasterIcons(3).BltTile Pictures.Canvas.hdc, CursorPos, -3
        End Select
        'Blt the sword icon, if the player
        'has the sword
        If Player.HasSword Then
            MasterIcons(4).BltTile Pictures.Canvas.hdc, frmMain.ScaleWidth - 130, 32
        Else
            TheBlitter.Blt Pictures.Canvas.hdc, frmMain.ScaleWidth - 130, 32, 32, 32, Pictures.picBlank.hdc, 0, 0, TheBlitter.SRCCOPY
        End If
        'Blt the Blue Potion icon, if the player
        'has it.
        If Player.HasPotion Then
            MasterIcons(5).BltTile Pictures.Canvas.hdc, frmMain.ScaleWidth - 96, 32
        Else
            TheBlitter.Blt Pictures.Canvas.hdc, frmMain.ScaleWidth - 96, 32, 32, 32, Pictures.picBlank.hdc, 0, 0, TheBlitter.SRCCOPY
        End If
        'And display the current HP. But first, we
        'must erase the old one.
        strTemp = "HP: " & Player.Stats("HP").CurrentValue & "/" & Player.Stats("HP").MaxValue
        TheBlitter.Blt Pictures.Canvas.hdc, frmMain.ScaleWidth - 130, 64, Len(strTemp) * Pictures.Canvas.Font.Size, Pictures.Canvas.Font.Size, Pictures.picBlank.hdc, 0, 0, TheBlitter.SRCCOPY
        
        WriteText strTemp, Pictures.Canvas.hdc, frmMain.ScaleWidth - 130, 64
End Sub

Public Sub DrawMap()
    'This draws the map, then the player (using transparency),
    'then the inventory.
    CurMap.BltMap Pictures.Canvas.hdc
    TheBlitter.TransparentBlt Pictures.Canvas.hdc, (Player.x - CurMap.TopX) * 32, (Player.y - CurMap.TopY) * 32, 32, 32, Pictures.Sprites(Player.Direction).hdc, 0, 0, vbBlack
    'Draw the key icon
    MasterIcons(0).BltTile Pictures.Canvas.hdc, frmMain.ScaleWidth - 130, 0
    'Draw the actual key count, and stock of items
    DrawInventory
    'Now Blt everything over from the back buffer to the
    'main screen so the player can see it.
    TheBlitter.Blt frmMain.hdc, 0, 0, frmMain.Width, _
            frmMain.Height, Pictures.Canvas.hdc, 0, 0, vbSrcCopy

End Sub


Public Sub InitData()
    'Initialize any data
    Set Flags = New GameFlags
    'Create a new player
    Set Player = New Character
    
    'Add a flag for the game
    Flags.Add False, "DestroyedBoss"
    'Create a new stat called HP
    Player.Stats.Add "HP", 0, 20, 10
End Sub

Public Sub InitSounds()
    'Load all the sounds in.
    Set Sounds = New clsSound
    
    Sounds.LoadSound App.Path & "\snd\seccave.wav", "GotTreasure"
    Sounds.LoadSound App.Path & "\snd\explode1.wav", "Explosion"
End Sub


Public Sub LoadGraphics()
    'Loads all of the graphics into the TileSet
    'object MasterTiles.
    Set TheBlitter = New clsBlitter
    Set MasterTiles(1) = New TileSet
    Set MasterTiles(2) = New TileSet
    Set MasterIcons = New TileSet
    Set ExplosionTiles = New TileSet
    Set NMEs = New TileSet
    
    'Set the index property to ensure correct loading.
    MasterTiles(1).Index = 1
    MasterTiles(2).Index = 2
    ExplosionTiles.Index = 3
    NMEs.Index = 4
    MasterIcons.Index = 5
    
    'Set the reference to the clsBlitter object
    Set MasterTiles(1).Blitter = TheBlitter
    Set MasterTiles(2).Blitter = TheBlitter
    Set MasterIcons.Blitter = TheBlitter
    Set ExplosionTiles.Blitter = TheBlitter
    Set NMEs.Blitter = TheBlitter
    
    'Call the load routines to get the gfx loaded into
    'each TileSet. I used the same file, with a
    'different index value, to show that you can
    'mix in multiple TileSet graphics in the same
    'file.
    MasterTiles(1).LoadTiles App.Path, "\main.txt"
    MasterTiles(2).LoadTiles App.Path, "\main.txt"
    ExplosionTiles.LoadTiles App.Path, "\main.txt"
    NMEs.LoadTiles App.Path, "\main.txt"
    MasterIcons.LoadTiles App.Path, "\main.txt"
    MasterIcons(1).CreateMask RGB(0, 255, 0)
    
End Sub

Public Sub LoadMaps()
    'Load all of the maps into memory.
    Dim I As Long
    Dim J As Long
    Dim tmpTile As Tile
    
    For I = 1 To UBound(Rooms)
        Set Rooms(I) = New TileMap
        Rooms(I).LoadMapFile App.Path & "\maps\map00" & I & ".map"
        Set Rooms(I).Exclusions = New ExclusionList
        Rooms(I).Exclusions.LoadExclusionFile App.Path & "\maps\map00" & I & ".exc"
        Rooms(I).DisplayHeight = 12
        Rooms(I).DisplayWidth = 12
        Set Rooms(I).Blitter = TheBlitter
        Rooms(I).TopX = 0
        Rooms(I).TopY = 0
        Set Rooms(I).MapSwitches = New Switches
        Rooms(I).MapSwitches.LoadSwitchFile App.Path & "\maps\map00" & I & ".stc"
        'Load the appropriate TileSet in.
        Set Rooms(I).TileMapSet = MasterTiles(I)
        'All of the MasterTile entries have a 1-to-1
        'relationship with their corresponding maps. This
        'just makes it easier to work with.
    Next I

    'Set the CurMap reference to the first TileMap in the array
    Set CurMap = Rooms(1)
    CurMapIndex = 1
    'Make sure the Player is on the same map.
    Set Player.CurrentMap = CurMap
End Sub

Sub Main()
    'Initialize the game here...
    InitSounds
    InitData
    LoadGraphics
    LoadMaps
    
    'Pre-load all of the forms for later.
    Load frmStart
    Load Dialog
    Load frmMain
    
    'This is our back buffer for the game.
    Pictures.Canvas.Width = frmMain.Width
    Pictures.Canvas.Height = frmMain.Height
    'This is used for the virtual cursor, which
    'prints all of the inventory information out.
    CursorPos = (frmMain.ScaleWidth - 130) + 35
    'Do the intro thingy...
    frmIntro.Show
End Sub




Public Sub WriteText(strText As String, DC As Long, x As Long, y As Long)
    'API call to print text out...
    TextOut DC, x, y, strText, Len(strText)
End Sub


