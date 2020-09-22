Attribute VB_Name = "mod_GFX"
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Const STRETCH_HALFTONE = 4
Public Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Type BMP
    hDC As Long
    obj As Long
    w As Long
    h As Long
    bmm As BITMAP
End Type

Public Type Bullet_ingun
    offsetX As Long
    offsetY As Long
    sX As Long
    sY As Long
    speedy As Double
    speedx As Double
    hDC As Long
    mskHDC As Long
    CX As Double
    CY As Double
    sDMG As Long
    FriendlyFire As Boolean
    isActive As Boolean
    STCHX As Long
    STCHY As Long
End Type
Public Type GUN_A
    offsetX As Long
    offsetY As Long
    bullets() As Bullet_ingun
    sName As String
    sID As Long
    sLevel As Byte
    sDamage As Double
    firePause As Long
    LOCKED As Boolean
    DLY As Long
    
    propExists As Boolean
    PropsX As Long
    PropsY As Long
    PropoffsetX As Long
    PropoffsetY As Long
    PropW As Long
    PropH As Long
    
    HasGunGFX As Boolean
    GGX As Long
    GGY As Long
    
    FM As Byte
End Type

Public Type Ship
    lHealth As Long
    lCurGun As GUN_A
    lCurGun2 As GUN_A
    PUTINSLOT1 As Boolean
    
    CX As Long
    CY As Long
    GBMP As BMP
    gMSK As BMP
    spdX As Double
    spdY As Double
    'gun effects
    Props As BMP
    propsMsk As BMP
    sScore As Long
    lCash As Long
    
    mxHlth As Long
    sLives As Long
    
    sExplT As Long
    sExplC As Long
End Type

Public Type DA
    backbuffeR As BMP
    FPSintrvl As Long
    tgtHdc As Long
    sW As Long
    sH As Long
    FPS As Long
    D_DEBUG As Boolean
End Type

Public Type SF_STAR
    sX As Double
    sY As Double
    spdX As Double
    spdY As Double
    sCol As Long
End Type

Public Type STARFIELD_A
    spdX As Double
    spdY As Double
    stars() As SF_STAR
End Type

Public Type POINTAPI_
        X As Double
        Y As Double
End Type

Public Type ENEMY_A
    eHealth As Long
    eDamage As Long
    eLevel As Long
    eHdc As Long
    eMSK As Long
    eW As Long
    eH As Long
    sX As Long
    sY As Long
    CX As Double
    CY As Double
    isActive As Boolean
    
    lCurGun As GUN_A
    cashValue As Long
    scoreValue As Long
    
    hasmoveroutine As Boolean
    moveRoutine() As POINTAPI_
    
    sName As String
    sID As Long
    
    spdX As Double
    spdY As Double
    
    wLock As Boolean
    wTime As Long
    wDelay As Long
End Type

'level 1 guns
'small red dot
Public GUN_L1_single As GUN_A
Public GUN_L1_double As GUN_A
Public GUN_L1_triple As GUN_A
Public GUN_L1_quad As GUN_A

'level 2 guns
'red spear
Public GUN_L2_single As GUN_A
Public GUN_L2_double As GUN_A
Public GUN_L2_triple As GUN_A
Public GUN_L2_quad As GUN_A

'level 3 guns
'large red fireball
Public GUN_L3_single As GUN_A
Public GUN_L3_double As GUN_A
Public GUN_L3_triple As GUN_A
Public GUN_L3_quad As GUN_A

'level 4 guns
'subnuke missles
Public GUN_L4_single As GUN_A
Public GUN_L4_double As GUN_A
Public GUN_L4_triple As GUN_A
Public GUN_L4_quad As GUN_A

'level 5 guns
'plasma spear
Public GUN_L5_single As GUN_A
Public GUN_L5_double As GUN_A
Public GUN_L5_triple As GUN_A
Public GUN_L5_quad As GUN_A

'level 6 guns
'plasma ball
Public GUN_L6_single As GUN_A
Public GUN_L6_double As GUN_A
Public GUN_L6_triple As GUN_A
Public GUN_L6_quad As GUN_A

'level 7 guns
'plasma wave
Public GUN_L7_single As GUN_A
Public GUN_L7_double As GUN_A
Public GUN_L7_triple As GUN_A
Public GUN_L7_quad As GUN_A

Public ENEMY_L1 As ENEMY_A
Public ENEMY_L2 As ENEMY_A

Public BULLETGFX As BMP
Public ENEMYGFX As BMP
Public ENEMYMSK As BMP
Public GUNGFX As BMP
Public BULLETGFXmsk As BMP
Public SPTH As String
Public GLBL_BUL() As Bullet_ingun 'all bullets being drawn
Public DrawArea As DA
Public GAME_EXIT As Boolean
Public MainShip As Ship
Public FIR As BMP

Public enemy_SHIPS() As ENEMY_A
Public starfield As STARFIELD_A
Public Const MASKCOLOR As Long = 16711935


'goodie system
Public Type GOODIE_A
    iActv As Boolean
    sX As Long
    sY As Long
    CX As Double
    CY As Double
    hDC As Long
    MSK As Long
    TYP As Byte
    spdX As Double
    spdY As Double
    
    sWPN As GUN_A
    TME As Long
    MXSTAY As Long
    
    CASHVAL As Long
    SCOREVAL As Long
    HEALTHVAL As Long
    DAMAGEVAL As Long
End Type

Public goodieMSK As BMP
Public goodieGFX As BMP
Public GDY_Cash As GOODIE_A
Public GDY_Score As GOODIE_A
Public GDY_Wpn As GOODIE_A
Public GDY_Health As GOODIE_A

Public GOODIES() As GOODIE_A


'fonts

Public Type FONT_A
    sCharW As Long
    sCharH As Long
    sCharHDC As Long
    sCharsPerRow As Long
    sCharsPerCol As Long
    sStretchDraw As Boolean
    sStretchW As Long
    sStretchH As Long
End Type

Public FONTGFX As BMP

Public MainFont As FONT_A
Public MainFontSmall As FONT_A

Public Type DIALOG_A
    sString As String
    sID As Long
    sFont As FONT_A
    
    lStay As Long
    lTime As Long
    
    lWordWrap As Boolean
    lWidth As Long
    lHeight As Long
    
    dX As Long
    dY As Long
    
    sTyperText As Boolean
    sCCchar As Long
    sDelay As Long
    sCnt As Long
    
    ldoneWriting As Boolean
    
End Type

Public Dialogs() As DIALOG_A

Public Type SCRIPT_ENTRY_A
    sCmd As String
    cmdID As Byte
    Params() As String
    Executed As Boolean
End Type

Public SCRIPT() As SCRIPT_ENTRY_A
Public SCRIPTWAIT As Boolean
Public SCRIPTTIMER As Long
Public SCRIPTWAITTIME As Long
Public ENEMYPROFILES() As ENEMY_A

Public Type GOBJECT_A
    sX As Double
    sY As Double
    spdX As Double
    spdY As Double
    sName As String
    sVisible As Boolean
    sBMP As BMP
    sMSK As BMP
    STCH As Boolean
    STW As Long
    STH As Long
End Type

Public Type EXPLOSION_A
    nFrames As Long
    lBMP As Long
    lMSK As Long
    lFW As Long
    lFH As Long
    lX As Long
    lY As Long
    lDraw As Boolean
    CurFrame As Long
End Type

Public EXP_LARGE As EXPLOSION_A
Public EXP_SMALL As EXPLOSION_A

Public EXPLODEBMP As BMP
Public EXPLODEMSk As BMP
Public EXPLODEsmlBMP As BMP
Public EXPLODEsmlMSk As BMP

Public EXPLOSIONS() As EXPLOSION_A
Public SWEAPONS() As GUN_A
Public OBJS() As GOBJECT_A

Public Sub LoadExplosionGFX(Optional sFile As String = "expld_lrg_4.bmp", Optional sFile2 As String = "expld_sml_4.bmp")
LoadBMP EXPLODEBMP, SPTH & "gfx\" & sFile
LoadBMP EXPLODEsmlBMP, SPTH & "gfx\" & sFile2
MakeMask EXPLODEBMP, EXPLODEMSk
MakeMask EXPLODEsmlBMP, EXPLODEsmlMSk
End Sub
Public Sub LoadExplosions()
With EXP_LARGE
    .lFH = 64
    .lFW = 64
    .lBMP = EXPLODEBMP.hDC
    .lMSK = EXPLODEMSk.hDC
    .nFrames = Int(EXPLODEBMP.w / .lFW)
    .lX = 0
    .lY = 0
    .CurFrame = 0
    .lDraw = False
End With

With EXP_SMALL
    .lFH = 32
    .lFW = 32
    .lBMP = EXPLODEsmlBMP.hDC
    .lMSK = EXPLODEsmlMSk.hDC
    .nFrames = Int(EXPLODEsmlBMP.w / .lFW)
    .lX = 0
    .lY = 0
    .CurFrame = 0
    .lDraw = False
End With
End Sub





Public Sub MakeMask(sBMP As BMP, dBMP As BMP)
'dynamically create masks
Dim CP As Long
With sBMP
    MakeBMP dBMP, .w, .h
    
    For xx = 0 To .w
    For yy = 0 To .h
        CP = GetPixel(.hDC, xx, yy)
        If CP = MASKCOLOR Or CP = RGB(252, 2, 252) Then
            'blt white to dest
            SetPixelV dBMP.hDC, xx, yy, vbWhite
            'blt black to src
            SetPixelV sBMP.hDC, xx, yy, vbBlack
        Else
            'set black pixel to dst
            SetPixelV dBMP.hDC, xx, yy, vbBlack
        End If
    Next
    Next
End With
End Sub

Public Sub InitStarfield(Optional sNumStars As Long = 120)
Dim CL As Long
With starfield
    ReDim .stars(sNumStars)
    .spdX = 0
    .spdY = 0
    For i = 0 To UBound(.stars)
        Randomize
        With .stars(i)
        CL = Int(Rnd * 255)
        .spdX = 0
        .spdY = Val("0.0" & CL)
        .sX = Rnd * DrawArea.sW
        .sY = Rnd * DrawArea.sH - (Rnd * 100)
        If CL > 200 Then CL = 200 'max brightness
        .sCol = RGB(Abs(255 - CL), Abs(255 - CL), Abs(255 - CL))
        End With
    Next
    
End With
End Sub

Public Sub Init()
ReDim enemy_SHIPS(0)
ReDim GOODIES(0)
ReDim GLBL_BUL(0)
ReDim Dialogs(0)
ReDim SCRIPT(0)
ReDim OBJS(0)
ReDim EXPLOSIONS(0)
SPTH = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
LoadBulletGraphics "bullets_4.bmp"
LoadShipGraphics "ship_4.bmp", "ship_guns_4.bmp"
LoadEnemyGFX "enemy_4.bmp"


MakeMask BULLETGFX, BULLETGFXmsk
MakeMask MainShip.GBMP, MainShip.gMSK
MakeMask MainShip.Props, MainShip.propsMsk

LoadGuns

DrawArea.sW = 630
DrawArea.sH = 512

MainShip.spdX = 5
MainShip.spdY = 4
MainShip.lCurGun = GUN_L1_single
MainShip.mxHlth = 200
MainShip.lHealth = MainShip.mxHlth
MainShip.sLives = 3

InitStarfield 2222

MainShip.CX = (DrawArea.sW / 2) - (MainShip.GBMP.w / 2)
MainShip.CY = DrawArea.sH - (MainShip.GBMP.h * 2)
MakeBMP DrawArea.backbuffeR, DrawArea.sW, DrawArea.sH
DrawArea.tgtHdc = frmMain.hDC
LoadGunGFX "gun_boxes_4.bmp"

LoadGoodieGFX "goodies_4.bmp"
LoadGoodies

LoadFontGFX
LoadFonts

LoadExplosionGFX
LoadExplosions

LoadEnemies
LoadBMP FIR, SPTH & "gfx\" & "fire_fx.bmp"
'debug data on
DrawArea.D_DEBUG = True

LoadScript
End Sub
Public Sub LoadFontGFX(Optional sFile As String = "font.bmp")
    LoadBMP FONTGFX, SPTH & "gfx\" & sFile
End Sub
Public Sub LoadBulletGraphics(Optional sFile As String = "Bullets_24.bmp")
    LoadBMP BULLETGFX, SPTH & "gfx\" & sFile
End Sub
Public Sub LoadShipGraphics(Optional sFile As String = "Ship_24.bmp", Optional sFile2 As String = "ship_guns_24.bmp")
With MainShip
    LoadBMP .GBMP, SPTH & "gfx\" & sFile
    LoadBMP .Props, SPTH & "gfx\" & sFile2
End With
End Sub
Public Sub LoadGunGFX(Optional sFile As String = "gun_boxes_24.bmp")
LoadBMP GUNGFX, SPTH & "gfx\" & sFile
End Sub
Public Sub LoadEnemyGFX(Optional sFile As String = "enemy_24.bmp")
LoadBMP ENEMYGFX, SPTH & "gfx\" & sFile
MakeMask ENEMYGFX, ENEMYMSK
End Sub

Public Sub LoadFonts()
With MainFont
    .sCharH = 16
    .sCharW = 12
    .sCharHDC = FONTGFX.hDC
    .sCharsPerRow = Int(FONTGFX.w / .sCharW)
    .sCharsPerCol = Int(FONTGFX.h / .sCharH)
End With
With MainFontSmall
    .sCharH = 16
    .sCharW = 12
    .sCharHDC = FONTGFX.hDC
    .sCharsPerRow = Int(FONTGFX.w / .sCharW)
    .sCharsPerCol = Int(FONTGFX.h / .sCharH)
    .sStretchDraw = True
    .sStretchH = 8
    .sStretchW = 6
End With
End Sub

Public Sub LoadBMP(sBMP As BMP, sFile As String)
With sBMP
    DeleteDC .hDC
    DeleteObject .obj
    .hDC = CreateCompatibleDC(GetDC(0))
    .obj = LoadImage(0, sFile, 0, 0, 0, &H10)
    If .hDC <> 0 And .obj <> 0 Then
        SelectObject .hDC, .obj
        GetObject .obj, Len(.bmm), .bmm
        .w = .bmm.bmWidth
        .h = .bmm.bmHeight
    Else
        MsgBox "Could not load Graphic: " & vbCrLf & sFile, vbCritical, "Error"
    End If
End With
End Sub

Public Sub DELBMP(sBMP As BMP)
With sBMP
    DeleteDC .hDC
    DeleteObject .obj
End With
End Sub

Public Sub ClearStuff()
DELBMP BULLETGFX
DELBMP DrawArea.backbuffeR
DELBMP EXPLODEBMP
DELBMP EXPLODEMSk
DELBMP EXPLODEsmlBMP
DELBMP EXPLODEsmlMSk
DELBMP MainShip.GBMP
DELBMP MainShip.gMSK
DELBMP MainShip.Props
DELBMP MainShip.propsMsk

DELBMP ENEMYGFX
DELBMP ENEMYMSK
DELBMP GUNGFX
DELBMP BULLETGFXmsk

DELBMP FIR
DELBMP goodieMSK
DELBMP goodieGFX

DELBMP FONTGFX

For i = 0 To UBound(OBJS)
    DELBMP OBJS(i).sBMP
    DELBMP OBJS(i).sMSK
Next

DELBMP EXPLODEBMP
DELBMP EXPLODEMSk
DELBMP EXPLODEsmlBMP
DELBMP EXPLODEsmlMSk
End Sub

Public Sub LoadGuns()
' o
With GUN_L1_single
    .offsetX = (MainShip.GBMP.w / 2) - (16 / 2)
    .firePause = 90 'ms
    .propExists = True
    .PropsX = 0
    .PropsY = 0
    .PropoffsetX = 0
    .PropoffsetY = 0
    .PropW = 48
    .PropH = 48
    .HasGunGFX = True
    .GGX = 0
    .GGY = 0
    .sDamage = 10
    ReDim .bullets(0)
    'set bullet info
    With .bullets(0)
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .sDMG = GUN_L1_single.sDamage
        .offsetX = 0
        .offsetY = 0
        .speedy = -4
        .speedx = 0
        .sX = 96
        .sY = 0
        
    End With
    
    .sID = 11
    .sLevel = 1
    .sName = "Red Plasma Flash: Single Fire"
End With

' o o
With GUN_L1_double
    .HasGunGFX = True
    .GGX = 32
    .GGY = 0
    .firePause = 100 'ms
    .sDamage = 11 'each bullet
    ReDim .bullets(1)
    'set bullet info
    With .bullets(0)
        .sDMG = GUN_L1_double.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = 0
        .offsetY = 0
        .speedy = -4.5
        .speedx = -0.08
        .sX = 96
        .sY = 0
    End With
    With .bullets(1)
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = 16
        .offsetY = 0
        .speedy = -4.5
        .speedx = 0.08
        .sX = 96
        .sY = 0
        .sDMG = GUN_L1_double.sDamage
    End With
    
    .sID = 12
    .sLevel = 1
    .sName = "Red Plasma Flash: Double Fire"
End With

'  o
' o o
With GUN_L1_triple
    .HasGunGFX = True
    .GGX = 64
    .GGY = 0
    .firePause = 90 'ms
    .offsetX = 16 / 2
    
    .propExists = True
    .PropsX = 48
    .PropsY = 0
    .PropH = 48
    .PropW = 48
    .PropoffsetX = 0
    .PropoffsetY = 0
    
    .sDamage = 12 'each bullet
    
    ReDim .bullets(2)
    'set bullet info
    With .bullets(0)
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .sDMG = GUN_L1_triple.sDamage
        .offsetX = 0
        .offsetY = 0
        .speedy = -3
        .speedx = -0.08
        .sX = 96
        .sY = 0
    End With
    With .bullets(1)
        .hDC = BULLETGFX.hDC
        .sDMG = GUN_L1_triple.sDamage
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = 16
        .offsetY = 0
        .speedy = -3
        .speedx = 0.08
        .sX = 96
        .sY = 0
    End With
    With .bullets(2)
        .sDMG = GUN_L1_triple.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = (16 / 2)
        .offsetY = -16
        .speedy = -4
        .speedx = 0
        .sX = 96
        .sY = 0
    End With
    
    .sID = 13
    .sLevel = 1
    .sName = "Red Plasma Flash: Triple Fire"
End With

' o o o o
With GUN_L1_quad
    .offsetX = 16
    .HasGunGFX = True
    .GGX = 96
    .GGY = 0
    .firePause = 102 'ms
    .sDamage = 13 'each bullet
    ReDim .bullets(3)
    'set bullet info
    With .bullets(0)
        .sDMG = GUN_L1_quad.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = -16
        .offsetY = 0
        .speedy = -4
        .speedx = -0.12
        .sX = 96
        .sY = 0
    End With
    With .bullets(1)
        .sDMG = GUN_L1_quad.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = -(16 / 2)
        .offsetY = 0
        .speedy = -3
        .speedx = -0.08
        .sX = 96
        .sY = 0
    End With
    With .bullets(2)
        .sDMG = GUN_L1_quad.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = (16 / 2)
        .offsetY = 0
        .speedy = -3
        .speedx = 0.08
        .sX = 96
        .sY = 0
    End With
    With .bullets(3)
        .sDMG = GUN_L1_quad.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = 16
        .offsetY = 0
        .speedy = -4
        .speedx = 0.12
        .sX = 96
        .sY = 0
    End With
    
    .sID = 14
    .sLevel = 1
    .sName = "Red Plasma Flash: Quad Fire"
End With



'RED SPEAR

With GUN_L2_single
    .offsetX = (MainShip.GBMP.w / 2) - (16 / 2)
    .firePause = 60 'ms
    '.propExists = True
    '.PropsX = 0
    '.PropsY = 0
    '.PropoffsetX = 0
    '.PropoffsetY = 0
    '.PropW = 48
    '.PropH = 48
    '.HasGunGFX = True
    '.GGX = 0
    '.GGY = 0
    .sDamage = 27
    .HasGunGFX = True
    .GGX = 32
    .GGY = 32
    ReDim .bullets(0)
    'set bullet info
    With .bullets(0)
        .sDMG = GUN_L2_single.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = 0
        .offsetY = 0
        .speedy = -5.3
        .speedx = 0
        .sX = 64
        .sY = 0
    End With
    
    .sID = 21
    .sLevel = 2
    .sName = "Red Plasma Arrow: Single Fire"
End With

With GUN_L2_double
    .HasGunGFX = True
    .GGX = 64
    .GGY = 32
    .offsetX = (MainShip.GBMP.w / 2) - (16 / 2)
    .firePause = 70 'ms
    '.propExists = True
    '.PropsX = 0
    '.PropsY = 0
    '.PropoffsetX = 0
    '.PropoffsetY = 0
    '.PropW = 48
    '.PropH = 48
    '.HasGunGFX = True
    '.GGX = 0
    '.GGY = 0
    .sDamage = 28
    ReDim .bullets(1)
    'set bullet info
    With .bullets(0)
        .sDMG = GUN_L2_double.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = -(16 / 2)
        .offsetY = 0
        .speedy = -4.78
        .speedx = -1
        .sX = 64
        .sY = 0
    End With
    With .bullets(1)
        .sDMG = GUN_L2_double.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = (16 / 2)
        .offsetY = 0
        .speedy = -4.78
        .speedx = 1
        .sX = 64
        .sY = 0
    End With
    
    .sID = 22
    .sLevel = 2
    .sName = "Red Plasma Arrow: Spreaders"
End With


With GUN_L2_triple
    .offsetX = (MainShip.GBMP.w / 2) - (16 / 2)
    .firePause = 50 'ms
    '.propExists = True
    '.PropsX = 0
    '.PropsY = 0
    '.PropoffsetX = 0
    '.PropoffsetY = 0
    '.PropW = 48
    '.PropH = 48
    '.HasGunGFX = True
    '.GGX = 0
    '.GGY = 0
    .HasGunGFX = True
    .GGX = 96
    .GGY = 32
    .sDamage = 18
    ReDim .bullets(2)
    'set bullet info
    With .bullets(0)
        .sDMG = GUN_L2_triple.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = -(16 / 2)
        .offsetY = 0
        .speedy = -4.7
        .speedx = -0.3
        .sX = 64
        .sY = 0
    End With
    With .bullets(1)
        .sDMG = GUN_L2_triple.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = 0
        .offsetY = 0
        .speedy = -5.5
        .speedx = 0
        .sX = 64
        .sY = 0
    End With
    With .bullets(2)
        .sDMG = GUN_L2_triple.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = 16 / 2
        .offsetY = 0
        .speedy = -4.7
        .speedx = 0.3
        .sX = 64
        .sY = 0
    End With
    
    .sID = 23
    .sLevel = 2
    .sName = "Red Plasma Arrow: Rapid Tri-Arrow"
End With

'level 7 plasma pulse
With GUN_L7_single
    .offsetX = (MainShip.GBMP.w / 2) - (16 / 2)
    .firePause = 22 'ms
    '.propExists = True
    '.PropsX = 0
    '.PropsY = 0
    '.PropoffsetX = 0
    '.PropoffsetY = 0
    '.PropW = 48
    '.PropH = 48
    '.HasGunGFX = True
    '.GGX = 0
    '.GGY = 0
    
    .sDamage = 45
    ReDim .bullets(0)
    'set bullet info
    With .bullets(0)
        .sDMG = GUN_L7_single.sDamage
        .hDC = BULLETGFX.hDC
        .mskHDC = BULLETGFXmsk.hDC
        .offsetX = -16
        .offsetY = 0
        '.STCHX = 32
        '.STCHY = 23
        .speedy = -6.7
        .speedx = 0
        .sX = 0
        .sY = 0
    End With
    .sID = 71
    .sLevel = 7
    .sName = "Blue Plasma Pulse Wave: Single Pulse Mode"
End With

ReDim SWEAPONS(7)

SWEAPONS(0) = GUN_L1_single
SWEAPONS(1) = GUN_L1_double
SWEAPONS(2) = GUN_L1_triple
SWEAPONS(3) = GUN_L1_quad

SWEAPONS(4) = GUN_L2_single
SWEAPONS(5) = GUN_L2_double
SWEAPONS(6) = GUN_L2_triple

SWEAPONS(7) = GUN_L7_single
End Sub


Public Function NewBullet() As Long
Dim cb As Long
For i = 0 To UBound(GLBL_BUL)
    With GLBL_BUL(i)
        If .CX >= DrawArea.sW Or .CX <= 0 Or .CY <= 0 Or .CY >= DrawArea.sH Then
            'bullet available
            clsBLT GLBL_BUL(i)
            .isActive = True
            NewBullet = i
            Exit Function
        End If
    End With
Next
cb = UBound(GLBL_BUL) + 1
ReDim Preserve GLBL_BUL(cb)
GLBL_BUL(cb).isActive = True

NewBullet = cb

End Function

Public Sub clsBLT(sBul As Bullet_ingun)
With sBul
    .CX = 0
    .CY = 0
    .hDC = 0
    .offsetX = 0
    .offsetY = 0
    .speedx = 0
    .speedy = 0
    .sX = 0
    .sY = 0
    .isActive = False
End With
End Sub

Public Sub MakeBMP(sBMP As BMP, sW As Long, sH As Long)
With sBMP
    DeleteDC .hDC
    DeleteObject .obj
    .hDC = CreateCompatibleDC(GetDC(0))
    .obj = CreateCompatibleBitmap(GetDC(0), sW, sH)
    If .hDC <> 0 And .obj <> 0 Then
        SelectObject .hDC, .obj
        .w = sW
        .h = sH
    Else
        If MsgBox("Could not initiate the backbuffer!" & vbCrLf & vbCrLf & "Memory resources may be low, try closing open programs and then try again" & vbCrLf & vbCrLf & "Would you like to continue?", vbCritical + vbYesNo, "Error") = vbNo Then
            ClearStuff
            Unload frmMain
            End
        End If
    End If
End With
End Sub


Public Sub GFX_MOVE_N_DRAW_BULLETS()
Dim MISS As Double
For i = 0 To UBound(GLBL_BUL)
    With GLBL_BUL(i)
        'move it
        Randomize
        .CX = .CX + .speedx '+ Val("0." & Int(Rnd * 999))
        .CY = .CY + .speedy + Rnd * 1
        
        'check for collisions on enemy ships
        'If .isActive Then
        If .FriendlyFire Then
        
        For ii = 0 To UBound(enemy_SHIPS)
            'With enemy_SHIPS(ii)
                If .CX >= enemy_SHIPS(ii).CX And .CX <= (enemy_SHIPS(ii).CX + (enemy_SHIPS(ii).eW)) Then
                    'collision X
                    If .CY >= enemy_SHIPS(ii).CY And .CY <= (enemy_SHIPS(ii).CY + enemy_SHIPS(ii).eH) Then
                        'collison y
                        ' a bullet hit this ship
                        
                        'random misses
                        Randomize
                        MISS = Rnd * 10
                        If MISS < 5 Then
                            BitBlt DrawArea.backbuffeR.hDC, .CX + FIR.w, .CY + FIR.h, FIR.w, FIR.h, FIR.hDC, 0, 0, vbSrcPaint
                            EnemyTookHit enemy_SHIPS(ii), GLBL_BUL(i)
                            clsBLT GLBL_BUL(i)
                            
                        End If
                        
                    End If
                End If
            'End With
        Next
        Else
            'non friendly fire
            If .CX >= MainShip.CX And .CX <= (MainShip.CX + (MainShip.GBMP.w - 16)) Then
                    'collision X
                    If .CY >= MainShip.CY And .CY <= (MainShip.CY + MainShip.GBMP.h) Then
                        'collison y
                        'random misses
                        Randomize
                        MISS = Rnd * 10
                        If MISS < 5 Then
                        MainShip.lHealth = MainShip.lHealth - .sDMG
                        If MainShip.lHealth <= 0 Then
                            KillShip
                        End If
                        NewExplosion EXP_SMALL, .CX, .CY
                        clsBLT GLBL_BUL(i)
                        End If
                    End If
            End If
        End If
        'draw it
            If .STCHX = 0 Or .STCHY = 0 Then
                BitBlt DrawArea.backbuffeR.hDC, .CX, .CY, 16, 16, .mskHDC, .sX, .sY, vbSrcAnd
                BitBlt DrawArea.backbuffeR.hDC, .CX, .CY, 16, 16, .hDC, .sX, .sY, vbSrcPaint
            Else
                StretchBlt DrawArea.backbuffeR.hDC, .CX, .CY, .STCHX, .STCHY, .mskHDC, .sX, .sY, 16, 16, vbSrcAnd
                StretchBlt DrawArea.backbuffeR.hDC, .CX, .CY, .STCHX, .STCHY, .hDC, .sX, .sY, 16, 16, vbSrcPaint
            End If
        'End If
    End With
Next
End Sub

Public Sub GFX_CLEAR_BUFFER()
With DrawArea.backbuffeR
    SetStretchBltMode .hDC, 4
    BitBlt .hDC, 0, 0, .w, .h, 0, 0, 0, vbWhiteness
End With
End Sub

Public Sub GFX_DRAWSHIP()
With MainShip
    BitBlt DrawArea.backbuffeR.hDC, .CX, .CY, .GBMP.w, .GBMP.h, .gMSK.hDC, 0, 0, vbSrcAnd
    BitBlt DrawArea.backbuffeR.hDC, .CX, .CY, .GBMP.w, .GBMP.h, .GBMP.hDC, 0, 0, vbSrcPaint

    'draw ship props
    If .lCurGun.propExists Then
        BitBlt DrawArea.backbuffeR.hDC, .CX + .lCurGun.PropoffsetX, .CY + .lCurGun.PropoffsetY, .lCurGun.PropW, .lCurGun.PropH, .propsMsk.hDC, .lCurGun.PropsX, .lCurGun.PropsY, vbSrcAnd
        BitBlt DrawArea.backbuffeR.hDC, .CX + .lCurGun.PropoffsetX, .CY + .lCurGun.PropoffsetY, .lCurGun.PropW, .lCurGun.PropH, .Props.hDC, .lCurGun.PropsX, .lCurGun.PropsY, vbSrcPaint
    End If
    
    If .lHealth <= 50 Then
        'sExplT As Long
    'sExplC As Long
    If .sExplC >= .sExplT Then
        Randomize
        NewExplosion EXP_SMALL, .CX + IIf(Rnd * 10 <= 5, -(Rnd * 32), (Rnd * 32)), .CY + IIf(Rnd * 10 <= 5, -(Rnd * 32), (Rnd * 32))
        Randomize
        .sExplT = Int(Rnd * 100)
        .sExplC = 0
    Else
        .sExplC = .sExplC + 10
    End If
    End If
End With
End Sub

Public Sub GFX_LOOP()
DoEvents
Do
    'do stuff
    GFX_CLEAR_BUFFER            'Clear the backbuffer
    GFX_MOVE_N_DRAW_STARFIELD   'Draw the starfield
    GFX_MOVE_N_DRAW_OBJECTS     'Draw objects
    GFX_MOVE_N_DRAW_ENEMIES     'Draw enemies and move them
    
    GFX_DRAWSHIP                'Draw the ship and its props
    
    GFX_MOVE_N_DRAW_BULLETS     'Draw bullets and check for bullet collisions
    GFX_MOVE_N_DRAW_GOODIES     'Draw helpers and check for helper collisions
    
    GFX_DRAW_EXPLOSIONS         'Draw graphical explosions
    GFX_BLT_HUD                 'Draw the heads-up-display
    GFX_DRAW_DIALOGS            'Draw Dialog text
    
    SPT_EXECUTE_SCRIPTS         'Execute pending scripts
    GET_INPUT                   'Get keys pressed
    
    If DrawArea.D_DEBUG Then GFX_DRAW_DEBUG     'Draw debug data
    
    'draw the backbuffer
    With DrawArea
        BitBlt .tgtHdc, 0, 0, .backbuffeR.w, .backbuffeR.h, .backbuffeR.hDC, 0, 0, vbSrcCopy
    End With
    
    DrawArea.FPS = DrawArea.FPS + 1
DoEvents
Sleep 10
Loop Until GAME_EXIT = True
ClearStuff
End
End Sub

Public Sub GET_INPUT()
Dim NB As Long
If GetAsyncKeyState(vbKeyEscape) < 0 Then
    GAME_EXIT = True
End If
If GetAsyncKeyState(vbKeyRight) < 0 Then
    MainShip.CX = MainShip.CX + MainShip.spdX
End If
If GetAsyncKeyState(vbKeyLeft) < 0 Then
    MainShip.CX = MainShip.CX - MainShip.spdX
End If

If GetAsyncKeyState(vbKeyUp) < 0 Then
    MainShip.CY = MainShip.CY - MainShip.spdY
End If
If GetAsyncKeyState(vbKeyDown) < 0 Then
    MainShip.CY = MainShip.CY + MainShip.spdY
End If

With MainShip.lCurGun
If .sName = "" Then GoTo ng1:
If .LOCKED Then
    'update lock status
    .DLY = .DLY - 10 'ms
    If .DLY <= 0 Then
        .LOCKED = False
        .DLY = 0
    End If
End If

If GetAsyncKeyState(vbKeySpace) < 0 Then
        If Not .LOCKED Then
        For i = 0 To UBound(.bullets)
            NB = NewBullet
            GLBL_BUL(NB) = .bullets(i)
            With GLBL_BUL(NB)
                .FriendlyFire = True
                .CX = MainShip.CX + MainShip.lCurGun.offsetX + .offsetX
                .CY = MainShip.CY + MainShip.lCurGun.offsetY + .offsetY
            
                BitBlt DrawArea.backbuffeR.hDC, .CX + FIR.w, .CY + FIR.h, FIR.w, FIR.h, FIR.hDC, 0, 0, vbSrcPaint
            End With
        Next
        .LOCKED = True
        .DLY = .firePause
        End If
End If
End With

ng1:

With MainShip.lCurGun2
If .sName = "" Then GoTo NG:
If .LOCKED Then
    'update lock status
    .DLY = .DLY - 10 'ms
    If .DLY <= 0 Then
        .LOCKED = False
        .DLY = 0
    End If
End If

If GetAsyncKeyState(vbKeySpace) < 0 Then
        If Not .LOCKED Then
        For i = 0 To UBound(.bullets)
            NB = NewBullet
            GLBL_BUL(NB) = .bullets(i)
            With GLBL_BUL(NB)
                .FriendlyFire = True
                .CX = MainShip.CX + MainShip.lCurGun2.offsetX + .offsetX
                .CY = MainShip.CY + MainShip.lCurGun2.offsetY + .offsetY
            
                BitBlt DrawArea.backbuffeR.hDC, .CX + FIR.w, .CY + FIR.h, FIR.w, FIR.h, FIR.hDC, 0, 0, vbSrcPaint
            End With
        Next
        .LOCKED = True
        .DLY = .firePause
        End If
End If
End With
NG:
End Sub

Public Sub GFX_DRAW_DEBUG()
Dim DBGDAT As String

With DrawArea
    DBGDAT = "FPS: " & .FPSintrvl & " | " & "B: " & UBound(GLBL_BUL) & " | GN: " & MainShip.lCurGun.sName & " | " & UBound(GOODIES) & " | SCR: " & MainShip.sScore & " | $: " & MainShip.lCash & " | HLT: " & MainShip.lHealth & " | SPT: " & DBG_SPTWAITING & "/" & UBound(SCRIPT)
    With .backbuffeR
        TextOut .hDC, 0, 0, DBGDAT, Len(DBGDAT)
    End With
End With


End Sub

Public Sub GFX_BLT_HUD()
On Error Resume Next
Dim GM As String

With MainShip.lCurGun
    GM = UBound(.bullets) + 1 & "x"
    If .HasGunGFX Then
        If DrawArea.D_DEBUG Then
        BitBlt DrawArea.backbuffeR.hDC, 0, 16, 32, 32, GUNGFX.hDC, .GGX, .GGY, vbSrcCopy
        TextOut DrawArea.backbuffeR.hDC, 0, 48, GM, Len(GM)
        Else
        BitBlt DrawArea.backbuffeR.hDC, 0, 0, 32, 32, GUNGFX.hDC, .GGX, .GGY, vbSrcCopy
        'TextOut DrawArea.backbuffeR.hDC, 0, 32, GM, Len(GM)
        End If
    Else
        If DrawArea.D_DEBUG Then
            'blank
        BitBlt DrawArea.backbuffeR.hDC, 0, 16, 32, 32, GUNGFX.hDC, 0, 32, vbSrcCopy
        TextOut DrawArea.backbuffeR.hDC, 0, 48, GM, Len(GM)
        Else
        BitBlt DrawArea.backbuffeR.hDC, 0, 0, 32, 32, GUNGFX.hDC, 0, 2, vbSrcCopy
        'TextOut DrawArea.backbuffeR.hDC, 0, 32, GM, Len(GM)
        End If
        
    End If
End With

With MainShip.lCurGun2
    GM = UBound(.bullets) + 1 & "x"
    If .HasGunGFX Then
        If DrawArea.D_DEBUG Then
        BitBlt DrawArea.backbuffeR.hDC, 32, 16, 32, 32, GUNGFX.hDC, .GGX, .GGY, vbSrcCopy
        TextOut DrawArea.backbuffeR.hDC, 32, 48, GM, Len(GM)
        Else
        BitBlt DrawArea.backbuffeR.hDC, 32, 0, 32, 32, GUNGFX.hDC, .GGX, .GGY, vbSrcCopy
        'TextOut DrawArea.backbuffeR.hDC, 32, 32, GM, Len(GM)
        End If
    Else
        If DrawArea.D_DEBUG Then
        BitBlt DrawArea.backbuffeR.hDC, 32, 16, 32, 32, GUNGFX.hDC, 0, 32, vbSrcCopy
        TextOut DrawArea.backbuffeR.hDC, 32, 48, GM, Len(GM)
        Else
        BitBlt DrawArea.backbuffeR.hDC, 32, 0, 32, 32, GUNGFX.hDC, 0, 32, vbSrcCopy
        'TextOut DrawArea.backbuffeR.hDC, 32, 32, GM, Len(GM)
        End If
    End If
End With
Dim TTO As String
TTO = "Health: " & MainShip.lHealth & "/" & MainShip.mxHlth
TextOut DrawArea.backbuffeR.hDC, 64, 0, TTO, Len(TTO)

End Sub

Public Sub GFX_MOVE_N_DRAW_STARFIELD()
Dim CL As Long
With starfield
    For i = 0 To UBound(.stars)
        With .stars(i)
            .sX = .sX + .spdX + starfield.spdX
            .sY = .sY + .spdY + starfield.spdY
            
            If .sX <= 0 Or .sX >= DrawArea.sW Or .sY >= DrawArea.sH Then
                Randomize
                
                CL = Int(Rnd * 255)
                .spdX = 0
                .spdY = Val("0.0" & CL)
                .sX = Rnd * DrawArea.sW
                .sY = -(Rnd * DrawArea.sH)
                If CL > 200 Then CL = 200
                .sCol = RGB(Abs(255 - CL), Abs(255 - CL), Abs(255 - CL))
        
            End If
            
           SetPixelV DrawArea.backbuffeR.hDC, .sX, .sY, .sCol
        End With
    Next
End With
End Sub

Public Sub LoadEnemies()
ReDim ENEMYPROFILES(1)
With ENEMY_L1
    'ReDim .lCurGun.bullets(0)
    .eDamage = 10
    .eHealth = 45
    .eLevel = 1
    .eW = 33
    .eH = 35
    .sX = 0
    .sY = 0
    .eHdc = ENEMYGFX.hDC
    .eMSK = ENEMYMSK.hDC
    .cashValue = 100
    .scoreValue = 10
    
    .lCurGun = GUN_L1_single

    For i = 0 To UBound(.lCurGun.bullets)
    .lCurGun.bullets(i).speedy = Abs(.lCurGun.bullets(i).speedy) 'opposite direction
    Next
    
    .wLock = 0
    .wDelay = 688 'ms delay of shooting
    .sName = "Organics-1 Scout"
    .sID = 1
    .spdX = 0
    .spdY = 1.333
End With

With ENEMY_L2
    'ReDim .lCurGun.bullets(0)
    .eDamage = 10
    .eHealth = 67
    .eLevel = 1
    .eW = 36
    .eH = 52
    .sX = 33
    .sY = 0
    .eHdc = ENEMYGFX.hDC
    .eMSK = ENEMYMSK.hDC
    .cashValue = 110
    .scoreValue = 14
    
    .lCurGun = GUN_L1_double

    For i = 0 To UBound(.lCurGun.bullets)
    .lCurGun.bullets(i).speedy = Abs(.lCurGun.bullets(i).speedy) 'opposite direction
    Next
    
    .wLock = 0
    .wDelay = 500 'ms delay of shooting
    .sName = "Organics-1 Transporter"
    .sID = 2
    .spdX = 0
    .spdY = 1.45
End With

ENEMYPROFILES(0) = ENEMY_L1
ENEMYPROFILES(1) = ENEMY_L2
End Sub

Public Sub Spawnenemy(sEnemyProfile As ENEMY_A, Optional ByVal sX As Double = 1, Optional ByVal sY As Double = 1)
For i = 0 To UBound(enemy_SHIPS)
    With enemy_SHIPS(i)
        If .CX + .eW <= 0 Or .CX >= DrawArea.sW Or .CY + .eH <= 0 Or .CY >= DrawArea.sH Then
            'free enemy
            enemy_SHIPS(i) = sEnemyProfile
            .isActive = True
            .CX = sX
            .CY = sY
            Exit Sub
        End If
    End With
Next
Dim UU As Long
UU = UBound(enemy_SHIPS) + 1
ReDim Preserve enemy_SHIPS(UU)
enemy_SHIPS(UU) = sEnemyProfile
enemy_SHIPS(UU).CX = sX
enemy_SHIPS(UU).CY = sY
enemy_SHIPS(UU).isActive = True
End Sub

Public Sub GFX_MOVE_N_DRAW_ENEMIES()

Dim NB As Long
For i = 0 To UBound(enemy_SHIPS)
    With enemy_SHIPS(i)
        .CX = .CX + .spdX
        .CY = .CY + .spdY
        If .CX + .eW <= 0 Or .CX >= DrawArea.sW Or .CY >= DrawArea.sH Then
        Else
            If .isActive Then
            BitBlt DrawArea.backbuffeR.hDC, .CX, .CY, .eW, .eH, .eMSK, .sX, .sY, vbSrcAnd
            BitBlt DrawArea.backbuffeR.hDC, .CX, .CY, .eW, .eH, .eHdc, .sX, .sY, vbSrcPaint
            
            If DrawArea.D_DEBUG Then
                TextOut DrawArea.backbuffeR.hDC, .CX, .CY, .eHealth, Len(Str(.eHealth)) - 1
            End If
            
            If .wLock Then
                'gun locked
                .wTime = .wTime - 10
                If .wTime <= 0 Then
                    .wLock = False
                End If
            Else
                'NewBullet
                With .lCurGun
                For ii = 0 To UBound(.bullets)
                    NB = NewBullet
                    GLBL_BUL(NB) = .bullets(ii)
                    With GLBL_BUL(NB)
                
                        .CX = enemy_SHIPS(i).CX + enemy_SHIPS(i).lCurGun.offsetX + .offsetX
                        .CY = enemy_SHIPS(i).CY + enemy_SHIPS(i).lCurGun.offsetY + .offsetY
            
                        'BitBlt DrawArea.backbuffeR.hdc, .cX + FIR.w, .cY + FIR.h, FIR.w, FIR.h, FIR.hdc, 0, 0, vbSrcPaint
                    End With
                Next
                End With
                .wTime = .wDelay
                .wLock = True
            End If
            End If
            
        End If
        
    End With
Next
End Sub

Public Sub EnemyTookHit(sEnemy As ENEMY_A, sBullet As Bullet_ingun)
Dim NIB As Long
Dim FF As Long
With sEnemy
    .eHealth = .eHealth - sBullet.sDMG
    If .eHealth <= 0 Then
    If .isActive Then
        .spdX = 0
        .spdY = 0
        .isActive = False
        MainShip.sScore = MainShip.sScore + .scoreValue
        'MainShip.lCash = MainShip.lCash + .cashValue
        
        NewExplosion EXP_LARGE, .CX, .CY
        NewExplosion EXP_LARGE, .CX - (Rnd * 10), .CY - (Rnd * 10)
        NewExplosion EXP_SMALL, .CX + (Rnd * 10), .CY + (Rnd * 10)
        
        'spawn goodies
        
        Randomize
        FF = Rnd * 10
        If FF <= 2 Then
        NIB = NewGoodie
        GOODIES(NIB) = GDY_Health
        'MsgBox GOODIES(NIB).MXSTAY
        With GOODIES(NIB)
            .HEALTHVAL = Int(Rnd * 10)
            .CX = sEnemy.CX
            .CY = sEnemy.CY
        End With
        
        End If
        
        
        Randomize
        FF = Rnd * 10
        If FF <= 5 Then
        NIB = NewGoodie
        GOODIES(NIB) = GDY_Cash
        'MsgBox GOODIES(NIB).MXSTAY
        With GOODIES(NIB)
            .CASHVAL = sEnemy.cashValue
            .CX = sEnemy.CX
            .CY = sEnemy.CY
        End With
        GoTo fff:
        End If
        
        Randomize
        FF = Rnd * 10
        If FF <= 3 Then
        NIB = NewGoodie
        GOODIES(NIB) = GDY_Score
        'MsgBox GOODIES(NIB).MXSTAY
        With GOODIES(NIB)
            .SCOREVAL = sEnemy.scoreValue
            .CX = sEnemy.CX
            .CY = sEnemy.CY
        End With
        GoTo fff:
        End If
        
        Randomize
        FF = Rnd * 10
        If FF <= 2 Then
            NIB = NewGoodie
            GOODIES(NIB) = GDY_Wpn
            'MsgBox GOODIES(NIB).MXSTAY
            With GOODIES(NIB)
                .sWPN = sEnemy.lCurGun
                .sX = sEnemy.lCurGun.GGX
                .sY = sEnemy.lCurGun.GGY
                
                .CX = sEnemy.CX
                .CY = sEnemy.CY
            End With
        End If
        
fff:
        .CX = 0
        .CY = 0
        
        
    End If
    End If
End With
End Sub

Public Sub LoadGoodieGFX(Optional sFile As String = "goodies_24.bmp")
LoadBMP goodieGFX, SPTH & "gfx\" & sFile
MakeMask goodieGFX, goodieMSK
End Sub
Public Sub LoadGoodies()
With GDY_Cash
    .hDC = goodieGFX.hDC
    .MSK = goodieMSK.hDC
    .MXSTAY = 2000
    .spdX = 0
    .spdY = -0.3
    .TYP = 1
    .sX = 0
    .sY = 0
End With
With GDY_Score
    .hDC = goodieGFX.hDC
    .MSK = goodieMSK.hDC
    .MXSTAY = 3000
    .spdX = 0
    .spdY = -0.5
    .sX = 16
    .sY = 0
    .TYP = 2
End With
With GDY_Wpn
    .hDC = GUNGFX.hDC
    .MXSTAY = 6000
    .spdX = 0
    .spdY = 1
    .TYP = 3
    .sX = 0
    .sY = 0
End With
With GDY_Health
    .hDC = goodieGFX.hDC
    .MSK = goodieMSK.hDC
    .MXSTAY = 2500
    .spdX = 0
    .spdY = 1.3
    .sX = 32
    .sY = 0
    .TYP = 4
End With
End Sub

Public Function NewGoodie() As Long
Dim UU As Long
Dim i As Long
For i = 0 To UBound(GOODIES)
    With GOODIES(i)
    If .CX <= 0 Or .CX >= DrawArea.sW Or .CY <= 0 Or .CY >= DrawArea.sH Then
        'avail
        ClsGoodie (i)
        DoEvents
        GOODIES(i).iActv = True
        NewGoodie = i
        Exit Function
    End If
    End With
ngl:
Next
UU = UBound(GOODIES) + 1
ReDim Preserve GOODIES(UU)
GOODIES(UU).iActv = True
NewGoodie = UU
End Function

Public Sub GFX_MOVE_N_DRAW_GOODIES()
For i = 0 To UBound(GOODIES)
    With GOODIES(i)
        'If .iActv Then
        .TME = .TME + 10
        
        If .TME >= .MXSTAY Then
            ClsGoodie i
            GoTo NXTG:
        End If
            .CX = .CX + .spdX
            .CY = .CY + .spdY
            
            If .CX >= MainShip.CX And .CX <= (MainShip.CX + MainShip.GBMP.w) Then
                    'collision X
                    If .CY >= MainShip.CY And .CY <= (MainShip.CY + MainShip.GBMP.h) Then
                        'collison y
                        'new goodie gotten
                        Select Case .TYP
                            Case 1 'cash
                                MainShip.lCash = MainShip.lCash + .CASHVAL
                            Case 2 'score
                                MainShip.sScore = MainShip.sScore + .SCOREVAL
                            Case 4 'health
                                MainShip.lHealth = MainShip.lHealth + .HEALTHVAL
                                If MainShip.lHealth > MainShip.mxHlth Then MainShip.lHealth = MainShip.mxHlth
                            Case 3 'weapon
                                If MainShip.PUTINSLOT1 Then
            
                            With MainShip
                                Select Case GOODIES(i).sWPN.sID
                                    Case 11
                                        .lCurGun = GUN_L1_single
                                    Case 12
                                        .lCurGun = GUN_L1_double
                                    Case 13
                                        .lCurGun = GUN_L1_triple
                                    Case 14
                                        .lCurGun = GUN_L1_quad
                                    Case 21
                                        .lCurGun = GUN_L2_single
                                    Case 22
                                        .lCurGun = GUN_L2_double
                                    Case 23
                                        .lCurGun = GUN_L2_triple
                                End Select
                                    .PUTINSLOT1 = False
                            End With
                                Else
                            With MainShip
                                    Select Case GOODIES(i).sWPN.sID
                                    Case 11
                                        .lCurGun2 = GUN_L1_single
                                    Case 12
                                        .lCurGun2 = GUN_L1_double
                                    Case 13
                                        .lCurGun2 = GUN_L1_triple
                                    Case 14
                                        .lCurGun2 = GUN_L1_quad
                                    Case 21
                                        .lCurGun2 = GUN_L2_single
                                    Case 22
                                        .lCurGun2 = GUN_L2_double
                                    Case 23
                                        .lCurGun2 = GUN_L2_triple
                                    End Select
                                    MainShip.PUTINSLOT1 = True
                            End With
                                End If
                        End Select
                        ClsGoodie i
                        GoTo NXTG:
                    End If
            End If
            
            'frmMain.Caption = .TYP
            If .TYP <> 3 Then
            
                BitBlt DrawArea.backbuffeR.hDC, .CX, .CY, 16, 16, .MSK, .sX, .sY, vbSrcAnd
                BitBlt DrawArea.backbuffeR.hDC, .CX, .CY, 16, 16, .hDC, .sX, .sY, vbSrcPaint
            Else
                'weapon
                StretchBlt DrawArea.backbuffeR.hDC, .CX, .CY, 16, 16, .hDC, .sWPN.GGX, .sWPN.GGY, 32, 32, vbSrcCopy
            End If
        'End If
    End With
NXTG:
Next
End Sub

Public Sub ClsGoodie(ByVal sINDX As Long)
With GOODIES(sINDX)
    .CX = -100
    .CY = -100
    .hDC = 0
    '.iActv = False
    .MSK = 0
    .sX = 0
    .sY = 0
End With
End Sub

Public Sub GFX_DRAW_DIALOGS()
Dim CurChar As String
Dim CCX As Long
Dim CCY As Long
Dim CR As Long
Dim CC As Long
For i = 0 To UBound(Dialogs)
    With Dialogs(i)
        If .lWordWrap Then
        Else
            'no word wrap
            If .lStay <> 0 Then
                If .sTyperText And .ldoneWriting Then
                    .lTime = .lTime + 10
                End If
                If .lTime >= .lStay Then
                    'drawn for time allotted
                    GoTo NXTT:
                Else
                    'draw dialog
                    If .sTyperText Then
                        If .sCnt >= .sDelay Then
                            .sCnt = 0
                            If .sCCchar < Len(.sString) - 1 Then
                                .sCCchar = .sCCchar + 1
                            Else
                                .ldoneWriting = True
                            End If
                        Else
                            .sCnt = .sCnt + 10
                        End If
                        ubc = .sCCchar
                    Else
                        ubc = Len(.sString) - 1
                    End If
                    For CC = 0 To ubc
                        CurChar = Mid(.sString, CC + 1, 1)
                        GetCharXY CurChar, .sFont, CCX, CCY
                        If Not .sFont.sStretchDraw Then
                            BitBlt DrawArea.backbuffeR.hDC, .dX + (.sFont.sCharW * CC), .dY, .sFont.sCharW, .sFont.sCharH, .sFont.sCharHDC, CCX, CCY, vbSrcPaint
                        Else
                            StretchBlt DrawArea.backbuffeR.hDC, .dX + (.sFont.sStretchW * CC), .dY, .sFont.sStretchW, .sFont.sStretchH, .sFont.sCharHDC, CCX, CCY, .sFont.sCharW, .sFont.sCharH, vbSrcPaint
                        End If
                    Next
                End If
            End If
        End If
    End With
NXTT:
Next
End Sub

Public Sub GetCharXY(sChar As String, sFont As FONT_A, Optional sX As Long, Optional sY As Long)
Dim RW As Long
Dim CL As Long
Dim AC As Long
Dim CX As Long
Dim CY As Long
If sChar = "" Then Exit Sub
AC = Asc(sChar)
AC = AC - 33 'offset
With sFont
    RW = Int((AC) / (.sCharsPerRow + 1))
    CL = (AC - (RW * .sCharsPerRow))
    
    CX = (CL * .sCharW) - (.sCharW * RW)
    CY = RW * .sCharH
End With

sX = CX
sY = CY
End Sub

Public Function NewDialog(ByRef SDLG As DIALOG_A) As Long
For i = 0 To UBound(Dialogs)
    With Dialogs(i)
        If .lTime >= .lStay Or .lStay = 0 Then
            'free dlg
            Dialogs(i) = SDLG
            NewDialog = i
            Exit Function
        End If
    End With
Next
Dim UU As Long
UU = UBound(Dialogs) + 1
ReDim Preserve Dialogs(UU)
Dialogs(UU) = SDLG
NewDialog = UU
End Function


'scripts
Public Sub LoadScript(Optional sFile As String = "main.spt")
Dim INDATA As String
Dim scptFIL As String
Dim CMD_ID As Long

Dim sENTRYS() As String
Dim PRMS() As String

Dim TSE As Long

scptFIL = SPTH & "scripts\" & sFile

INDATA = Space$(FileLen(scptFIL))
Open scptFIL For Binary Access Read As #1
    Get #1, , INDATA
Close #1

sENTRYS = Split(INDATA, vbCrLf)
If UBound(sENTRYS) <= 0 Then GoTo scptErr:

For i = 0 To UBound(SCRIPT)
    ClsScriptETY i
Next

For i = 0 To UBound(sENTRYS)

If Left(Trim(sENTRYS(i)), 1) = "#" Then GoTo ignore:
If Trim(sENTRYS(i)) = "" Then GoTo ignore:

PRMS = Split(sENTRYS(i), "|")

CMD_ID = 0

Select Case LCase(PRMS(0))
    Case "waittime"
        CMD_ID = 1
    Case "waitfor"
        CMD_ID = 2
    Case "spawn"
        CMD_ID = 3
    Case "helper"
        CMD_ID = 4
    Case "load"
        CMD_ID = 5
    Case "print"
        CMD_ID = 6
    Case "dialog"
        CMD_ID = 7
    Case "object"
        CMD_ID = 8
    Case "starfield"
        CMD_ID = 9
    Case Else
        CMD_ID = 0
        GoTo ignore:
End Select

TSE = NewScriptETY
With SCRIPT(TSE)
    .cmdID = CMD_ID
    .Executed = False
    DoEvents
    .Params = PRMS
    .sCmd = PRMS(0)
End With

ignore:
Next

Exit Sub
scptErr:
MsgBox "Could not load the script: " & sFile, vbCritical, "Script Error"
End Sub



Public Sub SPT_EXECUTE_SCRIPTS()
Dim PRNT As DIALOG_A
Dim HLP As GOODIE_A
Dim FW As Long
Dim HSNG As Long
Dim OB As GOBJECT_A

Dim SSSX As Long
Dim SSSY As Long
Dim FEM1 As Long
'Dim TMPSPL() As String
'Dim TMPSTR As String
If SCRIPTWAIT Then
    'waiting for script
    SCRIPTTIMER = SCRIPTTIMER + 10
    If SCRIPTTIMER >= SCRIPTWAITTIME Then
        SCRIPTWAIT = False
        GoTo EXECUTENEXT:
    End If
    Exit Sub
End If

EXECUTENEXT:
For i = 0 To UBound(SCRIPT)
    With SCRIPT(i)
        If .Executed Then GoTo nxtscpt:
        
        Select Case .cmdID
            Case 1 'wait for time allotted "WAITTIME"
                SCRIPTWAIT = True
                SCRIPTWAITTIME = Val(.Params(1))
                SCRIPTTIMER = 0
                .Executed = True
                Exit Sub
            Case 6 'print string "PRINT"
                With PRNT
                    .sFont = MainFontSmall
                    .sString = SCRIPT(i).Params(2)
                    Select Case LCase(Trim(SCRIPT(i).Params(1)))
                        Case "debug"
                            .dX = 32
                            .dY = 64 + Val(SCRIPT(i).Params(6)) * .sFont.sCharH
                        Case "custom"
                            .dX = Val(SCRIPT(i).Params(3))
                            .dY = Val(SCRIPT(i).Params(4))
                        Case "title"
                            .sFont = MainFont
                            If Val(SCRIPT(i).Params(3)) <> 0 Then
                                .dX = Val(SCRIPT(i).Params(3))
                            Else
                                'center screen
                                .dX = (DrawArea.sW / 2) - (((Len(.sString)) * .sFont.sCharW) / 2)
                            End If
                            If Val(SCRIPT(i).Params(4)) <> 0 Then
                                .dY = Val(SCRIPT(i).Params(4))
                            Else
                                'center screen
                                .dY = (DrawArea.sH / 2) - (.sFont.sCharH / 2)
                            End If
                    End Select
                    
                    .lStay = Val(SCRIPT(i).Params(5))
                    
                    .sTyperText = True
                    .sDelay = 10
                    .ldoneWriting = False
                    
                    'TMPSPL = Split(SCRIPT(i).Params(1), "_")
                    
                    'For c = 0 To UBound(TMPSPL)
                    '    TMPSTR = TMPSTR & TMPSPL(c) & " "
                    'Next
                    
                End With
                NewDialog PRNT
                
                .Executed = True
                Exit Sub
            Case 5 ' load script from file "LOAD"
                .Executed = True
                LoadScript .Params(1)
                Exit Sub
            Case 3 'spawn enemy "SPAWN"
                FEM1 = findEnemy(Val(.Params(1)))
                If FEM1 <> -1 Then
                    If Val(.Params(2)) = 0 Then
                        SSSX = Int(Rnd * DrawArea.sW)
                    Else
                        SSSX = Val(.Params(2))
                    End If
                    If Val(.Params(3)) = 0 Then
                        SSSY = Int(Rnd * DrawArea.sH)
                    Else
                        SSSY = Val(.Params(3))
                    End If
                    Spawnenemy ENEMYPROFILES(FEM1), SSSX, SSSY
                End If
                
                .Executed = True
                
                Exit Sub
            Case 4 'helper "HELPER"
                Select Case LCase(Trim(.Params(1)))
                    Case "cash"
                        HLP = GDY_Cash
                    Case "score"
                        HLP = GDY_Score
                    Case "weapon"
                        HLP = GDY_Wpn
                    Case "health"
                        HLP = GDY_Health
                End Select
                With HLP
                    .CASHVAL = Val(SCRIPT(i).Params(2))
                    .SCOREVAL = Val(SCRIPT(i).Params(3))
                    .HEALTHVAL = Val(SCRIPT(i).Params(4))
                    If Val(SCRIPT(i).Params(6)) <> 0 Then
                        .CX = Val(SCRIPT(i).Params(6))
                    Else
                        .CX = Int(Rnd * DrawArea.sW)
                    End If
                    If Val(SCRIPT(i).Params(7)) <> 0 Then
                        .CY = Val(SCRIPT(i).Params(7))
                    Else
                        .CY = Int(Rnd * DrawArea.sH)
                    End If
                    
                    FW = FindWeapon(Val(SCRIPT(i).Params(5)))
                    If FW <> -1 Then
                        .sWPN = SWEAPONS(FW)
                    End If
                End With
                
                HSNG = NewGoodie
                GOODIES(HSNG) = HLP
                
                .Executed = True
                Exit Sub
                
            Case 8 'object "OBJECT"
                With OB
                    LoadBMP .sBMP, SPTH & "gfx\" & SCRIPT(i).Params(1)
                    MakeMask .sBMP, .sMSK
                    .sX = Val(SCRIPT(i).Params(2))
                    .sY = Val(SCRIPT(i).Params(3))
                    .spdX = Val(SCRIPT(i).Params(4))
                    .spdY = Val(SCRIPT(i).Params(5))
                    .sName = SCRIPT(i).Params(6)
                    .sVisible = True
                    .STCH = IIf(SCRIPT(i).Params(7) = "1", True, False)
                    .STW = Val(SCRIPT(i).Params(8))
                    .STH = Val(SCRIPT(i).Params(9))
                End With
                AddObject OB
                .Executed = True
                Exit Sub
            Case 9 'Starfield "STARFIELD"
                With starfield
                    .spdX = Val(SCRIPT(i).Params(1))
                    .spdY = Val(SCRIPT(i).Params(2))
                End With
                .Executed = True
                Exit Sub
        End Select
    End With
    
nxtscpt:
Next
End Sub

Public Function NewScriptETY() As Long
Dim UU As Long
For i = 0 To UBound(SCRIPT)
    With SCRIPT(i)
        If .Executed = True Or .sCmd = "" Then
            'free script entry
            NewScriptETY = i
            Exit Function
        End If
    End With
Next
'make new
UU = UBound(SCRIPT) + 1
ReDim Preserve SCRIPT(UU)
NewScriptETY = UU
End Function

Public Sub ClsScriptETY(ByVal sINDX As Long)
With SCRIPT(sINDX)
    .Executed = True
    .sCmd = ""
    .cmdID = 0
    'Erase .Params
End With
End Sub


Public Function DBG_SPTWAITING() As Long
Dim TC As Long
For i = 0 To UBound(SCRIPT)
    If SCRIPT(i).Executed = False Then
        TC = TC + 1
    End If
Next
DBG_SPTWAITING = TC
End Function

Public Function findEnemy(ByVal sID1 As Long) As Long
For i = 0 To UBound(ENEMYPROFILES)
    With ENEMYPROFILES(i)
        If .sID = sID1 Then
            findEnemy = i
            Exit Function
        End If
    End With
Next
findEnemy = -1
End Function

Public Function FindWeapon(ByVal sWPNID As Long) As Long
For i = 0 To UBound(SWEAPONS)
    With SWEAPONS(i)
        If .sID = sWPNID Then
            FindWeapon = i
            Exit Function
        End If
    End With
Next
FindWeapon = -1
End Function

Public Function AddObject(sOBJ As GOBJECT_A)
Dim UU As Long
For i = 0 To UBound(OBJS)
    With OBJS(i)
        If .sVisible = False Or .sName = "" Then
            OBJS(i) = sOBJ
            .sVisible = True
            Exit Function
        End If
    End With
Next

UU = UBound(OBJS) + 1
ReDim Preserve OBJS(UU)
OBJS(UU) = sOBJ
OBJS(UU).sVisible = True

End Function

Public Sub GFX_MOVE_N_DRAW_OBJECTS()
For i = 0 To UBound(OBJS)
    With OBJS(i)
        If .sVisible = True Then
            .sX = .sX + .spdX
            .sY = .sY + .spdY
            
            If .sX >= DrawArea.sW Or .sY >= DrawArea.sH Then
                .sVisible = False
                GoTo NXTOBJ:
            End If
            
            'draw it
            If Not .STCH Then
                BitBlt DrawArea.backbuffeR.hDC, .sX, .sY, .sBMP.w, .sBMP.h, .sMSK.hDC, 0, 0, vbSrcAnd
                BitBlt DrawArea.backbuffeR.hDC, .sX, .sY, .sBMP.w, .sBMP.h, .sBMP.hDC, 0, 0, vbSrcPaint
            Else
                StretchBlt DrawArea.backbuffeR.hDC, .sX, .sY, .STW, .STH, .sMSK.hDC, 0, 0, .sMSK.w, .sMSK.h, vbSrcAnd
                StretchBlt DrawArea.backbuffeR.hDC, .sX, .sY, .STW, .STH, .sBMP.hDC, 0, 0, .sBMP.w, .sBMP.h, vbSrcPaint
            End If
        End If
    End With
NXTOBJ:
Next
End Sub

Public Sub GFX_DRAW_EXPLOSIONS()
For i = 0 To UBound(EXPLOSIONS)
    With EXPLOSIONS(i)
        If .lDraw Then
            If .CurFrame < .nFrames Then
                .CurFrame = .CurFrame + 1
            Else
                .lDraw = False
                GoTo NXTEXPLSN:
            End If
            BitBlt DrawArea.backbuffeR.hDC, .lX, .lY, .lFW, .lFH, .lMSK, (.lFW * .CurFrame), 0, vbSrcAnd
            BitBlt DrawArea.backbuffeR.hDC, .lX, .lY, .lFW, .lFH, .lBMP, (.lFW * .CurFrame), 0, vbSrcPaint
        End If
    End With
NXTEXPLSN:
Next
End Sub

Public Function NewExplosion(sExplProfile As EXPLOSION_A, Optional ByVal sX = 0, Optional ByVal sY = 0) As Long
Dim UU As Long
For i = 0 To UBound(EXPLOSIONS)
    With EXPLOSIONS(i)
        If .lDraw = False Then
            'free explosion
            EXPLOSIONS(i) = sExplProfile
            .lX = sX
            .lY = sY
            .lDraw = True
            NewExplosion = i
            Exit Function
        End If
    End With
Next
UU = UBound(EXPLOSIONS) + 1
ReDim Preserve EXPLOSIONS(UU)

EXPLOSIONS(i) = sExplProfile
With EXPLOSIONS(i)
    .lX = sX
    .lY = sY
    .lDraw = True
End With
NewExplosion = UU
End Function

Public Sub KillShip()
Dim FF As Long
Dim AA As Long

With MainShip
For i = 0 To 14
    Randomize
    FF = Rnd * 10
    If FF <= 5 Then
        Randomize
        AA = NewExplosion(EXP_SMALL, .CX + IIf(Rnd * 10 <= 5, -(Rnd * 32), (Rnd * 32)), .CY + IIf(Rnd * 10 <= 5, -(Rnd * 32), (Rnd * 32)))
    Else
        Randomize
        AA = NewExplosion(EXP_LARGE, .CX + IIf(Rnd * 10 <= 5, -(Rnd * 32), (Rnd * 32)), .CY + IIf(Rnd * 10 <= 5, -(Rnd * 32), (Rnd * 32)))
        
    End If
    
    Randomize
    With EXPLOSIONS(AA)
            .CurFrame = Int(Rnd * .nFrames)
        End With
    'Sleep Int(Rnd * 10)
Next

    .lCurGun = GUN_L1_single
    .CX = (DrawArea.sW / 2) - (.GBMP.w / 2)
    .CY = (DrawArea.sH - (.GBMP.h * 2))
    .lHealth = .mxHlth
    .PUTINSLOT1 = False
    .sScore = .sScore - 200
    .sLives = .sLives - 1
End With
End Sub
