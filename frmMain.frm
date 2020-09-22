VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   4005
   ClientTop       =   2430
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2760
      Top             =   2940
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_DblClick()
MainShip.lCurGun = GUN_L2_triple
MainShip.lCurGun2 = GUN_L7_single


End Sub

Private Sub Form_Load()
'Clipboard.Clear
'Clipboard.SetText RGB(255, 0, 255)
Me.Show
Me.Print "Loading Sprites, please wait..."
Init
Me.Move 0, 0, DrawArea.sW * 15, DrawArea.sH * 15
Me.Show
GFX_LOOP

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim aa As DIALOG_A
'With aa
'    .sFont = MainFontSmall
'    .lStay = 1000
'    .dX = X
'    .ldoneWriting = False
'    .dY = Y
'    .sString = "Hello World!"
''    .sTyperText = True
'    .sDelay = 10
'End With
'If Button = 1 Then
'Spawnenemy ENEMY_L1, X, Y
'NewDialog aa

'Else
'Spawnenemy ENEMY_L2, X, Y
'End If
NewExplosion EXP_LARGE, X, Y
KillShip
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ClearStuff
Unload Me
End
End Sub

Private Sub Timer1_Timer()
'Me.Caption = DrawArea.FPS & " " & UBound(GLBL_BUL)
DrawArea.FPSintrvl = DrawArea.FPS
DrawArea.FPS = 0
End Sub
