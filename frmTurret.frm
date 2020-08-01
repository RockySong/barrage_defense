VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "弹幕防御 - 庚子建军节献礼"
   ClientHeight    =   10260
   ClientLeft      =   660
   ClientTop       =   870
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   ScaleHeight     =   684
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   971
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00404040&
      Caption         =   "结束本局"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   12000
      MaskColor       =   &H00404040&
      TabIndex        =   14
      Top             =   9600
      Visible         =   0   'False
      Width           =   1417
   End
   Begin VB.CheckBox chkJoy 
      BackColor       =   &H00404040&
      Caption         =   "神助(不计分)"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      ToolTipText     =   "10倍子弹时间, 充足弹药, 受到1/10伤害"
      Top             =   9600
      Width           =   2175
   End
   Begin VB.CommandButton cmdUsers 
      BackColor       =   &H00404040&
      Caption         =   "英雄榜"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   10200
      MaskColor       =   &H00404040&
      TabIndex        =   5
      Top             =   9600
      Width           =   1417
   End
   Begin VB.PictureBox picHUD 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   598
      Left            =   117
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   984
      TabIndex        =   21
      Top             =   117
      Width           =   14755
      Begin VB.ComboBox cmbDifficulty 
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12570
         TabIndex        =   23
         Text            =   "Combo1"
         Top             =   120
         Width           =   1650
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "难度"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   11760
         TabIndex        =   24
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblAutoMode 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "连射"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10800
         TabIndex        =   22
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "弹幕散度"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   9240
         TabIndex        =   20
         Top             =   30
         Width           =   1305
      End
      Begin VB.Label lblFocus 
         BackColor       =   &H00000000&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   9240
         TabIndex        =   19
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "开火频率"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   7320
         TabIndex        =   18
         Top             =   30
         Width           =   1335
      End
      Begin VB.Label lblFPM 
         BackColor       =   &H00000000&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   405
         Left            =   7320
         TabIndex        =   17
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label lblHPTitle 
         BackColor       =   &H00000000&
         Caption         =   "生命值"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   360
         Left            =   4920
         TabIndex        =   15
         Top             =   60
         Width           =   1065
      End
      Begin VB.Label lblAmmo 
         BackColor       =   &H00000000&
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Left            =   3120
         TabIndex        =   13
         Top             =   30
         Width           =   1995
      End
      Begin VB.Label LabelAmmoTitle 
         BackColor       =   &H00000000&
         Caption         =   "弹药"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   2400
         TabIndex        =   12
         Top             =   60
         Width           =   705
      End
      Begin VB.Label lblPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "当前选手"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1905
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8671
      Left            =   117
      MouseIcon       =   "frmTurret.frx":0000
      MousePointer    =   2  'Cross
      ScaleHeight     =   576
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   950
      TabIndex        =   0
      Top             =   840
      Width           =   14287
      Begin VB.HScrollBar hsFPM 
         Height          =   247
         LargeChange     =   4
         Left            =   8115
         Max             =   10
         TabIndex        =   9
         Top             =   6090
         Value           =   5
         Visible         =   0   'False
         Width           =   2587
      End
      Begin VB.HScrollBar hsX 
         Height          =   247
         Left            =   8115
         Max             =   25
         Min             =   1
         TabIndex        =   8
         Top             =   5760
         Value           =   20
         Visible         =   0   'False
         Width           =   2587
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00404040&
         Caption         =   "点击开始 "
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   598
         Left            =   6435
         MaskColor       =   &H00404040&
         TabIndex        =   3
         Top             =   1989
         Width           =   1417
      End
      Begin VB.Timer tmrDraw 
         Enabled         =   0   'False
         Interval        =   33
         Left            =   5967
         Top             =   2106
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "弹幕分散"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   6600
         TabIndex        =   11
         Top             =   5760
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "开火频率"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   6615
         TabIndex        =   10
         Top             =   6045
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "成绩"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   1665
         Left            =   3840
         TabIndex        =   2
         Top             =   3120
         Visible         =   0   'False
         Width           =   6480
      End
   End
   Begin VB.Label lblStat 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   13440
      TabIndex        =   16
      Top             =   9720
      Width           =   840
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "帮助"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   9600
      Width           =   945
   End
   Begin VB.Label lblProj 
      Caption         =   "Label1"
      Height          =   247
      Left            =   10881
      TabIndex        =   1
      Top             =   9594
      Visible         =   0   'False
      Width           =   3523
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrRetumString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Function playalarmsound(mp3filename, flag)
    Dim ret As Long
    If Dir(mp3filename) <> "" Then
        If flag = 1 Then
            ret = mciSendString("play " & mp3filename & " repeat", vbNullString, 0, 0)
        Else
            ret = mciSendString("close " & mp3filename, 0&, 0, 0)
        End If
    End If
    ret = ret
End Function


Private Sub cmbDifficulty_Click()
    gv.turret.projVelMo0 = gv.dfcltLv(cmbDifficulty.ListIndex)
    
End Sub

Private Sub cmdNew_Click()
    Dim i As Long, n As Long
    If gv.state = STATE_PLAYING Then
        gv.state = STATE_SCORE
    End If
End Sub

Private Sub cmdStart_Click()
    gv.state = STATE_PLAYING
    PlaySound vbNullString, 0, 0
    gv.scoreBonus = 0
    ResetTurretAmmo gv.turret
    n = MAX_PROJ_CNT - 1
    For i = 0 To n
        gv.projs(i).leftticks = 0
    Next i
    n = MAX_TGT_CNT - 1
    For i = 0 To n
        gv.tgts(i).deadTicks = 0
        gv.tgts(i).leftticks = 0
    Next i
    
    n = MAX_TGT_CNT - 1
    For i = 0 To n
        gv.tgts(i).leftticks = 0
    Next i
    gv.tgtCnt = 0
    gv.projCnt = 0
    gv.killedCnt = 0
    gv.escapeCnt = 0
    gv.hitCnt = 0
    gv.newHitTick = 0
    gv.myHP = 100
    cmdStart.Visible = False
    gv.gameRemainTick = (CLng(100) + Form1.cmbDifficulty.ListIndex * 10) * 1000 + 200
    Form1.cmdNew.Visible = True
    Form1.cmdUsers.Visible = False
    Form1.cmbDifficulty.Enabled = False
    Form1.chkJoy.Enabled = False
    Form1.pic.SetFocus
    tmrDraw.Interval = 1000 / DRAW_PER_SEC
    ShowCursor 1
    'gv.ts0 = GetTickCount()
End Sub

Private Sub cmdUsers_Click()
    frmPlayer.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlaySound vbNullString, 0, 0
    'Call Terminate(pic.hwnd)
    ShowCursor 1
End Sub

Private Sub hsFPM_Change()
    gv.turret.fpm = gv.fpmSounds(hsFPM.Value)
    gv.turret.tickToNextFire = 1
    gv.turret.tickReload = 1000# * 60 / CSng(gv.turret.fpm)
    gv.turret.sFireFile = GetFpmSoundFile(gv.turret.fpm)
    lblFPM.Caption = Format(gv.turret.fpm, "# RPM")
    If gv.turret.isPlaying = True Then
        PlaySound vbNullString, 0, 0
        PlaySound gv.turret.sFireFile, 0, SND_ASYNC Or SND_FILENAME Or SND_LOOP
    End If
End Sub

Private Sub ProcFireCmd(X As Single, Y As Single)
    Dim xRatio As Single, yRatio As Single
    Dim halfW As Single, halfH As Single
    
    halfW = pic.ScaleWidth / 2
    halfH = pic.ScaleHeight / 2
    xRatio = (X - halfW) / halfW * 1.25 * 1.1
    yRatio = -(Y - halfH) / halfH * 0.83 * 1.1
    
    gv.turret.cam.vecN.X = xRatio
    gv.turret.cam.vecN.Y = yRatio
    gv.turret.cam.vecN.z = 1
    SetupCamera gv.turret.cam
End Sub


Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label6.BackColor = 0
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label6.BackColor = rgb(63, 63, 63)
    Form2.Show 1
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ProcFireCmd X, Y
    
    If Button = 2 Then
        gv.turret.fastFireFactor = 5
        gv.turret.burstRem = 5
    Else
        gv.turret.fastFireFactor = 1#
    End If
    If gv.turret.autoMode = 0 Then
        gv.isFiring = True
    Else
        gv.turret.tickToNextFire = 1
        FireTurret gv.turret.cam, gv.turret.fastFireFactor
        PlaySound App.Path & "\fire_single_2.wav", 0, SND_ASYNC Or SND_FILENAME
    End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        If Button = 2 Then
            gv.turret.fastFireFactor = 5
        Else
            gv.turret.fastFireFactor = 1
        End If
    End If
    ProcFireCmd X, Y

End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        gv.turret.fastFireFactor = 1
    End If
    gv.turret.tickToNextFire = 1
    gv.isFiring = False
End Sub

Private Sub ProcKeyCmd()
    Dim keyCode As Long
    gv.isButtletTimeOn = False
    If gv.keyCmd.isdown = False Then Exit Sub
    keyCode = gv.keyCmd.keyCode
    With Form1
    On Error GoTo ExitSub
    Select Case keyCode
    Case Asc("D")
        .hsX.Value = .hsX.Value + 1
    Case Asc("A")
        .hsX.Value = .hsX.Value - 1
    Case Asc("S")
        .hsFPM.Value = .hsFPM.Value - 1
    Case Asc("W")
        .hsFPM.Value = .hsFPM.Value + 1
    Case Asc("Q")
        If gv.keyCmd.isNewDown = True Then
            gv.turret.autoMode = 1 - gv.turret.autoMode
            ShowAutoMode
        End If
    Case Asc("E")
        If gv.keyCmd.isNewDown = True Then
            gv.isShowFloatingStat = Not gv.isShowFloatingStat
        End If
    Case Asc("R")
        With gv.turret
        If .ammo.reloadTickRem = 0 And .ammo.clipAmmoRemCnt < .ammo.clipSize And .ammo.ammoRemCnt <> 0 Then
            .ammo.ammoRemCnt = .ammo.ammoRemCnt + .ammo.clipAmmoRemCnt
            .ammo.clipAmmoRemCnt = 0
            .ammo.reloadTickRem = .ammo.reloadTickCnt
        End If
        End With
    Case Asc("T")
        If gv.keyCmd.isNewDown = True Then
            gv.isShowTgtDist = Not gv.isShowTgtDist
        End If
    Case Asc(" ")
        If gv.bulletTimeTick > 0 Then
            gv.isButtletTimeOn = True
        End If
    End Select
    keyCode = keyCode
    End With
    gv.keyCmd.isNewDown = False
ExitSub:
End Sub

Private Sub Form_Load()
    Dim i As Long
    ShowAutoMode
    Form1.tmrDraw.Interval = 1000 / DRAW_PER_SEC
    lblPlayer.Caption = gv.players(gv.playerNdx).sName
    hsX_Change
    gv.ts0 = GetTickCount
    Form1.tmrDraw.Enabled = True
    Form1.hsFPM.Min = 0
    Form1.hsFPM.Max = gv.fpmSoundCnt - 1
    Form1.hsFPM.Value = Form1.hsFPM.Max \ 2
    hsFPM_Change
    cmbDifficulty.AddItem "D-入门"
    cmbDifficulty.AddItem "C-普通"
    cmbDifficulty.AddItem "B-专家"
    cmbDifficulty.AddItem "A-精英"
    cmbDifficulty.AddItem "S-王牌"
    cmbDifficulty.AddItem "SS-传奇"
    cmbDifficulty.ListIndex = 1
    cmbDifficulty_Click
    pic.Font.Size = 16
    pic.Font.Name = "黑体"
    pic.ForeColor = rgb(0, 255, 0)
    'Call InitWndProc(pic.hwnd)
End Sub

Private Sub hsX_Change()
    gv.turret.accuracyErrDiv = ((hsX.Value + 48) ^ 0.5 - 7) * 50 + 1
    
    lblFocus.Caption = Format(CalcFocus, "0.#")
End Sub

Private Sub pic_KeyDown(keyCode As Integer, Shift As Integer)
    gv.keyCmd.isNewDown = True
    gv.keyCmd.isdown = True
    
    gv.keyCmd.keyCode = keyCode
End Sub

Private Sub pic_KeyUp(keyCode As Integer, Shift As Integer)
    gv.keyCmd.isdown = False
    gv.keyCmd.keyCode = keyCode
End Sub

Private Sub Form1_KeyDown(keyCode As Integer, Shift As Integer)
    gv.keyCmd.isdown = True
    gv.keyCmd.keyCode = keyCode
End Sub

Private Sub Form1_KeyUp(keyCode As Integer, Shift As Integer)
    gv.keyCmd.isdown = False
    gv.keyCmd.keyCode = keyCode
End Sub

Private Sub tmrDraw_Timer()
    Dim i As Long
    Dim ts As Long, dt As Long
    Static cnt As Long
    ts = GetTickCount()
    dt = ts - gv.ts0
    If dt < 2 Then Exit Sub
    gv.ts0 = ts
    If dt > 100 Then dt = 100
    If gv.isButtletTimeOn = True Then
        gv.bulletTimeTick = gv.bulletTimeTick - dt
        If gv.bulletTimeTick <= 0 Then
            gv.isButtletTimeOn = False
        End If
    End If
    ProcKeyCmd
    If gv.isButtletTimeOn = True Then
        dt = dt / 3 + 1
    End If
    For i = 1 To dt
        GameStep
    Next i
    lblProj.Caption = CStr(gv.projCnt) + " ; " + Format(gv.projs(0).vecVel.X, "#.#") + ", " + Format(gv.projs(0).vecVel.Y, "#.#") + ", " + Format(gv.projs(0).vecVel.z, "#.#")
    If gv.newHitTick <> 0 Then
        gv.newHitTick = gv.newHitTick - 1
    End If
    pic.BackColor = rgb(gv.newHitTick * 16, 0, 0)
    
    If cnt Mod 3 = 0 Then
        'pic.Cls
        picHUD.Cls
    End If
    cnt = cnt + 1
    Render

    If gv.state = STATE_PLAYING Or gv.state = STATE_INIT Then
        pic.Font.Name = "黑体"
        pic.Font.Bold = False
        pic.CurrentX = pic.ScaleWidth / 2 - 90
        pic.CurrentY = 7
        pic.Font.Size = 18
        pic.Print "救援到达 "
        pic.CurrentX = pic.ScaleWidth / 2 + 20
        pic.CurrentY = 4
        pic.Font.Size = 24
        pic.ForeColor = rgb(255, 255, 0)
        pic.Print Format(gv.gameRemainTick / 1000, "0.0")
        If gv.isButtletTimeOn = True Then
            pic.CurrentX = 10
            pic.CurrentY = 10
            pic.Font.Size = 18
            pic.ForeColor = rgb(255, 155, 0)
            pic.Print Format(gv.bulletTimeTick / 1000, "子弹时间 0.0")
        End If
    End If
    
End Sub
