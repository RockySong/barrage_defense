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
   Begin VB.CheckBox chkJoy 
      BackColor       =   &H00404040&
      Caption         =   "练习(不计分)"
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
      TabIndex        =   27
      Top             =   9600
      Width           =   2415
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
      Left            =   9000
      MaskColor       =   &H00404040&
      TabIndex        =   25
      Top             =   9600
      Width           =   1417
   End
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
      Left            =   936
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   9594
      Width           =   3090
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
      TabIndex        =   2
      Top             =   117
      Width           =   14755
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
         Left            =   8160
         MaskColor       =   &H00404040&
         TabIndex        =   6
         Top             =   0
         Width           =   1417
      End
      Begin VB.HScrollBar hsX 
         Height          =   247
         Left            =   5486
         Max             =   25
         Min             =   1
         TabIndex        =   5
         Top             =   52
         Value           =   20
         Width           =   2587
      End
      Begin VB.HScrollBar hsFPM 
         Height          =   247
         LargeChange     =   4
         Left            =   5486
         Max             =   10
         TabIndex        =   3
         Top             =   364
         Value           =   5
         Width           =   2587
      End
      Begin VB.Label lblPlayer 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   120
         TabIndex        =   26
         Top             =   120
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
         Left            =   3990
         TabIndex        =   20
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "救援到达"
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
         Left            =   10680
         TabIndex        =   16
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label lblHP 
         BackColor       =   &H00000000&
         Caption         =   "-"
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
         Height          =   345
         Left            =   2730
         TabIndex        =   14
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "生命值"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   240
         Left            =   1800
         TabIndex        =   13
         Top             =   165
         Width           =   945
      End
      Begin VB.Label lblTimeRemaining 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   12120
         TabIndex        =   12
         Top             =   120
         Width           =   825
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
         Left            =   9720
         TabIndex        =   7
         Top             =   120
         Width           =   600
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
         Height          =   286
         Left            =   3978
         TabIndex        =   4
         Top             =   39
         Width           =   1300
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11349
      Top             =   1638
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8671
      Left            =   117
      MousePointer    =   2  'Cross
      ScaleHeight     =   576
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   950
      TabIndex        =   0
      Top             =   819
      Width           =   14287
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
         TabIndex        =   9
         Top             =   1989
         Width           =   1417
      End
      Begin VB.Timer tmrDraw 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   5967
         Top             =   2106
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
         Left            =   6360
         TabIndex        =   24
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label LabelAmmoTitle 
         BackColor       =   &H00000000&
         Caption         =   "剩余弹药"
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
         Left            =   2640
         TabIndex        =   23
         Top             =   120
         Width           =   1305
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
         Left            =   3930
         TabIndex        =   22
         Top             =   75
         Width           =   1995
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
         Left            =   12720
         TabIndex        =   21
         Top             =   480
         Width           =   1185
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
         Left            =   12720
         TabIndex        =   19
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblFPM 
         BackColor       =   &H00000000&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   405
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1665
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
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Width           =   1335
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
         TabIndex        =   8
         Top             =   3120
         Visible         =   0   'False
         Width           =   6480
      End
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "帮助"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   4200
      TabIndex        =   15
      Top             =   9600
      Width           =   945
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
      Left            =   120
      TabIndex        =   10
      Top             =   9600
      Width           =   600
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
    gv.isNewHit = False
    gv.myHP = 100
    cmdStart.Visible = False
    gv.gameRemainTick = CLng(150) * 1000
    Form1.lblTimeRemaining.Visible = True
    Form1.cmdUsers.Visible = False
    Form1.cmbDifficulty.Enabled = False
    Form1.chkJoy.Enabled = False
    Form1.pic.SetFocus
End Sub

Private Sub cmdUsers_Click()
    frmPlayer.Show 1
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
    xRatio = (X - halfW) / halfW * 1.25 * 1.2
    yRatio = -(Y - halfH) / halfH * 0.83 * 1.2
    
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
    gv.isFiring = True
    If Button = 2 Then
        gv.turret.fastFireFactor = 5
    Else
        gv.turret.fastFireFactor = 1#
    End If
    ProcFireCmd X, Y
    
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        If Button = 2 Then
            gv.turret.fastFireFactor = 11
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
    gv.isFiring = False
End Sub

Private Sub Form_Load()
    Dim i As Long
    'Form1.Timer1.Interval = 1000 / DRAW_PER_SEC
    Form1.tmrDraw.Interval = 1000 / DRAW_PER_SEC
    lblPlayer.Caption = gv.players(gv.playerNdx).sName
    hsX_Change
    Form1.tmrDraw.Enabled = True
    Form1.hsFPM.Min = 0
    Form1.hsFPM.Max = gv.fpmSoundCnt - 1
    Form1.hsFPM.Value = Form1.hsFPM.Max \ 2
    hsFPM_Change
    cmbDifficulty.AddItem "入门"
    cmbDifficulty.AddItem "普通"
    cmbDifficulty.AddItem "专家"
    cmbDifficulty.AddItem "精英"
    cmbDifficulty.AddItem "王牌"
    cmbDifficulty.AddItem "传奇"
    cmbDifficulty.ListIndex = 1
    cmbDifficulty_Click
End Sub

Private Sub hsX_Change()
    gv.turret.accuracyErrDiv = ((hsX.Value + 35) ^ 0.5 - 6) * 25 + 1
    lblFocus.Caption = Format(100 / gv.turret.accuracyErrDiv, "0.#")
End Sub

Private Sub pic_KeyDown(keyCode As Integer, Shift As Integer)
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
    ProcKeyCmd
    For i = 1 To tmrDraw.Interval
        GameStep
    Next i
    
    lblProj.Caption = CStr(gv.projCnt) + " ; " + Format(gv.projs(0).vecVel.X, "#.#") + ", " + Format(gv.projs(0).vecVel.Y, "#.#") + ", " + Format(gv.projs(0).vecVel.z, "#.#")
    If gv.isNewHit = True Then
        pic.BackColor = rgb(255, 0, 0)
        gv.isNewHit = 0
    Else
        pic.BackColor = 0
    End If
    pic.Cls
    Render
End Sub
