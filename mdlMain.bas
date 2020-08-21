Attribute VB_Name = "mdlMain"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As String) As Long

Public Declare Function GetWindowLong Lib "user32" Alias _
    "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias _
        "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex _
        As Long, ByVal dwNewLong As Long) As Long
        
Public Declare Function CallWindowProc Lib "user32" Alias _
        "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal _
        hwnd As Long, ByVal Msg As Long, ByVal wParam As _
        Long, ByVal lParam As Long) As Long
  
Const GWL_WNDPROC = (-4&)

Dim PrevWndProc&

Private Const WM_DESTROY = &H2


Public Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long

Public Const TME_CANCEL = &H80000000
Public Const TME_HOVER = &H1&
Public Const TME_LEAVE = &H2&
Public Const TME_NONCLIENT = &H10&
Public Const TME_QUERY = &H40000000

Private Const WM_MOUSELEAVE = &H2A3&

Public Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Public bTracking As Boolean
'----------------------------------constants-----------------------------
Public Const DEGREE_PER_ARC As Single = 57.29578
Public Const MAX_PROJ_CNT As Long = 600
Public Const MAX_TGT_CNT As Long = 16
Public Const DRAW_PER_SEC As Long = 30
Public Const PROJ_LIFETIME_SEC As Single = 4
Public Const PROJ_TAIL_CNT As Long = 8
'----------------------------------types-----------------------------
Public Type State_t
    st As Long
    val As Long
End Type

' x轴向右，y轴向上，z轴向前

Public Type Point3D_t
    X As Single
    Y As Single
    z As Single
End Type

Public Type Vector3D_t
    X As Single
    Y As Single
    z As Single
End Type

Public Type ColHSL_t
    h As Single
    S As Single
    L As Single
End Type

Public Type COlRGB_t
    R As Long
    G As Long
    B As Long
End Type

Public Type Plane_t
    A As Single
    B As Single
    C As Single
    d As Single
End Type

Public Type KeyCmd_t
    keySensitivity As Single
    isdown As Boolean
    isNewDown As Boolean
    keyCode As Long
End Type

Public Type Camera_t
    viewDist As Single  ' camera成像平面距离camera镜头坐标的长度
    vecN As Vector3D_t
    pos As Point3D_t
    plane As Plane_t
    fovDgr As Single
    WvsH As Single
End Type

Public Type Ammo_t
    clipSize As Long
    ammoRemCnt As Long
    clipAmmoRemCnt As Long
    reloadTickRem As Long
    reloadTickCnt As Long
    cooldownTickRem As Long
End Type

Public Type Turret_t
    fastFireFactor As Long
    isFiring As Boolean
    autoMode As Long ' 0=auto, 1 = semi
    isPlaying As Boolean
    accuracyErrDiv As Single
    projVelMo0 As Single
    cam As Camera_t
    fpm As Single
    sFireFile As String
    fpmCnt As Long
    fpmNdx As Long
    fpmAry(0 To 20 - 1) As Single
    tickToNextFire As Single
    tickReload As Single
    ammo As Ammo_t
    burstRem As Long
End Type

Public Type Target_t
    maxTicks As Long
    leftticks As Long
    deadTicks As Long
    distToHit As Single
    vecA As Vector3D_t
    vecV As Vector3D_t
    ptPos As Point3D_t
    hp As Single
    hp0 As Single
End Type

Public Type Projectile_t
    isShowDist As Boolean
    leftticks As Long
    vecVel As Vector3D_t
    ptPos As Point3D_t
    posHist(0 To PROJ_TAIL_CNT - 1) As Point3D_t
    histCnt As Long
    histNdx As Long
    ptPosPrev As Point3D_t
    '先假设子弹是圆球
    radius As Single
    color As COlRGB_t
End Type

Public Const STATE_STARTUP = 0
Public Const STATE_PLAYING = 1
Public Const STATE_SCORE = 2
Public Const STATE_INIT = 3
Public Type Player_t
    sName As String
    sPassword As String
    playCnt As Long
    winCnt As Long
    scoreAcc As Single
    scoreHigh As Single
End Type

Public Type Global_t
    state As Long
    tickCnt As Long
    gameRemainTick As Long
    gameTotalTick As Long
    dfcltLv(0 To 6) As Long
    keyCmd As KeyCmd_t
    isFiring As Boolean
    cam As Camera_t
    turret As Turret_t
    projCnt As Long
    
    tgtCnt As Long
    isShowTgtDist As Boolean
    isShowFloatingStat As Boolean
    tgts(0 To MAX_TGT_CNT - 1) As Target_t
    deads(0 To MAX_TGT_CNT - 1) As Target_t
    fpmSounds(0 To 100 - 1) As Long
    fpmSoundCnt As Long
    killedCnt As Long
    escapeCnt As Long
    hitCnt As Long
    newHitTick As Long
    myHP As Single
    scoreBonus As Single
    players(0 To 100 - 1) As Player_t
    playerNdx As Long
    playerCnt As Long
    highestScoreNdx As Long
    ts0 As Long
    bulletTimeTick As Long
    isButtletTimeOn As Boolean
    zoomFactor As Single
End Type
'----------------------------------globals-----------------------------
Public gv As Global_t
Public g_projs(0 To MAX_PROJ_CNT - 1) As Projectile_t
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
    (ByVal lpszName As String, _
     ByVal hModule As Long, _
     ByVal dwFlags As Long) As Long

Public Const SND_APPLICATION As Long = &H80
Public Const SND_ALIAS As Long = &H10000
Public Const SND_ALIAS_ID As Long = &H110000
Public Const SND_ASYNC As Long = &H1
Public Const SND_FILENAME As Long = &H20000
Public Const SND_LOOP As Long = &H8
Public Const SND_MEMORY As Long = &H4
Public Const SND_NODEFAULT As Long = &H2
Public Const SND_NOSTOP As Long = &H10
Public Const SND_NOWAIT As Long = &H2000
Public Const SND_PURGE As Long = &H40
Public Const SND_RESOURCE As Long = &H40004
Public Const SND_SYNC As Long = &H0

Public Sub GetBestPlayerNdx()
    Dim i As Long
    Dim highScore As Long
    Dim highNdx As Long
    highScore = -1
    For i = 1 To gv.playerCnt
        With gv.players(i - 1)
            If .scoreHigh > highScore Then
                highNdx = i - 1
                highScore = .scoreHigh
            End If
        End With
    Next i
    gv.highestScoreNdx = highNdx
End Sub

Public Sub LoadPlayers()
    Dim X As Control
    Dim sName As String
    Dim playerCnt As Long
    Dim sRead As String * 64, sVal As String, sFile As String, sKey As String, sWr As String
    Dim read_ok As Long, i As Long
    sFile = App.Path & "\scores.cfg"
    playerCnt = 20
    read_ok = GetPrivateProfileString("Main", "Me", "deadbeef", sRead, 256, sFile)
    For i = 1 To playerCnt
        sKey = "Player" & CStr(i)
        With gv.players(i - 1)
        read_ok = GetPrivateProfileString("Names", sKey, "deadbeef", sRead, 256, sFile)
        Form1.lblProj.Caption = sRead
        .sName = Form1.lblProj.Caption
        read_ok = GetPrivateProfileString("Passwords", sKey, "deadbeef", sRead, 256, sFile)
        Form1.lblProj.Caption = sRead
        .sPassword = Form1.lblProj.Caption
        read_ok = GetPrivateProfileString("ScoresAcc", sKey, "deadbeef", sRead, 256, sFile)
        .scoreAcc = CLng(Left(sRead, read_ok))
        read_ok = GetPrivateProfileString("ScoresHigh", sKey, "deadbeef", sRead, 256, sFile)
        .scoreHigh = CLng(Left(sRead, read_ok))
        read_ok = GetPrivateProfileString("Battles", sKey, "deadbeef", sRead, 256, sFile)
        .playCnt = CLng(Left(sRead, read_ok))
        read_ok = GetPrivateProfileString("Wins", sKey, "deadbeef", sRead, 256, sFile)
        .winCnt = CLng(Left(sRead, read_ok))
        End With
    Next i
    read_ok = GetPrivateProfileString("LastPlayer", "Index", "deadbeef", sRead, 256, sFile)
    gv.playerNdx = CLng(Left(sRead, read_ok))
    gv.playerCnt = playerCnt
    GetBestPlayerNdx
    Form1.lblPlayer.Caption = gv.players(gv.playerNdx).sName
End Sub

Public Sub UpdateCurrentPlayer(ndx As Long)
    WritePrivateProfileString "LastPlayer", "Index", CStr(ndx), App.Path & "\scores.cfg"
End Sub

Public Sub UpdatePlayer(ndx As Long)
    Dim sVal As String
    Dim sKey As String
    sKey = "Player" & CStr(ndx + 1)
    sFile = App.Path & "\scores.cfg"
    With gv.players(ndx)
        write1 = WritePrivateProfileString("Names", sKey, .sName, sFile)
        write1 = WritePrivateProfileString("ScoresHigh", sKey, CStr(.scoreHigh), sFile)
        write1 = WritePrivateProfileString("ScoresAcc", sKey, CStr(.scoreAcc), sFile)
        write1 = WritePrivateProfileString("Battles", sKey, CStr(.playCnt), sFile)
        write1 = WritePrivateProfileString("Wins", sKey, CStr(.winCnt), sFile)
    End With
    GetBestPlayerNdx
    Form1.lblPlayer.Caption = gv.players(gv.playerNdx).sName
End Sub

Public Function Hue2RGB(v1 As Single, v2 As Single, vH As Single) As Single
    If vH < 0 Then vH = vH + 1
    If vH > 1 Then vH = vH - 1
    Hue2RGB = v1
    If 6# * vH < 1 Then Hue2RGB = v1 + (v2 - v1) * 6# * vH
    If 2# * vH < 1 Then Hue2RGB = v2
    If 3# * vH < 2 Then Hue2RGB = v1 + (v2 - v1) * ((2# / 3#) - vH) * 6#
End Function

Public Sub HSL2RGB(hsl As ColHSL_t, rgb As COlRGB_t)
    Dim v1 As Single, v2 As Single
    With hsl
        If .S = 0 Then
            rgb.R = .L * 255 + 0.5
            rgb.G = .L
            rgb.B = .L
        Else
            If .L < 0.5 Then v2 = .L * (1 + .S) Else v2 = .L + .S - .L * .S
            v1 = .L * 2 - v2
            rgb.R = Hue2RGB(v1, v2, .h + 1# / 3#)
            rgb.G = Hue2RGB(v1, v2, .h)
            rgb.B = Hue2RGB(v1, v2, .h - 1# / 3#)
        End If
    End With
End Sub

Public Sub GetUnitVector(vecIn As Vector3D_t, vecOut As Vector3D_t)
    Dim mo As Single
    mo = Sqr(vecIn.X * vecIn.X + vecIn.Y * vecIn.Y + vecIn.z * vecIn.z)
    vecOut.X = vecIn.X / mo: vecOut.Y = vecIn.Y / mo: vecOut.z = vecIn.z / mo
End Sub

Public Sub SetupPlane(pt As Point3D_t, vecNorm As Vector3D_t, plane As Plane_t)
    plane.A = vecNorm.X
    plane.B = vecNorm.Y
    plane.C = vecNorm.z
    plane.d = -1 * (vecNorm.X * pt.X + vecNorm.Y * pt.Y + vecNorm.z * pt.z)
End Sub


Public Function M3D_CalcDotPlaneDistance(pt As Point3D_t, plane As Plane_t) As Single
    Dim ret As Single
    ret = Abs(plane.A * pt.X + plane.B * pt.Y + plane.C * pt.z + plane.d)
    ' 平面的A, B, C是按单位法向量处理过的，故为1
    M3D_CalcDotPlaneDistance = ret
End Function

Public Function M3D_CalcDotDotDistance(pt1 As Point3D_t, pt2 As Point3D_t) As Single
    Dim d As Single
    d = (pt1.X - pt2.X) ^ 2 + (pt1.Y - pt2.Y) ^ 2 + (pt1.z - pt2.z) ^ 2
    M3D_CalcDotDotDistance = Sqr(d)
End Function

Public Function VectorDot(v1 As Vector3D_t, v2 As Vector3D_t) As Single
    VectorDot = v1.X * v2.X + v1.Y * v2.Y + v1.z * v2.z
End Function

Public Function ACos(X As Single) As Single
    ACos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function

Public Function CalcVectorAngle(vec1 As Vector3D_t, vec2 As Vector3D_t) As Single
    Dim v1unit As Vector3D_t, v2unit As Vector3D_t
    Dim cosVal As Single
    GetUnitVector vec1, v1unit
    GetUnitVector vec2, v2unit
    cosVal = v1unit.X * v2unit.X + v1unit.Y * v2unit.Y + v1unit.z * v2unit.z
    If cosVal >= 1 Then
        CalcVectorAngle = 0
    ElseIf cosVal <= -1 Then
        CalcVectorAngle = 180 / degree_perarc
    Else
        CalcVectorAngle = ACos(cosVal) * DEGREE_PER_ARC
    End If
End Function

Public Sub SetupCamera(cam As Camera_t)
    GetUnitVector cam.vecN, cam.vecN
    SetupPlane cam.pos, cam.vecN, cam.plane
End Sub

Public Sub MakeVectorFromUnitVector(vecUnit As Vector3D_t, mo As Single, vecOut As Vector3D_t)
    vecOut.X = vecUnit.X * mo: vecOut.Y = vecUnit.Y * mo: vecOut.z = vecUnit.z * mo
End Sub

Public Sub Turret_FpmToTickPerFire(turret As Turret_t)
    turret.tickReload = turret.fpm / 60# * DRAW_PER_SEC
End Sub

Public Sub DrawAmmo(pic As PictureBox, Optional barLen As Long = 150)
    Dim i As Long
    Dim y0 As Long
    Dim col As Long
    Dim X As Long
    With gv.turret.ammo
    If .reloadTickRem = 0 Then
        col = rgb(0, 205, 0)
        i = .clipAmmoRemCnt * barLen / .clipSize
    Else
        col = rgb(255, 245, 64)
        i = (.reloadTickCnt - .reloadTickRem) * barLen / .reloadTickCnt
    End If
    
    y0 = Form1.lblAmmo.Top + Form1.lblAmmo.Height + 3
    X = Form1.LabelAmmoTitle.Left
    
    pic.Line (X, y0 + 6 / 2)-(X + barLen, y0 + 6 / 2), rgb(225, 255, 225)
    pic.Line (X, y0)-(X + i, y0 + 6), col, BF
    End With
End Sub

Private Sub DrawHPBar(pic As PictureBox, Optional barLen As Long = 100)
    Dim X As Long, y0 As Long
    X = Form1.lblHPTitle.Left
    y0 = Form1.lblHPTitle.Top + Form1.lblHPTitle.Height
    pic.DrawStyle = 0
    col = rgb(255, 160, 225)
    pic.Line (X, y0 + 6 / 2)-(X + barLen, y0 + 6 / 2), rgb(225, 255, 225)
    pic.Line (X, y0)-(X + gv.myHP, y0 + 6), col, BF
End Sub
Public Sub Render()
    Dim prjPt As Point3D_t, prjPt2 As Point3D_t, prjPtUnzoomed As Point3D_t
    Dim vecToCam As Vector3D_t
    Dim angle As Single
    Dim aimPos As Point3D_t
    Dim i As Long, j As Long, k As Long, k2 As Long, n As Long
    Dim dist As Single
    Dim isNewHighScore As Boolean
    Dim oldHighScore As Long
    Dim cam As Camera_t
    Dim xMax As Single
    Dim yMax As Single
    Dim rasterX As Long, rasterX2 As Long
    Dim rasterY As Long, rasterY2 As Long
    Dim score As Long
    Dim txtCol As Long
    Dim rasterW As Long, rasterH As Long
    Dim projR As Single
    Dim col As Long
    Dim bri As Long
    Dim X As Long, Y As Long
    Dim pic As PictureBox
    Dim fX As Single
    Dim intProjR As Long
    Dim focal As Single
    Dim isArc As Boolean
    Set pic = Form1.pic
    If Form1.chkArc.Value <> 0 Then isArc = True Else isArc = False
    Form1.lblAmmo.Caption = gv.turret.ammo.clipAmmoRemCnt & "/" & gv.turret.ammo.ammoRemCnt
    If gv.turret.ammo.reloadTickRem <> 0 Then
    Form1.lblAmmo.ForeColor = rgb(255, 255, 0)
    Else
        If gv.turret.ammo.clipAmmoRemCnt > 0 Then
            Form1.lblAmmo.ForeColor = rgb(0, 255, 0)
        Else
            Form1.lblAmmo.ForeColor = rgb(255, 0, 0)
        End If
    End If
    
    If gv.state = STATE_PLAYING Then
        With gv.turret.ammo
            If gv.gameRemainTick = 0 Then
                gv.state = STATE_SCORE
            End If
        End With
    End If
    
    If gv.state = STATE_SCORE Then
        If True Then 'Form1.chkJoy.Value = 0 Then
            score = (gv.killedCnt + gv.killedCnt * gv.killedCnt / (gv.killedCnt + gv.escapeCnt + 0.0001) + gv.scoreBonus) _
                * (2 ^ Form1.cmbDifficulty.ListIndex)
            score = score * (1 - Form1.chkJoy.Value)
            score = score * CSng(gv.gameTotalTick - gv.gameRemainTick) / gv.gameTotalTick
            If Form1.chkJoy.Value = 0 Then
                With gv.players(gv.playerNdx)
                    .playCnt = .playCnt + 1
                    If gv.myHP > 0 Then
                        score = score * 2
                        .winCnt = .winCnt + 1
                    Else
                    End If
                    .scoreAcc = .scoreAcc + score
                    .playCnt = .playCnt + 1
                    If score > .scoreHigh Then
                        isNewHighScore = True
                        oldHighScore = .scoreHigh
                        .scoreHigh = score
                    Else
                        isNewHighScore = False
                    End If
                    UpdatePlayer gv.playerNdx
                End With
            End If
            
            With gv.players(gv.playerNdx)
            If gv.myHP > 0 And gv.gameRemainTick = 0 Then
                PlaySound App.Path & "\victory.wav", 0, SND_ASYNC Or SND_FILENAME Or SND_LOOP
                Form1.lblScore.BackColor = rgb(0, 50, 0)
                Form1.lblScore.ForeColor = rgb(150, 255, 150)
                Form1.lblScore.Caption = "得救了！我们爱你！" & gv.players(gv.playerNdx).sName & Chr$(13) & Chr$(10)
            Else
                Form1.lblScore.BackColor = rgb(50, 0, 0)
                Form1.lblScore.ForeColor = rgb(255, 150, 150)
                Form1.lblScore.Caption = "请再接再厉吧！" & Chr$(13) & Chr$(10)
                PlaySound App.Path & "\defeat.wav", 0, SND_ASYNC Or SND_FILENAME Or SND_LOOP
            End If
            If isNewHighScore = True Then
                Form1.lblScore.Caption = Form1.lblScore.Caption & "新高分！" & CStr(oldHighScore) & " -> " & CStr(score) & Chr$(13) & Chr$(10)
            End If
            Form1.lblScore.Caption = Form1.lblScore.Caption & "歼敌 " & _
                gv.killedCnt & " / " & (gv.killedCnt + gv.escapeCnt) & ", 分数: " & CStr(score)
            End With
        Else
                Form1.lblScore.BackColor = rgb(50, 50, 50)
                Form1.lblScore.ForeColor = rgb(250, 255, 250)
                Form1.lblScore.Caption = "练习结束" & Chr$(13) & Chr$(10) & "接下来挑战一下自己吧" & Chr$(13) & Chr$(10)
                Form1.lblScore.Caption = Form1.lblScore.Caption & "歼敌 " & _
                gv.killedCnt & " / " & (gv.killedCnt + gv.escapeCnt)
        End If
        If Form1.lblScore.Visible = False Then
            Form1.lblScore.Visible = True
            Form1.cmdNew.Visible = False
            Form1.cmdStart.Visible = True
            Form1.cmdUsers.Visible = True
            Form1.chkJoy.Enabled = True
            Form1.cmbDifficulty.Enabled = True
        End If
        gv.state = STATE_INIT
        'gv.gameRemainTick = 0
    End If
    
    Form1.lblStat = CStr(gv.killedCnt) & "/" & CStr(gv.killedCnt + gv.escapeCnt)
    cam = gv.cam
    focal = cam.viewDist - cam.pos.z
    cam.pos.X = 0: cam.pos.Y = 0: cam.pos.z = 0
    cam.vecN.X = 0: cam.vecN.Y = 0: cam.vecN.z = 1
    cam.viewDist = 1
    SetupCamera cam
    n = MAX_PROJ_CNT - 1
    xMax = cam.viewDist * Tan(cam.fovDgr / DEGREE_PER_ARC / 2) / 2
    yMax = xMax / cam.WvsH
    rasterW = Form1.pic.ScaleWidth
    rasterH = Form1.pic.ScaleHeight
    If Form1.lblScore.Visible = True Then GoTo PostRender
    'render hp
    If gv.state = STATE_PLAYING Then
        DrawHPBar Form1.picHUD
        DrawAmmo Form1.picHUD
    End If
    With gv
        pic.DrawStyle = 0
        pic.DrawWidth = 2
        n = MAX_TGT_CNT - 1
        pic.Font.Size = 10
        pic.Font.Bold = True
        pic.Font.Name = "Consolas"
        For i = 0 To n
            With .tgts(i)
                If .leftticks <> 0 Or .deadTicks <> 0 Then
                    vecToCam.X = .ptPos.X - cam.pos.X
                    vecToCam.Y = .ptPos.Y - cam.pos.Y
                    vecToCam.z = .ptPos.z - cam.pos.z
                    dist = M3D_CalcDotPlaneDistance(.ptPos, cam.plane)
                        prjPt.X = .ptPos.X * focal / dist / 2
                        prjPt.Y = .ptPos.Y * focal / dist / 2
                        If .ptPos.z > cam.viewDist Then
                            angle = CalcVectorAngle(vecToCam, cam.vecN)
                            ' now only supports (0,0,0) position (0,0,1) norm vector camera
                            prjPt.X = .ptPos.X * focal / dist / 2 * gv.zoomFactor
                            prjPt.Y = .ptPos.Y * focal / dist / 2 * gv.zoomFactor
                            If Abs(prjPt.X) < xMax And Abs(prjPt.Y) < yMax Then
                                rasterX = rasterW / 2 * prjPt.X / xMax + rasterW / 2
                                rasterY = -rasterH / 2 * prjPt.Y / yMax + rasterH / 2
                                
                                If dist - cam.viewDist > 1 Then
                                    projR = 15 * 60 / ((dist - cam.viewDist))
                                Else
                                    projR = 15 * 60
                                End If
                                projR = projR * gv.zoomFactor
                                If projR < 1 Then
                                    bri = projR * 255
                                    projR = 1
                                    If bri < 35 Then
                                        .leftticks = 0
                                    End If
                                Else
                                    bri = 255
                                End If
                                If .leftticks > 0 Then
                                    bri = bri / (5 - CSng(.leftticks) / .maxTicks * 4)
                                    col = rgb(bri, bri / (2 - (.hp / .hp0)), bri / (1 + 3 * (.hp0 - .hp) / .hp0))
                                Else
                                    bri = bri * (120 + .deadTicks) / 1200
                                    col = rgb(bri, 0, 0)
                                End If
                                If .distToHit < 255 And .leftticks > 0 Then
                                    If .leftticks Mod 200 < 100 Then
                                        col = rgb(255, 255, 255)
                                    End If
                                End If
                                If .leftticks > 0 Then
                                    pic.Circle (rasterX, rasterY), projR, col
                                    pic.Circle (rasterX, rasterY), projR / 3, col
                                    If gv.isShowTgtDist = True Then
                                        pic.CurrentX = rasterX + projR + 2
                                        pic.CurrentY = rasterY - 5
                                        If .distToHit < 255 * gv.zoomFactor Then
                                            txtCol = CLng(.distToHit)
                                            pic.ForeColor = rgb(255, txtCol, 0)
                                            pic.Print Format(.distToHit, "0")
                                        End If
                                    End If
                                Else
                                    pic.Line (rasterX - projR, rasterY - projR)-(rasterX + projR, rasterY + projR), col
                                    pic.Line (rasterX + projR, rasterY - projR)-(rasterX - projR, rasterY + projR), col
                                End If
                                'pic.Line (rasterX - projR, rasterY - projR)-(rasterX + projR, rasterY + projR), col, BF
                            End If
                        End If
                    End If
            End With
        Next i
                
        pic.DrawWidth = 1
        '=======================================================
        n = MAX_PROJ_CNT - 1
        pic.ForeColor = rgb(255, 255, 0)
        For i = 0 To n
            With g_projs(i)
                If .leftticks <> 0 Then
                    vecToCam.X = .ptPos.X - cam.pos.X
                    vecToCam.Y = .ptPos.Y - cam.pos.Y
                    vecToCam.z = .ptPos.z - cam.pos.z
                    If .ptPos.z > cam.viewDist Then
                        'angle = CalcVectorAngle(vecToCam, cam.vecN)
                        ' now only supports (0,0,0) position (0,0,1) norm vector camera

                        dist = M3D_CalcDotPlaneDistance(.ptPos, cam.plane)
                        prjPt.X = .ptPos.X * focal / dist / 2 * gv.zoomFactor
                        prjPt.Y = .ptPos.Y * focal / dist / 2 * gv.zoomFactor
                        If Abs(prjPt.X) < xMax And Abs(prjPt.Y) < yMax Then
                            rasterX = rasterW / 2 * prjPt.X / xMax + rasterW / 2
                            rasterY = -rasterH / 2 * prjPt.Y / yMax + rasterH / 2
                            
                            If dist - cam.viewDist > 1 Then
                                projR = 33 / ((dist - cam.viewDist)) ^ 0.78
                            Else
                                projR = 33
                            End If
                            If projR < 1 Then
                                bri = projR * 255
                                If bri < 32 Then
                                    .leftticks = 0
                                End If
                                bri = projR ^ 0.5 * 255
                                'projR = 1
                            Else
                                bri = 255
                            End If
                            intProjR = Int(projR)
                            If intProjR > 0 Then
                               pic.Line (rasterX - intProjR, rasterY - intProjR)-(rasterX + intProjR, rasterY + intProjR), rgb(bri, bri, 0), BF
                            Else
                                pic.PSet (rasterX, rasterY), rgb(bri, bri, 0)
                            End If
                            bri = 255 * (projR - intProjR)
                            If bri > 48 Then
                                intProjR = intProjR + 1
                                pic.Line (rasterX - intProjR, rasterY - intProjR)-(rasterX + intProjR, rasterY + intProjR), rgb(bri, bri, 0), B
                            End If
                            If isArc = True And projR > 0.15 Then
                                ' >>> ----------------绘制弹道曲线-------------------
                                pic.DrawWidth = 1
                                If .histCnt < PROJ_TAIL_CNT Then
                                    k = 0
                                Else
                                    k = (.histNdx + 1) Mod PROJ_TAIL_CNT
                                End If
                                For j = 1 To .histCnt - 1
                                    k2 = (k + 1) Mod PROJ_TAIL_CNT
                                    If .posHist(k).z > cam.viewDist Then
                                        dist = M3D_CalcDotPlaneDistance(.posHist(k), cam.plane)
                                        prjPt.X = .posHist(k).X * focal / dist / 2 * gv.zoomFactor
                                        prjPt.Y = .posHist(k).Y * focal / dist / 2 * gv.zoomFactor
                                        rasterX = rasterW / 2 * prjPt.X / xMax + rasterW / 2
                                        rasterY = -rasterH / 2 * prjPt.Y / yMax + rasterH / 2
                                        dist = M3D_CalcDotPlaneDistance(.posHist(k2), cam.plane)
                                        prjPt2.X = .posHist(k2).X * focal / dist / 2 * gv.zoomFactor
                                        prjPt2.Y = .posHist(k2).Y * focal / dist / 2 * gv.zoomFactor
                                        rasterX2 = rasterW / 2 * prjPt2.X / xMax + rasterW / 2
                                        rasterY2 = -rasterH / 2 * prjPt2.Y / yMax + rasterH / 2
                                        bri = j * 255 / .histCnt
                                        pic.Line (rasterX, rasterY)-(rasterX2, rasterY2), rgb(bri, bri, 0)
                                    End If
                                    k = k2
                                Next j
                                ' <<<
                            End If
                            
                            If .isShowDist = True Then
                                If PROJ_LIFETIME_SEC * 1000 - .leftticks <= 3000 Then
                                    pic.Print CLng(.ptPos.z)
                                End If
                            End If
                        End If
                    End If
                    If isArc = True And projR > 0.15 Then
                        .histNdx = .histNdx + 1
                        If .histNdx >= PROJ_TAIL_CNT Then .histNdx = 0
                        If .histCnt < PROJ_TAIL_CNT Then .histCnt = .histCnt + 1
                        .posHist(.histNdx) = .ptPos
                    End If
                End If
            End With
        Next i
        'render aim helper
        fX = CalcFocus + 4
        X = -1
        Y = -1
        If gv.state = STATE_PLAYING Then
            pic.DrawStyle = 0
            For i = 2 To 4
                aimPos = gv.turret.cam.pos
                aimPos.X = aimPos.X + .turret.cam.vecN.X * (30 + 3 ^ i)
                aimPos.Y = aimPos.Y + .turret.cam.vecN.Y * (30 + 3 ^ i)
                aimPos.z = aimPos.z + .turret.cam.vecN.z * (30 + 3 ^ i)
                dist = M3D_CalcDotPlaneDistance(aimPos, cam.plane)
                prjPt.X = aimPos.X * focal / dist / 2 * gv.zoomFactor
                prjPt.Y = aimPos.Y * focal / dist / 2 * gv.zoomFactor

                If Abs(prjPt.X) < xMax And Abs(prjPt.Y) < yMax Then
                    rasterX = rasterW / 2 * prjPt.X / xMax + rasterW / 2
                    rasterY = -rasterH / 2 * prjPt.Y / yMax + rasterH / 2
                    If dist - cam.viewDist > 1 Then
                        projR = fX * 80 / ((dist - cam.viewDist))
                    Else
                        projR = fX * 80
                    End If
                    projR = projR * gv.zoomFactor
                    bri = i * 51
                    X = rasterX
                    Y = rasterY
                    pic.Line (rasterX - projR, rasterY - projR)-(rasterX + projR, rasterY + projR), rgb(gv.turret.autoMode * 255, bri, 0), B
                End If
            Next i
            If gv.isShowFloatingStat = True Then
                If X = -1 Then
                    pic.CurrentX = pic.ScaleWidth / 2
                    pic.CurrentY = pic.ScaleHeight / 2
                End If
                X = pic.CurrentX + 30
                Y = pic.CurrentY
                pic.Font.Size = 14
                pic.Font.Bold = False
                If .myHP > 40 Or .ts0 Mod 400 < 250 Then
                    pic.ForeColor = rgb(255, 158, 235) 'rgb(255, .turret.ammo.clipAmmoRemCnt * 255 / .turret.ammo.clipSize, 0)
                Else
                    pic.ForeColor = rgb(255, 0, 0) 'rgb(255, .turret.ammo.clipAmmoRemCnt * 255 / .turret.ammo.clipSize, 0)
                End If
                pic.CurrentX = X
                pic.CurrentY = pic.CurrentY - 33
                pic.Print "HP: " & Format(.myHP, "0")

                
                If .turret.ammo.clipAmmoRemCnt * 100# / .turret.ammo.clipSize > 25 Or .ts0 Mod 400 < 250 Then
                    If .turret.ammo.reloadTickRem <> 0 Then
                        pic.ForeColor = rgb(255, 255, 0) 'rgb(255, .turret.ammo.clipAmmoRemCnt * 255 / .turret.ammo.clipSize, 0)
                    Else
                        pic.ForeColor = rgb(0, 255, 0) 'rgb(255, .turret.ammo.clipAmmoRemCnt * 255 / .turret.ammo.clipSize, 0)
                    End If
                Else
                    pic.ForeColor = rgb(0, 8 * .turret.ammo.clipAmmoRemCnt * 100# / .turret.ammo.clipSize, 0)
                End If
                If .turret.ammo.clipAmmoRemCnt = 0 And .turret.ammo.ammoRemCnt = 0 Then pic.ForeColor = rgb(255, 0, 0)
                pic.CurrentX = X
                If .turret.ammo.reloadTickRem = 0 Then
                    pic.Print CStr(.turret.ammo.clipAmmoRemCnt) & " / " & CStr(.turret.ammo.ammoRemCnt)
                Else
                    pic.Print "RLD..."
                End If
                
                pic.ForeColor = rgb(255, 255, 0)
                pic.CurrentX = X
                pic.Print CStr(.turret.fpm) & " RPM"
                'pic.ForeColor = rgb(0, 255, 255)
                'pic.CurrentX = X
                'pic.Print Format(100 / gv.turret.accuracyErrDiv, "0.#")
                pic.Font.Size = 10
            End If
        End If
        
        If gv.zoomFactor > 1 Then
            n = MAX_TGT_CNT - 1
            pic.Line (0, rasterH * 3 / 4)-(rasterW / 4, rasterH - 2), rgb(44, 44, 44), BF
            pic.DrawWidth = 2
            pic.Line (0, rasterH * 3 / 4)-(rasterW / 4, rasterH - 2), rgb(144, 144, 144), B
            pic.DrawWidth = 1
            For i = 0 To n
                 With .tgts(i)
                     If .leftticks <> 0 Or .deadTicks <> 0 Then
                         vecToCam.X = .ptPos.X - cam.pos.X
                         vecToCam.Y = .ptPos.Y - cam.pos.Y
                         vecToCam.z = .ptPos.z - cam.pos.z
                         dist = M3D_CalcDotPlaneDistance(.ptPos, cam.plane)
                             If .ptPos.z > cam.viewDist Then
                                 angle = CalcVectorAngle(vecToCam, cam.vecN)
                                 ' now only supports (0,0,0) position (0,0,1) norm vector camera
                                 prjPt.X = .ptPos.X * focal / dist / 2
                                 prjPt.Y = .ptPos.Y * focal / dist / 2
                                 If Abs(prjPt.X) < xMax And Abs(prjPt.Y) < yMax Then
                                     rasterX = (rasterW / 2 * prjPt.X / xMax + rasterW / 2) / 4
                                     rasterY = (-rasterH / 2 * prjPt.Y / yMax + rasterH / 2) / 4 + rasterH * 3 / 4
                                     
                                     If dist - cam.viewDist > 1 Then
                                         projR = 15 * 60 / ((dist - cam.viewDist))
                                     Else
                                         projR = 15 * 60
                                     End If
                                     projR = projR / 4
                                     If projR < 1 Then
                                         bri = projR * 255
                                         projR = 1
                                         If bri < 35 Then
                                             .leftticks = 0
                                         End If
                                     Else
                                         bri = 255
                                     End If
                                     If .leftticks > 0 Then
                                         bri = bri / (5 - CSng(.leftticks) / .maxTicks * 4)
                                         col = rgb(bri, bri / (2 - (.hp / .hp0)), bri / (1 + 3 * (.hp0 - .hp) / .hp0))
                                     Else
                                         bri = bri * (120 + .deadTicks) / 1200
                                         col = rgb(bri, 0, 0)
                                     End If
                                     If .distToHit < 255 And .leftticks > 0 Then
                                         If .leftticks Mod 200 < 100 Then
                                             col = rgb(255, 255, 255)
                                         End If
                                     End If
                                     If .leftticks > 0 Then
                                         pic.Circle (rasterX, rasterY), projR, col
                                         pic.Circle (rasterX, rasterY), projR / 3, col
                                         If gv.isShowTgtDist = True Then
                                             pic.CurrentX = rasterX + projR + 2
                                             pic.CurrentY = rasterY - 5
                                             If .distToHit < 255 Then
                                                 txtCol = CLng(.distToHit)
                                                 pic.ForeColor = rgb(255, txtCol, 0)
                                                 pic.Print Format(.distToHit, "0")
                                             End If
                                         End If
                                     Else
                                         pic.Line (rasterX - projR, rasterY - projR)-(rasterX + projR, rasterY + projR), col
                                         pic.Line (rasterX + projR, rasterY - projR)-(rasterX - projR, rasterY + projR), col
                                     End If
                                     'pic.Line (rasterX - projR, rasterY - projR)-(rasterX + projR, rasterY + projR), col, BF
                                 End If
                             End If
                         End If
                 End With
             Next i
        End If
PostRender:
        If gv.zoomFactor > 1 Then
            pic.CurrentX = 30
            pic.CurrentY = 30
            pic.Font.Size = 24
            pic.Font.Name = "黑体"
            pic.ForeColor = rgb(0, 128, 255)
            pic.Print Format(gv.zoomFactor, "#X倍放大")
        End If

        If True Then 'gv.state = STATE_PLAYING Or gv.state = STATE_INIT Then
            pic.Font.Name = "黑体"
            pic.ForeColor = rgb(255, 255, 0)
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
        
        

    End With
End Sub
Public Function CalcFocus() As Single
    Dim f As Single
    f = (81 / gv.turret.accuracyErrDiv) ^ 0.55 * 8 - 2
    CalcFocus = f
End Function
Public Sub FireTurret(cam As Camera_t, Optional cnt As Long = 1)
    Dim ndx As Long, i As Long
    Dim divVal As Single
    Dim n As Long
    Dim velOfs As Single
    With gv
    If gv.turret.burstRem <> 0 Then
        gv.turret.burstRem = gv.turret.burstRem - 1
        If .turret.tickToNextFire > 1 Then
            .turret.tickToNextFire = .turret.tickToNextFire - 1
        End If
    Else
        .turret.tickToNextFire = .turret.tickToNextFire - 1
        If .turret.tickToNextFire > 0 Then Exit Sub
        .turret.tickToNextFire = .turret.tickToNextFire + .turret.tickReload
        If gv.turret.burstRem = 0 And gv.turret.fastFireFactor > 1 Then
            gv.turret.burstRem = gv.turret.fastFireFactor
        End If
    End If
    
    cnt = ConsumeTurretAmmo(gv.turret, cnt)
    If cnt < 1 Then
        gv.turret.isFiring = False
        Exit Sub
    End If
    divVal = gv.turret.accuracyErrDiv ^ 1.7 / 35 + 1.3
    velOfs = gv.turret.projVelMo0 / divVal
    n = MAX_PROJ_CNT - 1
    For i = 1 To cnt
        For ndx = 0 To n
            If g_projs(ndx).leftticks = 0 Then
                .turret.isFiring = True
                With g_projs(ndx)
                    .color.R = 0
                    .color.G = 255
                    .color.B = 0
                    If gv.turret.autoMode = 1 And i = 1 Then .isShowDist = True Else .isShowDist = False
                    .leftticks = PROJ_LIFETIME_SEC * 1000
                    .ptPos = cam.pos
                    .histCnt = 1
                    .histNdx = 0
                    .posHist(.histNdx) = .ptPos
                    MakeVectorFromUnitVector cam.vecN, gv.turret.projVelMo0, .vecVel
                    .vecVel.X = .vecVel.X * (1 + Rnd / divVal - 1 / 2 / divVal) + Rnd * velOfs - velOfs / 2
                    .vecVel.Y = .vecVel.Y * (1 + Rnd / divVal - 1 / 2 / divVal) + Rnd * velOfs - velOfs / 2
                    .vecVel.z = .vecVel.z * (1 + Rnd / divVal - 1 / 2 / divVal) '+ Rnd * velOfs
                End With
                Exit For
            End If
        Next ndx
        If ndx = MAX_PROJ_CNT Then Exit For
    Next i
    End With
End Sub

Public Sub ResetTurretAmmo(turret As Turret_t)
    Form1.lblScore.Visible = False
    With turret.ammo
        .cooldownTickRem = 0
        .reloadTickRem = 0
        .ammoRemCnt = 3000 - Form1.cmbDifficulty.ListIndex * 250
        .clipAmmoRemCnt = .clipSize
        .reloadTickCnt = 1150
        If Form1.chkJoy.Value <> 0 Then
            .ammoRemCnt = .clipSize * 10
        End If
    End With
    gv.bulletTimeTick = CLng(10) * 1000 * (1 + Form1.chkJoy * 50)
End Sub

Public Sub ChangeFPM()
    gv.turret.fpm = gv.fpmSounds(Form1.hsFPM.Value)
    gv.turret.tickToNextFire = 1
    gv.turret.tickReload = 1000# * 60 / CSng(gv.turret.fpm)
    If gv.isButtletTimeOn = False Then
        gv.turret.sFireFile = GetFpmSoundFile(gv.turret.fpm)
    Else
        gv.turret.sFireFile = GetFpmSoundFile(gv.turret.fpm / 3.084)
    End If
    
    If gv.turret.isFiring = True Then
        PlaySound vbNullString, 0, 0
        PlaySound gv.turret.sFireFile, 0, SND_ASYNC Or SND_FILENAME Or SND_LOOP
    End If
End Sub

Public Sub ReloadTurret(turret As Turret_t)
    With turret.ammo
        If .reloadTickRem <> 0 Then
            .reloadTickRem = .reloadTickRem - 1
            If .reloadTickRem = 0 Then
                If .ammoRemCnt > .clipSize Then
                    .clipAmmoRemCnt = .clipSize
                    .ammoRemCnt = .ammoRemCnt - .clipSize
                Else
                    .clipAmmoRemCnt = .ammoRemCnt
                    .ammoRemCnt = 0
                End If
            End If
        End If
    End With
End Sub

Public Function ConsumeTurretAmmo(turret As Turret_t, cnt As Long) As Long
    ConsumeTurretAmmo = 0
    With turret.ammo
        If .reloadTickRem = 0 Then
            If .clipAmmoRemCnt >= cnt Then
                .clipAmmoRemCnt = .clipAmmoRemCnt - cnt
                ConsumeTurretAmmo = cnt
            Else
                ConsumeTurretAmmo = .clipAmmoRemCnt
                .clipAmmoRemCnt = 0
            End If
            
            If .clipAmmoRemCnt = 0 Then
                If .ammoRemCnt > 0 Then
                    .reloadTickRem = .reloadTickCnt
                End If
            End If
        End If
    End With
End Function

Public Sub SpawnTarget()
    Dim ptPos As Point3D_t
    Dim i As Long, n As Long
    
    ptPos.z = 475 + Rnd * 600
    ptPos.X = (ptPos.z * Rnd - ptPos.z / 2) / 3
    ptPos.Y = (ptPos.z * Rnd - ptPos.z * 3 / 4) / 2 - 150
    
    n = MAX_TGT_CNT - 1
    
    For i = 0 To n
        With gv.tgts(i)
        If .leftticks = 0 And .deadTicks = 0 Then
            .maxTicks = CLng(11000 + Rnd * 16500) * 5
            .leftticks = .maxTicks / 1.3
            .ptPos = ptPos
            .vecV.X = Rnd * 1
            .vecV.Y = 150 + Rnd * 50
            .vecV.z = Rnd * 1
            .vecA.X = 0: .vecA.Y = 0: .vecA.z = 0
            .hp = 100 + Rnd * 400 + Form1.cmbDifficulty.ListIndex * 75
            .hp0 = .hp
            .distToHit = M3D_CalcDotDotDistance(.ptPos, gv.turret.cam.pos)
            gv.tgtCnt = gv.tgtCnt + 1
            Exit For
        End If
        End With
    Next i
    
End Sub

' 返回被命中的次数
Public Function CheckHit(tgt As Target_t) As Long
    CheckHit = 0
    Dim dist As Single
    Dim m As Long, j As Long
    m = MAX_PROJ_CNT - 1
    For j = 0 To m
        With g_projs(j)
            If .leftticks <> 0 Then
                dist = M3D_CalcDotDotDistance(.ptPos, tgt.ptPos)
                If .ptPos.z >= tgt.ptPos.z And .ptPosPrev.z <= tgt.ptPos.z Or dist < 0.5 Then
                    If dist < 5 Then
                        dist = Sqr((.ptPos.X - tgt.ptPos.X) ^ 2 + (.ptPos.Y - tgt.ptPos.Y) ^ 2)
                        .leftticks = 0 ' 炮弹击中后消失
                        tgt.hp = tgt.hp - (100 + 700 / (dist + 1))
                        If tgt.hp <= 0 Then
                            tgt.leftticks = 0
                            tgt.deadTicks = 1200
                            gv.killedCnt = gv.killedCnt + 1
                            gv.scoreBonus = gv.scoreBonus + M3D_CalcDotDotDistance(tgt.ptPos, gv.turret.cam.pos) / 500
                        End If
                        Exit For
                    End If
                End If
            End If
        End With
    Next j
End Function

Public Sub ProcTargets()
    Dim n As Long
    Dim i As Long
    Dim vecErr As Vector3D_t
    Dim rndFac As Single
    Dim decay As Single
    Dim damage As Single
    Dim distToHit As Single
    Dim xyRndFac As Single, zRndFac As Single
    Dim isGod As Boolean
    decay = 1 - 7 / DRAW_PER_SEC
    n = MAX_TGT_CNT - 1
    rndFac = (1 + Form1.cmbDifficulty.ListIndex / 5)
    isGod = False
    If Form1.cmbDifficulty.ListIndex = Form1.cmbDifficulty.ListCount - 1 Then
        rndFac = rndFac * 1.33
        isGod = True
    End If
    For i = 0 To n
        With gv.tgts(i)
        If .leftticks <> 0 Then
            If .leftticks Mod 10 = 0 Then
                .ptPos.X = .ptPos.X + .vecV.X / 1000
                .ptPos.Y = .ptPos.Y + .vecV.Y / 1000
                .ptPos.z = .ptPos.z + .vecV.z / 1000
                .distToHit = M3D_CalcDotDotDistance(.ptPos, gv.cam.pos)
                If .distToHit < 3.6 Then
                    gv.hitCnt = gv.hitCnt + 1
                    'If Form1.chkJoy.Value = 0 Then
                        damage = 20 + Rnd * 20 + Form1.cmbDifficulty.ListIndex * 6
                        gv.newHitTick = damage * 255 \ 600 + 4
                        If Form1.chkJoy.Value <> 0 Then damage = damage / 10
                        gv.myHP = gv.myHP - damage
                        If gv.myHP < 0 Then
                            gv.myHP = 0
                            gv.newHitTick = 42
                            gv.state = STATE_SCORE
                        End If
                    'End If
                    .leftticks = 0
                    gv.escapeCnt = gv.escapeCnt + 1
                    GoTo NextLoop
                End If
                If .ptPos.z < gv.cam.viewDist Then
                    .leftticks = 0
                    gv.escapeCnt = gv.escapeCnt + 1
                    GoTo NextLoop
                End If
                                
                distToHit = M3D_CalcDotDotDistance(.ptPos, gv.cam.pos)
                xyRndFac = 1
                zRndFac = rndFac
                If distToHit < 250 Then
                    xyRndFac = (700 / (100 + distToHit) * rndFac) ^ 0.75
                    'zRndFac = 350 / (100 + distToHit) / rndFac
                End If
                .vecA.z = .vecA.z * 0.72 + (Rnd * 2 - 1) * rndFac
                .vecA.X = .vecA.X * 0.75 + (Rnd * 4 - 2) * rndFac
                .vecA.Y = .vecA.Y * 0.75 + (Rnd * 4 - 2) * rndFac
                
                .vecV.X = .vecV.X + .vecA.X
                .vecV.Y = .vecV.Y + .vecA.Y
                .vecV.z = .vecV.z + .vecA.z
                
                vecErr.X = (gv.turret.cam.pos.X - .ptPos.X)
                vecErr.Y = (gv.turret.cam.pos.Y - .ptPos.Y)
                vecErr.z = (gv.turret.cam.pos.z - .ptPos.z)
                GetUnitVector vecErr, vecErr
                ' magic guide
                .ptPos.X = .ptPos.X + vecErr.X * 0.5 * xyRndFac
                .ptPos.Y = .ptPos.Y + vecErr.Y * 0.5 * xyRndFac
                .ptPos.z = .ptPos.z + vecErr.z * 0.5 * rndFac
                If .ptPos.z > 20 Then
                    .vecA.z = .vecA.z + vecErr.z * rndFac * 0.05
                End If
            End If
            ' 检查是否被射中
            For j = 0 To m
                CheckHit gv.tgts(i)
            Next j
            If .leftticks = 0 Then
                gv.tgtCnt = gv.tgtCnt - 1
            Else
                .leftticks = .leftticks - 1
                If .leftticks = 0 Then
                    gv.escapeCnt = gv.escapeCnt + 1
                End If
            End If
        End If
        If .deadTicks <> 0 Then
            .deadTicks = .deadTicks - 1
            If .deadTicks Mod 10 = 0 Then
                .ptPos.X = .ptPos.X + .vecV.X / 1000
                .ptPos.Y = .ptPos.Y + .vecV.Y / 1000
                .ptPos.z = .ptPos.z + .vecV.z / 1000
                
                .vecA.X = .vecV.X * decay
                .vecA.Y = .vecV.Y * decay
                .vecA.z = .vecV.z * decay
                
                .vecA.Y = .vecV.Y - 98 / 1000
            End If
        End If
        End With
NextLoop:
    Next i
    n = Rnd * 100000
    If n < 105 + 25 * Form1.cmbDifficulty.ListIndex Then
        If gv.state = STATE_PLAYING Then
            SpawnTarget
        End If
    End If
End Sub

Public Sub GameStep()
    Dim i As Long, projCnt As Long
    Dim decay As Single

    If gv.myHP < 100 Then
        gv.myHP = gv.myHP + 1 / 1024
    End If

    If gv.gameRemainTick <> 0 And gv.state = STATE_PLAYING Then
        gv.gameRemainTick = gv.gameRemainTick - 1
        If gv.gameRemainTick = 0 Then
        End If
    End If
    If gv.gameRemainTick = 0 Or gv.state = STATE_SCORE Or gv.state = STATE_INIT Then
        If gv.turret.isFiring = True Then
            PlaySound vbNullString, 0, 0
            gv.turret.isFiring = False
        End If
        Exit Sub
    End If
    ProcTargets
    ReloadTurret gv.turret
    If gv.isFiring And gv.turret.ammo.reloadTickRem = 0 And gv.turret.ammo.clipAmmoRemCnt <> 0 Then
        FireTurret gv.turret.cam, 1
        If gv.turret.isFiring = True Then
            If gv.turret.isPlaying = False Then
                PlaySound gv.turret.sFireFile, 0, SND_ASYNC Or SND_FILENAME Or SND_LOOP
                gv.turret.isPlaying = True
            End If
        Else
            If gv.turret.isPlaying = True Then
                PlaySound vbNullString, 0, 0
                gv.turret.isFiring = False
            End If
            
        End If
    Else
        gv.turret.isFiring = False
        If gv.turret.isPlaying = True Then
            PlaySound vbNullString, 0, 0
            gv.turret.isPlaying = False
        End If
    End If
        
    decay = 1 - 0.011 / DRAW_PER_SEC
    projCnt = 0
    For i = 0 To MAX_PROJ_CNT - 1
        With g_projs(i)
            If .leftticks = 0 Then
                GoTo NextLoop
            End If
            .leftticks = .leftticks - 1
            projCnt = projCnt + 1
            .ptPos.X = .ptPos.X + .vecVel.X / 1000
            .ptPos.Y = .ptPos.Y + .vecVel.Y / 1000
            .ptPos.z = .ptPos.z + .vecVel.z / 1000
            .vecVel.X = .vecVel.X * decay
            .vecVel.Y = .vecVel.Y * decay
            .vecVel.z = .vecVel.z * decay
            
            ' apply gravity
            .vecVel.Y = .vecVel.Y - 9.8 / 1000 * 2
            
            If .ptPos.Y < -500 Then
                .leftticks = 0
            End If
        End With
NextLoop:
    Next i
    gv.projCnt = projCnt
End Sub
Public Sub ShowAutoMode()
    If gv.turret.autoMode = 0 Then
        Form1.lblAutoMode.Caption = "连射"
        Form1.lblAutoMode.BackColor = rgb(0, 130, 0)
    Else
        Form1.lblAutoMode.Caption = "单射"
        Form1.lblAutoMode.BackColor = rgb(130, 130, 0)
    End If
End Sub

Public Function GetFpmSoundFile(fpm As Single) As String
    Dim sFile As String
    Dim fileFPM As Single
    Dim minErr As Single
    Dim err As Single
    Dim sRetFile As String
    minErr = 100000
    sFile = Dir(App.Path & "\fpm*.wav")
    sFile = Left$(sFile, Len(sFile) - 4)
    Do
        If Left$(sFile, 4) = "fpm_" Then
            fileFPM = val(Mid$(sFile, 5))
            If fileFPM > fpm Then
                err = fileFPM / fpm
            Else
                err = fpm / fileFPM
            End If
            If err < minErr Then
                minErr = err
                sRetFile = sFile & ".wav"
            End If
        End If
    sFile = Dir
    If sFile = "" Then GoTo AfterDo
    sFile = Left$(sFile, Len(sFile) - 4)
    Loop
AfterDo:
    GetFpmSoundFile = App.Path & "\" & sRetFile
End Function

Public Sub EnumFpmSoundFiles()
    Dim sFile As String
    Dim fileFPM As Single
    Dim minErr As Single
    Dim fpmCnt As Long, fpm As Single
    Dim i As Long, j As Long
    minErr = 100000
    sFile = Dir(App.Path & "\fpm*.wav")
    sFile = Left$(sFile, Len(sFile) - 4)
    fpmCnt = 0
    Do
        If Left$(sFile, 4) = "fpm_" Then
            fileFPM = val(Mid$(sFile, 5))
            gv.fpmSounds(fpmCnt) = fileFPM
            fpmCnt = fpmCnt + 1
        End If
        sFile = Dir
        If sFile = "" Then GoTo AfterDo
    Loop
AfterDo:
    gv.fpmSoundCnt = fpmCnt
    For i = 0 To fpmCnt - 2
        For j = i + 1 To fpmCnt - 1
            If gv.fpmSounds(j) < gv.fpmSounds(i) Then
                fpm = gv.fpmSounds(i)
                gv.fpmSounds(i) = gv.fpmSounds(j)
                gv.fpmSounds(j) = fpm
            End If
        Next j
    Next i
End Sub

Public Sub Main()
    Dim i As Long, n As Long
    Dim sFile As String
    Randomize Timer
    gv.state = STATE_INIT
    gv.isShowFloatingStat = True
    gv.keyCmd.keySensitivity = 0.01
    gv.cam.fovDgr = 100#
    gv.cam.pos.X = 0
    gv.cam.pos.Y = 0.1
    gv.cam.pos.z = 0.5
    gv.cam.vecN.X = 0
    gv.cam.vecN.Y = 0.5
    gv.cam.vecN.z = 10
    gv.cam.viewDist = gv.cam.pos.z + 0.5
    gv.zoomFactor = 1
    sFile = Dir(App.Path & "\fpm*.wav")
    For i = 0 To 100 - 1
        
    Next i

    gv.dfcltLv(0) = 1680
    gv.dfcltLv(1) = 1200
    gv.dfcltLv(2) = 1111
    gv.dfcltLv(3) = 1028
    gv.dfcltLv(4) = 952
    gv.dfcltLv(5) = 882
    gv.dfcltLv(6) = 441
    n = MAX_PROJ_CNT - 1
    For i = 0 To n
        g_projs(i).leftticks = 0
    Next i
    
    gv.turret.accuracyErrDiv = 3
    gv.turret.cam = gv.cam
    gv.turret.cam.pos.X = 2
    gv.turret.cam.pos.Y = -1
    gv.turret.cam.pos.z = -1
    
    gv.turret.tickToNextFire = 1
    'gv.turret.tickReload = 1000# * 60 / CSng(gv.turret.fpm)
    gv.turret.autoMode = 0
    
    SetupCamera gv.cam
    SetupCamera gv.turret.cam
    
    EnumFpmSoundFiles
    ' must be put at the last, it will cause Form1 to load!
    gv.cam.WvsH = CSng(Form1.pic.ScaleWidth) / CSng(Form1.pic.ScaleHeight)
    gv.turret.cam.WvsH = gv.cam.WvsH
    LoadPlayers
    gv.turret.ammo.clipSize = 500
    
    Form1.Show
End Sub

