VERSION 5.00
Begin VB.Form frmPlayer 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�û�"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9840
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9840
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ListBox lstPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4020
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "ѡ��״̬"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblBest 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   2295
      Left            =   6600
      TabIndex        =   9
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label lblBestName 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "״Ԫ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblChangeName 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "���ĵ�¼����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   3
      Left            =   6240
      TabIndex        =   7
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label lblChangeName 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   2
      Left            =   8760
      TabIndex        =   6
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label lblCur 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "��ǰѡ��ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label lblChangeName 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "ѡ�������ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lblChangeName 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "����ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   0
      Left            =   3600
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblStat 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2295
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ShowPlayer(ndx As Long, lbl As Label)
    Dim strSt As String
    
    With gv.players(ndx)
    strSt = "ѡ������: " & .sName
    strSt = strSt & Chr$(13) & Chr$(10) & "ս������ : " & CStr(.playCnt)
    strSt = strSt & Chr$(13) & Chr$(10) & "ʤ������ : " & CStr(.winCnt)
    If .playCnt <> 0 Then
        strSt = strSt & Chr$(13) & Chr$(10) & "ʤ������ : " & Format("0.#%", CDbl(.winCnt) / .playCnt)
    Else
        strSt = strSt & Chr$(13) & Chr$(10) & "ʤ������ : -"
    End If
    strSt = strSt & Chr$(13) & Chr$(10) & "  �ܷ���: " & CStr(.scoreAcc)
    strSt = strSt & Chr$(13) & Chr$(10) & "  ��߷�: " & CStr(.scoreHigh)
    End With
    lbl.Caption = strSt
End Sub

Private Sub Form_Load()
    Dim i As Long
    For i = 1 To gv.playerCnt
        Me.lstPlayer.AddItem gv.players(i - 1).sName
    Next i
    Me.lstPlayer.ListIndex = gv.playerNdx
    ShowPlayer gv.playerNdx, Me.lblStat
    lblCur.Caption = lstPlayer.Text
    ShowBestPlayer
End Sub

Private Sub Label2_Click()

End Sub

Private Sub lblChangeName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblChangeName(Index).BackColor = 0
End Sub

Private Sub ShowBestPlayer()
    ShowPlayer gv.highestScoreNdx, Me.lblBest
End Sub

Private Sub ChangePassword(ndx As Long)
    Dim sInPass As String, sIn2Pass As String
    With gv.players(ndx)
        sInPass = "123456": sIn2Pass = "123456"
        If .sPassword <> "123456" Then
            sInPass = InputBox("������" & .sName & "�ĵ�¼����")
        End If
        If sInPass = .sPassword Then
            sInPass = InputBox("��Ϊ" & .sName & "�����µĵ�¼����")
            If sInPass = "" Then sInPass = "123456"
            sIn2Pass = InputBox("���ٴ�Ϊ" & .sName & "�����µĵ�¼����")
            If sIn2Pass = "" Then sIn2Pass = "123456"
            If sInPass = .sPassword Then
                MsgBox "������;����벻����ͬ��"
            Else
                If sInPass <> sIn2Pass Then
                    MsgBox "������������벻һ�£�"
                Else
                    .sPassword = sInPass
                    UpdatePlayer ndx
                    MsgBox "���µ�¼����ɹ���"
                    
                End If
            End If
        End If
    End With
End Sub

Private Sub lblChangeName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sInPass As String, sIn2Pass As String, sOldName As String
    Dim i As Long
    sInPass = "123456"
    lblChangeName(Index).BackColor = rgb(63, 63, 63)
    If Index = 0 Then
        If Len(txtName.Text) <> 0 And txtName.Text <> lstPlayer.Text Then
            For i = 1 To gv.playerCnt
                If txtName.Text = gv.players(i - 1).sName Then
                    MsgBox "��ѡ�ֺ͵�ǰѡ���б��е�����!"
                    Exit Sub
                End If
            Next i
            If MsgBox("����ѡ�ֻ�ɾ����ǰѡ�ֵ����ϣ���ȷ����", vbYesNo, "����ѡ��") = vbYes Then
                With gv.players(lstPlayer.ListIndex)
                    If .sPassword <> "123456" Then
                        sInPass = InputBox("�����뵱ǰѡ�ֵĵ�¼����")
                    End If
                    If sInPass = .sPassword Then
                        sInPass = InputBox("��������ѡ�ֵĵ�¼����")
                        If sInPass = "" Then sInPass = "123456"
                        sIn2Pass = InputBox("���ٴ�������ѡ�ֵĵ�¼����")
                        If sIn2Pass = "" Then sIn2Pass = "123456"
                        If sInPass = sIn2Pass Then
                            sOldName = .sName
                            .sName = txtName.Text
                            .scoreAcc = 0
                            .scoreHigh = 0
                            .winCnt = 0
                            .playCnt = 0
                            .sPassword = sInPass
                            UpdatePlayer lstPlayer.ListIndex
                            lstPlayer.List(lstPlayer.ListIndex) = .sName
                            MsgBox "����ѡ�ֳɹ��� " & sOldName & " -> " & .sName
                            GetBestPlayerNdx
                            ShowBestPlayer
                        Else
                            MsgBox "�������벻һ�£��޷��½��û���"
                        End If
                    Else
                        MsgBox "��������޷�ɾ����ǰ�û�" & .sName
                    End If
                End With
            End If
        End If
    ElseIf Index = 1 Then
        With gv.players(lstPlayer.ListIndex)
            sInPass = "123456"
            If .sPassword <> "123456" Then
                sInPass = InputBox("������" & lstPlayer.Text & "�ĵ�¼����")
            End If
            If sInPass = .sPassword Then
                gv.playerNdx = lstPlayer.ListIndex
                Form1.lblPlayer.Caption = gv.players(gv.playerNdx).sName
                lblCur.Caption = lstPlayer.Text
                UpdateCurrentPlayer lstPlayer.ListIndex
            Else
                MsgBox "��¼�������"
            End If
        End With
    ElseIf Index = 2 Then '�˳�
        Unload Me
    ElseIf Index = 3 Then '���ĵ�¼����
        ChangePassword gv.playerNdx
    End If
End Sub

Private Sub lstPlayer_Click()
    ShowPlayer lstPlayer.ListIndex, Me.lblStat
End Sub
