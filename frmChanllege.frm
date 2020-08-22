VERSION 5.00
Begin VB.Form frmChanllege 
   BackColor       =   &H00404040&
   Caption         =   "挑战项"
   ClientHeight    =   4395
   ClientLeft      =   18300
   ClientTop       =   9540
   ClientWidth     =   5100
   LinkTopic       =   "Form3"
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   Begin VB.CheckBox chkCnlg 
      BackColor       =   &H00404040&
      Caption         =   "敌导弹更致命"
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
      Index           =   3
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "10倍子弹时间, 充足弹药, 受到1/10伤害"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CheckBox chkCnlg 
      BackColor       =   &H00404040&
      Caption         =   "敌导弹更机动"
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
      Index           =   2
      Left            =   360
      TabIndex        =   8
      ToolTipText     =   "10倍子弹时间, 充足弹药, 受到1/10伤害"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CheckBox chkCnlg 
      BackColor       =   &H00404040&
      Caption         =   "救援到达更慢"
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
      Index           =   1
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "10倍子弹时间, 充足弹药, 受到1/10伤害"
      Top             =   960
      Width           =   2295
   End
   Begin VB.CheckBox chkCnlg 
      BackColor       =   &H00404040&
      Caption         =   "炮弹初速降低"
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
      Index           =   0
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "10倍子弹时间, 充足弹药, 受到1/10伤害"
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "黑体"
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
      Left            =   2280
      TabIndex        =   17
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label lblScoreFac 
      BackColor       =   &H00404040&
      Caption         =   "得分加成"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   360
      TabIndex        =   16
      Top             =   2880
      Width           =   4200
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   11
      Left            =   3360
      TabIndex        =   15
      Top             =   2040
      Width           =   195
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   10
      Left            =   3000
      TabIndex        =   14
      Top             =   2040
      Width           =   195
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   9
      Left            =   2640
      TabIndex        =   13
      Top             =   2040
      Width           =   195
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   8
      Left            =   3480
      TabIndex        =   11
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   7
      Left            =   3120
      TabIndex        =   10
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   6
      Left            =   2760
      TabIndex        =   9
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   5
      Left            =   3840
      TabIndex        =   7
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   4
      Left            =   3480
      TabIndex        =   6
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   3
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   195
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   195
   End
   Begin VB.Label lblChlng 
      BackColor       =   &H00808080&
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
      Index           =   2
      Left            =   3600
      TabIndex        =   1
      Top             =   480
      Width           =   195
   End
End
Attribute VB_Name = "frmChanllege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RenderChlng()
    Dim i As Long, ndx As Long, j As Long
    Dim cols(0 To 5) As Long
    cols(0) = rgb(0, 255, 0)
    cols(1) = rgb(255, 255, 0)
    cols(2) = rgb(255, 0, 0)
    cols(3) = rgb(112, 144, 112)
    cols(4) = rgb(144, 144, 112)
    cols(5) = rgb(144, 112, 112)
    With gv.chlng
    For i = 1 To CHLNG_COUNT
        chkCnlg(i - 1).Left = 32
        chkCnlg(i - 1).Top = 20 + (i - 1) * 30
        
        ndx = (i - 1) * 3
        Me.chkCnlg(i - 1).Value = .isEn(i - 1)
        For j = 0 To 3 - 1
            With Me.lblChlng(ndx + j)
                .Left = 200 + j * 20
                .Top = 20 + (i - 1) * 30
                If j < gv.chlng.lvs(i - 1) Then
                    If gv.chlng.isEn(i - 1) = 1 Then
                        .BackColor = cols(j)
                    Else
                        .BackColor = cols(j + 3)
                    End If
                Else
                    .BackColor = rgb(100, 100, 100)
                End If
            End With
        Next j
    Next i
    End With
    lblScoreFac.Caption = "得分加成: " & Format(CalcChlngScoreFac, "#%")
End Sub
Private Sub chkCnlg_Click(Index As Integer)
    Dim ndx0 As Long
    If chkCnlg(Index).Value = 0 Then gv.chlng.isEn(Index) = 0 Else gv.chlng.isEn(Index) = 1
    RenderChlng
End Sub

Private Sub Form_Load()
RenderChlng
End Sub

Private Sub lblChlng_Click(Index As Integer)
    Dim row As Long, col As Long
    Dim ndx0 As Long, ndx As Long
    Dim i As Long
    row = Index \ 3
    gv.chlng.isEn(row) = 1
    col = Index Mod 3
    gv.chlng.lvs(row) = col + 1
    RenderChlng
End Sub

Private Sub lblOK_Click(Index As Integer)
    CalcProjVelMo0
    Unload Me
End Sub
