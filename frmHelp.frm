VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "弹幕防御 说明"
   ClientHeight    =   8205
   ClientLeft      =   2190
   ClientTop       =   1725
   ClientWidth     =   11040
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   547
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   736
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   5520
      ScaleHeight     =   231
      ScaleMode       =   0  'User
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   3840
      Width           =   5415
   End
   Begin VB.PictureBox picCIWS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   5520
      ScaleHeight     =   231
      ScaleMode       =   0  'User
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   7935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelp.frx":0000
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label lblAck 
      BackColor       =   &H00404040&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   3
      Top             =   7440
      Width           =   5295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim str1 As String
    picCIWS.Picture = LoadPicture(App.Path & "\CIWS_1130.bmp")
    pic2.Picture = LoadPicture(App.Path & "\CIWS_1130_B.bmp")
    lblAck.ForeColor = rgb(255, 160, 185)
    lblAck.Caption = "致敬最可爱的人！有你才有岁月静好" & Chr$(13) & Chr$(10) & "瘐子年八一献礼"
End Sub

