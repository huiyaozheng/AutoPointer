VERSION 5.00
Begin VB.Form frmword 
   Caption         =   "添加文字"
   ClientHeight    =   2175
   ClientLeft      =   5325
   ClientTop       =   2640
   ClientWidth     =   3915
   Icon            =   "frmword.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   3915
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   2475
      TabIndex        =   5
      Top             =   1215
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   1125
      TabIndex        =   4
      Top             =   1215
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   82
      TabIndex        =   1
      Top             =   810
      Width           =   3750
   End
   Begin VB.Frame Frame1 
      Caption         =   "字体"
      Height          =   690
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3795
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmword.frx":030A
         Left            =   2340
         List            =   "frmword.frx":0365
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   1365
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmword.frx":03DA
         Left            =   90
         List            =   "frmword.frx":03F3
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   2130
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "用鼠标左键可移动文字"
      Height          =   195
      Index           =   1
      Left            =   675
      TabIndex        =   7
      Top             =   1890
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "注意：建立后用鼠标右键单击文字可清除文字"
      Height          =   195
      Index           =   0
      Left            =   157
      TabIndex        =   6
      Top             =   1665
      Width           =   3600
   End
End
Attribute VB_Name = "frmword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() '确定添加
  word = Text1.Text
  wfont = Combo1.Text
  size = Val(Combo2.Text)
  Unload Me
End Sub

Private Sub Command2_Click() '取消添加
  word = ""
  Unload Me
End Sub

Private Sub Form_Load() '初始化
  Combo1.Text = "宋体"
  Combo2.Text = 20
End Sub

Private Sub Text1_Change() '处理Enter
  If KeyAscii = 13 Then Call Command1_Click
End Sub
