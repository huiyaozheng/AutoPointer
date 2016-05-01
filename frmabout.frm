VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "关于AutoPointer"
   ClientHeight    =   3705
   ClientLeft      =   5415
   ClientTop       =   3600
   ClientWidth     =   5220
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   90
      Top             =   2340
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "frmabout.frx":030A
      Top             =   0
      Width           =   5250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "特别鸣谢： 无锡市大桥实验中学"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   465
      TabIndex        =   3
      Top             =   2835
      Width           =   4290
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "感谢支持本软件开发、改进的朋友们！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   450
      TabIndex        =   2
      Top             =   2565
      Width           =   4290
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "QQ: 1410246660"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   465
      TabIndex        =   1
      Top             =   2295
      Width           =   4290
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Email: huiyaozheng@gmail.com"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   465
      TabIndex        =   0
      Top             =   2025
      Width           =   4290
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Image1_Click()
  Unload Me
End Sub

Private Sub Label1_Click(Index As Integer)
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Dim i As Integer
  For i = 2 To 5
    Label1(i).Top = Label1(i).Top - 10
    If Label1(i).Top <= 0 Then Label1(i).Top = Me.Height + 500: Label1(i).Visible = True
    If Label1(i).Top + Label1(i).Height < Image1.Height Then Label1(i).Visible = False
  Next i
End Sub
