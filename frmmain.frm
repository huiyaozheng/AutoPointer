VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoPointer"
   ClientHeight    =   1290
   ClientLeft      =   13470
   ClientTop       =   4995
   ClientWidth     =   1605
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   1605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "���ڱ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   30
      TabIndex        =   2
      Top             =   540
      Width           =   1545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   30
      TabIndex        =   1
      Top             =   855
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʾ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   1545
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() '��ʾ��Ļ
  frmmain.Hide
  DoEvents
  frmpic.Show
End Sub

Private Sub Command2_Click() '����
  End
End Sub

Private Sub Command3_Click() '����
  frmabout.Show
End Sub

Private Sub Form_Load() '�����ʼ��
  If App.PrevInstance = True Then  '�ж��Ƿ�������
     MsgBox "�Բ��𣬳������һ��ʵ�������У�", , "AutoPointer"
     End
  End If
  cpencil = RGB(255, 0, 0) '��ʼ������
  cline = RGB(255, 0, 0)
  cword = RGB(0, 0, 0)
  pencil = 2
  rubber = 48
  l = 2
  Dim fh, fw As Integer '���������ö�
  fh = frmmain.Height
  fw = frmmain.Width
  Dim r As Long
  r = SetWindowPos(frmmain.hwnd, -1, 0, 0, 0, 0, wflages)
  frmmain.Width = fw
  frmmain.Height = fh
  frmmain.Top = Screen.Height / 2 - frmmain.Height
  frmmain.Left = Screen.Width - frmmain.Width
  iswhite = False
  Load frmback
  frmback.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer) '����
  End
End Sub
