VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmtool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   4005
   ClientLeft      =   7065
   ClientTop       =   2850
   ClientWidth     =   1605
   Icon            =   "frmtool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   107
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1260
      ScaleHeight     =   285
      ScaleWidth      =   375
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����ͼ��"
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
      Left            =   45
      TabIndex        =   9
      Top             =   3105
      Width           =   1500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ǧ��"
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
      Left            =   52
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   45
      Width           =   1500
   End
   Begin VB.OptionButton Option2 
      Caption         =   "��Ƥ"
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
      Left            =   52
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   495
      Width           =   1500
   End
   Begin VB.OptionButton Option3 
      Caption         =   "ֱ��"
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
      Left            =   52
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   945
      Width           =   1500
   End
   Begin VB.OptionButton Option4 
      Caption         =   "����"
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
      Left            =   52
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1395
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   990
      Top             =   2430
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�װ�"
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
      Left            =   52
      TabIndex        =   3
      Top             =   2655
      Width           =   1500
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   315
      Top             =   1035
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   255
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
      Left            =   52
      TabIndex        =   4
      Top             =   3555
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ѡ����ɫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   585
      TabIndex        =   2
      Top             =   2205
      Width           =   960
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmtool.frx":030A
      Left            =   90
      List            =   "frmtool.frx":033E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1845
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Top             =   2205
      Width           =   465
   End
End
Attribute VB_Name = "frmtool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click() '�ʴ���С
  If c = 1 Then pencil = Val(Combo1.Text)
  If c = 2 Then rubber = Val(Combo1.Text)
  If c = 3 Then l = Val(Combo1.Text)
End Sub

Private Sub Command1_Click() 'ѡ����ɫ
  On Error GoTo a1
  cd1.ShowColor
  If c = 1 Then cpencil = cd1.Color
  If c = 3 Then cline = cd1.Color
  If c = 4 Then cword = cd1.Color
  Label1.BackColor = cd1.Color
a1: End Sub

Private Sub Command2_Click() '�˳�
  frmtool.Command3.Caption = "�װ�"
  iswhite = False
  Unload frmpic
  frmtool.Hide
  frmmain.Show
End Sub

Private Sub Command3_Click() '�װ塢��Ļ
  If Command3.Caption = "�װ�" Then
    Command3.Caption = "��Ļ"
    iswhite = True
    frmpic.PaintPicture frmpic.pc3.Image, 0, 0
  Else
    Command3.Caption = "�װ�"
    iswhite = False
    frmpic.PaintPicture frmpic.pc2.Image, 0, 0
  End If
End Sub

Private Sub Command4_Click() '����ͼ��
  cd1.FileName = ""
  cd1.Filter = "*.bmp|*.bmp"
  cd1.ShowSave
  If cd1.FileName <> "" Then
    frmtool.Hide
    DoEvents
    Dim Screendc As Long '����
    Dim ret
    Screendc = GetDC(0)
    ret = BitBlt(frmpic.pc4.hdc, 0, 0, Screen.Width, Screen.Height, Screendc, 0, 0, vbSrcCopy)
    ret = ReleaseDC(0, screenhdc)
    Set frmpic.pc4.Picture = frmpic.pc4.Image
    SavePicture frmpic.pc4.Picture, cd1.FileName
    frmtool.Show
  End If
End Sub

Private Sub Form_Load() '�����ʼ��
  Option1.Value = True
  Label1.BackColor = cpencil
  Command3.Caption = "�װ�"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) '��������
  frmtool.Height = 4455
End Sub

Private Sub Form_Unload(Cancel As Integer) '�˳�
  Call Command2_Click
End Sub

Private Sub Label1_DblClick() 'ѡ����ɫ
  Call Command1_Click
End Sub

Private Sub Option1_Click() 'ѡ��Ǧ��
  c = 1
  Combo1.Text = Mid(Str(pencil), 2)
  Label1.BackColor = cpencil
  frmpic.Line1.Visible = False
End Sub

Private Sub Option2_Click() 'ѡ����Ƥ
  c = 2
  Combo1.Text = Mid(Str(rubber), 2)
End Sub

Private Sub Option3_Click() 'ѡ��ֱ��
  c = 3
  Combo1.Text = Mid(Str(l), 2)
  Label1.BackColor = cline
  If drawl Then frmpic.Line1.Visible = True
End Sub

Private Sub Option4_Click() 'ѡ������
  c = 4
  Label1.BackColor = cword
End Sub

Private Sub Timer1_Timer() '��������
  Dim m As POINTAPI, re As Long, pptx As Long, ppty As Long
  pptx = Screen.TwipsPerPixelX
  ppty = Screen.TwipsPerPixelY
  re = GetCursorPos(m)
  If Not (m.x >= frmtool.Left / pptx And m.x <= frmtool.Left / pptx + frmtool.Width / pptx And m.y >= frmtool.Top / ppty And m.y <= frmtool.Top / ppty + frmtool.Height / ppty) Then
    frmtool.Height = 20
  End If
End Sub
