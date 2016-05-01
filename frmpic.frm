VERSION 5.00
Begin VB.Form frmpic 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "AutoPointer"
   ClientHeight    =   3570
   ClientLeft      =   4365
   ClientTop       =   7110
   ClientWidth     =   6810
   Icon            =   "frmpic.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmpic.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   3570
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pc4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   180
      ScaleHeight     =   825
      ScaleWidth      =   1455
      TabIndex        =   3
      Top             =   315
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox pc3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1590
      Left            =   1485
      Picture         =   "frmpic.frx":045C
      ScaleHeight     =   1590
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   1170
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.PictureBox pc2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   3510
      ScaleHeight     =   825
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   585
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   1485
      X2              =   5805
      Y1              =   2925
      Y2              =   2925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   0
      Left            =   1125
      TabIndex        =   1
      Top             =   2700
      Width           =   45
   End
End
Attribute VB_Name = "frmpic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim drawit As Boolean, ru As Boolean
Dim linex As Integer, liney As Integer
Dim t As Integer

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single) '拖拽文字
  Source.Move (x - Source.Width / 2), (y - Source.Height / 2)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer) '如果是Esc则退出
  If KeyAscii = 27 Then frmtool.Command2.Value = True
End Sub

Private Sub Form_Load() '窗体初始化
  t = 0
  While frmmain.Visible <> False '隐藏主窗体
    DoEvents
  Wend
  frmpic.Top = 0 '将屏幕窗体最大化
  frmpic.Left = 0
  frmpic.Width = Screen.Width
  frmpic.Height = Screen.Height
  pc2.Top = 0 '将屏幕备份最大化
  pc2.Left = 0
  pc2.Width = Screen.Width
  pc2.Height = Screen.Height
  pc3.Top = 0 '将白板备份最大化
  pc3.Left = 0
  pc3.Width = Screen.Width
  pc3.Height = Screen.Height
  pc4.Top = 0 '将截屏储存最大化
  pc4.Left = 0
  pc4.Width = Screen.Width
  pc4.Height = Screen.Height
  Dim Screendc As Long '截屏
  Dim ret
  Screendc = GetDC(0)
  ret = BitBlt(frmpic.hdc, 0, 0, Screen.Width, Screen.Height, Screendc, 0, 0, vbSrcCopy)
  ret = ReleaseDC(0, screenhdc)
  pc2.PaintPicture frmpic.Image, 0, 0
  Dim f As Long, fw As Long, fh As Long '初始化工具窗体
  frmtool.Show
  fw = frmtool.Width
  fh = frmtool.Height
  frmtool.Visible = False
  f = SetWindowPos(frmtool.hwnd, -1, 0, 0, 150, 300, wflage)
  frmtool.Width = fw
  frmtool.Height = fh
  frmtool.Top = Screen.Height / 2 - fh
  frmtool.Left = Screen.Width - fw
  frmtool.Visible = True
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) '按下鼠标
  If c = 1 Then '铅笔
    drawit = True
    CurrentX = x
    CurrentY = y
    frmpic.DrawWidth = pencil
    frmpic.ForeColor = cpencil
    frmpic.PSet (x, y)
  End If
  If c = 3 Then '直线
    If drawl = False Then '起点
      linex = x
      liney = y
      drawl = True
      Line1.X1 = x
      Line1.X2 = x
      Line1.Y1 = y
      Line1.Y2 = y
      Line1.BorderColor = cline
      Line1.Visible = True
    Else
      frmpic.DrawWidth = l '终点
      frmpic.Line (linex, liney)-(x, y), cline
      drawl = False
      Line1.Visible = False
    End If
  End If
  If c = 2 Then '橡皮
    ru = True
  End If
  If c = 4 Then '文字
    word = ""
    frmword.Show 1
    If word <> "" Then
      t = t + 1
      Load Label1(t)
      Label1(t).Caption = word
      Label1(t).Font.Name = wfont
      Label1(t).Font.size = size
      Label1(t).Left = x
      Label1(t).Top = y
      Label1(t).ForeColor = cword
      Label1(t).Visible = True
    End If
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) '移动鼠标
  If drawit Then '铅笔
    If c = 1 Then
      frmpic.DrawWidth = pencil
      frmpic.Line (CurrentX, CurrentY)-(x, y), cpencil
      CurrentX = x
      CurrentY = y
    End If
  End If
  If ru = True Then '橡皮
    Dim ret
    Dim xx, xy
    Dim pt32 As POINTAPI
    Call GetCursorPos(pt32)
    xx = pt32.x - rubber \ 2
    xy = pt32.y - rubber \ 2
    If iswhite = False Then
      ret = BitBlt(frmpic.hdc, xx, xy, rubber, rubber, pc2.hdc, xx, xy, vbSrcCopy)
    Else
      ret = BitBlt(frmpic.hdc, xx, xy, rubber, rubber, frmback.picwhite.hdc, 1, 1, Not (rop))
    End If
    frmpic.Refresh
  End If
  If c = 3 Then '直线
    Line1.X2 = x
    Line1.Y2 = y
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) '鼠标弹起
  If c = 1 Then '铅笔
    drawit = False
  End If
  If c = 2 Then ru = False '橡皮
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, x As Single, y As Single) '拖拽文字
  Source.Move Label1(Index).Left + (x - Source.Width / 2), Label1(Index).Top + (y - Source.Height / 2)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single) '删除文字
  If Button = 2 Then Unload Label1(Index)
End Sub
