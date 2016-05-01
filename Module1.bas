Attribute VB_Name = "Module1"
Type POINTAPI '指针类型
    x As Long
    y As Long
End Type
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Const wFlags = &H10 Or &H40
Global cpencil As Long, cline As Long, cword As Long '笔触颜色
Global pencil As Double, rubber As Double, l As Double '笔触大小
Global c As Integer '当前工具
Global word As String '添加的文字
Global wfont As String '添加的文字字体
Global size As Integer '添加的文字字体
Global iswhite As Boolean '是否是白板
Global drawl As Boolean '引导直线
