Attribute VB_Name = "Module1"
Type POINTAPI 'ָ������
    x As Long
    y As Long
End Type
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Const wFlags = &H10 Or &H40
Global cpencil As Long, cline As Long, cword As Long '�ʴ���ɫ
Global pencil As Double, rubber As Double, l As Double '�ʴ���С
Global c As Integer '��ǰ����
Global word As String '��ӵ�����
Global wfont As String '��ӵ���������
Global size As Integer '��ӵ���������
Global iswhite As Boolean '�Ƿ��ǰװ�
Global drawl As Boolean '����ֱ��
