VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   4650
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "自动签到启动"
      Height          =   2295
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpclassname As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpclassname As String, ByVal nMaxCount As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim clName As String
Dim jg As Long
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Private Type mType
fhwnd As Long
fText As String * 255
fRect As RECT
pHwnd As Long
pText As String * 255
End Type
Private Sub mGetAllWindow(m_Type() As mType)
Dim Wndback As Long
Dim i As Long
Do
ReDim Preserve m_Type(i)
DoEvents

m_Type(i).fhwnd = FindWindowEx(0, Wndback, vbNullString, vbNullString)
If m_Type(i).fhwnd = 0 Then
Exit Sub
Else
GetWindowText m_Type(i).fhwnd, m_Type(i).fText, 255

GetWindowRect m_Type(i).fhwnd, m_Type(i).fRect

m_Type(i).pHwnd = GetParent(m_Type(i).fhwnd)

GetWindowText m_Type(i).pHwnd, m_Type(i).pText, 255
End If
Wndback = m_Type(i).fhwnd
i = i + 1
Loop
End Sub
Sub PlayWavFile(strFileName As String, PlayCount As Long, JianGe As Long)


If Len(Dir(strFileName)) = 0 Then Exit Sub
If PlayCount = 0 Then Exit Sub
If JianGe < 1000 Then JianGe = 1000
DoEvents
sndPlaySound strFileName, 16 + 1
Sleep JianGe
Call PlayWavFile(strFileName, PlayCount - 1, JianGe)
End Sub

Private Sub Command1_Click()
Dim cType() As mType

mGetAllWindow cType()
Dim lpclassname As String
Dim rh As Long
Dim lhwnd As Long
Dim i As Long
Dim isVisible As Long
Dim a As RECT
Dim bili As Double
Dim wxblong As Long, wxblength As Long
Dim presswxbX As Long, presswxbY As Long
Dim j As Long

For j = 1 To 100
Sleep jg * 1000
j = 1
For i = LBound(cType) To UBound(cType)

lpclassname = Space(6)
GetClassName cType(i).fhwnd, lpclassname, 255
If lpclassname = "#32770" Then
If IsWindowVisible(cType(i).fhwnd) = 1 Then
GetWindowRect cType(i).fhwnd, a
wxblong = a.Right - a.Left
wxblength = a.Bottom - a.Top
bili = wxblong / wxblength
If bili > 0.9 Then
If bili < 0.95 Then
MsgBox "签到来了"
PlayWavFile "leile.wav", 1, 1000
Sleep 2000
presswxbX = (a.Right + a.Left) / 2
presswxbY = a.Bottom + wxblength / 8
'AutoPressMouse presswxbX, presswxbY
PlayWavFile "cgl.wav", 1, 1000
End If
End If
End If
End If
Next
Next j
End Sub
Private Sub AutoPressMouse(x As Long, y As Long)
SetCursorPos x, y
' mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
Private Sub Form_Load()
jg = InputBox("请输入循环检测签到框间隔(不需要输单位，单位为秒,不要输入为空或者0、负数之类的，程序卡死或者不检测本人一概不负责！）:", "请输入")
MsgBox "确认你的签到间隔为" & jg & " 秒"
Command1.Caption = "刷新"
End Sub
