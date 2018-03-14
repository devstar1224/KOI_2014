VERSION 5.00
Begin VB.UserControl VSKIN 
   BackColor       =   &H002D2D30&
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "맑은 고딕"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   270
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   255
   Begin VB.PictureBox picButtons 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   1200
      Picture         =   "VSKIN.ctx":0000
      ScaleHeight     =   130
      ScaleMode       =   0  '사용자
      ScaleWidth      =   78
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox picLayout 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   120
      Picture         =   "VSKIN.ctx":781C
      ScaleHeight     =   64
      ScaleMode       =   0  '사용자
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "VSKIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'jn4kim의 폼스킨 기반
'http://jnakim.com/

' 2014 - 02 - 14  초콜릿 못받은 날 ㅠㅠ
'Piero UL(the_xma_n) 제작
'http://cafe.naver.com/gogoomas 활동중
'내가 만든거 내가 배포함.. 돈주고 거래는 안될 저급 폼스킨이지만 그래도 걸리면 죽는다

'modFade : 페이드인/페이드아웃 효과
'modMinSize : 폼 최소사이즈 적용

'---[Properties]
'   Resize : 리사이징 허용
'   Caption : 폼 제목



'############폼스킨에따라 바꿔줘야 할 상수############

'--LAYOUT--
Private Const cWidth = 10    ' 모서리 너비
Private Const uHeight = 10   ' 상단 높이
Private Const dHeight = 18   ' 하단 높이

'--BUTTON--
Const bHeight = 26           ' 버튼 높이
Const bMargin = 0            ' 버튼 오른쪽 여백
Const bMinWidth = 26        ' 최소화버튼 너비
Const bMaxWidth = 26         ' 최대화버튼 너비
Const bExitWidth = 26        ' 종료버튼 너비

'--CAPTION--
Const cX = 9                 '캡션 X 좌표
Const cY = 4                 '캡션 Y 좌표

'############폼스킨에따라 바꿔줘야 할 상수 끝############


Private bLeft As Integer

'--폼 리사이징--
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const SC_DRAG_RESIZEL As Long = 10& ' 왼쪽
Private Const SC_DRAG_RESIZER As Long = 11& ' 오른쪽
Private Const SC_DRAG_RESIZEU As Long = 12& ' 위쪽
Private Const SC_DRAG_RESIZEUL As Long = 13& ' 위쪽 왼쪽으로 늘리기
Private Const SC_DRAG_RESIZEUR As Long = 14& ' 위쪽 오른쪽으로 늘리기
Private Const SC_DRAG_RESIZED As Long = 15& ' 아래로 늘리기
Private Const SC_DRAG_RESIZEDL As Long = 16& ' 아래 왼쪽으로
Private Const SC_DRAG_RESIZEDR As Long = 17& ' 아래 오른쪽으로
Private Const SC_DRAG_MOVE As Long = 2& ' 움직이기


'--유저컨트롤 설정값--
Dim cName As String
Dim bResize As Boolean


'--WORK AREA 구하기--
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Dim wArea As Rect


'--형식--
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type BtnPosition
    X As Single
    Y As Single
End Type


'--그래픽 API--
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'--최대화 관련--
Dim bMaxed As Boolean
Dim originalRect As Rect ' 최대화 하기 전 폼사이즈


'--마우스 관련 API--
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As PointAPI) As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long
Dim iPA As PointAPI, iRect As Rect, iBP As BtnPosition

'--마우스위치가 해당 버튼위에 있는지/없는지
Dim bButtons As Boolean
Dim bMinimize As Boolean, bMaximize As Boolean, bExit As Boolean
Dim g_MouseIn As Boolean, g_MouseDown As Boolean
Dim iMouseDown As Integer

' 화면 위치
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As Rect) As Long



'--jn4kim의 폼스킨 기반
'--http://jnakim.com/

Sub drawForm()
    On Error Resume Next
    
    
    Cls
    
    '위
    BitBlt hdc, 0, 0, cWidth, uHeight, picLayout.hdc, 0, 0, vbSrcCopy
    StretchBlt hdc, cWidth, 0, Width / 15 - 2 * cWidth, uHeight, picLayout.hdc, cWidth, 0, picLayout.Width - 2 * cWidth, uHeight, vbSrcCopy
    BitBlt hdc, Width / 15 - cWidth, 0, cWidth, uHeight, picLayout.hdc, picLayout.Width - cWidth, 0, vbSrcCopy
    
    ' 중간
    StretchBlt hdc, 0, uHeight, cWidth, Height / 15 - uHeight - dHeight, picLayout.hdc, 0, uHeight, cWidth, picLayout.Height - uHeight - dHeight, vbSrcCopy
    StretchBlt hdc, cWidth, uHeight, Width / 15 - 20, Height / 15 - uHeight - dHeight, picLayout.hdc, cWidth, uHeight, picLayout.Width - 2 * cWidth, picLayout.Height - uHeight - dHeight, vbSrcCopy
    StretchBlt hdc, Width / 15 - cWidth, uHeight, cWidth, Height / 15 - uHeight - dHeight, picLayout.hdc, picLayout.Width - cWidth, uHeight, cWidth, picLayout.Height - uHeight - dHeight, vbSrcCopy
    
    ' 아래
    BitBlt hdc, 0, Height / 15 - dHeight, cWidth, dHeight, picLayout.hdc, 0, picLayout.Height - dHeight, vbSrcCopy
    StretchBlt hdc, cWidth, Height / 15 - dHeight, Width / 15 - 2 * cWidth, dHeight, picLayout.hdc, cWidth, picLayout.Height - dHeight, cWidth, dHeight, vbSrcCopy
    BitBlt hdc, Width / 15 - cWidth, Height / 15 - dHeight, cWidth, dHeight, picLayout.hdc, picLayout.Width - cWidth, picLayout.Height - dHeight, vbSrcCopy


    ' 버튼(default)
    BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, 0, vbSrcCopy
    If bResize = True Then
        If bMaxed = True Then
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, 0, bHeight * 4, vbSrcCopy
        Else
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, 0, vbSrcCopy
        End If
    Else
        BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight * 3, vbSrcCopy
    End If
    BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, 0, vbSrcCopy

        
        '폼캡션 입력
        TextOut hdc, cX, cY, cName, LenB(StrConv(cName, vbFromUnicode))
    
End Sub

Private Sub UserControl_DblClick()
If Resize = False Then Exit Sub

If (iRect.Right - iPA.X > bLeft And iPA.Y - iRect.Top < uHeight) Then
        If bMaxed = False Then
            '//최대화 전의 설정 변수에 담아둠
            originalRect.Left = Parent.Left
            originalRect.Right = Parent.Width
            originalRect.Top = Parent.Top
            originalRect.Bottom = Parent.Height
            
            '//작업영역(작업표시줄을 제외한 영역)
            
            SystemParametersInfo 48, 0, wArea, 0
            
            '//작업영역에따라 폼의 사이즈와 폼스킨 사이즈 설정
            bMaxed = True
            Parent.Left = wArea.Left * 15
            Parent.Top = wArea.Top * 15
            Parent.Width = (wArea.Right - wArea.Left) * 15
            Parent.Height = (wArea.Bottom - wArea.Top) * 15
            Width = Parent.Width
            Height = Parent.Height
            drawForm
            Exit Sub
        Else
            '//최대화 이전의 상태로 설정
            Parent.Left = originalRect.Left
            Parent.Width = originalRect.Right
            Parent.Top = originalRect.Top
            Parent.Height = originalRect.Bottom
            Width = Parent.Width
            Height = Parent.Height
            bMaxed = False
            drawForm
            
        End If
End If
End Sub

Sub drawButtons()
'//현재 마우스 위치 가져옴
GetCursorPos iPA
'//현재 창 위치 가져옴
GetWindowRect hWnd, iRect



Dim a, B, C, d
a = iPA.X - iRect.Left
B = iPA.Y - iRect.Top
C = iRect.Right - iPA.X
d = iRect.Bottom - iPA.Y

Dim intRight As Integer
intRight = bLeft - bMinWidth - bMaxWidth - bExitWidth

'Debug.Print a & " " & b & " " & c & " " & d




If (0 < B And B < bHeight) And (intRight < C And C < bLeft) Then
    bButtons = True
    
    If C - intRight <= bExitWidth Then
        bExit = True ' 종료
        bMinimize = False
        bMaximize = False
        
        If iMouseDown = 1 Then
           BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, 2 * bHeight, vbSrcCopy
        Else
            ' 종료버튼 롤오버
            BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, bHeight, vbSrcCopy
        End If
        
        '다른 버튼은 원상복구
        If bResize = True Then
            If bMaxed = True Then
                BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, 0, bHeight * 4, vbSrcCopy
            Else
                BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, 0, vbSrcCopy
            End If
        Else
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight * 3, vbSrcCopy
        End If
        BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, 0, vbSrcCopy
        
    
    ElseIf C - intRight <= bExitWidth + bMaxWidth Then
        bMaximize = True
        bMinimize = False
        bExit = False
        
        
        '최대화버튼 롤오버
        If bResize = True Then
            If iMouseDown = 1 Then
                If bMaxed = True Then
                     BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMaxWidth, bHeight * 4, vbSrcCopy
                Else
                     BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight * 2, vbSrcCopy
                End If
            Else
                If bMaxed = True Then
                     BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMaxWidth * 2, bHeight * 4, vbSrcCopy
                Else
                     BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight, vbSrcCopy
                End If
                
            End If
        End If
        
        '다른 버튼은 원상복구
        BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, 0, vbSrcCopy
        BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, 0, vbSrcCopy
    ElseIf C - intRight <= bExitWidth + bMaxWidth + bMinWidth Then
        bMinimize = True
        bMaximize = False
        bExit = False
        
        
        '최소화버튼 롤오버
        If iMouseDown = 1 Then
            BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, bHeight * 2, vbSrcCopy
        Else
            BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, bHeight, vbSrcCopy
        End If
        
        '다른 버튼은 원상복구
        If bResize = True Then
            If bMaxed = True Then
                BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, 0, bHeight * 4, vbSrcCopy
            Else
                BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, 0, vbSrcCopy
            End If
        Else
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight * 3, vbSrcCopy
        End If
        BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, 0, vbSrcCopy
    End If
    
Else
    bButtons = False
    BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, 0, vbSrcCopy
    'ReleaseCapture
    If bResize = True Then
        If bMaxed = True Then
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, 0, bHeight * 4, vbSrcCopy
        Else
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, 0, vbSrcCopy
        End If
    Else
        BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight * 3, vbSrcCopy
    End If
    BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, 0, vbSrcCopy
    bMinimize = False
    bMaximize = False
    bExit = False
End If
End Sub
Private Sub UserControl_Initialize()
            '//작업영역(작업표시줄을 제외한 영역)
            SystemParametersInfo 48, 0, wArea, 0
            bLeft = bMinWidth + bMaxWidth + bExitWidth + bMargin
End Sub



Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' 클릭 이미지 변경

iMouseDown = 1



    
    ' 폼 리사이징 부분.
    Dim SysCommWparam As Integer
    If Button = 1 And bResize = True And bMaxed = False Then
        If (X < 3&) And (Y < 3&) Then
            SysCommWparam = SC_DRAG_RESIZEUL

        ElseIf (X > ScaleWidth - 3&) And (Y > ScaleHeight - 3&) Then
            SysCommWparam = SC_DRAG_RESIZEDR

        ElseIf (X < 3&) And (Y > ScaleHeight - 3&) Then
            SysCommWparam = SC_DRAG_RESIZEDL

        ElseIf (X > ScaleWidth - 3&) And (Y < 3&) Then
            SysCommWparam = SC_DRAG_RESIZEUR

        ElseIf (X < 3&) Then
            SysCommWparam = SC_DRAG_RESIZEL

        ElseIf (X > ScaleWidth - 3&) Then
            SysCommWparam = SC_DRAG_RESIZER

        ElseIf (Y < 3&) Then
            SysCommWparam = SC_DRAG_RESIZEU

        ElseIf (Y > ScaleHeight - 3&) Then
            SysCommWparam = SC_DRAG_RESIZED

        Else
            'SysCommWparam = SC_DRAG_MOVE
        End If

        ReleaseCapture
        SendMessage Parent.hWnd, &HA1, SysCommWparam, 0&
    End If
    
    ' 폼드래그
    If Button = 1 And (Y < 30) And bButtons = False And bMaxed = False Then
        SysCommWparam = SC_DRAG_MOVE
        Call ReleaseCapture
        SendMessage Parent.hWnd, &HA1, SysCommWparam, 0&
        Exit Sub
    End If
    
    g_MouseDown = True
    g_MouseIn = True
    
        
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

If Button = 0 And iMouseDown = 1 Then iMouseDown = 0



    If bResize = True And bMaxed = False Then
        If ((X < 3&) And (Y < 3&)) Or ((X > ScaleWidth - 3&) And (Y > ScaleHeight - 3&)) Then
            MousePointer = 8& ' NWSE
    
        ElseIf ((X < 3&) And (Y > ScaleHeight - 3&)) Or ((X > ScaleWidth - 3&) And (Y < 3&)) Then
            MousePointer = 6& ' NESW
    
        ElseIf ((X < 3&) Or (X > ScaleWidth - 3&)) Then
            MousePointer = 9& ' WE
    
        ElseIf ((Y < 3&) Or (Y > ScaleHeight - 3&)) Then
            MousePointer = 7& ' NS
    
        Else
            MousePointer = 0& ' Default
        End If
    End If
    
    




If Y < 16 Then
    SetCapture hWnd
Else
    ReleaseCapture
End If




If X >= 0 And Y >= 0 And X <= Parent.Width / 15 And Y <= Parent.Height / 15 Then

        
        'If Button = 0 And Not GetCapture = hWnd Then SetCapture hWnd
          drawButtons
          
Else
    ReleaseCapture
    'If Button = 0 Then ReleaseCapture
    drawButtons
    
End If

End Sub







Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bMinimize = True Then
    Parent.WindowState = 1
    
End If

If bResize = True And bMaximize = True Then
        If bMaxed = False Then
            '//최대화 전의 설정 변수에 담아둠
            originalRect.Left = Parent.Left
            originalRect.Right = Parent.Width
            originalRect.Top = Parent.Top
            originalRect.Bottom = Parent.Height
            
            '//작업영역(작업표시줄을 제외한 영역)
            
            SystemParametersInfo 48, 0, wArea, 0
            
            '//작업영역에따라 폼의 사이즈와 폼스킨 사이즈 설정
            bMaxed = True
            Parent.Left = wArea.Left * 15
            Parent.Top = wArea.Top * 15
            Parent.Width = (wArea.Right - wArea.Left) * 15
            Parent.Height = (wArea.Bottom - wArea.Top) * 15
            Width = Parent.Width
            Height = Parent.Height
            drawForm
        Else
            '//최대화 이전의 상태로 설정
            bMaxed = False
            Parent.Left = originalRect.Left
            Parent.Width = originalRect.Right
            Parent.Top = originalRect.Top
            Parent.Height = originalRect.Bottom
            Width = Parent.Width
            Height = Parent.Height
            drawForm
        End If
End If

If bExit = True Then
    Unload Parent
End If
End Sub

Private Sub UserControl_Paint(): drawForm
End Sub

Private Sub UserControl_Resize()
drawForm
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    cName = PropBag.ReadProperty("Caption", UserControl.Name)
    bResize = PropBag.ReadProperty("Resize", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Caption", cName, Empty)
Call PropBag.WriteProperty("Resize", bResize, True)
End Sub

Public Property Get Caption() As String
Caption = cName
End Property

Public Property Let Caption(Str As String)
cName = Str
drawForm
End Property

Public Property Get Resize() As Boolean:
Resize = bResize
End Property

Public Property Let Resize(value As Boolean)
bResize = value
PropertyChanged "Resize"
drawForm

End Property



