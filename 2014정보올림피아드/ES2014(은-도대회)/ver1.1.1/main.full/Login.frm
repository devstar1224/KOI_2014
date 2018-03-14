VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   0  '없음
   Caption         =   "Login"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleMode       =   0  '사용자
   ScaleWidth      =   3942.857
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  '사용 못함
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin Project1.VSKIN VSKIN1 
      Height          =   3975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7011
      Caption         =   "Login"
      Resize          =   0   'False
      Begin VB.CommandButton Command3 
         Caption         =   "비회원접속"
         Height          =   735
         Left            =   2040
         TabIndex        =   9
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "회원가입"
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00404040&
         Caption         =   "비밀번호 저장"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00404040&
         Caption         =   "아이디 저장"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "※회원과 비회원의 차이는 문의의 차이입니다."
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   3840
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   3840
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   3840
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   375
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SaveSetting "id", "2", "3", Text1.Text
If Check1.value = 0 Then
DeleteSetting "id", "2", "3"
Else
SaveSetting "id", "2", "3", Text1.Text
End If

SaveSetting "pw", "2", "3", Text2.Text
If Check2.value = 0 Then
DeleteSetting "pw", "2", "3"
Else
SaveSetting "pw", "2", "3", Text2.Text
End If

Dim WinHttp As New WinHttpRequest
WinHttp.Open "GET", "http://dltkddlr789.dothome.co.kr/login.php?id=" & Text1.Text & "&pw=" & Text2.Text
WinHttp.Send
If InStr(WinHttp.ResponseText, "1") Then
MsgBox "로그인 성공", vbInformation, "알림"
Login.Hide
Main.Show
Else
MsgBox "아이디 또는 비밀번호가 맞지않습니다.", vbExclamation, "오류"
End If
End Sub

Private Sub Command2_Click()
Join.Show
Unload Me
End Sub

Private Sub Command3_Click()
Text1.Text = "비회원"
support.Command1.Enabled = False
support.Label1.Caption = "비회원"
Login.Hide
Main.Show
End Sub

Private Sub Form_Load()
Text1.Text = GetSetting("id", "2", "3")
If Not GetSetting("id", "2", "3") = "" Then
Check1.value = 1
Else
Check1.value = 0
End If

Text2.Text = GetSetting("pw", "2", "3")
If Not GetSetting("pw", "2", "3") = "" Then
Check2.value = 1
Else
Check2.value = 0
End If

End Sub

