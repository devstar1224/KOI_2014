VERSION 5.00
Begin VB.Form Join 
   BorderStyle     =   0  '없음
   Caption         =   "Join"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleMode       =   0  '사용자
   ScaleWidth      =   3797.248
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4471
      Caption         =   "Join"
      Resize          =   0   'False
      Begin VB.TextBox Text3 
         Height          =   375
         IMEMode         =   3  '사용 못함
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         IMEMode         =   3  '사용 못함
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Join"
         Height          =   615
         Left            =   3000
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   4560
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   4560
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "PW확인:"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
   End
End
Attribute VB_Name = "Join"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WinHttp As New WinHttpRequest
Private Sub Command1_Click()
If Text2.Text = Text3.Text Then
Else
MsgBox "비밀번호가 다릅니다", vbExclamation, "알림"
GoTo a
End If

    WinHttp.Open "GET", "http://dltkddlr789.dothome.co.kr/register.php?id=" & Text1.Text & "&pw=" & Text2.Text
    WinHttp.Send
    If WinHttp.ResponseText = "FAIL" Then
        MsgBox "이미 존재하는 아이디 입니다.", vbInformation, "알림"
    Else
        MsgBox "회원가입에 성공하셨습니다.", vbInformation, "알림"
        Unload Me
    End If
a:
End Sub

Private Sub Form_Unload(Cancel As Integer)
Login.Show
End Sub
