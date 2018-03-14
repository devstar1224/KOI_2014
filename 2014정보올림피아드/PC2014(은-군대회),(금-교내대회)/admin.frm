VERSION 5.00
Begin VB.Form admin 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5530
      Caption         =   "Admin - PC2014"
      Resize          =   0   'False
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         Height          =   615
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         IMEMode         =   3  '사용 못함
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "Hosting By : dothome"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   3480
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   $"admin.frx":0000
         ForeColor       =   &H8000000B&
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   3480
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "PW"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "ID"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
   End
End
Attribute VB_Name = "admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim WinHttp As New WinHttpRequest
WinHttp.Open "GET", "http://dltkddlr789.dothome.co.kr/login.php?id=" & Text1.Text & "&pw=" & Text2.Text
WinHttp.Send
If InStr(WinHttp.ResponseText, "1") Then
MsgBox "로그인 성공", vbInformation, ""
consol.Show
Unload Me
Else
MsgBox "아이디 또는 비밀번호가 맞지않습니다.", vbExclamation, ""
End If
End Sub

