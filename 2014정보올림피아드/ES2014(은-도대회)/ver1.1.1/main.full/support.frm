VERSION 5.00
Begin VB.Form support 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   8493
      Caption         =   "Support"
      Resize          =   0   'False
      Begin VB.CommandButton Command1 
         Caption         =   "전송"
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   2895
         Left            =   960
         TabIndex        =   4
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "내용:"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "Label2"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "아이디:"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "support"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WinHttp As New WinHttpRequest
Private Sub Command1_Click()

   WinHttp.Open "GET", "http://dltkddlr789.dothome.co.kr/support/support.php?id=" & Label2.Caption & "&pw=" & Text1.Text
    WinHttp.Send
    If WinHttp.ResponseText = "FAIL" Then
   
    Else
        MsgBox "전송성공.", vbInformation, "알림"
    End If
End Sub

Private Sub Form_Load()
Label2.Caption = Login.Text1.Text
If Label2.Caption = "비회원" Then
MsgBox "비회원은 문의는 불가합니다. 문의를 하실경우 회원가입후 이용해주세요", vbInformation, "알림"
End If
End Sub

