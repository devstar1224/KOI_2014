VERSION 5.00
Begin VB.Form update 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8493
      Caption         =   "UpDate - PC2014"
      Resize          =   0   'False
      Begin VB.CommandButton Command1 
         Caption         =   "업데이트"
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "최신버전"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   855
         Left            =   1440
         TabIndex        =   3
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "1.0.1"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   735
         Left            =   1440
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "현재버전"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WinHttp As New WinHttpRequest
Private update_new As Byte
Private update_now As Byte

Private Sub Command1_Click()
WinHttp.Open "GET", "http://dltkddlr789.dothome.co.kr/"
WinHttp.Send
update_new = StrConv(WinHttp.ResponseBody, vbUnicode)
update_now = "1.0.0"
If update_new = update_now Then
ElseIf update_new <> update_now Then
    If MsgBox("최신 버전을 다운받을까요?" & vbCr & "현제버전: " & update_now & " 최신버전: " & update_new & "", 32 + 4, "정보") = vbYes Then
    Shell "explorer http://http://dltkddlr789.dothome.co.kr/PC2014.exe"
    Unload Me
    Else
    End If
Else
    MsgBox "잘못된 버전입니다." & vbCr & "현제버전: " & update_now & " 최신버전: " & update_new & "", 0 + 16, "정보"
End If
End Sub

Private Sub Form_Load()
WinHttp.Open "GET", "http://dltkddlr789.dothome.co.kr/"
WinHttp.Send
Label3.Caption = StrConv(WinHttp.ResponseBody, vbUnicode)
If Label2.Caption = Label3.Caption Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub


