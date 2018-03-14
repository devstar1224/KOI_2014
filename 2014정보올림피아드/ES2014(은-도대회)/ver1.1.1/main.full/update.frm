VERSION 5.00
Begin VB.Form update 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   9551
      Caption         =   "Update"
      Resize          =   0   'False
      Begin VB.CommandButton Command1 
         Caption         =   "업데이트"
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "????"
         BeginProperty Font 
            Name            =   "나눔손글씨 붓"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1095
         Left            =   1800
         TabIndex        =   4
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "최신버전"
         BeginProperty Font 
            Name            =   "나눔손글씨 붓"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "1.1.1"
         BeginProperty Font 
            Name            =   "나눔손글씨 붓"
            Size            =   48
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   975
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "현재버전"
         BeginProperty Font 
            Name            =   "나눔손글씨 붓"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   1920
         TabIndex        =   1
         Top             =   840
         Width           =   1455
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
WinHttp.Open "GET", "http://dltkddlr789.dothome.co.kr/ES2014/ver.htm"
WinHttp.Send
    If MsgBox("최신 버전을 다운받을까요?" & vbCr) = vbYes Then
    Shell "explorer http://http://dltkddlr789.dothome.co.kr/ES2014.exe"
    Unload Me
    Else
    End If

End Sub

Private Sub Form_Load()
WinHttp.Open "GET", "http://dltkddlr789.dothome.co.kr/ES2014/ver.htm"
WinHttp.Send
Label4.Caption = StrConv(WinHttp.ResponseBody, vbUnicode)
If Label2.Caption = Label4.Caption Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub

