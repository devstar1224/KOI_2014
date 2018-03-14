VERSION 5.00
Begin VB.Form midte 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   Icon            =   "midte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      _ExtentX        =   20135
      _ExtentY        =   12726
      Caption         =   "문장-mid"
      Resize          =   0   'False
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "체점"
         Height          =   615
         Left            =   9960
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   4800
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "메인으로"
         Height          =   615
         Left            =   480
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Text            =   "s"
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Text            =   "j"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Text            =   "a"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '투명
         Caption         =   "a fast bird. 난 그렇게 빠른 새를 결코 본 적이 없다"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   4920
         TabIndex        =   18
         Top             =   4080
         Width           =   7815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '투명
         Caption         =   "I have never seen"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   4080
         Width           =   5175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '투명
         Caption         =   "like her mother 그녀는 어머니를 꼭 닮았다"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3240
         TabIndex        =   15
         Top             =   3480
         Width           =   8055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "She looks"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   3480
         Width           =   4455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "go to the same store 난 늘 같은 가게에 간다"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   2160
         TabIndex        =   12
         Top             =   2880
         Width           =   7335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "? 열차를 타고 갈 거니 비행기로 갈거니?"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   5520
         TabIndex        =   9
         Top             =   2280
         Width           =   5775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "Will you take the train or"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   5535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "about war 많은사람들이 전쟁에 대해 걱정한다"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   1560
         Width           =   6615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "Many people"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "s at 9:00 1교시는 9시에 시작된다."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "The first class"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
   End
End
Attribute VB_Name = "midte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell App.Path & "\" & "main.exe"
End
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
If Text1.Text = "begin" Then
Text1.Text = Text1.Text & "정답"
Text1.Enabled = False
Else
Text1.Text = "오답"
End If

If Text2.Text = "worry" Then
Text2.Text = Text2.Text & "정답"
Text2.Enabled = False
Else
Text2.Text = "오답"
End If

If Text3.Text = "fly" Then
Text3.Text = Text3.Text & "정답"
Text3.Enabled = False
Else
Text3.Text = "오답"
End If

If Text4.Text = "always" Then
Text4.Text = Text4.Text & "정답"
Text4.Enabled = False
Else
Text4.Text = "오답"
End If

If Text5.Text = "just" Then
Text5.Text = Text5.Text & "정답"
Text5.Enabled = False
Else
Text5.Text = "오답"
End If

If Text6.Text = "such" Then
Text6.Text = Text6.Text & "정답"
Text6.Enabled = False
Else
Text6.Text = "오답"
End If

End Sub
