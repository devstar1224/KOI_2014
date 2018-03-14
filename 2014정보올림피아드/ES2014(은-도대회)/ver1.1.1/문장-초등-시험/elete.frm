VERSION 5.00
Begin VB.Form elete 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8310
   Icon            =   "elete.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9763
      Caption         =   "문장-test-ele"
      Resize          =   0   'False
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "메인으로"
         Height          =   615
         Left            =   360
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "체점"
         Height          =   615
         Left            =   5880
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   4680
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Text            =   "????? ??"
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Text            =   "W??? ???? ???? ?????? ??"
         Top             =   3360
         Width           =   3855
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Text            =   "??? ????"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Text            =   "???? ?? ??"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Text            =   "g?? ??????"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Text            =   "o? ??? ??????"
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '투명
         Caption         =   "school. 그녀는 학교에 걸어서 간다."
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
         Left            =   3000
         TabIndex        =   18
         Top             =   3960
         Width           =   5175
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '투명
         Caption         =   "She"
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
         TabIndex        =   16
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '투명
         Caption         =   "너네 어머니 뭐하시니?"
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
         Left            =   4320
         TabIndex        =   14
         Top             =   3360
         Width           =   4215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "father work? 너네 아버지 일하시니?"
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
         Left            =   2400
         TabIndex        =   13
         Top             =   2760
         Width           =   5655
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "D"
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
         Left            =   480
         TabIndex        =   11
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "?  뭐가되고 싶니?"
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
         Left            =   4560
         TabIndex        =   10
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "What do you"
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
         Left            =   480
         TabIndex        =   8
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "?  학교에 어떻게가지?"
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
         Left            =   3000
         TabIndex        =   7
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "How do I"
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
         Left            =   480
         TabIndex        =   5
         Top             =   1560
         Width           =   4815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "우리집은 4층이다"
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
         Left            =   5400
         TabIndex        =   4
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "floor."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   4560
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "My apartment "
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
         Left            =   480
         TabIndex        =   1
         Top             =   960
         Width           =   4455
      End
   End
End
Attribute VB_Name = "elete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
If Text1.Text = "on the fourth" Then
Text1.Text = Text1.Text & "정답"
Text1.Enabled = False
Else
Text1.Text = "오답"
End If

If Text2.Text = "get school" Then
Text2.Text = Text2.Text & "정답"
Text2.Enabled = False
Else
Text2.Text = "오답"
End If

If Text3.Text = "want to be" Then
Text3.Text = Text3.Text & "정답"
Text3.Enabled = False
Else
Text3.Text = "오답"
End If

If Text4.Text = "oes your" Then
Text4.Text = Text4.Text & "정답"
Text4.Enabled = False
Else
Text4.Text = "오답"
End If

If Text5.Text = "What does your mother do?" Then
Text5.Text = Text5.Text & "정답"
Text5.Enabled = False
Else
Text5.Text = "오답"
End If


If Text6.Text = "walks to" Then
Text6.Text = Text6.Text & "정답"
Text6.Enabled = False
Else
Text6.Text = "오답"
End If



End Sub

Private Sub Command2_Click()
Shell App.Path & "\" & "main.exe"
End

End Sub
