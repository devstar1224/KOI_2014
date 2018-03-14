VERSION 5.00
Begin VB.Form elest 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   Icon            =   "elest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9551
      Caption         =   "문장-ele"
      Resize          =   0   'False
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "메인으로"
         Height          =   495
         Left            =   5760
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "※추후 매 업데이트시 문장이 추가됩니다."
         BeginProperty Font 
            Name            =   "나눔손글씨 붓"
            Size            =   18
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   600
         TabIndex        =   10
         Top             =   4800
         Width           =   4575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "These are my friend. 얘네는 내 친구야."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
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
         Top             =   4200
         Width           =   6255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "She walks to school. 그녀는 학교에 걸어서 간다"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   3720
         Width           =   6495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "What does your mother do?  너네 어머니 뭐하시니?"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   480
         TabIndex        =   6
         Top             =   3240
         Width           =   6375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "Doses your father work?  너네 아버지 일하시니?"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
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
         Top             =   2760
         Width           =   7215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "What do you want to be?  뭐가되고싶니?"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   2280
         Width           =   6375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "It takes 20 minutes. 20분걸려"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1800
         Width           =   6735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "How do I get school?  학교에 어떻게가지?"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1320
         Width           =   6615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "My apartment is on the fourth floor. 우리집은 4층이다"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   7335
      End
   End
End
Attribute VB_Name = "elest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell App.Path & "\" & "main.exe"
End
End Sub
